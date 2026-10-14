<#
Checks the Docker packaging of the sourcing tool. For Windows PowerShell 5.1 with Docker Desktop.

    $env:OBS_ENABLED = "0"
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -WriteConstraints   # first build only
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Only V04,V09       # some checks; what they need is still prepared

  -WriteConstraints  writes docker\constraints.txt from the first image's installed packages, then builds again with it
  -Only              ids of the checks to run (default: all)
  -Root              verify the Docker files in another folder (the mutation run uses this); default: this project
  -Keep              leave this run's containers, images, network and temp folder in place
  -NoLangfuseGet     skip the one read-only GET of the Langfuse health address

What it uses: fake keys only; a throwaway data folder and a staged build context under %TEMP% (the real .env is never
inside a build context); names that start with lcom-verify; the label com.zync.lcom=1.
What it never does: touch a container, image, network or volume that is not its own; run a prune; push an image; call a
model, Nimble or a mail server; send a trace; print the value of a setting.
Needs: python with the project's requirements on this computer (docker\verify_helper.py drives the UI over HTTP).
#>
param(
    [string]$Root = "",
    [string[]]$Only = @(),
    [switch]$WriteConstraints,
    [switch]$Keep,
    [switch]$NoLangfuseGet
)
Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
$env:OBS_ENABLED = "0"
$Only = @($Only | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })      # -File passes "V04,V09" as one word

$Label = "com.zync.lcom=1"
$Real = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path                  # the real project: only ever hashed and listed
if ($Root) { $Root = (Resolve-Path $Root).Path } else { $Root = $Real }     # where the Docker files under test are
$HelperPy = Join-Path $Real "docker\verify_helper.py"
$Id = Get-Date -Format "yyyyMMddHHmmss"
$Work = Join-Path $env:TEMP "lcom-verify-$Id"
$Ctx = "$Work\ctx"; $Proj = "$Work\compose"; $Data = "$Work\data"
$Name = "lcom-verify-ui"
$Image = "lcom-ui:verify-$Id"; $TestImage = "lcom-ui:verify-test-$Id"; $ProbeImage = "lcom-ui:verify-ctx-$Id"
$Needle = "sk-ant-FAKE-verify-PLANTED-$Id"                                  # planted in the build context; must never be in an image
$StateNeedle = "planted-state-file-$Id"
$Zone = "Asia/Kolkata"; $ZoneOffset = "+0530"; $ZoneLabel = "UTC+05:30"
$Suites = @("test_pricing.py", "test_email_agent.py", "test_decision_agent.py", "test_gates.py", "test_order_sheet.py", "test_content_agent.py",
            "test_observability.py", "test_ui.py", "tests/test_ui_pipeline.py")
# Parts of protected suites that only run when real output files exist. An image holds none, so they do not run there.
$DataDependent = @("test_email_agent.py:309 (drafts on the newest real results file)", "test_email_agent.py:365 (real near_misses.csv with the newest real results file)",
                   "test_email_agent.py:518 (leak scan of the newest real results file)", "test_content_agent.py:639 (the real approved_orders.csv, if there is one)")
$Protected = @("sourcing_agent.py", "email_agent.py", "decision_agent.py", "order_sheet.py", "content_agent.py", "observability.py", "statefile.py",
               "requirements.txt", ".env", ".gitignore", "test_pricing.py", "test_email_agent.py", "test_decision_agent.py", "test_gates.py",
               "test_order_sheet.py", "test_content_agent.py", "test_observability.py",
               "email_state.json", "decision_state.json", "approvals_log.jsonl", "approved_orders.csv", "approved_orders.json", "product_facts.csv")
$RealFolders = @("Excel Output Sheets", "Reports", "Order Sheets", "Walmart Listings", "Uploads", "ui_data", "data")
$Results = New-Object System.Collections.ArrayList
$State = @{ Staged = $false; Built = $false; Up = $false; Port = 0; Stem = ""; Before = $null; Others = $null }
$Utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------- small helpers ----------

function Say($text) { Write-Host $text }

function Add-Result($id, $check, $ok, $detail) {
    $r = if ($null -eq $ok) { "INFO" } elseif ($ok) { "PASS" } else { "FAIL" }
    [void]$Results.Add([pscustomobject]@{ Id = $id; Check = $check; Result = $r; Detail = "$detail" })
    Say ("{0}  {1}  {2,-24} {3}" -f $r, $id, $check, $detail)
}

function Pass($detail) { $script:Ok = $true; $script:Detail = $detail }
function Fail($detail) { $script:Ok = $false; $script:Detail = $detail }
function Info($detail) { $script:Ok = $null; $script:Detail = $detail }
function Judge($problems, $good) { if (@($problems).Count -gt 0) { Fail (@($problems) -join "; ") } else { Pass $good } }

function Check($id, $check, [scriptblock]$body) {
    if ($Only.Count -gt 0 -and ($Only -notcontains $id)) { return }
    $script:Ok = $false; $script:Detail = "the check set no result"
    try { & $body | Out-Null } catch { Fail ("could not run: " + $_.Exception.Message) }
    Add-Result $id $check $script:Ok $script:Detail
}

function Same-List($a, $b) { return ((@($a) -join "`n") -ceq (@($b) -join "`n")) }      # also true for two empty lists

function Prop($object, $name) {
    if ($null -ne $object -and $object.PSObject.Properties[$name]) { return $object.$name }
    return $null
}

function Write-Text($path, $text) {
    New-Item -ItemType Directory -Force (Split-Path $path) | Out-Null
    [System.IO.File]::WriteAllText($path, $text.Replace("`r`n", "`n"), $Utf8)     # no BOM, LF: as Docker and Python expect
}

# docker.exe is called by its file name so these wrappers never call themselves.
function Dk {
    $out = & docker.exe @args
    if ($LASTEXITCODE -ne 0) { throw ("docker " + ($args -join " ") + " ended with code $LASTEXITCODE") }
    return $out
}

function Compose {
    $out = & docker.exe compose -p lcom-verify --project-directory $Proj -f "$Proj\docker-compose.yml" --env-file "$Proj\.env" @args
    if ($LASTEXITCODE -ne 0) { throw ("docker compose " + ($args -join " ") + " ended with code $LASTEXITCODE") }
    return $out
}

function Test-Ours($container) {
    $names = @(& docker.exe ps -a --filter "name=$container" --filter "label=$Label" --format "{{.Names}}")
    return ($names -contains $container) -and ($container -like "lcom-verify*")
}

function Assert-Ours($container) {
    if (-not (Test-Ours $container)) { throw "'$container' is not one of this run's containers (label $Label, name lcom-verify*): not touched" }
}

function From-Json($lines) { return ((@($lines) | Select-Object -Last 1) | ConvertFrom-Json) }

function Helper { return (From-Json (& python -B $HelperPy @args)) }

function In-Container($code) {
    Assert-Ours $Name
    $out = $code | & docker.exe exec -i $Name python -
    if ($LASTEXITCODE -ne 0) { throw "python inside the container ended with code $LASTEXITCODE" }
    return (From-Json $out)
}

function In-Image($image, $code, [string[]]$dockerArgs = @(), [string[]]$scriptArgs = @()) {
    $out = $code | & docker.exe run --rm -i --network none --label $Label --entrypoint python @dockerArgs $image - @scriptArgs
    if ($LASTEXITCODE -ne 0) { throw "python inside $image ended with code $LASTEXITCODE" }
    return (From-Json $out)
}

function Logs($container) { Assert-Ours $container; return (@(cmd /c "docker logs --tail 40 $container 2>&1") -join "`n") }

function Get-FreePort {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $l.Start(); $p = $l.LocalEndpoint.Port; $l.Stop()
    return $p
}

function Get-Tree($folder) {
    if (-not (Test-Path -LiteralPath $folder)) { return @() }
    return @(Get-ChildItem -LiteralPath $folder -Recurse -File | ForEach-Object { "{0}|{1}" -f $_.FullName.Substring($folder.Length + 1), (Get-FileHash -LiteralPath $_.FullName).Hash } | Sort-Object)
}

function Get-RealState {
    # Hashes of the files that must not change, and a listing of the real output folders. Read-only. Values are never printed.
    $out = @{}
    foreach ($f in $Protected) {
        $p = Join-Path $Real $f
        $out[$f] = if (Test-Path -LiteralPath $p) { (Get-FileHash -LiteralPath $p).Hash.Substring(0, 16).ToLower() } else { "absent" }
    }
    foreach ($d in $RealFolders) {
        $p = Join-Path $Real $d
        $out["$d\"] = if (Test-Path -LiteralPath $p) { "{0} files, listing {1}" -f @(Get-ChildItem -LiteralPath $p -Recurse -File).Count, (Get-ListingHash $p) } else { "absent" }
    }
    $out["git status --short"] = Get-TextHash ((@(& git -C $Real status --short) -join "`n"))
    return $out
}

function Get-TextHash($text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return ([System.BitConverter]::ToString($sha.ComputeHash($Utf8.GetBytes("$text"))) -replace "-", "").Substring(0, 16).ToLower()
}

function Get-ListingHash($folder) {
    return (Get-TextHash ((@(Get-ChildItem -LiteralPath $folder -Recurse -File | ForEach-Object { "{0}|{1}|{2}" -f $_.FullName, $_.Length, $_.LastWriteTimeUtc.Ticks } | Sort-Object) -join "`n")))
}

function Get-Others {
    # Every container (name and state) and volume that does NOT carry this tool's label: the proof nothing else was touched.
    $mine = @(& docker.exe ps -a --filter "label=$Label" --format "{{.ID}}")
    $containers = @(& docker.exe ps -a --format "{{.ID}}|{{.Names}} ({{.State}})" | Where-Object { $mine -notcontains ($_ -split "\|")[0] } | ForEach-Object { ($_ -split "\|", 2)[1] } | Sort-Object)
    $myVolumes = @(& docker.exe volume ls --filter "label=$Label" --format "{{.Name}}")
    $volumes = @(& docker.exe volume ls --format "{{.Name}}" | Where-Object { $myVolumes -notcontains $_ } | Sort-Object)
    return @{ Containers = $containers; Volumes = $volumes }
}

# ---------- preparations (each runs once, when a check first needs it) ----------

function Stage {
    if ($State.Staged) { return }
    foreach ($need in @("scratch\pre_gates\decision_agent.py", "scratch\pre_gates\email_agent.py")) {
        if (-not (Test-Path -LiteralPath (Join-Path $Root $need))) {
            throw "$need is missing. The test image needs it: test_gates.py compares today's agents with these two earlier copies. Nothing was built."
        }
    }
    New-Item -ItemType Directory -Force $Ctx, $Proj, $Data | Out-Null
    $allowed = @(Get-Content (Join-Path $Root "docker\image_files.txt")) + @(Get-Content (Join-Path $Root "docker\image_files_test.txt")) | Where-Object { $_ }
    # the allowed files, the two Docker files, and the project's other Docker-side files (which must NOT get into an image)
    $copy = $allowed + @("Dockerfile", ".dockerignore", "docker-compose.yml", ".env.docker.example", "docker/verify.ps1", "docker/verify_helper.py",
                         "docker/image_files.txt", "docker/image_files_test.txt")
    foreach ($f in $copy) {
        $src = Join-Path $Root $f
        if (-not (Test-Path -LiteralPath $src)) { throw "$f is listed but missing in $Root" }
        $dst = Join-Path $Ctx $f
        New-Item -ItemType Directory -Force (Split-Path $dst) | Out-Null
        Copy-Item -LiteralPath $src -Destination $dst
    }
    # planted: everything a careless build could pick up. None of it may reach Docker.
    Write-Text "$Ctx\.env" "ANTHROPIC_API_KEY=$Needle`nSMTP_PASSWORD=$Needle`n"
    Write-Text "$Ctx\secret_key.txt" $Needle
    foreach ($f in @("email_state.json", "decision_state.json", "approvals_log.jsonl", "approved_orders.csv", "approved_orders.json", "product_facts.csv",
                     "Excel Output Sheets\planted.xlsx", "Reports\planted.md", "Order Sheets\planted.xlsx", "Walmart Listings\planted.md", "Uploads\planted.xlsx",
                     "ui_data\jobs\planted.log", "data\planted.txt", ".git\config", "README.md", "notes.xlsx", "scratch\other.py", "scratch\pre_gates\README.md",
                     "ui\__pycache__\planted.pyc", "ui\static\planted.txt", "ui\templates\planted.txt", "tests\__pycache__\planted.pyc", "tests\shim\planted.txt",
                     "docker\planted.txt")) {
        Write-Text "$Ctx\$f" $StateNeedle
    }
    Copy-Item -LiteralPath (Join-Path $Root "docker-compose.yml") -Destination "$Proj\docker-compose.yml"
    $State.Port = Get-FreePort
    $envText = @"
TZ=$Zone
LCOM_PORT=$($State.Port)
LCOM_DATA=$($Data -replace '\\', '/')
LCOM_NAME=$Name
LCOM_IMAGE=$Image
ANTHROPIC_API_KEY="sk-ant-FAKE-verify-000111222333"
NIMBLE_API_KEY=nimble-FAKE-verify-444555
SMTP_HOST=smtp.example.invalid
SMTP_USER=sender@example.invalid
SMTP_PASSWORD='FAKE-pa`$`$word-with#hash and space'
SHIPPING_ADDRESS="Test Receiving\n1 Test Street\nTestville, TX 00000"
BRAND_NAME=TestBrand
SENDER_NAME="Test Sender"
OBS_ENABLED=1
LANGFUSE_BASE_URL=http://localhost:3000
LANGFUSE_HOST=http://localhost:3000
LANGFUSE_PUBLIC_KEY=pk-lf-FAKE-verify
LANGFUSE_SECRET_KEY=sk-lf-FAKE-verify
UI_HOST=10.9.8.7
UI_ENABLE_STUB=1
PROJECT_DIR=/tmp/elsewhere
HTTPS_PROXY=http://127.0.0.1:9
HTTP_PROXY=http://127.0.0.1:9
"@
    Write-Text "$Proj\.env" $envText          # fake values; the last two lines make any accidental outbound call fail at once
    $State.Staged = $true
}

function Build {
    Stage
    if ($State.Built) { return }
    Say "building the runtime and test images (the first build downloads the base image and the packages; several minutes) ..."
    Dk build -q --target runtime -t $Image --label $Label $Ctx | Out-Null
    Dk build -q --target test -t $TestImage --label $Label $Ctx | Out-Null
    $State.Built = $true
}

function Wait-Healthy($seconds = 150) {
    $end = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $end) {
        $s = @(& docker.exe inspect --format "{{.State.Status}}|{{if .State.Health}}{{.State.Health.Status}}{{end}}" $Name)[0]
        if ($s -eq "running|healthy") { return }
        if ($s -like "exited*") { throw ("the container stopped. Its last lines:`n" + (Logs $Name)) }
        Start-Sleep -Seconds 2
    }
    throw ("the container did not become healthy in $seconds seconds. Its last lines:`n" + (Logs $Name))
}

function Up {
    Build
    if ($State.Up) { return }
    $State.Stem = (Helper fixture $Data).stem
    Compose up -d --no-build | Out-Null
    Wait-Healthy
    $State.Up = $true
}

# ---------- V01: the engine ----------

$server = @(& docker.exe info --format "{{.ServerVersion}}")
if (-not ($server -and "$($server[0])".Trim())) {
    Say "STOPPED: the Docker engine does not answer (docker info gave no server version). Start Docker Desktop, wait until it is running, and run this again. Nothing was built or changed."
    exit 2
}
Add-Result "V01" "engine" $true ("Docker engine {0}, client {1}, {2}" -f $server[0], (& docker.exe version --format "{{.Client.Version}}"), (@(& docker.exe compose version)[0]))

# ---------- V02: before ----------

$State.Before = Get-RealState
$State.Others = Get-Others
Add-Result "V02" "before" $null ("recorded {0} items of the real project; containers not from this tool: {1}; volumes not from this tool: {2}" -f $State.Before.Count, (@($State.Others.Containers) -join ", "), (@($State.Others.Volumes) -join ", "))

try {
    # ---------- V03: the staged build context ----------
    Check "V03" "staged-context" {
        Stage
        $p = @()
        foreach ($f in @("Dockerfile", ".dockerignore", "docker-compose.yml")) {
            if ((Get-FileHash -LiteralPath (Join-Path $Root $f)).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Ctx $f)).Hash) { $p += "$f in the staged copy differs from the one under test" }
        }
        if (Test-Path -LiteralPath (Join-Path $Real ".env")) {
            if ((Get-FileHash -LiteralPath (Join-Path $Real ".env")).Hash -eq (Get-FileHash -LiteralPath "$Ctx\.env").Hash) { $p += "the staged .env is the real one" }
        }
        Judge $p ("{0} files staged under {1}; a fake .env, a fake key file and 24 decoy files planted; the real .env is not in it" -f @(Get-ChildItem $Ctx -Recurse -File).Count, $Ctx)
    }

    # ---------- V04: what Docker receives ----------
    $ListCode = @'
import json, os, sys
root = sys.argv[1] if len(sys.argv) > 1 else '/app'
out = []
for d, _, files in os.walk(root):
    out += [os.path.relpath(os.path.join(d, f), root).replace(os.sep, '/') for f in files]
print(json.dumps(sorted(out)))
'@
    Check "V04" "context-exact" {
        Stage
        $from = (@(Get-Content "$Ctx\Dockerfile" | Where-Object { $_ -match "^FROM\s+\S+" })[0] -replace "^FROM\s+(\S+).*$", '$1')
        Write-Text "$Work\probe.Dockerfile" "FROM $from`nCOPY . /ctx`n"
        Dk build -q -f "$Work\probe.Dockerfile" -t $ProbeImage --label $Label $Ctx | Out-Null
        $got = @(In-Image $ProbeImage $ListCode @() @("/ctx"))
        $want = @(@(Get-Content "$Ctx\docker\image_files.txt") + @(Get-Content "$Ctx\docker\image_files_test.txt") | Where-Object { $_ } | Sort-Object)
        $extra = @($got | Where-Object { $want -notcontains $_ }); $missing = @($want | Where-Object { $got -notcontains $_ })
        $p = @()
        if ($extra.Count) { $p += "Docker received files that are not on the allowlist: " + ($extra -join ", ") }
        if ($missing.Count) { $p += "allowed files did not reach Docker: " + ($missing -join ", ") }
        Judge $p ("Docker received exactly the {0} allowed files; none of the planted ones" -f $want.Count)
    }

    # ---------- V05: build, and the installed packages ----------
    Check "V05" "build" {
        Build
        $freeze = @(Dk run --rm --network none --label $Label --entrypoint python $Image -m pip freeze | Where-Object { $_ } | Sort-Object)
        $pinsFile = Join-Path $Root "docker\constraints.txt"
        if ($WriteConstraints) {
            $head = @(Get-Content $pinsFile | Where-Object { $_ -like "#*" -and $_ -notlike "# NOT FILLED IN*" -and $_ -notlike "# and docker*" })
            Write-Text $pinsFile ((@($head) + $freeze) -join "`n")
            Copy-Item -LiteralPath $pinsFile -Destination "$Ctx\docker\constraints.txt" -Force
            $State.Built = $false
            Build
            $freeze = @(Dk run --rm --network none --label $Label --entrypoint python $Image -m pip freeze | Where-Object { $_ } | Sort-Object)
        }
        $pins = @(Get-Content $pinsFile | Where-Object { $_ -and $_ -notlike "#*" } | Sort-Object)
        $from = (@(Get-Content "$Ctx\Dockerfile" | Where-Object { $_ -match "^FROM\s+\S+" })[0] -replace "^FROM\s+(\S+).*$", '$1')
        $size = Dk image inspect --format "{{.Size}}" $Image
        $p = @()
        if ($from -notmatch "^python:3\.13\.5-slim-bookworm@sha256:[0-9a-f]{64}$") { $p += "the base image is not pinned by tag and digest: $from" }
        if (-not $pins.Count) { $p += "docker\constraints.txt has no pins yet: run once with -WriteConstraints" }
        elseif (Compare-Object $pins $freeze) { $p += "the installed packages differ from docker\constraints.txt: " + ((Compare-Object $pins $freeze | ForEach-Object { $_.SideIndicator + $_.InputObject }) -join " ") }
        Judge $p ("runtime and test images built from $from; {0} packages installed, all pinned; runtime image {1:N0} MB" -f $freeze.Count, ([double]$size / 1MB))
    }

    # ---------- V06: the files in the images ----------
    Check "V06" "image-files" {
        Build
        $runtime = @(Get-Content "$Ctx\docker\image_files.txt" | Where-Object { $_ } | Sort-Object)
        $test = @($runtime + @(Get-Content "$Ctx\docker\image_files_test.txt" | Where-Object { $_ }) | Sort-Object)
        $p = @()
        foreach ($pair in @(@($Image, $runtime, "runtime"), @($TestImage, $test, "test"))) {
            $listed = @($pair[1])
            $got = @(In-Image $pair[0] $ListCode)
            $absent = @($listed | Where-Object { $got -notcontains $_ }); $unlisted = @($got | Where-Object { $listed -notcontains $_ })
            if ($absent.Count -or $unlisted.Count) { $p += "$($pair[2]) image: listed but absent: [" + ($absent -join ", ") + "]; present but not listed: [" + ($unlisted -join ", ") + "]" }
        }
        Judge $p ("/app holds exactly the {0} listed files in the runtime image and {1} in the test image; no tests and no stub in runtime" -f $runtime.Count, $test.Count)
    }

    # ---------- V07: nothing secret in any layer ----------
    Check "V07" "no-secrets-in-layers" {
        Build
        $p = @(); $told = @()
        foreach ($pair in @(@($Image, "runtime"), @($TestImage, "test"))) {
            $tar = "$Work\$($pair[1]).tar"
            Dk save $pair[0] -o $tar | Out-Null
            $r = Helper scan $tar $pair[1] $Needle $StateNeedle
            Remove-Item -LiteralPath $tar -Force
            if (Prop $r "error") { $p += "$($pair[1]): the scan failed: $($r.error)" }
            elseif (-not $r.ok) { $p += "$($pair[1]): forbidden names: " + (@($r.bad_names) -join ", ") + "; planted text found in: " + (@($r.needle_hits) -join ", ") }
            else { $told += "{0}: {1} layers, {2} files" -f $pair[1], $r.layers, $r.files }
        }
        Judge $p ("no .env, state file, output folder or planted text in any layer or in the image settings (" + ($told -join "; ") + ")")
    }

    # ---------- V08: the image's own settings ----------
    Check "V08" "image-config" {
        Build
        $c = ((Dk image inspect $Image) -join "`n" | ConvertFrom-Json)[0].Config
        $p = @()
        if ((Prop $c "User") -ne "10001:10001") { $p += "the image's user is '$(Prop $c 'User')', not 10001:10001" }
        $envNames = @(@(Prop $c "Env") | ForEach-Object { ("$_" -split "=", 2)[0] })
        $keyLike = @(@(Prop $c "Env") | Where-Object { $_ -match "(?i)(api_?key|secret|password|token)[^=]*=.+" -or $_ -match "sk-ant-|sk-lf-|pk-lf-" } | ForEach-Object { ("$_" -split "=", 2)[0] })
        if ($keyLike.Count) { $p += "key-like settings are baked into the image: " + ($keyLike -join ", ") }
        if ((@(Prop $c "Entrypoint") -join " ") -ne "python /app/docker/entrypoint.py") { $p += "the entrypoint is '" + (@(Prop $c "Entrypoint") -join " ") + "'" }
        if (-not (Prop (Prop $c "Healthcheck") "Test")) { $p += "no health check" }
        if ((Prop (Prop $c "Labels") "com.zync.lcom") -ne "1") { $p += "the label is missing" }
        if (Prop $c "Volumes") { $p += "the image declares a volume, so a start without a data folder would hide the data" }
        Judge $p ("user 10001:10001, entrypoint and health check as designed, label present, no declared volume; settings in the image by name: " + ($envNames -join ", "))
    }

    # ---------- V09: the compose file, resolved ----------
    Check "V09" "compose-config" {
        Stage
        $svc = ((Compose config --format json) -join "`n" | ConvertFrom-Json).services.ui      # never printed whole: it holds every setting
        $envs = Prop $svc "environment"
        $p = @()
        $ports = @(Prop $svc "ports")
        if ($ports.Count -ne 1) { $p += "expected one published port, found $($ports.Count)" }
        else {
            if ((Prop $ports[0] "host_ip") -ne "127.0.0.1") { $p += "the port is published on '$(Prop $ports[0] 'host_ip')', not on 127.0.0.1 only" }
            if ("$(Prop $ports[0] 'published')" -ne "$($State.Port)" -or "$(Prop $ports[0] 'target')" -ne "$($State.Port)") { $p += "the port is not the same number inside and outside" }
        }
        if ((Prop $svc "read_only") -ne $true) { $p += "the root filesystem is not read-only" }
        if (@(Prop $svc "cap_drop") -notcontains "ALL") { $p += "capabilities are not all dropped" }
        if (@(Prop $svc "security_opt") -notcontains "no-new-privileges:true") { $p += "no-new-privileges is not set" }
        if ((Prop $svc "user") -ne "10001:10001") { $p += "the user is '$(Prop $svc 'user')'" }
        if ((Prop $svc "init") -ne $true) { $p += "init is off" }
        if ((Prop $svc "restart") -ne "no") { $p += "the restart policy is '$(Prop $svc 'restart')', not 'no'" }
        if (((Prop $svc "extra_hosts") | ConvertTo-Json -Compress) -notmatch "host\.docker\.internal") { $p += "host.docker.internal is not mapped" }
        if (-not (@(Prop $svc "tmpfs") -join " ").StartsWith("/tmp")) { $p += "/tmp is not a memory folder" }
        $mounts = @(Prop $svc "volumes")
        if ($mounts.Count -ne 1) { $p += "expected exactly one mount (the data folder), found $($mounts.Count)" }
        elseif ((Prop $mounts[0] "type") -ne "bind" -or (Prop $mounts[0] "target") -ne "/data") { $p += "the one mount is not a folder bound at /data" }
        elseif (-not (Test-Path -LiteralPath (Prop $mounts[0] "source") -PathType Container)) { $p += "the mounted path is not a folder (a single file cannot be replaced by a rename)" }
        $want = @{ UI_HOST = "0.0.0.0"; UI_ALLOW_NON_LOOPBACK = "1"; UI_ENABLE_STUB = "0"; PROJECT_DIR = "/data"; UI_DATA_DIR = "/data/ui_data"; UI_PORT = "$($State.Port)";
                   OBS_ENABLED = "0"; LANGFUSE_BASE_URL = "http://host.docker.internal:3000"; LANGFUSE_HOST = "http://host.docker.internal:3000"; TZ = $Zone }
        foreach ($k in $want.Keys) { if ("$(Prop $envs $k)" -ne $want[$k]) { $p += "$k is not what the compose file must set (the settings file must not be able to change it)" } }
        $log = Prop (Prop $svc "logging") "options"
        if ((Prop $log "max-size") -ne "10m" -or "$(Prop $log 'max-file')" -ne "3") { $p += "the log limits are not 10m and 3 files" }
        if ((Prop (Prop $svc "build") "target") -ne "runtime") { $p += "compose would build '$(Prop (Prop $svc 'build') 'target')', not the runtime target" }
        if ((Prop (Prop $svc "labels") "com.zync.lcom") -ne "1") { $p += "the container label is missing" }
        Judge $p "published on 127.0.0.1 only, same port inside and outside; read-only; no capabilities; no-new-privileges; user 10001; init; restart no; one folder mount; log limits; the 10 pinned settings win over the settings file; builds runtime only"
    }

    # ---------- V10: it runs, as designed ----------
    $WhoCode = @'
import json, os
status = dict(line.split(':', 1) for line in open('/proc/self/status') if ':' in line)
def writable(path):
    try:
        open(path, 'w').close()
        os.remove(path)
        return True
    except OSError:
        return False
print(json.dumps({'uid': os.getuid(), 'gid': os.getgid(), 'app': writable('/app/_verify'), 'root': writable('/_verify'), 'tmp': writable('/tmp/_verify'),
                  'data': writable('/data/_verify_w'), 'nonewprivs': status.get('NoNewPrivs', '').strip(), 'capbnd': status.get('CapBnd', '').strip(),
                  'pid1': open('/proc/1/comm').read().strip(), 'ui_host': os.environ.get('UI_HOST'), 'project': os.environ.get('PROJECT_DIR'),
                  'stub': os.environ.get('UI_ENABLE_STUB')}))
'@
    Check "V10" "up-healthy" {
        Up
        $w = In-Container $WhoCode
        $p = @()
        if ($w.uid -ne 10001 -or $w.gid -ne 10001) { $p += "runs as uid $($w.uid) gid $($w.gid), not 10001" }
        if ($w.app -or $w.root) { $p += "the program folder or the root filesystem is writable" }
        if (-not $w.tmp -or -not $w.data) { $p += "/tmp or the data folder is not writable" }
        if ($w.nonewprivs -ne "1") { $p += "NoNewPrivs is '$($w.nonewprivs)'" }
        if ($w.capbnd -notmatch "^0+$") { $p += "capabilities are available: $($w.capbnd)" }
        if ($w.pid1 -notmatch "init|tini") { $p += "process 1 is '$($w.pid1)', not an init" }
        if ($w.ui_host -ne "0.0.0.0" -or $w.project -ne "/data" -or $w.stub -ne "0") { $p += "the settings file changed a pinned setting" }
        Judge $p "healthy; uid and gid 10001; /app and / read-only; /tmp and /data writable; NoNewPrivs 1; no capabilities; an init is process 1"
    }

    # ---------- V11: where the port listens on this computer ----------
    Check "V11" "host-port" {
        Up
        $addresses = @(Get-NetTCPConnection -State Listen -LocalPort $State.Port -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LocalAddress -Unique)
        if ($addresses.Count -eq 1 -and $addresses[0] -eq "127.0.0.1") { Pass "port $($State.Port) listens on 127.0.0.1 only" }
        else { Fail ("port $($State.Port) listens on: " + ($addresses -join ", ")) }
    }

    # ---------- V12: Host, Origin and CSRF through the published port ----------
    Check "V12" "host-check" {
        Up
        $r = Helper hosts $State.Port $Name
        if (Prop $r "error") { Fail $r.error } elseif ($r.ok) { Pass "$($r.checked) requests: the two right hosts answered; every wrong Host got 400 (0.0.0.0, host.docker.internal, wrong port, no port, upper case, trailing dot, the container's name); every wrong Origin and a missing or wrong token got 403" }
        else { Fail (@($r.failures) -join "; ") }
    }

    # ---------- V13: the settings reach the container as python-dotenv would read them ----------
    $EnvNames = @("ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SHIPPING_ADDRESS", "BRAND_NAME", "SENDER_NAME",
                  "LANGFUSE_PUBLIC_KEY", "LANGFUSE_SECRET_KEY")
    $HashCode = @'
import hashlib, json, os
names = [NAMES]
print(json.dumps({n: (hashlib.sha256(os.environ[n].encode('utf-8')).hexdigest() if n in os.environ else None) for n in names}))
'@
    Check "V13" "env-parity" {
        Up
        $want = Helper envhash "$Proj\.env" @EnvNames
        $got = In-Container ($HashCode.Replace("NAMES", (($EnvNames | ForEach-Object { "'$_'" }) -join ", ")))
        $differ = @($EnvNames | Where-Object { (Prop $want $_) -ne (Prop $got $_) -or -not (Prop $got $_) })
        $page = Helper pages $State.Port $State.Stem
        $notSet = @(@("ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SHIPPING_ADDRESS", "BRAND_NAME") | Where-Object { @(Prop $page "set") -notcontains $_ })
        $p = @()
        if ($differ.Count) { $p += "these settings differ between the container and python-dotenv's reading of the same file (names only): " + ($differ -join ", ") }
        if ($notSet.Count) { $p += "the Health page does not show these as Set: " + ($notSet -join ", ") }
        Judge $p ("{0} settings (double quotes, single quotes with `$ and #, \n, spaces) are identical by SHA-256 to python-dotenv's reading; Health shows them as Set" -f $EnvNames.Count)
    }

    # ---------- V14: work done through the UI lands in the host folder ----------
    Check "V14" "data-on-host" {
        Up
        $r = Helper flow $State.Port $Data $State.Stem
        if (Prop $r "error") { Fail $r.error }
        elseif ($r.ok) { Pass ("{0} steps through the UI, each file found in {1}: upload, approval request, a refused and an accepted record, export, download, order sheet" -f @($r.steps).Count, $Data) }
        else { Fail ((@($r.steps) | Where-Object { -not $_.ok } | ForEach-Object { $_.step + " (" + $_.detail + ")" }) -join "; ") }
    }

    # ---------- V15: the data survives stop, start and recreate ----------
    Check "V15" "survives" {
        Up
        $before = Get-Tree $Data
        Compose stop | Out-Null
        Compose start | Out-Null
        Wait-Healthy
        $afterRestart = Get-Tree $Data
        Compose down | Out-Null                    # removes the container and its network; there is no volume to remove
        Compose up -d --no-build | Out-Null
        Wait-Healthy
        $afterRecreate = Get-Tree $Data
        $page = Helper pages $State.Port $State.Stem
        $p = @()
        if (-not (Same-List $before $afterRestart)) { $p += "files changed across stop and start" }
        if (-not (Same-List $before $afterRecreate)) { $p += "files changed across remove and recreate" }
        if (-not $before.Count) { $p += "the data folder is empty, so there was nothing to compare" }
        if (-not $page.ok) { $p += "after recreate the pages lost something: " + (@($page.problems) -join ", ") }
        Judge $p ("{0} files on the host have the same hashes after stop/start and after remove/recreate; the UI still shows the results file, the request, the approval and the order sheet" -f $before.Count)
    }

    # ---------- V16: atomic state writes on the mounted folder ----------
    $SingleCode = @'
import json, sys
sys.path.insert(0, '/app')
import statefile
try:
    statefile.atomic_write('/probe/state.json', 'x')
    print(json.dumps({'replaced': True}))
except OSError as e:
    print(json.dumps({'replaced': False, 'error': type(e).__name__, 'errno': e.errno}))
'@
    Check "V16" "atomic-on-mount" {
        Up
        $r = Helper atomic $Name $Data
        $single = "$Work\single"
        Write-Text "$single\state.json" "{}"
        $s = In-Image $Image $SingleCode @("--tmpfs", "/probe:mode=1777", "-v", "$single\state.json:/probe/state.json")
        $control = if ($s.replaced) { "a single mounted file WAS replaced by a rename here" } else { "a rename onto a single mounted file is refused ($($s.error) $($s.errno)), which is why the data is one mounted folder" }
        if (Prop $r "error") { Fail $r.error }
        elseif ($r.ok) { Pass ("200 locked atomic writes while this computer read the file {0} times: never partial, nothing left behind; {1}" -f $r.reads, $control) }
        else { Fail ("writer exit {0} ({1}); partial reads {2}; last value {3}; left behind: {4}; {5}" -f $r.writer_exit, $r.writer_said, $r.partial, $r.last, (@($r.left_behind) -join ", "), $control) }
    }

    # ---------- V17: the time zone ----------
    $ZoneCode = @'
import json, os, time
print(json.dumps({'tz': os.environ.get('TZ'), 'offset': time.strftime('%z')}))
'@
    Check "V17" "time-zone" {
        Up
        $z = In-Container $ZoneCode
        $page = Helper pages $State.Port $State.Stem
        $p = @()
        if ($z.tz -ne $Zone -or $z.offset -ne $ZoneOffset) { $p += "inside the container TZ is '$($z.tz)' and the offset is $($z.offset); wanted $Zone and $ZoneOffset" }
        if (-not "$(Prop $page 'zone')".StartsWith($ZoneLabel)) { $p += "the Health page says '$(Prop $page 'zone')', not $ZoneLabel" }
        Judge $p "with TZ=$Zone the container's offset is $ZoneOffset and the Health page says $(Prop $page 'zone')"
    }

    # ---------- V18: tracing is off by default and points at the host ----------
    $TraceCode = @'
import json, os, socket, urllib.request
out = {'obs': os.environ.get('OBS_ENABLED'), 'base': os.environ.get('LANGFUSE_BASE_URL'), 'host': os.environ.get('LANGFUSE_HOST'), 'get': 'not tried'}
try:
    out['resolves'] = bool(socket.gethostbyname('host.docker.internal'))
except OSError:
    out['resolves'] = False
if DO_GET:
    try:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open((out['base'] or '').rstrip('/') + '/api/public/health', timeout=4) as r:
            out['get'] = 'reachable (HTTP %d)' % r.status
    except Exception as e:
        out['get'] = 'not reachable (%s)' % type(e).__name__
print(json.dumps(out))
'@
    Check "V18" "tracing-default" {
        Up
        $t = In-Container ($TraceCode.Replace("DO_GET", $(if ($NoLangfuseGet) { "False" } else { "True" })))
        $p = @()
        if ($t.obs -ne "0") { $p += "OBS_ENABLED is '$($t.obs)' although the settings file says 1 and LCOM_TRACING is not set" }
        if ($t.base -ne "http://host.docker.internal:3000" -or $t.host -ne "http://host.docker.internal:3000") { $p += "the Langfuse address is not http://host.docker.internal:3000 in both variables (the settings file says localhost)" }
        if (-not $t.resolves) { $p += "host.docker.internal does not resolve inside the container" }
        Judge $p ("tracing off although the settings file switches it on; both Langfuse addresses are http://host.docker.internal:3000; the name resolves; one read-only GET of the health address: $($t.get). No trace was sent.")
    }

    # ---------- V19: the nine suites inside the test image, with no network ----------
    Check "V19" "suites-in-container" {
        Build
        $p = @(); $told = @(); $skipped = @(); $n = 0
        foreach ($suite in $Suites) {
            $n++
            Say "  running $suite inside the test image (no network) ..."
            $out = @(& docker.exe run --rm --init --network none --read-only --tmpfs "/tmp:size=768m" --user 10001:10001 --cap-drop ALL --security-opt no-new-privileges `
                         --label $Label --name "lcom-verify-test-$n" -e OBS_ENABLED=0 --workdir /app --entrypoint python $TestImage -B $suite)
            $code = $LASTEXITCODE
            $summary = @($out | Where-Object { $_ -match "passed|^ok$" } | Select-Object -Last 2) -join " / "
            $skipped += @($out | Where-Object { $_ -like "skip *" } | ForEach-Object { "$suite $_" })
            if ($code -ne 0) { $p += "$suite ended with code $code (" + (@($out | Where-Object { $_ -like "FAIL *" }) -join ", ") + ")" }
            $told += "$suite [$summary]"
        }
        $note = "skipped tests: " + $(if ($skipped.Count) { $skipped -join "; " } else { "none" }) + ". Parts of protected suites that need real output files and so do not run in an image: " + ($DataDependent -join "; ")
        if ($p.Count) { Fail (($p -join "; ") + ". " + $note) } else { Pass ("all nine suites exit 0 as uid 10001 with no network and a read-only system: " + ($told -join "; ") + ". " + $note) }
    }

    # ---------- V20: the source folder as the data folder is refused ----------
    Check "V20" "refuses-source-folder" {
        Build
        $src = "$Work\srcdata"
        New-Item -ItemType Directory -Force "$src\.git" | Out-Null
        Write-Text "$src\email_agent.py" "MY OWN WORKING COPY - DO NOT REPLACE"
        $hash = (Get-FileHash -LiteralPath "$src\email_agent.py").Hash
        & docker.exe run --name lcom-verify-src --label $Label --network none --read-only --tmpfs /tmp -e "TZ=$Zone" -v "${src}:/data" $Image | Out-Null
        $code = $LASTEXITCODE
        $log = Logs lcom-verify-src
        Dk rm lcom-verify-src | Out-Null
        $p = @()
        if ($code -eq 0) { $p += "the container started" }
        if ($log -notmatch "looks like the source folder") { $p += "the message does not say why" }
        if ((Get-FileHash -LiteralPath "$src\email_agent.py").Hash -ne $hash -or @(Get-ChildItem $src -Recurse -File).Count -ne 1) { $p += "the folder was changed" }
        Judge $p "a data folder holding .git: exit code $code, the message names the reason, the planted email_agent.py is byte-identical and nothing was added"
    }

    # ---------- V21: no data folder is refused ----------
    Check "V21" "refuses-no-mount" {
        Build
        & docker.exe run --name lcom-verify-nomount --label $Label --network none --read-only --tmpfs /tmp -e "TZ=$Zone" $Image | Out-Null
        $code = $LASTEXITCODE
        $log = Logs lcom-verify-nomount
        Dk rm lcom-verify-nomount | Out-Null
        if ($code -ne 0 -and $log -match "no data folder is mounted") { Pass "started without a data folder: exit code $code and a message that says so" }
        else { Fail "exit code $code; message: $log" }
    }

    # ---------- V22: a second container with the same name ----------
    Check "V22" "one-name" {
        Up
        & docker.exe run -d --name $Name --label $Label --network none -e "TZ=$Zone" $Image | Out-Null
        $code = $LASTEXITCODE
        $count = @(& docker.exe ps -a --filter "name=$Name" --format "{{.Names}}" | Where-Object { $_ -eq $Name }).Count
        if ($code -ne 0 -and $count -eq 1) { Pass "Docker refused a second container named $Name (exit code $code); still exactly one" }
        else { Fail "exit code $code; containers with that name: $count" }
    }

    # ---------- V23: a state file held open on this computer (information only) ----------
    $LockCode = @'
import json, sys
sys.path.insert(0, '/data')
import statefile
try:
    statefile.atomic_write('/data/decision_state.json', '{"verify": "host lock"}')
    print(json.dumps({'written': True}))
except OSError as e:
    print(json.dumps({'written': False, 'error': type(e).__name__, 'errno': e.errno}))
'@
    Check "V23" "host-lock" {
        Up
        $target = "$Data\decision_state.json"
        $saved = [System.IO.File]::ReadAllBytes($target)
        $hashBefore = (Get-FileHash -LiteralPath $target).Hash
        $hold = [System.IO.File]::Open($target, "Open", "ReadWrite", "None")
        try { $r = In-Container $LockCode } finally { $hold.Close() }
        $same = ((Get-FileHash -LiteralPath $target).Hash -eq $hashBefore)
        $left = @(Get-ChildItem $Data -File | Where-Object { $_.Name -like "decision_state.json.*" } | ForEach-Object { $_.Name })
        if (-not $same) { [System.IO.File]::WriteAllBytes($target, $saved) }
        Info ("while this computer held decision_state.json open: the container's write " + $(if ($r.written) { "went through" } else { "was refused ($($r.error) $($r.errno))" }) +
              "; the file is " + $(if ($same) { "unchanged" } else { "CHANGED (restored for the rest of this run)" }) + "; left behind: " + $(if ($left.Count) { $left -join ", " } else { "nothing" }))
    }
}
finally {
    # ---------- V24: after, and cleanup ----------
    $after = Get-RealState
    $changed = @($State.Before.Keys | Where-Object { $State.Before[$_] -ne $after[$_] } | Sort-Object)
    $othersAfter = Get-Others
    $sameContainers = Same-List $State.Others.Containers $othersAfter.Containers
    $sameVolumes = Same-List $State.Others.Volumes $othersAfter.Volumes
    Say ""
    Say "The real project, before and after (sha256, first 16 characters; values are never read into this report):"
    foreach ($k in ($State.Before.Keys | Sort-Object)) { Say ("  {0,-28} {1}  ->  {2}  {3}" -f $k, $State.Before[$k], $after[$k], $(if ($State.Before[$k] -eq $after[$k]) { "same" } else { "CHANGED" })) }
    Say "Containers that are not this tool's (name and state), before:  $(@($State.Others.Containers) -join ', ')"
    Say "Containers that are not this tool's (name and state), after:   $(@($othersAfter.Containers) -join ', ')"
    Say "Volumes that are not this tool's, before:  $(@($State.Others.Volumes) -join ', ')"
    Say "Volumes that are not this tool's, after:   $(@($othersAfter.Volumes) -join ', ')"
    Add-Result "V24" "after" ($changed.Count -eq 0 -and $sameContainers -and $sameVolumes) $(
        if ($changed.Count -eq 0 -and $sameContainers -and $sameVolumes) { "the real project's {0} items are unchanged; containers and volumes of other projects are identical before and after" -f $State.Before.Count }
        else { "changed in the real project: " + ($changed -join ", ") + "; other containers identical: $sameContainers; other volumes identical: $sameVolumes" })

    $containers = @(& docker.exe ps -a --filter "label=$Label" --format "{{.Names}}")
    $images = @(& docker.exe images --filter "label=$Label" --format "{{.Repository}}:{{.Tag}}")
    $networks = @(& docker.exe network ls --filter "label=$Label" --format "{{.Name}}")
    $volumes = @(& docker.exe volume ls --filter "label=$Label" --format "{{.Name}}")
    $mineC = @($containers | Where-Object { $_ -like "lcom-verify*" })
    $mineI = @($images | Where-Object { $_ -in @($Image, $TestImage, $ProbeImage) })
    $mineN = @($networks | Where-Object { $_ -like "lcom-verify*" })
    Say ""
    Say "Docker objects carrying the label $($Label):"
    Say "  containers: $($containers -join ', ')"
    Say "  images:     $($images -join ', ')"
    Say "  networks:   $($networks -join ', ')"
    Say "  volumes:    $($volumes -join ', ')"
    if ($Keep) {
        Say "Kept (-Keep). To remove this run's objects later:"
        Say "  docker rm -f $($mineC -join ' ') ; docker rmi $($mineI -join ' ') ; docker network rm $($mineN -join ' ')"
        Say "  Remove-Item -Recurse -Force '$Work'"
    }
    else {
        Say "Removing only this run's own: containers [$($mineC -join ', ')], images [$($mineI -join ', ')], networks [$($mineN -join ', ')], and the folder $Work"
        foreach ($c in $mineC) { if (Test-Ours $c) { & docker.exe rm -f $c | Out-Null } }
        foreach ($i in $mineI) { & docker.exe rmi $i | Out-Null }
        foreach ($n in $mineN) { & docker.exe network rm $n | Out-Null }
        if (Test-Path -LiteralPath $Work) { Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue }
        Say "Left in place on purpose: the base image and the build cache (they cannot carry the label), and any other container of this tool that is not from this run."
    }
}

Say ""
$Results | Format-Table Id, Check, Result -AutoSize | Out-String | Write-Host
$failed = @($Results | Where-Object { $_.Result -eq "FAIL" })
Say ("{0} checks: {1} passed, {2} failed, {3} for information" -f $Results.Count, @($Results | Where-Object { $_.Result -eq "PASS" }).Count, $failed.Count, @($Results | Where-Object { $_.Result -eq "INFO" }).Count)
if ($failed.Count) { exit 1 }
