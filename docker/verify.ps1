<#
Checks the Docker packaging of the sourcing tool, in four stages. For Windows PowerShell 5.1 with Docker Desktop.

    $env:OBS_ENABLED = "0"
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 1 -WriteConstraints   # V01-V06; -WriteConstraints the first time only
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 2                      # V07-V13
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 3                      # V14-V19
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Stage 4                      # V20-V24, then cleans up
    powershell -NoProfile -ExecutionPolicy Bypass -File docker\verify.ps1 -Cleanup                      # V24 now: the after-check, then cleanup

  Stage 1 starts a run; stages 2 to 4 continue it, in order, each once. Between stages the run's images, its container
  and its temp folder stay in place so the next stage can use them. A failing check stops its stage: the check's name and
  the reason are printed, nothing is cleaned up (so it can be looked at), and the run can only be removed with -Cleanup.

  -Stage             1, 2, 3 or 4
  -Cleanup           the after-check (V24) and removal of this run's objects, at any point
  -WriteConstraints  stage 1: writes docker\constraints.txt from the first image's installed packages, then builds again with it
  -Only              one run of just these checks, with cleanup at the end (for the mutation run); not with -Stage
  -Root              stage 1 or -Only: verify the Docker files in another folder; default: this project
  -Keep              with -Only or -Cleanup: leave the objects and the temp folder in place
  -NoLangfuseGet     skip the one read-only GET of the Langfuse health address (V18)
  -ImageRepo         with -Only: the images' repository name, lcom-<something> (the mutation run uses lcom-mut-<n>); default lcom-ui
  -FakeEnvWithoutZone  with -Only: the fake .env has no TZ line (mutation 9b)

What it uses: fake keys only; a throwaway data folder and a staged build context under %TEMP% (the real .env is never
inside a build context); names that start with lcom-verify; the label com.zync.lcom=1.
What it never does: touch a container, image, network or volume that is not its own; run a prune; push an image; call a
model, Nimble or a mail server; send a trace; print the value of a setting. It reads the real project only to hash it.
Needs: python with the project's requirements on this computer (docker\verify_helper.py drives the UI over HTTP).
#>
param(
    [ValidateSet("", "1", "2", "3", "4")][string]$Stage = "",
    [switch]$Cleanup,
    [string[]]$Only = @(),
    [string]$Root = "",
    [switch]$WriteConstraints,
    [switch]$Keep,
    [switch]$NoLangfuseGet,
    [ValidatePattern('^lcom-[a-z0-9-]+$')][string]$ImageRepo = "lcom-ui",
    [switch]$FakeEnvWithoutZone
)
Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
$env:OBS_ENABLED = "0"
$Only = @($Only | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })      # -File passes "V04,V09" as one word

$Label = "com.zync.lcom=1"
$Real = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path                  # the real project: only ever hashed and listed
$HelperPy = Join-Path $Real "docker\verify_helper.py"
$RunFile = Join-Path $env:TEMP "lcom-verify-run.json"                       # the open staged run, between stages
$Name = "lcom-verify-ui"
$Zone = "Asia/Kolkata"; $ZoneOffset = "+0530"; $ZoneLabel = "UTC+05:30"
$StageOf = @{ V01 = 1; V02 = 1; V03 = 1; V04 = 1; V05 = 1; V06 = 1; V07 = 2; V08 = 2; V09 = 2; V10 = 2; V11 = 2; V12 = 2; V13 = 2;
              V14 = 3; V15 = 3; V16 = 3; V17 = 3; V18 = 3; V19 = 3; V20 = 4; V21 = 4; V22 = 4; V23 = 4; V24 = 4 }
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
# The single approved change to a protected file (2026-10-14: workflow_id treats \ like /, so a Windows path gives the same
# id on Linux). V24 accepts this file differing from a run's start only when it now has exactly this hash.
$ApprovedChanges = @{ "observability.py" = "d0b4b1ee38e0ac52b2f31000c871512c8a989c27f169a54947f2251ad1d2a1b4" }
$Results = New-Object System.Collections.ArrayList
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:Stopped = ""

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

function Check($checkId, $checkName, [scriptblock]$body) {
    # (parameter names must not repeat a script variable: PowerShell ignores case, so a parameter $id would hide the run's $Id
    # inside every check body)
    if ($Mode -eq "stage" -and $StageOf[$checkId] -ne [int]$Stage) { return }
    if ($Mode -eq "only" -and $Only -notcontains $checkId) { return }
    if ($Mode -eq "cleanup" -or $script:Stopped) { return }
    $script:Ok = $false; $script:Detail = "the check set no result"
    try { & $body | Out-Null } catch { Fail ("could not run: " + $_.Exception.Message) }
    Add-Result $checkId $checkName $script:Ok $script:Detail
    if ($Mode -eq "stage" -and $script:Ok -eq $false) {
        $script:Stopped = "$checkId $checkName"
        Say ""
        Say "STAGE $Stage STOPPED at $checkId ($checkName): $($script:Detail)"
    }
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

function Invoke-Compose {
    # folder, project, then the compose arguments. No named parameters on purpose: PowerShell would bind a docker flag such
    # as -d to a parameter whose name starts with d.
    $dir = $args[0]; $project = $args[1]
    $rest = @($args | Select-Object -Skip 2)          # always a list: one word such as "stop" would otherwise be spread into letters
    $out = & docker.exe compose -p $project --project-directory $dir -f "$dir\docker-compose.yml" --env-file "$dir\.env" @rest
    if ($LASTEXITCODE -ne 0) { throw ("docker compose " + (@($rest) -join " ") + " ended with code $LASTEXITCODE") }
    return $out
}

function Compose { return (Invoke-Compose $Proj "lcom-verify" @args) }

function Split-Docker {
    # exit code, output lines and error lines of one docker command, kept apart; never throws
    $old = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try { $all = @(& docker.exe @args 2>&1) } finally { $ErrorActionPreference = $old }
    return [pscustomobject]@{ Code = $LASTEXITCODE
                              Out = @($all | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" })
                              Err = @($all | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" }) }
}

function Quiet-Docker {
    # stdout of a docker command whose error text is not wanted; never throws. (In Windows PowerShell, redirecting a
    # native command's error output while errors stop the script would turn the first error line into a stop.)
    $old = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try { $out = & docker.exe @args 2>$null } finally { $ErrorActionPreference = $old }
    return $out
}

function Test-Ours($container) {
    $names = @(& docker.exe ps -a --filter "name=$container" --filter "label=$Label" --format "{{.Names}}")
    return ($names -contains $container) -and ($container -like "lcom-verify*")
}

function Assert-Ours($container) {
    if (-not (Test-Ours $container)) { throw "'$container' is not one of this tool's verify containers (label $Label, name lcom-verify*): not touched" }
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

function Get-FileSha($path) { if (Test-Path -LiteralPath $path) { return (Get-FileHash -LiteralPath $path).Hash.ToLower() } else { return "absent" } }

function Get-TextHash($text) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    return ([System.BitConverter]::ToString($sha.ComputeHash($Utf8.GetBytes("$text"))) -replace "-", "").Substring(0, 16).ToLower()
}

function Get-ListingHash($folder) {
    return (Get-TextHash ((@(Get-ChildItem -LiteralPath $folder -Recurse -File | ForEach-Object { "{0}|{1}|{2}" -f $_.FullName, $_.Length, $_.LastWriteTimeUtc.Ticks } | Sort-Object) -join "`n")))
}

function Get-RealState {
    # Hashes of the files that must not change, and a listing of the real output folders. Read-only. Values are never printed.
    $out = [ordered]@{}
    foreach ($f in $Protected) {
        $p = Join-Path $Real $f
        $out[$f] = if (Test-Path -LiteralPath $p) { (Get-FileHash -LiteralPath $p).Hash.Substring(0, 16).ToLower() } else { "absent" }
    }
    foreach ($d in $RealFolders) {
        $p = Join-Path $Real $d
        $out["$d\"] = if (Test-Path -LiteralPath $p) { "{0} files, listing {1}" -f @(Get-ChildItem -LiteralPath $p -Recurse -File).Count, (Get-ListingHash $p) } else { "absent" }
    }
    $out["git status --short"] = Get-TextHash ((@(& git -C $Real status --short) -join "`n"))
    return [pscustomobject]$out
}

function Get-Others {
    # Every container (name, image and state; never the uptime text) and volume that does NOT carry this tool's label: the
    # proof nothing else was touched.
    $mine = @(& docker.exe ps -a --filter "label=$Label" --format "{{.ID}}")
    $containers = @(& docker.exe ps -a --format "{{.ID}}|{{.Names}} | {{.Image}} | {{.State}}" | Where-Object { $mine -notcontains ($_ -split "\|")[0] } | ForEach-Object { ($_ -split "\|", 2)[1] } | Sort-Object)
    $myVolumes = @(& docker.exe volume ls --filter "label=$Label" --format "{{.Name}}")
    $volumes = @(& docker.exe volume ls --format "{{.Name}}" | Where-Object { $myVolumes -notcontains $_ } | Sort-Object)
    return [pscustomobject]@{ Containers = $containers; Volumes = $volumes }
}

function Save-Run { Write-Text $RunFile ($Run | ConvertTo-Json -Depth 6) }

# ---------- which run, which mode ----------

$Mode = if ($Cleanup) { "cleanup" } elseif ($Stage) { "stage" } elseif ($Only.Count) { "only" } else { "" }
if (-not $Mode -or ($Stage -and ($Only.Count -or $Cleanup)) -or (($ImageRepo -ne "lcom-ui" -or $FakeEnvWithoutZone) -and $Mode -ne "only")) {
    Say "Run one stage at a time:  -Stage 1 (add -WriteConstraints the first time), then -Stage 2, -Stage 3, -Stage 4.  -Cleanup removes a run at any point."
    exit 1
}
$open = Test-Path -LiteralPath $RunFile
if (($Mode -eq "stage" -and $Stage -eq "1") -or $Mode -eq "only") {
    if ($open) {
        $prev = Get-Content -LiteralPath $RunFile -Raw | ConvertFrom-Json
        Say "STOPPED: run $($prev.Id) is still open (stages done: $(@($prev.Done) -join ', '); stopped at: $(if ($prev.Stopped) { $prev.Stopped } else { 'nothing' })). Finish it with the next stage, or remove it with -Cleanup. Nothing was changed."
        exit 1
    }
    $rootDir = if ($Root) { (Resolve-Path $Root).Path } else { $Real }
    $Run = [pscustomobject]@{ Id = (Get-Date -Format "yyyyMMddHHmmss"); Root = $rootDir; Port = (Get-FreePort); Done = @(); Stopped = ""; Before = $null; Others = $null
                              Head = ""; ConstraintsStart = ""; ConstraintsWritten = ""; RecordedLater = "" }
}
else {
    if (-not $open) { Say "There is no open run. Start one with -Stage 1. Nothing was changed."; exit 1 }
    $Run = Get-Content -LiteralPath $RunFile -Raw | ConvertFrom-Json
    if ($Root -and (Resolve-Path $Root).Path -ne $Run.Root) { Say "STOPPED: this run verifies $($Run.Root); -Root cannot change between stages."; exit 1 }
    if ($Mode -eq "stage") {
        $need = [int]$Stage - 1
        if ($Run.Stopped) { Say "STOPPED: run $($Run.Id) stopped at $($Run.Stopped). Remove it with -Cleanup, then start again with -Stage 1."; exit 1 }
        if (@($Run.Done) -contains [int]$Stage) { Say "STOPPED: stage $Stage of run $($Run.Id) has already run. Each stage runs once; -Cleanup, then -Stage 1, to run again."; exit 1 }
        if (@($Run.Done) -notcontains $need) { Say "STOPPED: stage $need of run $($Run.Id) has not passed yet. Run -Stage $need first."; exit 1 }
    }
}
$Root = $Run.Root; $Id = $Run.Id
$Work = Join-Path $env:TEMP "lcom-verify-$Id"
$Ctx = "$Work\ctx"; $Proj = "$Work\compose"; $Data = "$Work\data"
$Image = "${ImageRepo}:verify-$Id"; $TestImage = "${ImageRepo}:verify-test-$Id"; $ProbeImage = "${ImageRepo}:verify-ctx-$Id"
$Needle = "sk-ant-FAKE-verify-PLANTED-$Id"                                  # planted in the build context; must never be in an image
$StateNeedle = "planted-state-file-$Id"
$ProxyUser = "fakeproxyuser$Id"; $ProxyPass = "FAKE-proxy-pass-$Id"          # in the fake proxy settings; must appear nowhere
$State = @{ Built = $false; Up = $false; Stem = "sourcing_results_20260101_100000" }

# ---------- preparations (each runs once, when a check first needs it; a later stage finds them done) ----------

function New-FakeEnv($dir, $port, $container, $data, [switch]$NoZone) {
    # Fake values only. Some lines try to change what the container must decide itself; the proxy lines make any
    # accidental outbound call fail at once.
    $lines = @()
    if (-not $NoZone) { $lines += "TZ=$Zone" }
    $lines += @("LCOM_PORT=$port", "LCOM_DATA=$($data -replace '\\', '/')", "LCOM_NAME=$container", "LCOM_IMAGE=$Image",
                'ANTHROPIC_API_KEY="sk-ant-FAKE-verify-000111222333"', "NIMBLE_API_KEY=nimble-FAKE-verify-444555", "SMTP_HOST=smtp.example.invalid",
                "SMTP_USER=sender@example.invalid", "SMTP_PASSWORD='FAKE-pa`$`$word-with#hash and space'",
                'SHIPPING_ADDRESS="Test Receiving\n1 Test Street\nTestville, TX 00000"', "BRAND_NAME=TestBrand", 'SENDER_NAME="Test Sender"',
                "OBS_ENABLED=1", "LANGFUSE_BASE_URL=http://localhost:3000", "LANGFUSE_HOST=http://localhost:3000", "LANGFUSE_PUBLIC_KEY=pk-lf-FAKE-verify",
                "LANGFUSE_SECRET_KEY=sk-lf-FAKE-verify", "UI_HOST=10.9.8.7", "UI_ENABLE_STUB=1", "PROJECT_DIR=/tmp/elsewhere",
                "HTTPS_PROXY=http://${ProxyUser}:${ProxyPass}@127.0.0.1:9", "HTTP_PROXY=http://${ProxyUser}:${ProxyPass}@127.0.0.1:9",
                "NO_PROXY=internal.example")
    Write-Text "$dir\.env" (($lines -join "`n") + "`n")
    Copy-Item -LiteralPath (Join-Path $Root "docker-compose.yml") -Destination "$dir\docker-compose.yml" -Force
}

function Prepare-Context {
    if (Test-Path -LiteralPath "$Ctx\Dockerfile") { return }
    foreach ($need in @("scratch\pre_gates\decision_agent.py", "scratch\pre_gates\email_agent.py")) {
        if (-not (Test-Path -LiteralPath (Join-Path $Root $need))) {
            throw "$need is missing. The test image needs it: test_gates.py compares today's agents with these two earlier copies. Nothing was built."
        }
    }
    New-Item -ItemType Directory -Force $Ctx, $Proj, $Data | Out-Null
    $allowed = @(Get-Content (Join-Path $Root "docker\image_files.txt")) + @(Get-Content (Join-Path $Root "docker\image_files_test.txt")) | Where-Object { $_ }
    # the allowed files, the Docker files, and the project's other Docker-side files (which must NOT get into an image)
    $copy = $allowed + @("Dockerfile", ".dockerignore", "docker-compose.yml", ".env.docker.example", "DOCKER.md", "docker/verify.ps1",
                         "docker/verify_helper.py", "docker/image_files.txt", "docker/image_files_test.txt")
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
    New-FakeEnv $Proj $Run.Port $Name $Data -NoZone:$FakeEnvWithoutZone
}

function Build-Images {
    Prepare-Context
    if ($State.Built) { return }
    Quiet-Docker image inspect $Image $TestImage --format "{{.Id}}" | Out-Null
    if ($LASTEXITCODE -eq 0) { $State.Built = $true; return }            # built by an earlier stage of this run
    Say "building the runtime and test images (the first build downloads the base image and the packages; several minutes) ..."
    Dk build -q --target runtime -t $Image --label $Label $Ctx | Out-Null
    Dk build -q --target test -t $TestImage --label $Label $Ctx | Out-Null
    $State.Built = $true
}

function Wait-Healthy($container = $Name, $seconds = 150) {
    $end = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $end) {
        $s = @(& docker.exe inspect --format "{{.State.Status}}|{{if .State.Health}}{{.State.Health.Status}}{{end}}" $container)[0]
        if ($s -eq "running|healthy") { return }
        if ($s -like "exited*") { throw ("the container stopped. Its last lines:`n" + (Logs $container)) }
        Start-Sleep -Seconds 2
    }
    throw ("the container did not become healthy in $seconds seconds. Its last lines:`n" + (Logs $container))
}

function Start-Ui {
    Build-Images
    if ($State.Up) { return }
    if (-not (Test-Path -LiteralPath "$Data\Excel Output Sheets\$($State.Stem).xlsx")) { $State.Stem = (Helper fixture $Data).stem }
    Compose up -d --no-build | Out-Null                                   # starts it, or leaves the running one as it is
    Wait-Healthy
    $State.Up = $true
}

function Finish-Run {
    # V24: the real project and other projects' containers and volumes are as they were; then only this run's objects go.
    # Pass or fail rests on the hashes of the protected files, the real state files and output folders, and
    # docker\constraints.txt. git status and HEAD are secondary evidence: reported, never a failure on their own.
    $after = Get-RealState
    $gitKey = "git status --short"
    $changed = @($Run.Before.PSObject.Properties | Where-Object { $_.Name -ne $gitKey -and $_.Value -ne (Prop $after $_.Name) } | ForEach-Object { $_.Name })
    $approvedNote = @()
    foreach ($name in $ApprovedChanges.Keys) {
        $now = Get-FileSha (Join-Path $Real $name)
        $approvedNote += "$name is " + $(if ($now -eq $ApprovedChanges[$name]) { "the approved version (sha256 $now)" } else { "NOT the approved version (sha256 $now)" })
        if ($changed -contains $name -and $now -eq $ApprovedChanges[$name]) { $changed = @($changed | Where-Object { $_ -ne $name }) }   # the one approved exception
    }
    $approvedNote = $approvedNote -join "; "
    $cNow = Get-FileSha (Join-Path $Root "docker\constraints.txt")
    $cStart = Prop $Run "ConstraintsStart"; $cWritten = Prop $Run "ConstraintsWritten"
    if ($cStart -and $cNow -eq $cStart) { $cNote = "docker\constraints.txt is unchanged since the run started" }
    elseif ($cWritten -and $cNow -eq $cWritten) { $cNote = "docker\constraints.txt differs from the start only by what this run wrote with -WriteConstraints (same hash as recorded then)" }
    else { $changed += "docker\constraints.txt"; $cNote = "docker\constraints.txt changed, and not to what this run wrote" }
    $headStart = Prop $Run "Head"; $headNow = "$(& git -C $Real rev-parse HEAD)".Trim()
    $statusNow = @(& git -C $Real status --short)
    if ($headStart -and $headNow -ne $headStart) { $gitNote = "HEAD changed during the run: $headStart -> $headNow (a commit made outside this run); git status is not compared" }
    elseif ((Prop $Run.Before $gitKey) -eq (Prop $after $gitKey)) { $gitNote = "HEAD unchanged; git status --short is the same as at the start" }
    else { $gitNote = "HEAD unchanged; git status --short differs from the start (secondary evidence, not a failure). Now: " + ($statusNow -join "; ") }
    $othersAfter = Get-Others
    $sameContainers = Same-List $Run.Others.Containers $othersAfter.Containers
    $sameVolumes = Same-List $Run.Others.Volumes $othersAfter.Volumes
    Say ""
    Say "The real project, before (stage 1) and now (sha256, first 16 characters; values are never read into this report):"
    foreach ($p in $Run.Before.PSObject.Properties) {
        $tag = if ($p.Name -eq $gitKey) { "(secondary)" } elseif ($p.Value -eq (Prop $after $p.Name)) { "same" } else { "CHANGED" }
        Say ("  {0,-28} {1}  ->  {2}  {3}" -f $p.Name, $p.Value, (Prop $after $p.Name), $tag)
    }
    Say "  $cNote"
    Say "  approved exception: $approvedNote"
    Say "  $gitNote"
    if (Prop $Run "RecordedLater") { Say "  Note: $($Run.RecordedLater)" }
    Say "Containers that are not this tool's (name | image | state), before:  $(@($Run.Others.Containers) -join ', ')"
    Say "Containers that are not this tool's (name | image | state), now:     $(@($othersAfter.Containers) -join ', ')"
    Say "Volumes that are not this tool's, before:  $(@($Run.Others.Volumes) -join ', ')"
    Say "Volumes that are not this tool's, now:     $(@($othersAfter.Volumes) -join ', ')"
    $ok = $changed.Count -eq 0 -and $sameContainers -and $sameVolumes
    Add-Result "V24" "after" $ok $(if ($ok) { "all {0} recorded items (protected files, real state files, output folders) are unchanged or the approved exception; {1}; {2}; containers and volumes of other projects identical; {3}" -f (@($Run.Before.PSObject.Properties).Count - 1), $approvedNote, $cNote, $gitNote }
                                   else { "changed: " + ($changed -join ", ") + "; other containers identical: $sameContainers; other volumes identical: $sameVolumes; $gitNote" })

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
        if ($Mode -ne "only" -and (Test-Path -LiteralPath $RunFile)) { Remove-Item -LiteralPath $RunFile -Force }
        Say "Left in place on purpose: the base image and the build cache (they cannot carry the label), and every container of this tool that is not from a verify run."
    }
}

# ---------- V01: the engine (every stage needs it) ----------

$server = @(Quiet-Docker info --format "{{.ServerVersion}}")
if (-not ($server -and "$($server[0])".Trim())) {
    Say "STOPPED: the Docker engine does not answer (docker info gave no server version). Start Docker Desktop, wait until it is running, and run this again. Nothing was built or changed."
    exit 2
}
if ($Mode -eq "only" -or ($Mode -eq "stage" -and $Stage -eq "1")) {
    Add-Result "V01" "engine" $true ("Docker engine {0}, client {1}, {2}" -f $server[0], (& docker.exe version --format "{{.Client.Version}}"), (@(& docker.exe compose version)[0]))
}

# ---------- V02: before ----------

if ($Mode -eq "only" -or ($Mode -eq "stage" -and $Stage -eq "1")) {
    $Run.Before = Get-RealState
    $Run.Others = Get-Others
    $Run.Head = "$(& git -C $Real rev-parse HEAD)".Trim()
    $Run.ConstraintsStart = Get-FileSha (Join-Path $Root "docker\constraints.txt")
    Add-Result "V02" "before" $null ("recorded {0} items of the real project; containers not from this tool: {1}; volumes not from this tool: {2}" -f @($Run.Before.PSObject.Properties).Count, (@($Run.Others.Containers) -join ", "), (@($Run.Others.Volumes) -join ", "))
    if ($Mode -eq "stage") { Save-Run }
}

$ListCode = @'
import json, os, sys
root = sys.argv[1] if len(sys.argv) > 1 else '/app'
out = []
for d, _, files in os.walk(root):
    out += [os.path.relpath(os.path.join(d, f), root).replace(os.sep, '/') for f in files]
print(json.dumps(sorted(out)))
'@

try {
    # ================= stage 1: V03-V06 =================

    Check "V03" "staged-context" {
        Prepare-Context
        $p = @()
        foreach ($f in @("Dockerfile", ".dockerignore", "docker-compose.yml")) {
            if ((Get-FileHash -LiteralPath (Join-Path $Root $f)).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Ctx $f)).Hash) { $p += "$f in the staged copy differs from the one under test" }
        }
        if (Test-Path -LiteralPath (Join-Path $Real ".env")) {
            if ((Get-FileHash -LiteralPath (Join-Path $Real ".env")).Hash -eq (Get-FileHash -LiteralPath "$Ctx\.env").Hash) { $p += "the staged .env is the real one" }
        }
        Judge $p ("{0} files staged under {1}; a fake .env, a fake key file and 24 decoy files planted; the real .env is not in it" -f @(Get-ChildItem $Ctx -Recurse -File).Count, $Ctx)
    }

    Check "V04" "context-exact" {
        Prepare-Context
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

    Check "V05" "build" {
        Build-Images
        $freeze = @(Dk run --rm --network none --label $Label --entrypoint python $Image -m pip freeze | Where-Object { $_ } | Sort-Object)
        $pinsFile = Join-Path $Root "docker\constraints.txt"
        if ($WriteConstraints) {
            $head = @(Get-Content $pinsFile | Where-Object { $_ -like "#*" -and $_ -notlike "# NOT FILLED IN*" -and $_ -notlike "# and docker*" })
            Write-Text $pinsFile ((@($head) + $freeze) -join "`n")
            $Run.ConstraintsWritten = Get-FileSha $pinsFile                 # V24 accepts this file changing only to exactly this
            Copy-Item -LiteralPath $pinsFile -Destination "$Ctx\docker\constraints.txt" -Force
            foreach ($i in @($Image, $TestImage)) { & docker.exe rmi $i | Out-Null }
            $State.Built = $false
            Build-Images
            $freeze = @(Dk run --rm --network none --label $Label --entrypoint python $Image -m pip freeze | Where-Object { $_ } | Sort-Object)
        }
        $pins = @(Get-Content $pinsFile | Where-Object { $_ -and $_ -notlike "#*" } | Sort-Object)
        $from = (@(Get-Content "$Ctx\Dockerfile" | Where-Object { $_ -match "^FROM\s+\S+" })[0] -replace "^FROM\s+(\S+).*$", '$1')
        $size = Dk image inspect --format "{{.Size}}" $Image
        $p = @()
        if ($from -notmatch "^python:3\.13\.5-slim-bookworm@sha256:[0-9a-f]{64}$") { $p += "the base image is not pinned by tag and digest: $from" }
        if (-not $pins.Count) { $p += "docker\constraints.txt has no pins yet: run stage 1 with -WriteConstraints" }
        elseif (-not (Same-List $pins $freeze)) { $p += "the installed packages differ from docker\constraints.txt: " + ((Compare-Object $pins $freeze | ForEach-Object { $_.SideIndicator + $_.InputObject }) -join " ") }
        Judge $p ("runtime and test images built from $from; {0} packages installed, all pinned; runtime image {1:N0} MB" -f $freeze.Count, ([double]$size / 1MB))
    }

    Check "V06" "image-files" {
        Build-Images
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

    # ================= stage 2: V07-V13 =================

    Check "V07" "no-secrets-in-layers" {
        Build-Images
        $p = @(); $told = @()
        foreach ($pair in @(@($Image, "runtime"), @($TestImage, "test"))) {
            $tar = "$Work\$($pair[1]).tar"
            Dk save $pair[0] -o $tar | Out-Null
            $r = Helper scan $tar $pair[1] $Needle $StateNeedle $ProxyUser $ProxyPass
            Remove-Item -LiteralPath $tar -Force
            if (Prop $r "error") { $p += "$($pair[1]): the scan failed: $($r.error)" }
            elseif (-not $r.ok) { $p += "$($pair[1]): forbidden names: " + (@($r.bad_names) -join ", ") + "; planted text found in: " + (@($r.needle_hits) -join ", ") }
            else { $told += "{0}: {1} layers, {2} files" -f $pair[1], $r.layers, $r.files }
        }
        Judge $p ("no .env, state file, output folder, planted text or fake proxy credentials in any layer or in the image settings (" + ($told -join "; ") + ")")
    }

    Check "V08" "image-config" {
        Build-Images
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

    Check "V09" "compose-config" {
        Prepare-Context
        $svc = ((Compose config --format json) -join "`n" | ConvertFrom-Json).services.ui      # never printed whole: it holds every setting
        $envs = Prop $svc "environment"
        $p = @()
        $ports = @(Prop $svc "ports")
        if ($ports.Count -ne 1) { $p += "expected one published port, found $($ports.Count)" }
        else {
            if ((Prop $ports[0] "host_ip") -ne "127.0.0.1") { $p += "the port is published on '$(Prop $ports[0] 'host_ip')', not on 127.0.0.1 only" }
            if ("$(Prop $ports[0] 'published')" -ne "$($Run.Port)" -or "$(Prop $ports[0] 'target')" -ne "$($Run.Port)") { $p += "the port is not the same number inside and outside" }
        }
        if ((Prop $svc "read_only") -ne $true) { $p += "the root filesystem is not read-only" }
        if (@(Prop $svc "cap_drop") -notcontains "ALL") { $p += "capabilities are not all dropped" }
        if (@(Prop $svc "security_opt") -notcontains "no-new-privileges:true") { $p += "no-new-privileges is not set" }
        if ((Prop $svc "user") -ne "10001:10001") { $p += "the user is '$(Prop $svc 'user')'" }
        if ((Prop $svc "init") -ne $true) { $p += "init is off" }
        if ((Prop $svc "restart") -ne "no") { $p += "the restart policy is '$(Prop $svc 'restart')', not 'no'" }
        $extraHosts = "$((Prop $svc 'extra_hosts') | ConvertTo-Json -Compress)"    # as text: a list that is missing altogether must fail too
        if ($extraHosts -notmatch "host\.docker\.internal") { $p += "host.docker.internal is not mapped" }
        if (-not (@(Prop $svc "tmpfs") -join " ").StartsWith("/tmp")) { $p += "/tmp is not a memory folder" }
        $mounts = @(Prop $svc "volumes")
        if ($mounts.Count -ne 1) { $p += "expected exactly one mount (the data folder), found $($mounts.Count)" }
        elseif ((Prop $mounts[0] "type") -ne "bind" -or (Prop $mounts[0] "target") -ne "/data") { $p += "the one mount is not a folder bound at /data" }
        elseif (-not (Test-Path -LiteralPath (Prop $mounts[0] "source") -PathType Container)) { $p += "the mounted path is not a folder (a single file cannot be replaced by a rename)" }
        $want = @{ UI_HOST = "0.0.0.0"; UI_ALLOW_NON_LOOPBACK = "1"; UI_ENABLE_STUB = "0"; PROJECT_DIR = "/data"; UI_DATA_DIR = "/data/ui_data"; UI_PORT = "$($Run.Port)";
                   OBS_ENABLED = "0"; LANGFUSE_BASE_URL = "http://host.docker.internal:3000"; LANGFUSE_HOST = "http://host.docker.internal:3000"; TZ = $Zone }
        foreach ($k in $want.Keys) { if ("$(Prop $envs $k)" -ne $want[$k]) { $p += "$k is not what the compose file must set (the settings file must not be able to change it)" } }
        $log = Prop (Prop $svc "logging") "options"
        if ((Prop $log "max-size") -ne "10m" -or "$(Prop $log 'max-file')" -ne "3") { $p += "the log limits are not 10m and 3 files" }
        if ((Prop (Prop $svc "build") "target") -ne "runtime") { $p += "compose would build '$(Prop (Prop $svc 'build') 'target')', not the runtime target" }
        if ((Prop (Prop $svc "labels") "com.zync.lcom") -ne "1") { $p += "the container label is missing" }
        Judge $p "published on 127.0.0.1 only, same port inside and outside; read-only; no capabilities; no-new-privileges; user 10001; init; restart no; one folder mount; log limits; the pinned settings win over the settings file; TZ passed through; builds runtime only"
    }

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
def root_read_only():
    # the mount options of / as the kernel has them (the last entry for / is the one in use); file rights cannot hide them
    flags = None
    for line in open('/proc/self/mountinfo'):
        parts = line.split()
        if len(parts) > 5 and parts[4] == '/':
            flags = parts[5].split(',')
    return flags is not None and 'ro' in flags
print(json.dumps({'uid': os.getuid(), 'gid': os.getgid(), 'app': writable('/app/_verify'), 'root': writable('/_verify'), 'tmp': writable('/tmp/_verify'),
                  'rootro': root_read_only(),
                  'data': writable('/data/_verify_w'), 'nonewprivs': status.get('NoNewPrivs', '').strip(), 'capbnd': status.get('CapBnd', '').strip(),
                  'pid1': open('/proc/1/comm').read().strip(), 'ui_host': os.environ.get('UI_HOST'), 'project': os.environ.get('PROJECT_DIR'),
                  'stub': os.environ.get('UI_ENABLE_STUB')}))
'@
    Check "V10" "up-healthy" {
        Start-Ui
        $w = In-Container $WhoCode
        $p = @()
        if ($w.uid -ne 10001 -or $w.gid -ne 10001) { $p += "runs as uid $($w.uid) gid $($w.gid), not 10001" }
        if ($w.app -or $w.root) { $p += "the program folder or the root filesystem is writable" }
        if ($w.rootro -ne $true) { $p += "the root filesystem is not mounted read-only" }
        if (-not $w.tmp -or -not $w.data) { $p += "/tmp or the data folder is not writable" }
        if ($w.nonewprivs -ne "1") { $p += "NoNewPrivs is '$($w.nonewprivs)'" }
        if ($w.capbnd -notmatch "^0+$") { $p += "capabilities are available: $($w.capbnd)" }
        if ($w.pid1 -notmatch "init|tini") { $p += "process 1 is '$($w.pid1)', not an init" }
        if ($w.ui_host -ne "0.0.0.0" -or $w.project -ne "/data" -or $w.stub -ne "0") { $p += "the settings file changed a pinned setting" }
        Judge $p "healthy; uid and gid 10001; / mounted read-only, and /app and / not writable; /tmp and /data writable; NoNewPrivs 1; no capabilities; an init is process 1"
    }

    Check "V11" "host-port" {
        Start-Ui
        $addresses = @(Get-NetTCPConnection -State Listen -LocalPort $Run.Port -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LocalAddress -Unique)
        if ($addresses.Count -eq 1 -and $addresses[0] -eq "127.0.0.1") { Pass "port $($Run.Port) listens on 127.0.0.1 only" }
        else { Fail ("port $($Run.Port) listens on: " + ($addresses -join ", ")) }
    }

    Check "V12" "host-check" {
        Start-Ui
        $r = Helper hosts $Run.Port $Name
        if (Prop $r "error") { Fail $r.error } elseif ($r.ok) { Pass "$($r.checked) requests: the two right hosts answered; every wrong Host got 400 (0.0.0.0, host.docker.internal, wrong port, no port, upper case, trailing dot, the container's name); every wrong Origin and a missing or wrong token got 403" }
        else { Fail (@($r.failures) -join "; ") }
    }

    $EnvNames = @("ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SHIPPING_ADDRESS", "BRAND_NAME", "SENDER_NAME",
                  "LANGFUSE_PUBLIC_KEY", "LANGFUSE_SECRET_KEY")
    $HashCode = @'
import hashlib, json, os
names = [NAMES]
print(json.dumps({n: (hashlib.sha256(os.environ[n].encode('utf-8')).hexdigest() if n in os.environ else None) for n in names}))
'@
    Check "V13" "env-parity" {
        Start-Ui
        $want = Helper envhash "$Proj\.env" @EnvNames
        $got = In-Container ($HashCode.Replace("NAMES", (($EnvNames | ForEach-Object { "'$_'" }) -join ", ")))
        $differ = @($EnvNames | Where-Object { (Prop $want $_) -ne (Prop $got $_) -or -not (Prop $got $_) })
        $page = Helper pages $Run.Port $State.Stem
        $notSet = @(@("ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SHIPPING_ADDRESS", "BRAND_NAME") | Where-Object { @(Prop $page "set") -notcontains $_ })
        $p = @()
        if ($differ.Count) { $p += "these settings differ between the container and python-dotenv's reading of the same file (names only): " + ($differ -join ", ") }
        if ($notSet.Count) { $p += "the Health page does not show these as Set: " + ($notSet -join ", ") }
        Judge $p ("{0} settings (double quotes, single quotes with `$ and #, \n, spaces) are identical by SHA-256 to python-dotenv's reading; Health shows them as Set" -f $EnvNames.Count)
        # Information only, no pass or fail: how compose and python-dotenv read four shapes the settings above do not use.
        # Each shape gets its own made-up settings file, so one shape that compose refuses cannot hide another.
        try {
            $shapes = [ordered]@{
                "a multi-line value in double quotes" = @("SHAPE_MULTILINE", "SHAPE_MULTILINE=`"line one`nline two`"")
                "a `${...} reference to another line" = @("SHAPE_REF", "SHAPE_REF_BASE=base-value`nSHAPE_REF=`${SHAPE_REF_BASE}-suffix")
                "an export prefix" = @("SHAPE_EXPORT", "export SHAPE_EXPORT=exported-value")
                "an unquoted `$" = @("SHAPE_DOLLAR", "SHAPE_DOLLAR=pa`$word")
            }
            $i = 0
            foreach ($shape in $shapes.Keys) {
                $i++
                $varName, $line = $shapes[$shape]
                $dir = "$Work\shape$i"
                New-Item -ItemType Directory -Force $dir | Out-Null
                Copy-Item -LiteralPath (Join-Path $Root "docker-compose.yml") -Destination "$dir\docker-compose.yml" -Force
                Write-Text "$dir\.env" ("TZ=$Zone`nLCOM_DATA=$($Data -replace '\\', '/')`nLCOM_NAME=lcom-verify-shape$i`nLCOM_IMAGE=$Image`n$line`n")
                $r = Split-Docker compose -p "lcom-verify-shape$i" --project-directory $dir -f "$dir\docker-compose.yml" --env-file "$dir\.env" config --format json
                $warn = (@($r.Err) | Where-Object { $_ }) -join " "
                if ($r.Code -ne 0) { $composeSays = "compose refused the file ($warn)" }
                else {
                    $v = Prop (Prop (($r.Out -join "`n" | ConvertFrom-Json).services.ui) "environment") $varName
                    $composeSays = $(if ($null -eq $v) { "(not set)" } else { ConvertTo-Json "$v" -Compress }) + $(if ($warn) { " with the warning: $warn" } else { "" })
                }
                $dot = Prop (Helper envvalues "$dir\.env" $varName) $varName
                Say ("  V13 observation, {0}: compose gives {1}; python-dotenv gives {2}" -f $shape, $composeSays, $(if ($null -eq $dot) { "(not set)" } else { ConvertTo-Json "$dot" -Compress }))
            }
        }
        catch { Say "  V13 observations could not run: $($_.Exception.Message)" }
    }

    # ================= stage 3: V14-V19 =================

    Check "V14" "data-on-host" {
        Start-Ui
        $r = Helper flow $Run.Port $Data $State.Stem
        if (Prop $r "error") { Fail $r.error }
        elseif ($r.ok) { Pass ("{0} steps through the UI, each file found in {1}: upload, approval request, a refused and an accepted record, export, download, order sheet, a product details upload kept byte for byte, and the Listings count made by the listing tool from the data folder (2 used, 2 ignored, 1 empty)" -f @($r.steps).Count, $Data) }
        else { Fail ((@($r.steps) | Where-Object { -not $_.ok } | ForEach-Object { $_.step + " (" + $_.detail + ")" }) -join "; ") }
    }

    $AppNameCode = @'
import json, os
print(json.dumps({'app_name': os.environ.get('UI_APP_NAME'), 'sender': os.environ.get('SENDER_NAME')}))
'@
    Check "V15" "survives" {
        Start-Ui
        $before = Get-Tree $Data
        Compose stop | Out-Null
        Compose start | Out-Null
        Wait-Healthy
        $afterRestart = Get-Tree $Data
        Compose down | Out-Null                    # removes the container and its network; there is no volume to remove
        Compose up -d --no-build | Out-Null
        Wait-Healthy
        $afterRecreate = Get-Tree $Data
        $page = Helper pages $Run.Port $State.Stem
        $p = @()
        if (-not (Same-List $before $afterRestart)) { $p += "files changed across stop and start" }
        if (-not (Same-List $before $afterRecreate)) { $p += "files changed across remove and recreate" }
        if (-not $before.Count) { $p += "the data folder is empty, so there was nothing to compare" }
        if (-not $page.ok) { $p += "after recreate the pages lost something: " + (@($page.problems) -join ", ") }
        # Does a plain `docker compose up -d` pick up a changed settings file? Settled by doing it (DOCKER.md says what this finds).
        $marker = "Verify Settings Change $Id"
        $first = (In-Container $AppNameCode).app_name
        Write-Text "$Proj\.env" ([System.IO.File]::ReadAllText("$Proj\.env") + "UI_APP_NAME=$marker`n")
        Compose up -d --no-build | Out-Null
        Wait-Healthy
        $plain = (In-Container $AppNameCode).app_name
        $plainPage = (Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$($Run.Port)/health").Content -match [regex]::Escape($marker)
        $forced = "not needed"
        if ($plain -ne $marker) {
            Compose up -d --no-build --force-recreate | Out-Null
            Wait-Healthy
            $forced = if ((In-Container $AppNameCode).app_name -eq $marker) { "picked it up" } else { "did NOT pick it up" }
        }
        if ($first -eq $marker) { $p += "the new setting was already there before the settings file changed" }
        if ($plain -ne $marker -and $forced -ne "picked it up") { $p += "neither plain 'up -d' nor 'up -d --force-recreate' picked up a changed settings file" }
        if ($plain -eq $marker -and -not $plainPage) { $p += "plain 'up -d' set the new value in the container, but the page does not show it" }
        $settled = if ($plain -eq $marker) { "plain 'docker compose up -d' DID pick up an ADDED setting (container recreated; page shows it)" } else { "plain 'docker compose up -d' did NOT pick up an added setting; 'up -d --force-recreate' $forced" }
        # Then CHANGE settings that already exist: SENDER_NAME (there from the start) and UI_APP_NAME (added above), plain `up -d` again.
        $marker2 = "Verify Settings Changed Again $Id"; $sender2 = "Verify Sender $Id"
        $before2 = In-Container $AppNameCode
        $text = [System.IO.File]::ReadAllText("$Proj\.env") -replace "(?m)^UI_APP_NAME=.*$", "UI_APP_NAME=$marker2"
        $text = $text -replace "(?m)^SENDER_NAME=.*$", "SENDER_NAME=`"$sender2`""
        Write-Text "$Proj\.env" $text
        Compose up -d --no-build | Out-Null
        Wait-Healthy
        $after2 = In-Container $AppNameCode
        $page2 = (Invoke-WebRequest -UseBasicParsing -Uri "http://127.0.0.1:$($Run.Port)/health").Content -match [regex]::Escape($marker2)
        $changedOk = $after2.app_name -eq $marker2 -and $after2.sender -eq $sender2
        if ($before2.sender -eq $sender2 -or $before2.app_name -eq $marker2) { $p += "a changed value was already there before the settings file changed" }
        if (-not $changedOk) { $p += "plain 'up -d' did not pick up CHANGED settings (container shows the old values)" }
        if ($changedOk -and -not $page2) { $p += "plain 'up -d' changed the value in the container, but the page does not show it" }
        $settled += $(if ($changedOk -and $page2) { "; and DID pick up CHANGED existing settings (SENDER_NAME and UI_APP_NAME new in the container; the page shows the new name)" } else { "; did NOT pick up changed existing settings" })
        Judge $p ("{0} files on the host have the same hashes after stop/start and after remove/recreate; the UI still shows the results file, the request, the approval and the order sheet. Settings file: {1}" -f $before.Count, $settled)
    }

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
        Start-Ui
        $r = Helper atomic $Name $Data
        $single = "$Work\single"
        Write-Text "$single\state.json" "{}"
        $s = In-Image $Image $SingleCode @("--tmpfs", "/probe:mode=1777", "-v", "$single\state.json:/probe/state.json")
        $control = if ($s.replaced) { "a single mounted file WAS replaced by a rename here" } else { "a rename onto a single mounted file is refused ($($s.error) $($s.errno)), which is why the data is one mounted folder" }
        if (Prop $r "error") { Fail $r.error }
        elseif ($r.ok) { Pass ("200 locked atomic writes while this computer read the file {0} times: never partial, nothing left behind; {1}" -f $r.reads, $control) }
        else { Fail ("writer exit {0} ({1}); partial reads {2}; last value {3}; left behind: {4}; {5}" -f $r.writer_exit, $r.writer_said, $r.partial, $r.last, (@($r.left_behind) -join ", "), $control) }
    }

    $ZoneCode = @'
import json, os, time
print(json.dumps({'tz': os.environ.get('TZ'), 'offset': time.strftime('%z')}))
'@
    Check "V17" "time-zone" {
        Start-Ui
        $z = In-Container $ZoneCode
        $page = Helper pages $Run.Port $State.Stem
        $p = @()
        if ($z.tz -ne $Zone -or $z.offset -ne $ZoneOffset) { $p += "inside the container TZ is '$($z.tz)' and the offset is $($z.offset); wanted $Zone and $ZoneOffset" }
        if (-not "$(Prop $page 'zone')".StartsWith($ZoneLabel)) { $p += "the Health page says '$(Prop $page 'zone')', not $ZoneLabel" }
        Judge $p "with TZ=$Zone the container's offset is $ZoneOffset and the Health page says $(Prop $page 'zone')"
    }

    # Runs inside the container as the container user. It takes the UI process's own environment from /proc (the UI is
    # started by the entrypoint as a new process carrying the merged NO_PROXY), switches to exactly that environment, and
    # makes its requests the way the libraries do by default, with the settings file's dead proxy still set. It returns
    # setting NAMES, the NO_PROXY lists (made-up test values) and outcomes; never a proxy value.
    $TraceCode = @'
import json, os, socket, urllib.request
def ui_pid():
    for pid in os.listdir('/proc'):
        if pid.isdigit():
            try:
                parts = open('/proc/%s/cmdline' % pid, 'rb').read().split(b'\0')
            except OSError:
                continue
            if any(p.endswith(b'run_ui.py') for p in parts):
                return pid
pid = ui_pid()
env = {}
if pid:
    raw = open('/proc/%s/environ' % pid, 'rb').read().decode('utf-8', 'replace')
    env = dict(item.split('=', 1) for item in raw.split('\0') if '=' in item)
out = {'pid': pid, 'obs': env.get('OBS_ENABLED'), 'base': env.get('LANGFUSE_BASE_URL'), 'host': env.get('LANGFUSE_HOST'),
       'no_proxy': env.get('NO_PROXY'), 'no_proxy_lower': env.get('no_proxy'),
       'proxy_names': sorted(n for n in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy') if env.get(n))}
try:
    out['resolves'] = bool(socket.gethostbyname('host.docker.internal'))
except OSError:
    out['resolves'] = False
os.environ.clear()
os.environ.update(env)
def attempt(fn):
    try:
        return 'direct %s' % fn()
    except Exception as e:
        reason = getattr(e, 'reason', None)
        refused = isinstance(reason, ConnectionRefusedError) or isinstance(e, ConnectionRefusedError) or 'refused' in str(e).lower()
        return 'refused by the dead proxy' if refused else 'failed (%s)' % type(e).__name__
def by_urllib(url):
    with urllib.request.build_opener().open(url, timeout=4) as r:
        return r.status
import httpx, requests
local = 'http://127.0.0.1:' + env.get('UI_PORT', '8000') + '/static/app.css'
out['local'] = {'urllib': attempt(lambda: by_urllib(local)), 'httpx': attempt(lambda: httpx.get(local, timeout=4).status_code),
                'requests': attempt(lambda: requests.get(local, timeout=4).status_code)}
out['langfuse_bypass'] = bool(urllib.request.proxy_bypass_environment('host.docker.internal'))
out['langfuse_get'] = attempt(lambda: by_urllib((env.get('LANGFUSE_BASE_URL') or '').rstrip('/') + '/api/public/health')) if DO_GET else 'not tried'
out['control_bypass'] = bool(urllib.request.proxy_bypass_environment('example.invalid'))
out['control_get'] = attempt(lambda: by_urllib('http://example.invalid/'))
print(json.dumps(out))
'@
    Check "V18" "tracing-and-proxy" {
        Start-Ui
        $t = In-Container ($TraceCode.Replace("DO_GET", $(if ($NoLangfuseGet) { "False" } else { "True" })))
        # Information only, never part of the verdict: if the GET from inside failed, the same read-only GET at once from
        # Windows (no proxy), with how long it took and the memory available on this computer at that moment.
        $fromWindows = ""
        if (-not $NoLangfuseGet -and $t.langfuse_get -ne "direct 200") {
            $clock = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $req = [System.Net.WebRequest]::Create("http://127.0.0.1:3000/api/public/health"); $req.Proxy = $null; $req.Timeout = 30000
                $resp = $req.GetResponse(); $got = "HTTP $([int]$resp.StatusCode)"; $resp.Close()
            } catch {
                $e = if ($_.Exception.InnerException) { $_.Exception.InnerException } else { $_.Exception }
                $got = if ($e -is [System.Net.WebException] -and $e.Response) { "HTTP $([int]$e.Response.StatusCode)" }
                       elseif ($e -is [System.Net.WebException]) { "failed ($($e.Status))" } else { "failed ($($e.GetType().Name))" }
            }
            $clock.Stop()
            $free = try { "$([int](Get-Counter '\Memory\Available MBytes').CounterSamples[0].CookedValue) MB" } catch { "unknown" }
            $fromWindows = " [information only: at once from Windows the same GET of http://127.0.0.1:3000/api/public/health gave " + $got +
                           " in " + $clock.Elapsed.TotalSeconds.ToString("0.00") + " s; memory available on this computer then: " + $free + "]"
        }
        $merged = "internal.example,localhost,127.0.0.1,host.docker.internal"
        $p = @()
        if ($t.obs -ne "0") { $p += "OBS_ENABLED is '$($t.obs)' although the settings file says 1 and LCOM_TRACING is not set" }
        if ($t.base -ne "http://host.docker.internal:3000" -or $t.host -ne "http://host.docker.internal:3000") { $p += "the Langfuse address is not http://host.docker.internal:3000 in both variables (the settings file says localhost)" }
        if (-not $t.resolves) { $p += "host.docker.internal does not resolve inside the container" }
        # step 1: the UI process's own environment
        if (-not $t.pid) { $p += "the UI process (run_ui.py) was not found" }
        if ($t.no_proxy -ne $merged -or $t.no_proxy_lower -ne $merged) { $p += "the UI process's NO_PROXY and no_proxy are not exactly the person's entry followed by the three added names: [$($t.no_proxy)] [$($t.no_proxy_lower)]" }
        if ((@($t.proxy_names) -join ",") -ne "HTTPS_PROXY,HTTP_PROXY") { $p += "the dead proxy is not set in the UI process (names: $(@($t.proxy_names) -join ', ')), so the proof would prove nothing" }
        # step 2: this container's own address, by urllib, httpx and requests, with that environment
        foreach ($lib in @("urllib", "httpx", "requests")) { if ((Prop $t.local $lib) -ne "direct 200") { $p += "${lib} to 127.0.0.1: $(Prop $t.local $lib)" } }
        # step 3: Langfuse on this computer, bypassing the proxy (the one read-only GET)
        if (-not $t.langfuse_bypass) { $p += "host.docker.internal would go through the proxy" }
        if (-not $NoLangfuseGet -and $t.langfuse_get -ne "direct 200") { $p += "the GET of Langfuse's health address: $($t.langfuse_get)$fromWindows" }
        # step 4: the control: everything else still goes to the proxy
        if ($t.control_bypass -or $t.control_get -ne "refused by the dead proxy") { $p += "a request to example.invalid did not go to the proxy: $($t.control_get)" }
        # the fake proxy credentials appear nowhere: container log (the entrypoint's lines included), stored job logs, pages
        # an address with a user name and password, planted in a real job's output (the email agent's drafts print seller
        # names and listing addresses), and the container's own proxy setting next to it: masked in the stored log and on pages
        $plantUser = "plantuser$Id"; $plantPass = "FAKE-plant-pass-$Id"
        $plant = Helper plantjob $Run.Port $Data $plantUser $plantPass "http://${ProxyUser}:${ProxyPass}@127.0.0.1:9"
        if (Prop $plant "error") { $p += "the planted address: $($plant.error)" }
        elseif (-not $plant.ok) { $p += "a planted address was not masked: job $($plant.job); masked addresses in the log $($plant.masked_addresses_in_log); found in the stored log [$(@($plant.found_in_stored_log) -join ',')], on pages [$(@($plant.found_on_pages) -join ',')]" }
        $needles = @($ProxyUser, $ProxyPass, $plantUser, $plantPass)
        $log = (@(cmd /c "docker logs $Name 2>&1") -join "`n")
        if (@($needles | Where-Object { $log.Contains($_) }).Count) { $p += "the container log holds the fake proxy credentials" }
        $want = "proxy settings present: HTTPS_PROXY, HTTP_PROXY; reached directly: 1 name(s) from your settings and 3 added (localhost, 127.0.0.1, host.docker.internal)"
        if (-not $log.Contains($want)) { $p += "the entrypoint's proxy line is missing or different" }
        $jobFiles = @(Get-ChildItem -LiteralPath "$Data\ui_data\jobs" -File -ErrorAction SilentlyContinue)
        $inJobs = @($jobFiles | Where-Object { $text = [System.IO.File]::ReadAllText($_.FullName); @($needles | Where-Object { $text.Contains($_) }).Count } | ForEach-Object { $_.Name })
        if ($inJobs.Count) { $p += "stored job files hold the fake proxy credentials: " + ($inJobs -join ", ") }
        $scan = Helper pagescan $Run.Port $State.Stem @needles
        if (-not $scan.ok) { $p += "pages hold the fake proxy credentials: " + (@($scan.hits) -join ", ") }
        Judge $p ("tracing off although the settings file switches it on; both Langfuse addresses are http://host.docker.internal:3000 and the name resolves. " +
                  "With the dead proxy set: the UI process's NO_PROXY and no_proxy are '$merged'; urllib, httpx and requests reach 127.0.0.1 directly; " +
                  "Langfuse's health address bypasses the proxy (GET: $($t.langfuse_get)); example.invalid still goes to the proxy and is refused there. " +
                  "A drafts job whose output carried a planted user:password address and the proxy setting showed both masked ($(Prop $plant 'masked_addresses_in_log') masked addresses in its log). " +
                  "The fake proxy and planted credentials are in none of: the container log with the entrypoint's lines, $($jobFiles.Count) stored job files, $($scan.pages) pages. No trace was sent.")
    }

    Check "V19" "suites-in-container" {
        Build-Images
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
        # Skips: in the image exactly one is expected (the compose-file test: the image has no compose file); anything else fails.
        $expectedSkips = @("test_ui.py skip compose_never_fails_a_command_because_a_setting_is_missing: docker-compose.yml is not in this image; the check is about the source folder")
        if (-not (Same-List ($skipped | Sort-Object) ($expectedSkips | Sort-Object))) { $p += "the skipped tests in the image are not exactly the one expected: [" + ($skipped -join "; ") + "]" }
        # Windows side: on the source folder the same suite must skip nothing (the compose-file test runs there).
        Say "  running test_ui.py on this computer (Windows), where nothing may be skipped ..."
        $win = @(& python -B (Join-Path $Real "test_ui.py") 2>&1 | ForEach-Object { "$_" })
        $winCode = $LASTEXITCODE
        $winSummary = @($win | Where-Object { $_ -cmatch "^\d+/\d+ passed" })[-1]
        if ($winCode -ne 0 -or $winSummary -notmatch "^(\d+)/\1 passed, 0 skipped$") { $p += "on Windows test_ui.py did not pass with 0 skipped: exit $winCode, '$winSummary'" }
        $note = "skipped tests in the image: " + $(if ($skipped.Count) { $skipped -join "; " } else { "none" }) + ". On Windows: test_ui.py '$winSummary'. Parts of protected suites that need real output files and so do not run in an image: " + ($DataDependent -join "; ")
        if ($p.Count) { Fail (($p -join "; ") + ". " + $note) } else { Pass ("all nine suites exit 0 as uid 10001 with no network and a read-only system: " + ($told -join "; ") + ". " + $note) }
    }

    # ================= stage 4: V20-V23 (V24 follows) =================

    Check "V20" "refuses-source-folder" {
        Build-Images
        $src = "$Work\srcdata"
        New-Item -ItemType Directory -Force "$src\.git" | Out-Null
        Write-Text "$src\email_agent.py" "MY OWN WORKING COPY - DO NOT REPLACE"
        $hash = (Get-FileHash -LiteralPath "$src\email_agent.py").Hash
        # Started in the background and given 60 s: a container that does not refuse would otherwise run, and this check wait, for ever.
        & docker.exe run -d --name lcom-verify-src --label $Label --network none --read-only --tmpfs /tmp -e "TZ=$Zone" -v "${src}:/data" $Image | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "docker run could not create lcom-verify-src (exit code $LASTEXITCODE)" }
        $end = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $end -and @(& docker.exe inspect --format "{{.State.Status}}" lcom-verify-src)[0] -ne "exited") { Start-Sleep -Seconds 1 }
        $srcState = @(& docker.exe inspect --format "{{.State.Status}} {{.State.ExitCode}}" lcom-verify-src)[0]
        $stillRunning = $srcState -notlike "exited *"
        $code = if ($stillRunning) { "none (still running)" } else { [int]($srcState -split " ")[1] }
        $log = Logs lcom-verify-src
        Dk rm -f -v lcom-verify-src | Out-Null              # -v: also this container's own unnamed volumes (mutation 18b)
        $p = @()
        if ($stillRunning) { $p += "the container started and was still running after 60 s" } elseif ($code -eq 0) { $p += "the container started" }
        if ($log -notmatch "looks like the source folder") { $p += "the message does not say why" }
        if ((Get-FileHash -LiteralPath "$src\email_agent.py").Hash -ne $hash -or @(Get-ChildItem $src -Recurse -File).Count -ne 1) { $p += "the folder was changed" }
        Judge $p "a data folder holding .git: exit code $code, the message names the reason, the planted email_agent.py is byte-identical and nothing was added"
    }

    Check "V21" "refuses-no-mount-no-tz" {
        Build-Images
        $p = @()
        # (a) no data folder at all. Started in the background and given 60 s, as in V20.
        & docker.exe run -d --name lcom-verify-nomount --label $Label --network none --read-only --tmpfs /tmp -e "TZ=$Zone" $Image | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "docker run could not create lcom-verify-nomount (exit code $LASTEXITCODE)" }
        $end = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $end -and @(& docker.exe inspect --format "{{.State.Status}}" lcom-verify-nomount)[0] -ne "exited") { Start-Sleep -Seconds 1 }
        $nmState = @(& docker.exe inspect --format "{{.State.Status}} {{.State.ExitCode}}" lcom-verify-nomount)[0]
        $stillRunning = $nmState -notlike "exited *"
        $code = if ($stillRunning) { "none (still running after 60 s)" } else { [int]($nmState -split " ")[1] }
        $log = Logs lcom-verify-nomount
        Dk rm -f -v lcom-verify-nomount | Out-Null          # -v: also this container's own unnamed volumes (mutation 18b)
        if ($stillRunning -or $code -eq 0 -or $log -notmatch "no data folder is mounted") { $p += "without a data folder: exit code $code, message: $log" }
        # (b) the compose file with a settings file that has no TZ: every compose command works, the container refuses to start the UI
        $dir = "$Work\compose-notz"; $data = "$Work\data-notz"
        New-Item -ItemType Directory -Force $dir, $data | Out-Null
        New-FakeEnv $dir (Get-FreePort) "lcom-verify-notz" $data -NoZone
        $c = "lcom-verify-notz"
        $cfg = (Invoke-Compose $dir "lcom-verify-notz" config --format json) -join "`n" | ConvertFrom-Json
        if ("$(Prop (Prop $cfg.services.ui 'environment') 'TZ')" -ne "") { $p += "compose config without TZ did not pass TZ through empty" }
        Invoke-Compose $dir "lcom-verify-notz" up -d --no-build | Out-Null
        $end = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $end -and @(& docker.exe inspect --format "{{.State.Status}}" $c)[0] -ne "exited") { Start-Sleep -Seconds 1 }
        $exit = @(& docker.exe inspect --format "{{.State.Status}} {{.State.ExitCode}}" $c)[0]
        $log = Logs $c
        if ($exit -notmatch "^exited [1-9]") { $p += "without TZ the container is '$exit', not exited with an error" }
        if ($log -notmatch "NOT STARTED: TZ is not set") { $p += "without TZ the log does not say why: $log" }
        if (@(Get-ChildItem -LiteralPath $data -Force).Count) { $p += "without TZ the container wrote into the data folder" }
        foreach ($cmd in @(@("ps", "-a"), @("logs", "--tail", "5", "ui"), @("stop"), @("down"))) {
            try { Invoke-Compose $dir "lcom-verify-notz" @cmd | Out-Null } catch { $p += "docker compose $($cmd -join ' ') failed without TZ" }
        }
        Judge $p "no data folder: refused with a message. No TZ in the settings file: compose config passes TZ through empty, up, ps, logs, stop and down all work, and the container exits at once with 'NOT STARTED: TZ is not set', writing nothing"
    }

    Check "V22" "one-name" {
        Start-Ui
        & docker.exe run -d --name $Name --label $Label --network none -e "TZ=$Zone" $Image | Out-Null
        $code = $LASTEXITCODE
        $count = @(& docker.exe ps -a --filter "name=$Name" --format "{{.Names}}" | Where-Object { $_ -eq $Name }).Count
        if ($code -ne 0 -and $count -eq 1) { Pass "Docker refused a second container named $Name (exit code $code); still exactly one" }
        else { Fail "exit code $code; containers with that name: $count" }
    }

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
        Start-Ui
        $target = "$Data\decision_state.json"
        $saved = [System.IO.File]::ReadAllBytes($target)
        $hashBefore = (Get-FileHash -LiteralPath $target).Hash
        $hold = [System.IO.File]::Open($target, "Open", "ReadWrite", "None")
        try { $r = In-Container $LockCode } finally { $hold.Close() }
        $same = ((Get-FileHash -LiteralPath $target).Hash -eq $hashBefore)
        $left = @(Get-ChildItem $Data -File | Where-Object { $_.Name -like "decision_state.json.*" } | ForEach-Object { $_.Name })
        if (-not $same) { [System.IO.File]::WriteAllBytes($target, $saved) }
        Info ("while this computer held decision_state.json open: the container's write " + $(if ($r.written) { "went through" } else { "was refused ($($r.error) $($r.errno))" }) +
              "; the file is " + $(if ($same) { "unchanged" } else { "CHANGED (restored)" }) + "; left behind: " + $(if ($left.Count) { $left -join ", " } else { "nothing" }))
    }
}
finally {
    if ($Mode -eq "stage") {
        if ($script:Stopped) { $Run.Stopped = $script:Stopped } else { $Run.Done = @(@($Run.Done) + [int]$Stage) }
        if ($Stage -eq "4" -and -not $script:Stopped) { Finish-Run } else { Save-Run }
    }
    elseif ($Mode -eq "cleanup" -or $Mode -eq "only") { Finish-Run }
}

Say ""
$Results | Format-Table Id, Check, Result -AutoSize | Out-String | Write-Host
$failed = @($Results | Where-Object { $_.Result -eq "FAIL" })
Say ("{0} checks: {1} passed, {2} failed, {3} for information" -f $Results.Count, @($Results | Where-Object { $_.Result -eq "PASS" }).Count, $failed.Count, @($Results | Where-Object { $_.Result -eq "INFO" }).Count)
if ($Mode -eq "stage") {
    if ($script:Stopped) { Say "Stage $Stage stopped at $($script:Stopped). Nothing was cleaned up, so it can be looked at. To start over: -Cleanup, then -Stage 1." }
    elseif ($Stage -ne "4") { Say "Stage $Stage passed. The run's images, container and temp folder are kept for the next stage: -Stage $([int]$Stage + 1)" }
    else { Say "Stage 4 passed and the run is cleaned up." }
}
if ($failed.Count) { exit 1 }
