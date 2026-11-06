# First live run: 1 to 2 products through the UI

The first run with real keys: real searches, a real email to yourself, then to real sellers. Every step below is a button
or command the tool already has. It assumes the Docker setup in `DOCKER.md` (UI at `http://127.0.0.1:8010/`, data folder
`data\`). Commands are PowerShell, run in the folder that holds `docker-compose.yml`.

Stop at the first check that fails. Nothing later depends on finishing in one sitting, except approvals: each one lasts
7 days, so the order sheet and listings must follow within 7 days of recording the manager's answer.

## Expected cost

| Step | Paid service | For 2 products |
|---|---|---|
| 1. Find suppliers | Claude Haiku, plus Nimble (up to 13 searches + page reads per product) | about $0.16 by the page's estimate ($0.08 and 56 s per product, from one 15-product batch); the 15-product run of 2026-09-29 measured about $0.10 per product. Nimble's charge is not in the tool's figure |
| 2. Email sellers: preview, send, mark | your mail server only | $0 |
| 2. Email sellers: record a reply | one Claude Haiku call per reply | not measured; one short call |
| 3. Get approval, 4. Order sheet | none | $0 |
| 5. Listings | Claude Haiku | the job log starts with `WARNING: worst-case cost ...` if the run could cost over $0.50 |

Expect well under $1 in model cost. Afterwards, compare with the Anthropic console and Nimble's usage page.

## 0. Before you start

- [ ] Docker Desktop is running; `.env` is filled in as in `DOCKER.md` → Settings (`TZ`, the API keys, the mail
      settings, `SHIPPING_ADDRESS`, `BRAND_NAME`, `SENDER_NAME`). A value with `$` is in single quotes.
- [ ] You have an email address of your own for the test send, and you know which 1 or 2 SKUs you will search.
- [ ] No native UI (`python run_ui.py`) is running on the same data folder.

## 1. Back up the data folder

From `DOCKER.md` → Backing up and restoring. Stop first, so no job writes while the copy is made.

```powershell
docker compose stop
Compress-Archive -Path .\data\* -DestinationPath ("lcom-data-{0:yyyyMMdd-HHmm}.zip" -f (Get-Date))
docker compose start
```

- [ ] The zip exists and opens. (A brand-new, empty data folder has nothing to back up: skip this once.)
- [ ] The zip stays out of git and out of anything you share. It holds no keys: `.env` is not in the data folder.

## 2. Start and check Setup

```powershell
docker compose up -d
docker compose ps                        # wait until STATUS says "healthy"
Start-Process "http://127.0.0.1:8010/"
```

- [ ] If the container stopped at once: `docker compose logs ui` shows a line starting `NOT STARTED:`; fix it with the
      table in `DOCKER.md` → If the container stops at once.
- [ ] **Setup** (top bar): "Finding suppliers", "Sending emails to sellers", "Reading seller replies" and "Writing
      listings" all say **Set up**. If not, its Technical details name the missing setting.
- [ ] The time zone shown is yours.
- [ ] Tracing says Off (it is off in the container unless `LCOM_TRACING=1`).
- [ ] Any reference file listed as different from this version's is one you edited on purpose.

## 3. Find suppliers (step 1)

1. Under **Upload a spreadsheet of your products**, choose your `.xlsx` and press **Upload and preview**.
2. Check **What was read from ...**: the SKUs and the descriptions the search will use are right.
3. Under **Which products**, pick **Choose products** and tick your 1 or 2 SKUs (**Find a product** filters the list).
   Leave **Also include the built-in products** unticked.
4. Press **Next: check the cost**. Under **Check the cost and confirm**, open **Products in this search** and check it
   lists only your SKUs, and that the estimate is for 1 or 2 products.
5. Type the number of products and press **Start search**.

While it runs and afterwards:

- [ ] The run ends without being shown as partial. (**Cancel this search** keeps the products finished so far.)
- [ ] **Activity** → the search's log has no `Error researching this product` line.
- [ ] The log's last cost line (`Estimated total cost: ~$...`) is near the estimate.
- [ ] **Results**: for each recommended pick, open the listing and check it is the same kind of product as the SKU
      (the 2026-09-29 run recommended a USB printer cable for an adapter), the price is per unit and not per pack or
      per MOQ, and the seller is named.
- [ ] **Files**: the results workbook and the report download and open.

## 4. Email sellers (step 2): yourself first

1. Press **Preview the emails** and read every draft: the product and SKU, the quantities asked about, your name and
   the company.
2. Under **Or send yourself a test copy of every email first**, type your own address and press **Review the test send**.
3. The page says **This is a TEST**: every email goes only to your address, with `[TEST]` in the subject; no seller is
   contacted and nothing is marked as sent. Type the number of emails and press **Send the TEST emails**.

In your inbox:

- [ ] One email per draft arrived, each subject starting with `[TEST]`, from the sender name and address you set.
- [ ] Each body says where it would have gone, and the text and the signature are right.
- [ ] **Activity** → the send log shows addresses as `[EMAIL]`, never in full.

**HARD STOP.** Show Neeraj the test email (or a demo) and get his go-ahead before contacting any real seller. He asked to
be included in supplier emails, and the tool cannot CC anyone: ask him how he wants to be kept in the loop.

- [ ] Neeraj has seen the test email (or a demo) and said go ahead.
- [ ] You know how he wants to be kept in the loop on supplier emails.

Only when all five are right, contact the real sellers:

4. Press **Review N emails to sellers**. Under **Check who will be emailed**, read every recipient (addresses are shown
   in full here on purpose). Type the number of emails and press **Send to sellers**.
5. A seller with only a website: open **Copy an inquiry to paste into a seller's website form**, paste it into their
   form yourself, then under **I contacted a seller through the contact form on their website** press
   **Mark as contacted**.

When a seller replies (by email, outside the tool):

6. Under **A seller replied**, pick the product, paste the reply and press **Record the reply** (one model call).
7. Under **Quote checks**, read the counts (blocking, to confirm, notes) against the reply itself. If there are
   questions, copy them from the read-only box, send them to the seller yourself, then press **Mark the questions as
   sent**.

## 5. Get approval (step 3)

1. Under **1. Create the request for your manager**, add a budget if you want one compared, and press
   **Create the request**. If it warns that no seller has been contacted yet, go back to step 4 first.
2. The request text is shown under **Your requests for this search**. The tool does not send it: send it to your
   manager yourself (or **Download this request**).
3. When the answer comes, under **2. Record your manager's answer**: pick the request, paste the answer exactly as sent,
   type the request number again, say how it reached you and who approved it, and press **Record the answer**.

- [ ] **Approvals so far** shows each approved SKU with the quantity the manager gave, who approved it and how, and an
      expiry 7 days after it was recorded.
4. Under **3. Save the approved list**, press **Save the approved list**.

## 6. Order sheet (step 4)

1. Press **Create the order sheet**. Nothing is typed; it only creates a file.

- [ ] The sheet marked **Newest, created after your latest approved list** downloads and opens.
- [ ] It lists only the approved SKUs, with the approved quantities and prices, and any quote-check flags as notes.
- [ ] Ordering and payment stay manual: nothing has been ordered.

## 7. Listings (step 5)

1. Under **1. Get the product details sheet**, press **Create the product details sheet** and download it.
2. Fill in the values and statuses yourself, from the sample you hold rather than the seller's listing. Under
   **2. Fill it in, then upload it**, upload it with **Upload the product details**.
3. Under **3. Check what was read**, press **Show** and check the counts of details used and ignored.
4. Under **4. Write the listings**, press **Next: check and confirm**, type the number of products and press
   **Write the listings**.

- [ ] The job log has no `WARNING: worst-case cost` line, or you expected it.
- [ ] Each product is **READY FOR REVIEW**, or **BLOCKED** with reasons you can read in the workbook's
      **Review notes** sheet.
- [ ] Nothing was published: the tool only writes a workbook to `Walmart Listings\`.

## 8. Afterwards: no secret in the logs

Job logs mask keys, passwords, the shipping address and every email address. This checks it on your real logs and the
container's own messages. It prints each setting's **name** and the files its value was found in; it never prints a value.

```powershell
docker compose logs ui 2>&1 | Out-File -Encoding utf8 "$env:TEMP\lcom-ui-log.txt"
$files = @(Get-ChildItem -LiteralPath .\data\ui_data\jobs -Recurse -File -ErrorAction SilentlyContinue) + @(Get-Item -LiteralPath "$env:TEMP\lcom-ui-log.txt")
$names = "ANTHROPIC_API_KEY", "NIMBLE_API_KEY", "SMTP_PASSWORD", "LANGFUSE_SECRET_KEY", "HTTP_PROXY", "HTTPS_PROXY", "SHIPPING_ADDRESS"
foreach ($line in Get-Content -LiteralPath .env) {
    if ($line -notmatch '^\s*(?:export\s+)?([A-Za-z_]+)\s*=\s*(.*)$' -or $names -notcontains $Matches[1]) { continue }
    $name = $Matches[1]; $value = $Matches[2].Trim()
    if ($value.Length -ge 2 -and $value[0] -eq $value[-1] -and "'`"".Contains($value[0])) { $value = $value.Substring(1, $value.Length - 2) }
    foreach ($part in @($value -split '\\n' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 6 })) {
        $hits = @($files | Select-String -SimpleMatch -Pattern $part -List | ForEach-Object { $_.Path })
        "{0}: {1}" -f $name, $(if ($hits) { "FOUND in " + ($hits -join ", ") } else { "not found" })
    }
}
Remove-Item -LiteralPath "$env:TEMP\lcom-ui-log.txt"
```

- [ ] Every line says `not found`. (`SHIPPING_ADDRESS` gets one line per address line of 6 characters or more.)
- [ ] For any `FOUND`: open that file and look. A short, common address line can match ordinary text; a key or
      password never should. Do not share the file.
- [ ] Do not run `docker compose config` or `docker inspect` on a shared screen: they print every setting.

## 9. Close out

- [ ] **Activity**: every job has ended; none is still running.
- [ ] Back up again (step 1), now with the first live data.
- [ ] Note the real costs (Anthropic console, Nimble usage) next to the estimate above.
- [ ] If a step fails with "the state file is busy", see `DOCKER.md` → After a hard stop: leftover lock files.
