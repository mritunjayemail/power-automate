# Weekly "Apply credit card" reminder — Power Automate (Microsoft 365 / Outlook)

Sends **one separate email per person** listed in an Excel sheet on OneDrive,
**every Monday at 10:00 AM**, from the Microsoft 365 cloud — so it is delivered even if
your laptop is switched off and Outlook is closed.

| Item | Value |
|---|---|
| Schedule | Weekly, Monday, 10:00 AM (your time zone) |
| Subject | `Apply credit card - <recipient name>` |
| Body | `Dear <recipient name>, Please note you didn't applied credit card yet, Please apply. Thanks - Rakhi` |
| Recipients | The rows of the `Recipients` table in `recipients.xlsx` on OneDrive |
| Runs on | Microsoft's servers (Power Automate cloud flow) — **not** your laptop |

---

## 0. Why a Power Automate *cloud flow* (and not Outlook itself)

| Approach | Works with laptop off? | Verdict |
|---|---|---|
| Outlook **Delay Delivery** / Drafts | No — classic Outlook must be running to release the mail | ✗ |
| Outlook **rule** or **VBA macro** | No — rules/macros run inside the Outlook client | ✗ |
| Windows **Task Scheduler** + script | No — machine must be on | ✗ |
| **Power Automate cloud flow** (this guide) | **Yes** — hosted by Microsoft | ✓ Recommended |

A *desktop flow* (Power Automate for desktop) is **not** the same thing — that one needs your
machine on. Build a **cloud flow**, as described below.

---

## 1. Prerequisites

1. A **Microsoft 365 work/school account** (e.g. `rakhi@yourcompany.com`) — the Office 365
   Outlook and Excel Online (Business) connectors do not work with a personal `@outlook.com`
   account.
2. Access to <https://make.powerautomate.com> with that account.
   The seeded Power Automate rights in most M365 plans are enough; no premium licence is
   needed for the Recurrence trigger, Excel Online (Business), or Office 365 Outlook.
3. **OneDrive for Business** (comes with the same account) to hold the sheet.
4. Permission to send mail from the mailbox you sign in with. Emails will show **From: you**.
   To send as a shared mailbox, see section 8.

---

## 2. Put the recipient sheet on OneDrive

`recipients.xlsx` in this folder is a ready-made starter file. It already contains a **named table**
called `Recipients`:

| Name | Email | Active |
|---|---|---|
| Mritunjay Kumar | mritunjayemail@gmail.com | Yes |

1. Open `recipients.xlsx`, add one row per person, and save.
2. Upload it to **OneDrive for Business**, somewhere stable — e.g. a folder called `Automation`,
   giving `/Automation/recipients.xlsx`.

### If you build the sheet yourself instead

1. **New** → **Excel workbook** in OneDrive.
2. `Name` in **A1**, `Email` in **B1**, `Active` in **C1**. Fill the rows underneath.
3. Select the filled range **including the headers** → **Insert** → **Table** →
   tick **My table has headers** → **OK**.
4. With the table selected: **Table Design** → **Table Name** → type `Recipients`.

Steps 3 and 4 are not optional. The Excel connector reads a **named table**, never a plain sheet —
a missing table is the single most common reason the flow finds nothing to send.

### Rules for the sheet

- Keep the file as `.xlsx`. A `.csv` does not work with this connector.
- Header names are used verbatim in the flow and are **case-sensitive**.
- No duplicate addresses — a repeat means that person gets the reminder twice.
- Watch for stray spaces in addresses; Excel will not warn you.

---

## 3. Create the flow (step by step)

### Step 3.1 — New scheduled cloud flow

1. Go to <https://make.powerautomate.com> and sign in.
2. Confirm the **environment** picker (top right) shows the right tenant/environment.
3. Left menu → **Create** → **Scheduled cloud flow**.
4. Fill the dialog:
   - **Flow name**: `Weekly Credit Card Reminder`
   - **Starting**: today's date, **10:00 AM**
   - **Repeat every**: `1` **Week**
   - **On these days**: select **M** (Monday) only
5. Click **Create**.

### Step 3.2 — Fix the time zone (important)

1. Open the **Recurrence** trigger card → **Show advanced options**.
2. **Time zone** → pick yours (e.g. `(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi`).
3. Confirm **At these hours** = `10`, **At these minutes** = `0`, **On these days** = `Monday`.

If you leave Time zone blank the flow runs in **UTC** and will arrive at the wrong local hour.
Setting the time zone also keeps 10:00 AM correct across daylight-saving changes.

### Step 3.3 — Read the sheet

1. **+ New step** → search **Excel Online (Business)** → **List rows present in a table**.
2. Sign in when prompted to create the connection (this stores a consented token in the cloud).
3. Fill the dropdowns — do not type these by hand, pick them so the ids resolve:
   - **Location**: `OneDrive for Business`
   - **Document Library**: `OneDrive`
   - **File**: browse to `/Automation/recipients.xlsx`
   - **Table**: `Recipients`
4. **Show advanced options** → **Filter Query** (only if you use the `Active` column):

```
Active eq 'Yes'
```

5. ⋯ → **Settings** → **Pagination** → **On**, **Threshold** = `5000`.
   Without this the connector stops at the first 256 rows.

### Step 3.4 — Loop over the people

1. **+ New step** → search **Apply to each** (Control).
2. Rename it (⋯ → **Rename**) to `Send one email per person`.
3. **Select an output from previous steps** → dynamic content → **value**
   (from *List rows present in a table*).

### Step 3.5 — Send the email (inside the loop)

1. Inside the *Apply to each*, click **Add an action** → search
   **Office 365 Outlook — Send an email (V2)**.
2. Sign in if prompted — this connection is what lets the flow send mail while your laptop is off.
3. Fill the fields — use the **Expression** tab for the expressions below, not plain text:

   **To**
   ```
   items('Send_one_email_per_person')?['Email']
   ```

   **Subject**
   ```
   concat('Apply credit card - ', items('Send_one_email_per_person')?['Name'])
   ```

   **Body** — click the `</>` (code view) button on the body editor and paste:
   ```html
   <p>Dear @{items('Send_one_email_per_person')?['Name']},</p>
   <p>Please note you didn't applied credit card yet, Please apply.</p>
   <p>Thanks<br>Rakhi</p>
   ```

> `Name` and `Email` inside `items(...)` are your **Excel column headers** and are case-sensitive.
> If you renamed the loop, the name inside `items('...')` must match it with spaces replaced by
> underscores. Easiest alternative: use **Dynamic content → Name / Email** from the Excel step and
> the designer inserts the correct reference for you.

### Step 3.6 — One email per person, not one group mail

Each loop iteration creates its own message with a single address in **To**, so nobody sees the
other recipients. Do **not** put multiple addresses in one **To** field.

### Step 3.7 — Send them one at a time (optional but recommended)

1. On the **Apply to each** card: ⋯ → **Settings**.
2. **Concurrency Control** → **On** → **Degree of Parallelism** = `1`.

This keeps the sending order predictable and stays well under Exchange Online's throttling limits
(~30 messages/minute, 10,000 recipients/day).

### Step 3.8 — Save

Click **Save**. The flow is live and owned by your account, running in Microsoft's cloud.

The finished shape is recorded in [`flow/definition.json`](flow/definition.json) — compare it with
**⋯ → Peek code** in the designer if something does not match.

---

## 4. Test before Monday

The starter sheet holds one recipient — you — so this rehearsal reaches nobody else.

1. In the flow designer click **Test** → **Manually** → **Test** → **Run flow**.
   (A Recurrence flow can still be triggered manually for testing.)
2. Watch each step turn green; open *List rows present in a table* and confirm it returned your
   rows, then open the *Apply to each* — one iteration per row.
3. Check your **Sent Items** in Outlook — one message per recipient.
4. Check your inbox for `Apply credit card - Mritunjay Kumar` and confirm the greeting, the body
   and the signature read the way you want.
5. Then add the rest of the people to the sheet (section 5) and run the test once more.

> Outside addresses such as Gmail are fine as recipients. Several mails sent to one Gmail inbox in
> the same minute may get threaded together — they are still separate messages.

---

## 5. Changing the list later

This is the point of keeping it in Excel: **you never open Power Automate again.**

- **Add someone**: type a new row directly under the last one, inside the table. The table border
  extends automatically. Fill `Name`, `Email`, and `Yes` in `Active`.
- **Pause someone**: set `Active` to `No`. The Filter Query skips them, and you keep the row.
- **Remove someone**: delete the whole row (right-click → **Delete** → **Table Rows**).

The next Monday run picks up whatever the sheet says at that moment. Editing it in Excel Online in
the browser is safest, because it saves in place.

### Things that catch people out

- **Never delete and re-upload the file.** Overwrite it in place. A new upload gets a new file ID
  and the flow loses its reference to it.
- **Do not rename the table or the columns.** The flow refers to `Recipients`, `Name` and `Email`
  by name.
- **Do not add rows below a blank line** — they fall outside the table and will be ignored. Check
  that new rows are inside the shaded table area.
- **Do not open the file in the desktop Excel and leave it locked** while the flow runs.

---

## 6. Make it robust (optional hardening)

- **Retry**: on the *Send an email (V2)* card → ⋯ → **Settings** → **Retry Policy** →
  *Exponential Interval*, Count `4`, Interval `PT20S`.
- **Failure alert**: after the *Apply to each*, add **Send an email (V2)** to yourself, then
  ⋯ → **Configure run after** → tick **has failed** and **has timed out**. Subject:
  `Credit card reminder flow FAILED`.
- **Turn it off for a week**: Power Automate → **My flows** → ⋯ → **Turn off**.
- **Ownership**: ⋯ → **Share** → add a co-owner. A flow owned by a single account stops working
  when that account is disabled — a co-owner prevents a surprise outage. Share the OneDrive file
  with them too, or keep the sheet on a SharePoint team site instead of personal OneDrive.
- **Idle suspension**: flows in trial/developer environments can be suspended after long periods
  of inactivity. A weekly run keeps this one active; you will get a warning email if it is ever
  about to be disabled.

---

## 7. Monitoring

- **My flows → Weekly Credit Card Reminder → 28-day run history**: every Monday run with its
  status and per-recipient detail.
- Failures also generate an automatic "flow failure" email from Microsoft.
- Recurrence runs are queued by the service; if the service is briefly busy the run starts a
  little late but is **not** skipped.

---

## 8. Sending from a shared mailbox instead of your own

1. Ask an admin to grant your account **Send As** (or *Send on Behalf*) on the shared mailbox.
2. Replace **Send an email (V2)** with **Office 365 Outlook — Send an email from a shared mailbox (V2)**.
3. Set **Original Mailbox Address** to the shared address, e.g. `cards@yourcompany.com`.

This is the better setup for a recurring team reminder: it survives you leaving or changing roles.

---

## 9. Where settings live in the flow

There is no config file and no secret to paste anywhere. Everything the flow needs is held in the
flow itself or in the connections:

| What | Where it lives |
|---|---|
| Mailbox credentials | The **Office 365 Outlook connection** you signed into once (step 3.5). It stores a consented token in the cloud, listed under **Data → Connections** — that is what lets the flow send while your laptop is off |
| OneDrive access | The **Excel Online (Business) connection** (step 3.3), same mechanism |
| Recipient list | `recipients.xlsx` on OneDrive |
| Subject / body text | The *Send an email (V2)* card |
| Schedule | The Recurrence trigger |

If you would rather keep the tunable values in one place instead of scattered across the cards,
add a `Settings` **Compose** action right after the trigger:

```json
{
  "senderName": "Rakhi",
  "subjectPrefix": "Apply credit card - "
}
```

and reference it downstream:

```
concat(outputs('Settings')?['subjectPrefix'], items('Send_one_email_per_person')?['Name'])
```

For moving a flow between dev/test/prod, package it in a **solution** and use solution
**environment variables** (**Solutions** → **New** → **More** → **Environment variable**) — this
needs a Dataverse-enabled environment and the flow must be created inside the solution.

> **Never put a client secret, password or API key in a Compose action or an environment
> variable** — both are readable by anyone who can open the flow. With the connections above you
> do not need one in the first place.

---

## 10. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Mail arrives at the wrong hour | Time zone not set on the Recurrence trigger (section 3.2) |
| **Table** dropdown is empty | The range was never turned into a named table — section 2 |
| Flow runs green, sends nothing | The table is empty, the new rows fell outside the table, or `Filter Query` excludes every row |
| `Email` is blank in the mail | Column name mismatch — `items(...)?['Email']` must match the header exactly, including case |
| Only 256 people get mail | Turn on **Pagination** with threshold `5000` (step 3.3.5) |
| "File not found" after re-upload | Deleting and re-uploading changes the file ID — reselect the file in the action, and overwrite in place next time |
| `items('Apply_to_each')` is invalid | The name in `items('…')` must match the loop name with underscores |
| Only one mail sent | The *Send an email* action is outside the **Apply to each** — drag it inside |
| Everyone sees each other | Multiple addresses in one **To** — the loop must send one mail per iteration |
| Flow ran, no mail | Check the connections under **Data → Connections**; reauthenticate if one shows an error |
| Mail goes to Junk | Ask recipients to allow your address, or send from a shared mailbox (section 8) |
| External address gets nothing | Some tenants block or quarantine outbound mail to personal domains — check the Microsoft 365 Defender message trace, or test with an internal address |
| Someone gets two mails | Their address appears twice in the sheet — remove the duplicate row |
| Throttling / 429 errors | Set concurrency to 1 (step 3.7) and keep the retry policy |

---

## Files in this folder

| File | Purpose |
|---|---|
| `recipients.xlsx` | Starter sheet with the `Recipients` named table — upload to OneDrive, then edit it there |
| `flow/definition.json` | Reference definition of the finished flow (compare with **Peek code**) |

Once the sheet is on OneDrive, that copy is the live list. The local `recipients.xlsx` is only the
starting point — do not keep editing both.

---

## Note on the wording

The body text is kept **exactly** as specified. If you want it grammatically clean:

> Dear `<name>`, Our records show you have not applied for your credit card yet. Please apply at
> your earliest convenience. Thanks — Rakhi

Change it in **Step 3.5** — the body field of the *Send an email (V2)* card.
