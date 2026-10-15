# Usher Attendance: Setup

This replaces the paper Church Attendance sheet with a web page ushers open on their phones. Each time an usher submits, it adds one row to a Google Sheet and emails a PDF report to the staff.

It has two parts:

- **The Sheet and its script** (`Code.gs` and `Report.gs`). The script lives inside the Google Sheet. It saves each submission as a row, tells the page which rooms to show, and emails the report.
- **The page** (`index.html`, plus the logo, icons, and `manifest.webmanifest`). It's hosted on GitHub Pages at <https://davidluttrull.github.io/grace-attendance/>. The church site forwards <https://gracetoanderson.com/ushercount> there, and that's the address ushers use.

The page sends its numbers to the script without any Google sign-in. Ushers don't need a Google account, and it works whatever account their phone happens to be signed in to. Instead, each submission must carry the **usher code**, a 6-digit number the script creates for you. Without it, a stranger who found the page couldn't add, change or look up attendance.

## 1. Create the Sheet

1. Sign in to the Google account that should own the attendance records. The church office account is a better choice than a personal one, because the page keeps working when people leave.
2. Create a new blank Google Sheet and name it something like **Church Attendance**.

## 2. Add the script

1. In the Sheet, open **Extensions → Apps Script**.
2. Click "Untitled project" at the top and rename it **Usher Attendance**.
3. You'll see a file called `Code.gs` with a few lines in it. Delete them, then paste in the entire contents of this folder's `Code.gs`.
4. Click the **+** next to "Files", choose **Script**, and name it `Report` (Apps Script adds the `.gs` itself). Delete what's inside and paste in the contents of `Report.gs`.
5. Click the **+** next to "Services", choose **Drive API**, and click **Add**. The report uses it to turn each report into a PDF.
6. Click the save icon (💾).

## 3. Run setup once

1. Click **Code.gs** in the Files list. The Run button and the function dropdown appear only while a `.gs` file is open. In the dropdown next to **Debug** (it probably shows `doGet`), pick **setup**, then click **Run**.
2. Google will ask for permission. Click **Review permissions** and choose your account.
3. You'll get a warning that says "Google hasn't verified this app". That's expected for a script you wrote yourself. Click **Advanced → Go to Usher Attendance (unsafe) → Allow**. The list of permissions includes Sheets, Drive, Docs and sending email: the script needs Drive and Docs to build each report as a PDF, and email to send it.
4. Switch back to the Sheet. You should see three tabs:
   - **Attendance**, with a dark green header row
   - **Summary**
   - **Report Emails**
5. On the **Report Emails** tab, type the staff email addresses in column A, one per row, starting in A2. A Google Group or routing address works too.
6. Back in Apps Script, click **Execution log** at the top. The last line reads **Usher code: 123456** with your own number. Give that code to the ushers. Running setup again shows the same code.

## 4. Deploy the script

1. In Apps Script, click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in:
   - **Description:** `v1`
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**. The page needs this setting to reach the script without a sign-in.
4. Click **Deploy** and copy the **Web app URL**. It ends in `/exec`.

To check the deployment, open that URL in a browser. You should see a short block of text that starts with `{"sections":`. That's the room list, and it means the script is working.

## 5. Connect the page to the script

The page has to know the script's address. Near the bottom of `index.html`, find the line that starts with `const API_URL =`, and put the `/exec` URL from step 4 between the quotes.

This only needs to be done once. Deploying a new version of the script later keeps the same URL (see "Changing rooms or areas" below).

## 6. Put it on the ushers' phones

Send the ushers the page's address, <https://gracetoanderson.com/ushercount>, and the usher code. Each usher should save the page to their home screen. It shows up as **Attendance** with the Grace cross icon. The page asks for the code under "Logged By" the first time, and the phone remembers it after that.

- **iPhone:** open the link in Safari, tap the Share button, then tap **Add to Home Screen**.
- **Android:** open the link in Chrome, tap **⋮**, then tap **Add to Home screen** (or **Install app**).

## How it works

- The page fills in today's date and picks the service from the day and time: Sunday before 3 PM is AM, Sunday after 3 PM is PM, and Wednesday is Wednesday. On other days, the usher picks one.
- Ushers can tap **+ / −** or tap a box and type a number. Subtotals and the total update as they go. Bulletins Left is not counted in the total.
- Usher Assignments (who stood at each door) is optional and starts collapsed.
- The phone remembers the usher's name. It also keeps unfinished counts for 12 hours, so a dropped connection or a locked screen won't lose anything. If saving fails, the numbers stay on the screen so the usher can try again.
- If someone submits a service that's already logged (same date and same service), the page shows who logged it and asks whether to replace it. Nothing gets counted twice.
- A wrong usher code is rejected and cleared from the phone. After 20 wrong codes within 10 minutes, from any phones, the script stops accepting submissions for 10 minutes, so nobody can guess the code by trying numbers.

## Changing the usher code

Change it if it gets passed around more than you'd like, or when an usher leaves the team. In Apps Script, click the gear icon (**Project Settings**) on the left, scroll to **Script Properties**, click **Edit script properties**, change the value of **USHER_CODE**, and click **Save script properties**. The new code works right away; you don't need to redeploy. Each phone asks for the new code the next time someone submits.

## Emailed reports

About a minute after each submission, the script emails a two-page PDF report to everyone on the **Report Emails** tab. The report has the totals, a breakdown by area, the auditorium by section, the full count, the last few services of the same kind, and the door ushers and notes. The email itself shows the total and the section subtotals, so staff can see the numbers without opening the attachment.

- The **Report Emailed** column at the right end of the Attendance tab shows what happened to each row: *Pending*, *Sent* with the time and number of recipients, *Not sent* if the list is empty, or *Failed* with the reason.
- If an usher replaces a service that was already logged, staff get a new report with "Updated" in the subject.
- To stop the emails, clear the addresses on the Report Emails tab.

The Sheet has an **Attendance** menu (it appears a few seconds after the Sheet opens):

- **Email report for selected row** sends that row's report to everyone on the list again. Use it after a *Failed* row, or to resend an old service.
- **Send me a test report for selected row** sends it only to you, so you can check how it looks without emailing the staff.

Each report is built in a temporary Google Doc that's moved to the Drive trash right after the PDF is made, so you may see "Attendance report (temporary)" files in the trash. Google empties the trash on its own after 30 days.

## The Sheet

- **Attendance:** one row per service, in this order: date, service, total, every area grouped by section with subtotals, bulletins, door ushers, who logged it, when, and notes.
- **Summary:** one row per date, with the total for each service in its own column. It updates automatically.
- **Report Emails:** the addresses that get the report.

You can sort, filter, chart, or add your own columns to the right. The script finds its columns by their header names. Leave the **Attendance** tab name and the existing header names unchanged.

## Updating the script

Whenever `Code.gs` or `Report.gs` changes:

1. Paste the new version into Apps Script and save.
2. Open **Code.gs** and run **setup** once (step 3 above). The dropdown lists only the open file's functions, so setup won't appear while Report.gs is open. If the new version needs a new permission, Google asks for it here. The ushers' page can't ask, so skipping this step makes submissions fail until someone runs setup.
3. Click **Deploy → Manage deployments**, then click the pencil icon. Under **Version**, choose **New version**, then click **Deploy**.

Don't use **New deployment**, because it creates a new URL and the page would stop reaching the script.

## Changing rooms or areas

Room names live near the top of `Code.gs`, in `SECTIONS`. To add, rename, or remove one, edit the list and follow "Updating the script" above.

The page itself doesn't need to change, because it reads the room list from the script. A phone that has the page open shows the new list a moment after the page loads. If someone submits from a page that still shows a removed room, the script refuses the numbers and asks them to reopen the page, so nothing is lost without a warning.

A new area gets its own column at the right end of the Attendance tab. Removing an area from the page leaves its old column in the Sheet, so past numbers aren't lost.
