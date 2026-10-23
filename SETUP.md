# Usher Attendance: Setup

This replaces the paper Church Attendance sheet with a web page ushers open on their phones. Each time an usher submits, it adds one row to a Google Sheet and emails a report to the staff. A second page, for the church office, takes the week's other numbers (Sunday School, Keenagers, AWANA) and emails a weekly report every Saturday morning.

It has three parts:

- **The Sheet and its script** (`Code.gs`, `Report.gs`, `History.gs`, `Office.gs` and `Weekly.gs`). The script lives inside the Google Sheet. It saves each submission as a row, tells the pages what to show, and emails the reports.
- **The ushers' page** (`index.html`, plus the logo, icons, and `manifest.webmanifest`). It's hosted on GitHub Pages at <https://davidluttrull.github.io/grace-attendance/>. The church site forwards <https://gracetoanderson.com/ushercount> there, and that's the address ushers use.
- **The office page** (`office/index.html`), at <https://davidluttrull.github.io/grace-attendance/office/>. See "The office page" below.

The ushers' page sends its numbers to the script without any Google sign-in. Ushers don't need a Google account, and it works whatever account their phone happens to be signed in to. Instead, each submission must carry the **usher code**, a 6-digit number the script creates for you. Without it, a stranger who found the page couldn't add, change or look up attendance.

## 1. Create the Sheet

1. Sign in to the Google account that should own the attendance records. The church office account is a better choice than a personal one, because the page keeps working when people leave.
2. Create a new blank Google Sheet and name it something like **Church Attendance**.

## 2. Add the script

1. In the Sheet, open **Extensions → Apps Script**.
2. Click "Untitled project" at the top and rename it **Usher Attendance**.
3. You'll see a file called `Code.gs` with a few lines in it. Delete them, then paste in the entire contents of this folder's `Code.gs`.
4. Click the **+** next to "Files", choose **Script**, and name it `Report` (Apps Script adds the `.gs` itself). Delete what's inside and paste in the contents of `Report.gs`. Do the same for `History.gs`, `Office.gs` and `Weekly.gs`, naming them `History`, `Office` and `Weekly`.
5. Click the save icon (💾).

## 3. Run setup once

1. Click **Code.gs** in the Files list. The Run button and the function dropdown appear only while a `.gs` file is open. In the dropdown next to **Debug** (it probably shows `doGet`), pick **setup**, then click **Run**.
2. Google will ask for permission. Click **Review permissions** and choose your account.
3. You'll get a warning that says "Google hasn't verified this app". That's expected for a script you wrote yourself. Click **Advanced → Go to Usher Attendance (unsafe) → Allow**. The list of permissions includes these:
   - **See, edit, create, and delete all your Google Sheets spreadsheets.** The script saves to this Sheet, and the history import reads the old attendance sheet.
   - **Send email as you**, for the reports.
   - **Allow this application to run when you are not present**, so the report can go out after an usher submits and the weekly report can go out on Saturday.
4. Switch back to the Sheet. You should see four tabs:
   - **Attendance**, with a dark green header row
   - **Summary**
   - **Weekly**, for the office page's numbers
   - **Report Emails**
5. On the **Report Emails** tab, type the staff email addresses in column A, one per row, starting in A2. A Google Group or routing address works too.
6. Back in Apps Script, click **Execution log** at the top. The last two lines read **Usher code: 123456** and **Office code: 654321**, with your own numbers. Give the usher code to the ushers and the office code to the office. Running setup again shows the same codes.

Setup also schedules the weekly report for Saturday mornings. Running it again doesn't schedule a second one.

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

The page has to know the script's address. Near the bottom of `index.html`, find the line that starts with `const API_URL =`, and put the `/exec` URL from step 4 between the quotes. `office/index.html` has the same line; give it the same URL.

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

## The office page

The office fills in the week's other numbers at <https://davidluttrull.github.io/grace-attendance/office/>. It's meant for a computer but works on a phone too. If you'd like a shorter address, have the church site forward something like `gracetoanderson.com/officecount` there, the way `ushercount` forwards to the ushers' page.

- The first time, it asks for a name and the **office code**, and that computer remembers both. The usher code doesn't open it: the office page can read and change everything the ushers logged.
- A wrong office code is cleared and the page asks again. After 20 wrong office codes within 10 minutes, the script stops accepting the office code for 10 minutes. This count is separate from the usher code's, so a locked-out office page doesn't stop the ushers.
- To sign in as someone else, click **Not (name)? Change name or code** at the bottom of the page.
- It opens on the current week, Sunday through Saturday. The arrows move a week at a time, and **Go to** jumps to the week containing any date. **This week** comes back to the current week. The right arrow stops at the current week. The address ends in the week's Sunday (`…/office/#2026-10-18`), so a bookmark goes back to that week.
- **Usher Counts** lists the week's services: Sunday AM, Sunday PM, Wednesday, and any others the ushers logged. Open one to see every area. Change a number to correct it, or fill in a service the ushers missed. Saving updates or adds that service's row on the Attendance tab and notes who changed it in the **Office Edit** column. It doesn't email the per-service report.
- A service the office fills in gets a new row on the Attendance tab with the office person's name under **Submitted By**, "Entered by" in **Office Edit**, and *Not emailed: entered on the office page* under **Report Emailed**. To send its report anyway, use the **Attendance** menu (see "Emailed reports").
- Below that are the numbers the ushers don't count: Sunday School adults and teens, the five Sunday School classes (Nursery, K4-K5, 1st-2nd, 3rd-4th, 5th-6th), Keenagers and AWANA. Children's attendance during services isn't here, because the ushers count it.
- Sunday School **Children** isn't typed in. It's the total of the classes, and the Sunday School **Total** is adults plus teens plus children. Both update as the boxes change.
- On the Weekly tab, Keenagers and AWANA are saved under the History tabs' column names, **Keenagers/HGM** and **K4T/AWANA**, so the weekly report can find last year's numbers.
- **Save** puts the week on the **Weekly** tab, one row per week. Save as often as you like through the week: each save replaces that week's row. Saving doesn't email anyone.
- For weeks before the office page, it shows the old sheet's numbers from the History tabs. Saving one of those weeks copies its numbers to the Weekly tab. Words such as "No Awana" aren't copied, because the boxes only take numbers, and saving the week clears them from the History tab as well. If a word like that matters, type it in **Notes** before you save.
- If you move to another week or close the page with unsaved changes, it asks first.

## Changing the usher code or the office code

Change the usher code if it gets passed around more than you'd like, or when an usher leaves the team. In Apps Script, click the gear icon (**Project Settings**) on the left, scroll to **Script Properties**, click **Edit script properties**, change the value of **USHER_CODE**, and click **Save script properties**. The new code works right away; you don't need to redeploy. Each phone asks for the new code the next time someone submits.

The office code is **OFFICE_CODE** in the same list, and it changes the same way. The office page asks for the new code the next time it opens.

## Emailed reports

About a minute after each submission, the script emails the report to everyone on the **Report Emails** tab. The report is the email itself, not an attachment, laid out like the ushers' page: the dark green header with the logo, then cards. It has the total and each section in boxes at the top, where everyone was, the auditorium as a floor plan, the other areas as bars, the full count, the last few services of the same kind, and the bulletins, door ushers and notes. On a computer some cards show two columns side by side; on a phone they stack. Email apps that can't show the layout get a plain-text version with the same numbers.

- The **Report Emailed** column at the right end of the Attendance tab shows what happened to each row: *Pending*, *Sent* with the time and number of recipients, *Not sent* if the list is empty, *Failed* with the reason, or *Not emailed: entered on the office page* for a service the office filled in.
- To send the report, the script sets itself a one-time timer for `sendPendingReports`. You may see it for a moment under **Triggers** in Apps Script. Leave it alone; it removes itself when it runs.
- If an usher replaces a service that was already logged, staff get a new report with "Updated" in the subject.
- To stop the emails, clear the addresses on the Report Emails tab.

The Sheet has an **Attendance** menu (it appears a few seconds after the Sheet opens):

- **Email report for selected row** sends that row's report to everyone on the list again. Use it after a *Failed* row, or to resend an old service.
- **Send me a test report for selected row** sends it only to you, so you can check how it looks without emailing the staff.
- **Email weekly report** and **Send me a test weekly report** are covered under "The weekly report" below.
- **Import history from the old attendance sheet** and **Fill in History tabs from Attendance and Weekly** are covered under "History from the old attendance sheet" below.

### The weekly report

Every Saturday at about 7 AM, the script emails the week just past to the same **Report Emails** list. It covers each service's count by section, then Sunday School, its classes, Keenagers and AWANA, each beside the same week last year. Office and usher notes come last.

- The **Weekly Report** column on the Weekly tab shows when each week's report went out, or why it didn't. If the office never saved a week, the report still goes out, and that week gets a row on the Weekly tab with only the date and this column filled in. The office page still treats that week as not saved.
- The menu has **Email weekly report** and **Send me a test weekly report**. Click a row on the Weekly tab first to send that week. Otherwise they send the week that started on the most recent Sunday, so early in the week that report is mostly empty.
- Only **Email weekly report** and the Saturday email write to the Weekly Report column. The test doesn't.
- The hour is `WEEKLY_HOUR` near the top of `Weekly.gs`, in the script's time zone. Google runs it some time within that hour. The day is the word `SATURDAY` a few lines below it, in `ensureWeeklyTrigger_`. To change either, open **Triggers** (the alarm clock on the left of Apps Script), delete the `sendWeeklyReport` trigger, edit `Weekly.gs`, and run setup again. The email's footer says "every Saturday morning", so change that wording in `Weekly.gs` too if you move the day.
- If the Saturday email stops coming, open **Triggers** and check that `sendWeeklyReport` is listed. If it isn't, run setup again.

Reports used to be sent as a PDF, made by way of a temporary Google Doc. They aren't anymore, so the script no longer needs the Drive API service. If it's listed under **Services** in Apps Script, you can leave it or remove it.

## History from the old attendance sheet

Attendance from July 2006 until the ushers' page was kept by hand in the **Attendance Report** spreadsheet. `History.gs` copies it into this Sheet, one tab per year: **History 2006**, **History 2007**, and so on.

- Each tab has the old sheet's columns: Date, Bus, Keenagers/HGM, K4T/AWANA, Sunday AM, Auditorium, Kidstown, Sunday PM, Wednesday, the four Sunday School columns, and Notes. From 2024, the Children's Ministries numbers are added on the right, on the row with the same date.
- Numbers stay numbers, and words like "No Service" or "Snow Day" stay as they were typed. A few dates in the old sheet had the wrong year (like "2/4/06" in the 2007 section); they're filed under the year they were listed in, with a note saying what was typed. Rows with nothing in them, and the monthly averages and summary tables, aren't copied.
- The import reads the old sheet directly, so the Google account that runs it needs at least view access to the **Attendance Report** spreadsheet. Which spreadsheet it reads is set near the top of `History.gs` (`OLD_SHEET_ID`).
- To import, open Apps Script, click **History.gs** in the Files list, pick **importHistory** in the dropdown, and click **Run**. It takes a minute and shows how many rows each year got. You can also run it from the Sheet's **Attendance** menu (**Import history from the old attendance sheet**).
- Nothing is left out silently. A **History Import Log** tab lists every row of the old sheet that wasn't copied, and any cells cut off the right of a copied row, with the reason and what was in the row. Children's rows whose date has no match in the main tab are listed too, because they're usually a typo. Look the log over after an import.
- Running it again rebuilds the History tabs and the log from the old sheet, then fills the new pages' weeks back in (see below). To correct a past number, fix it in the old sheet and import again, or edit the History tab and don't import again.

### History keeps going from the new pages

The History tabs stay the full record, one row per week, after the old sheet. Each time an usher submits or the office saves, the script fills in that week's row on the History tab for its year. It creates **History 2027** and later tabs when the year turns.

| History column | Comes from |
|---|---|
| Sunday AM, Sunday PM, Wednesday | That service's Total Attendance on the Attendance tab. Wednesday goes on the row of the Sunday before it, as in the old sheet. |
| Auditorium | Sunday AM's Auditorium Subtotal. |
| AM Nursery, PM Nursery, Wed Nursery | The three nurseries added together, at that service. |
| AM K4-K5, Wed K4-K5 | Pre-Primary (4-5 yrs), at that service. |
| AM 1st-3rd, Wed 1st-3rd | Junior Church (1st-3rd), at that service. |
| Keenagers/HGM, K4T/AWANA, the Sunday School columns, Notes | The week's row on the Weekly tab. SS Teens repeats SS Teen, as the old sheet did. |
| Bus, Kidstown, Wed 4th-6th, Wed Teens | Nothing any more. They stay blank, or keep what the old sheet had. |

- A week goes on the History row dated in that week, so a Sunday count an usher logged on Tuesday still lands on Sunday's row.
- Only columns that have something to fill them change. If the ushers haven't logged a service or the office hasn't saved the week, those columns keep whatever was there. A blank office note doesn't erase the old sheet's note.
- Edits typed straight into a History tab stay until that week is filled in again, which happens the next time an usher submits that week or the office saves it. To change a number for good, change it on the Attendance or Weekly tab (or on the office page).
- **Attendance → Fill in History tabs from Attendance and Weekly** fills in every week at once. Use it after editing the Attendance or Weekly tab by hand.

The emailed report reads the History tabs for **Recent services** and **Compared with last year**: the same week last year, this year's average so far, and last year's average over the same weeks. The Total Attendance box also says how this service compares with the same week last year. The Attendance tab counts first; history fills in the dates before it. For Wednesday, the old sheet kept the count on the Sunday row, so the report dates it to that week's Wednesday.

## The Sheet

- **Attendance:** one row per service, in this order: date, service, total, every area grouped by section with subtotals, bulletins, door ushers, who logged it, when, notes, whether the report was emailed, and any change made on the office page.
- **Summary:** one row per date, with the total for each service in its own column. It updates automatically.
- **Weekly:** the office page's numbers, one row per week (the Sunday's date), with the same column names as the History tabs, who saved it last, and when the weekly report went out.
- **Report Emails:** the addresses that get both reports.
- **History 2026** back to **History 2006:** one tab per year, newest first, after the other tabs. The old sheet's numbers, and from October 2026 the weeks from the ushers' page and the office page (see above).
- **History Import Log:** the last tab. Every row of the old sheet the import didn't copy, and why. Each import replaces it.

You can sort, filter, chart, or add your own columns to the right. The script finds its columns by their header names. Leave the **Attendance** and **Weekly** tab names and the existing header names unchanged.

## Updating the script

Whenever `Code.gs`, `Report.gs`, `History.gs`, `Office.gs` or `Weekly.gs` changes:

1. Paste the new version into Apps Script and save.
2. Open **Code.gs** and run **setup** once (step 3 above). The dropdown lists only the open file's functions, so setup won't appear while another file is open. If the new version needs a new permission, Google asks for it here. The pages can't ask, so skipping this step makes submissions fail until someone runs setup.
3. Click **Deploy → Manage deployments**, then click the pencil icon. Under **Version**, choose **New version**, then click **Deploy**.

Don't use **New deployment**, because it creates a new URL and the pages would stop reaching the script.

## Changing rooms or areas

Room names live near the top of `Code.gs`, in `SECTIONS`. To add, rename, or remove one, edit the list and follow "Updating the script" above.

The page itself doesn't need to change, because it reads the room list from the script. A phone that has the page open shows the new list a moment after the page loads. If someone submits from a page that still shows a removed room, the script refuses the numbers and asks them to reopen the page, so nothing is lost without a warning.

A new area gets its own column at the right end of the Attendance tab. Removing an area from the page leaves its old column in the Sheet, so past numbers aren't lost.

A few rooms are also named in other files, so renaming one of them takes an extra edit:

- The History tabs' children's columns look for the three nurseries, **Pre-Primary (4-5 yrs)** and **Junior Church (1st-3rd)** by name, in `historySources_` in `History.gs`. If you rename one, change it there too, or those History columns stop filling in.
- The auditorium floor plan in the emailed report looks for **South Wing**, **South Main**, **Platform**, **North Wing** and **North Main**. If they change, the report shows bars instead of the floor plan.

The office page's boxes are in `OFFICE_GROUPS` near the top of `Office.gs`. Each box's `key` is a History tab column name, so change the `label` (what people see) and leave the `key` alone. Leave the card titles (`Sunday School`, `Sunday School Classes`, `Ministries`) alone as well, because the weekly report finds the cards by those titles.
