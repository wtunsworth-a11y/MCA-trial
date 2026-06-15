# MCA Project Tracker — Setup Guide

## What This Builds

A Google Sheets workbook with 4 tabs:

| Tab | Purpose |
|-----|---------|
| **Dashboard** | Live view of progress (approved entries only), pending approvals, summary by Output/Partner/Schedule |
| **Activities** | Master list of all 95 activities — colour-coded by SO1/SO2/SO3, auto-updated % complete |
| **Tasks** | Breakdown of each Activity into smaller tasks (added by activity owners) |
| **Progress Log** | All progress submissions — Pending/Approved/Rejected workflow |

## 4-Level Hierarchy

```
Strategic Output (SO1 / SO2 / SO3)
  └── Sub-Output (1.1, 1.2 ... 3.4)
        └── Activity (1.1-1, 1.1-2 ... 3.4-8)
              └── Task (1.1-1.T1, 1.1-1.T2 ...)  ← you define these
```

Progress is logged at the **Task** level and rolls up automatically to Activities → Sub-Outputs → Strategic Outputs.

---

## One-Time Setup Steps

### 1. Create a new Google Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet. Name it **MCA Project Tracker**.

### 2. Open the Apps Script editor

`Extensions > Apps Script`

### 3. Paste the code

- Delete the default `myFunction()` in the editor
- Open `Code.gs` from this repository and paste the entire contents
- Before saving, find this line near the top and set the approver's email:

```javascript
APPROVER_EMAIL: 'approver@example.com',   // ← change this
```

### 4. Save and run setup

- Click **Save** (Ctrl+S / Cmd+S)
- From the function dropdown, select `setupWorkbook`
- Click **Run**
- Accept the permissions prompt (the script needs Gmail for notifications and Sheets access)
- Wait ~30 seconds — all 4 sheets will be created and populated

### 5. Share the spreadsheet

- **View/Comment only** for team members who will log progress (they use the dialog)
- **Editor** for the approver
- **Editor** for activity owners who define tasks

---

## Day-to-Day Workflow

### Step 1 — Activity owners define tasks
Each activity needs to be broken into concrete tasks before progress can be logged.

1. Open the spreadsheet
2. Go to **MCA Tracker > Add Tasks to an Activity**
3. Select the activity, fill in task name, who it's assigned to, dates, and optionally a weight (% of the activity this task represents — all tasks under one activity should sum to 100)
4. Repeat for all sub-tasks of that activity

Tasks land in the **Tasks** tab with auto-generated codes like `1.1-1.T1`, `1.1-1.T2`.

### Step 2 — Team members log progress
When work is done on a task:

1. Go to **MCA Tracker > Log Progress on a Task**
2. Select the activity, then the specific task
3. Enter % complete, a status update, and any evidence/notes
4. Click **Submit for Approval**

The entry is saved as **Pending** in the Progress Log. The approver receives an email notification immediately.

### Step 3 — Approver reviews and approves
1. Open the **Progress Log** tab
2. Find entries with status **Pending** (highlighted yellow)
3. Change the **Approval Status** dropdown to **Approved** or **Rejected**
4. Optionally add notes in the **Approver Notes** column
5. The dashboard refreshes automatically; the submitter receives an email notification

Only **Approved** entries affect the dashboard and activity % complete figures.

---

## Dashboard Sections

| Section | Shows |
|---------|-------|
| Pending Approvals | All entries awaiting review |
| Summary by Strategic Output | Count of activities, avg % complete, status |
| Activity Progress Table | All activities with live % complete |
| Progress by Responsible Partner | Breakdown by partner organisation |
| Schedule Status | Activities due each quarter — complete/in progress/not started |

Click **MCA Tracker > Refresh Dashboard** at any time to force a recalculation.

---

## Notes on % Complete Calculation

- Each task can have a **weight** (e.g., a task worth 40% of the activity gets weight 40)
- If no weights are set, all tasks are treated as equal
- The activity % complete = weighted average of the highest approved % for each task
- Activity status: `Not Started` (0%) → `In Progress` (1–99%) → `Complete` (100%)

---

## Sharing for Multiple Users

Since multiple people need to enter progress, the recommended approach is:

1. Share the spreadsheet with team members as **Editors**
2. They use the **MCA Tracker** menu in the sheet to submit entries
3. Alternatively, create a **Google Form** linked to the Progress Log sheet as an alternative entry method — ask if you want this added

The approver is the only person who should change the **Approval Status** column. You can protect that column for extra security:
`Data > Protect sheets and ranges > select column K in Progress Log > restrict to approver only`
