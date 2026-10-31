# Jira to Excel Report with GitHub Actions

Export every issue in a Jira Cloud project to a formatted Excel workbook. Run it on demand from GitHub Actions, with no server and no local setup.

Each report answers a few questions quickly:

- What is the status of every ticket, and who owns it?
- Which tickets are missing a **Description** or **Acceptance Criteria**?
- What are the start and due dates, and how is each ticket labeled?

## What you get

A single `.xlsx` file with one row per issue:

| Column | Source |
|---|---|
| Key, Summary, IssueType, Status, Assignee | Standard Jira fields |
| Labels | Joined into a comma-separated list |
| StartDate, DueDate | Start date custom field and due date |
| Description | `True` / `False`: is it filled in? |
| AcceptanceCriteria | `True` / `False`: is it filled in? |

The sheet comes ready to use:

- **Status colors:** Done (green), In Progress (yellow), Blocked (red), To Do / Backlog (gray), On Hold (blue)
- **Quality flags:** `True` in green, `False` in red, so tickets with missing details stand out
- Auto-filter, frozen header row, and table styling

## How it works

```
GitHub Actions (manual run)
   └─ Export-JiraReport.ps1  (PowerShell 7)
        ├─ Jira Cloud REST API v3  /rest/api/3/search/jql  (token-based pagination)
        └─ ImportExcel module  →  .xlsx
   └─ Uploaded as a workflow artifact (kept for 7 days)
```

## Setup in your own repo

1. **Fork or copy** this repository.
2. **Create a Jira API token** at <https://id.atlassian.com/manage-profile/security/api-tokens>.
3. In your repo, go to **Settings → Secrets and variables → Actions** and add:

   | Type | Name | Example |
   |---|---|---|
   | Variable | `JIRA_BASE_URL` | `https://your-site.atlassian.net` |
   | Variable | `JIRA_EMAIL` | `you@example.com` |
   | Secret | `JIRA_API_TOKEN` | *(your API token)* |

4. Go to **Actions → Jira Report → Run workflow**, enter a project key (default `KAN`), and run it.
5. Open the finished run and download the **`jira-report`** artifact.

> **Custom field IDs vary between Jira sites.** The script defaults to `customfield_10015` (Start date) and `customfield_10320` (Acceptance criteria). To find yours, open `https://your-site.atlassian.net/rest/api/3/field` in a browser and search for the field name. Then pass `-StartDateFieldId` and `-AcceptanceCriteriaFieldId`, or change the defaults in the script.

## Run it locally

Requires PowerShell 7+. The ImportExcel module installs automatically if it is missing.

```powershell
$env:JIRA_BASE_URL  = 'https://your-site.atlassian.net'
$env:JIRA_EMAIL     = 'you@example.com'
$env:JIRA_API_TOKEN = '<your-api-token>'

./Export-JiraReport.ps1 -ProjectKey KAN -OutputFolder ./reports
```

### Parameters

| Parameter | Default | Purpose |
|---|---|---|
| `-ProjectKey` | `KAN` | Jira project to export |
| `-OutputFolder` | `D:\Reports` | Where the workbook is saved |
| `-ExcelFileName` | `JIRA_Report_All.xlsx` | Output file name |
| `-WorksheetName` | `All_Issues` | Sheet name |
| `-StartDateFieldId` | `customfield_10015` | Start date custom field |
| `-AcceptanceCriteriaFieldId` | `customfield_10320` | Acceptance criteria custom field |
| `-JiraBaseUrl`, `-Email`, `-ApiToken` | From environment variables | Jira connection |

## Security notes

- The API token is stored only as a GitHub **secret** and never written to the repo or the report.
- The token can do anything your Jira account can do, so use an account with read access to the project.
- Report artifacts are deleted automatically after **7 days**.

## Tech stack

PowerShell 7 · Jira Cloud REST API v3 · [ImportExcel](https://github.com/dfinke/ImportExcel) · GitHub Actions

---

Questions or ideas? Open an issue or a pull request.
