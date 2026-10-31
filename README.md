<div align="center">

# 📊 Jira → Excel Reports with GitHub Actions

**One click in GitHub Actions turns any Jira Cloud project into a color-coded Excel report: no server, no local setup.**

[![Jira Report](https://github.com/21bshwjt/GHActions_JIRA/actions/workflows/jira-report.yml/badge.svg)](https://github.com/21bshwjt/GHActions_JIRA/actions/workflows/jira-report.yml)
![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-5391FE?logo=powershell&logoColor=white)
![Jira Cloud](https://img.shields.io/badge/Jira%20Cloud-REST%20API%20v3-0052CC?logo=jira&logoColor=white)
![GitHub Actions](https://img.shields.io/badge/GitHub-Actions-2088FF?logo=githubactions&logoColor=white)
![Excel](https://img.shields.io/badge/Output-Excel%20.xlsx-217346?logo=microsoftexcel&logoColor=white)
[![GitHub stars](https://img.shields.io/github/stars/21bshwjt/GHActions_JIRA?style=social)](https://github.com/21bshwjt/GHActions_JIRA/stargazers)
[![Last commit](https://img.shields.io/github/last-commit/21bshwjt/GHActions_JIRA)](https://github.com/21bshwjt/GHActions_JIRA/commits/main)

[What you get](#-what-you-get) •
[Quick start](#-quick-start) •
[Architecture](#-architecture) •
[Run locally](#-run-it-locally) •
[Security](#-security)

</div>

---

## 💡 Why this repo?

Teams spend time exporting Jira issues by hand, then cleaning them up in Excel to answer the same questions every sprint. This repo automates that with a **GitHub Actions workflow** and a **PowerShell script**.

- ✅ **One-click reports:** run the workflow, download the `.xlsx`
- ✅ **Spot gaps fast:** flags tickets with no **Description** or **Acceptance Criteria**
- ✅ **Color-coded status:** Done, In Progress, Blocked and more at a glance
- ✅ **No infrastructure:** runs on GitHub-hosted runners
- ✅ **Service account auth:** OAuth 2.0 client credentials, not a personal API token
- ✅ **Self-cleaning:** report artifacts are deleted after **7 days**

---

## 📦 What you get

An Excel workbook with **one row per issue**, plus auto-filter, a frozen header and table styling.

| Column | What it shows |
|---|---|
| **Key** · **Summary** · **IssueType** | Basic issue details |
| **Status** | 🟢 Done · 🟡 In Progress · 🔴 Blocked · ⚪ To Do / Backlog · 🔵 On Hold |
| **Assignee** | Who owns it |
| **Labels** | All labels, comma-separated |
| **StartDate** · **DueDate** | Planning dates |
| **Description** | 🟢 `True` if filled in · 🔴 `False` if empty |
| **AcceptanceCriteria** | 🟢 `True` if filled in · 🔴 `False` if empty |

---

## 🚀 Quick start

1. **Fork** this repository.
2. **Create an Atlassian service account with OAuth 2.0 credentials** (*admin.atlassian.com → Directory → Service accounts*). Give it the **`read:jira-work`** scope and **Browse** permission on your project. Note the **Client ID** and **Client secret**.
3. **Find your Cloud ID** by opening `https://your-site.atlassian.net/_edge/tenant_info` in a browser.
4. Go to **Settings → Secrets and variables → Actions** in your fork and add:

   | Type | Name | Example |
   |---|---|---|
   | 🔧 Variable | `JIRA_CLOUD_ID` | `a1b2c3d4-...` |
   | 🔒 Secret | `JIRA_CLIENT_ID` | *(service account Client ID)* |
   | 🔒 Secret | `JIRA_CLIENT_SECRET` | *(service account Client secret)* |

5. Open **Actions → Jira Report → Run workflow**, enter your **project key** (default `KAN`) and run it.
6. When it finishes, download the **`jira-report`** artifact from the run summary. 🎉

> [!IMPORTANT]
> **Custom field IDs differ on every Jira site.** The defaults are `customfield_10015` (Start date) and `customfield_10320` (Acceptance criteria).
> To find yours, open `https://your-site.atlassian.net/rest/api/3/field` in a browser while signed in and search for the field name. Then update the defaults in [`Export-JiraReport.ps1`](Export-JiraReport.ps1).

---

## 🏗️ Architecture

```mermaid
flowchart LR
    U["👤 You<br/>Run workflow"] --> GA["⚙️ GitHub Actions<br/>ubuntu-latest"]
    GA --> PS["📜 Export-JiraReport.ps1<br/>PowerShell 7"]
    S[("🔒 GitHub Secrets<br/>& Variables")] -. credentials .-> PS
    PS -- "OAuth 2.0<br/>client credentials" --> O["🔑 auth.atlassian.com"]
    PS -- "REST API v3 via api.atlassian.com<br/>/search/jql (paged)" --> J[("🟦 Jira Cloud")]
    PS -- "ImportExcel" --> X["📊 .xlsx report"]
    X --> A["📦 Workflow artifact<br/>(kept 7 days)"]
    A --> U
```

**How it works**

1. You trigger the workflow manually (`workflow_dispatch`) with a project key.
2. The script exchanges the service account's Client ID and secret for an OAuth access token.
3. It calls the Jira Cloud **REST API v3** through `api.atlassian.com/ex/jira/{cloudId}`, following `nextPageToken` until every issue is fetched.
4. Each issue becomes one row, with `True`/`False` flags for Description and Acceptance Criteria.
5. **ImportExcel** writes the styled workbook, and the workflow uploads it as an artifact.

---

## 💻 Run it locally

Requires **PowerShell 7+**. The [ImportExcel](https://github.com/dfinke/ImportExcel) module installs automatically on first run. Microsoft Excel is **not** needed to generate the report.

```powershell
$env:JIRA_CLOUD_ID      = '<your-cloud-id>'
$env:JIRA_CLIENT_ID     = '<your-client-id>'
$env:JIRA_CLIENT_SECRET = Read-Host 'Client secret'

./Export-JiraReport.ps1 -ProjectKey KAN -OutputFolder ./reports
```

<details>
<summary><b>All parameters</b></summary>

| Parameter | Default | Purpose |
|---|---|---|
| `-ProjectKey` | `KAN` | Jira project to export |
| `-OutputFolder` | `D:\Reports` | Where the workbook is saved |
| `-ExcelFileName` | `JIRA_Report_All.xlsx` | Output file name |
| `-WorksheetName` | `All_Issues` | Sheet name |
| `-StartDateFieldId` | `customfield_10015` | Start date custom field |
| `-AcceptanceCriteriaFieldId` | `customfield_10320` | Acceptance criteria custom field |
| `-CloudId` · `-ClientId` · `-ClientSecret` | `JIRA_CLOUD_ID` · `JIRA_CLIENT_ID` · `JIRA_CLIENT_SECRET` | OAuth 2.0 service account (preferred) |
| `-JiraBaseUrl` · `-Email` · `-ApiToken` | `JIRA_BASE_URL` · `JIRA_EMAIL` · `JIRA_API_TOKEN` | Fallback: basic auth with an API token |

</details>

---

## 🔐 Security

- 🔒 The Client ID and secret are stored only as **GitHub Secrets**. They never appear in the repo, the logs or the report.
- 🤖 A **service account** is used instead of a personal account, so access is limited to the `read:jira-work` scope and the projects you grant it.
- ⏱️ Each run exchanges the credentials for an access token that expires after **1 hour**.
- 🗑️ Report artifacts are **automatically deleted after 7 days**.

---

## 📁 Project structure

```
GHActions_JIRA/
├── .github/workflows/
│   └── jira-report.yml      # Manual workflow: runs the script, uploads the report
├── Export-JiraReport.ps1    # Fetches issues from Jira and builds the Excel file
└── README.md
```

---

## 🧰 Tech stack

**PowerShell 7** · **Jira Cloud REST API v3** · **[ImportExcel](https://github.com/dfinke/ImportExcel)** · **GitHub Actions**

---

<div align="center">

⭐ **If this saved you time, give the repo a star!** ⭐

Questions or ideas? [Open an issue](https://github.com/21bshwjt/GHActions_JIRA/issues) or send a pull request.

</div>
