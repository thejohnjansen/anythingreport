# Anything Report

A lightweight, locally-run web app that turns an Azure DevOps query into a PowerPoint-style slide deck.

## Quick start

```bash
npm install
npm start
```

Then open **http://localhost:3456**, paste your ADO query URL, and click **Run**.

## Prerequisites

| Requirement | Why |
|---|---|
| **Node.js** ≥ 18 | `fetch` is used natively (no extra HTTP lib) |
| **Azure CLI** (`az`) | Auth via `az account get-access-token` — run `az login` first |

## How it works

1. Paste a query URL like `https://microsoft.visualstudio.com/Edge/_queries/query-edit/<guid>`
2. Click **Run** — the server executes the query and fetches work-item details
3. The app renders a slide deck:
   - **Slide 1 — Top of Mind**: editable bullet points (persisted in localStorage)
   - **Remaining slides**: one per Level-2 item (child of the Objective), with a table of its child epics

### Columns

| Column | Source |
|---|---|
| ID | `System.Id` — links to the work item in ADO |
| Title | `System.Title` |
| Midpoint Risk | `OSG.RiskAssessment` — color-coded badge |
| Midpoint Details | `OSG.RiskAssessmentComment` |

### Risk assessment colors

- **On Track** → green ✔
- **At Risk** → orange ⚠
- **Off Track** → red ✘
