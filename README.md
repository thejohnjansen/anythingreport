# Anything Report

A lightweight, locally-run web app that turns an Azure DevOps query into a PowerPoint-style slide deck.

## How to run this tool

### To install locally (Windows)

1. Browse to: [My Onedrive](https://onedrive.cloud.microsoft/:f:/a@668pm6kq/r/_layouts/15/onedrive.aspx?id=%2Fa%40668pm6kq%2FDocuments%2Fdist&share=cgqrp5PxfDkFRLlImaHqItmpEgUCKRl%5FPwW8QK88udS0%2DN%2DT1Q)  
2. Double-click **Anything Report Setup 1.0.0**
   > You may need to approve the download a couple of times before it will run.
3. The app should launch automatically, if not, double-click **Anything Report** on your desktop.

### Running reports

1. Initially you will see **Top of Mind**
2. Add rich text — it will automatically be saved to a Cosmos DB backend
3. Paste in your team's ADO query URL — it should contain at least one Objective with *n* Epics under it, and *m* child Epics under those, e.g.:
   ```
   https://dev.azure.com/microsoft/Edge/_queries/query/c4e47939-b63c-4d6e-bcf4-811d15655294/
   ```
4. Click through to one of the **Pipeline** slides
5. Hover a Topic and click **Edit** to update its contents — changes are saved automatically to the backend

   > **NOTE:** If you and someone else are editing Top of Mind or Pipeline slides on the web site, you should see each other's updates in nearly real time.

6. Click to create a **PPTX** of the content
7. Verify everything looks correct and update as necessary

### On a Mac

```bash
# Clone or pull the repo, then:
npm install
npm run dist:mac   # produces dist/Anything Report-1.0.0.dmg
```

> **Note:** macOS Gatekeeper may warn that the app is from an unidentified developer.
> Right-click the `.dmg` → **Open** → **Open anyway** to proceed.

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

### Open any bugs or feature requests in ADO
- Area Path: Edge/Web Platform
- Assigned to: John Jansen