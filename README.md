# Receiving Dashboard-Shwapno

An executive, mobile-friendly GitHub Pages dashboard that reads the latest data from the supplied public Power BI model. It is a custom dashboard—not an embedded Power BI frame and not a static copy of the numbers.

## What management gets

- A one-sentence executive pulse showing whether receipts are building or drawing down inventory.
- Six headline KPIs: received units, sold units, receipt balance, inventory, over-receiving value and active outlets.
- Daily receiving-versus-sales trend.
- Category receipt-balance graph.
- Division exposure table.
- Ranked outlet action queue for over receiving, under receiving and stock cover.
- Cascading filters for period, business division, category, division/region, RHO, Zonal, outlet and article.
- The complete business-division list from the supplied Power BI view: Company Goods, Fresh Produce, General Merchandise, Lifestyle, Loose Commodity and Packed Commodity.
- Search by outlet code/name, article code/name, RHO or Zonal; every committed selection shortens the related choices.
- Clickable sorting on every useful column in the division, outlet and article-detail tables.
- Click-through detail from every KPI and numeric table value to live article-level records.
- Outlet code/name, RHO and Zonal displayed together throughout the outlet action view.
- A remembered light/dark colour-mode switch.
- Manual refresh plus automatic refresh every 15 minutes.
- A clear live-data error state with a direct link to the source report.

## Files to upload

Upload all six files to the repository root:

| File | Purpose |
|---|---|
| `index.html` | Dashboard structure and accessibility |
| `styles.css` | Responsive Shwapno executive design |
| `app.js` | Filters, charts, management signals and interactions |
| `powerbi.js` | Read-only live connection to the published Power BI model |
| `organization.js` | Read-only Zonal/RHO and outlet hierarchy connection |
| `README.md` | Deployment and maintenance guide |

No build command, package installation, server or secret is required.

## Publish at the required URL

The required URL only works when the project repository is named exactly `receiving-dashboard-shwapno`.

### If your repository is currently named `Receiving-Dashboard`

1. Open the repository on GitHub.
2. Go to **Settings → General**.
3. Under **Repository name**, change it to `receiving-dashboard-shwapno` and confirm the rename.
4. Return to the **Code** tab.

### Replace the dashboard files

1. Extract the supplied ZIP.
2. In the repository, choose **Add file → Upload files**.
3. Drag the six files listed above onto the upload page.
4. Confirm that the files are at the repository root—not inside another folder.
5. Choose **Commit directly to the `main` branch**, then click **Commit changes**.
6. Open **Settings → Pages**.
7. Under **Build and deployment**, choose **Deploy from a branch**.
8. Select **main** and **/(root)**, then click **Save**.
9. Wait for the Pages deployment to finish, then open:
   `https://aftabz-lab.github.io/receiving-dashboard-shwapno/`
10. If the old page is cached, use a hard refresh: **Ctrl+Shift+R** on Windows or **Cmd+Shift+R** on Mac.

## How live updating works

On every page load, manual refresh and scheduled 15-minute refresh, the site:

1. Reads the current model/report identifiers from the public Power BI publication.
2. Queries the same live semantic model used by the supplied report.
3. Reads the newest published Zone Distribution snapshot for outlet, RHO and Zonal names.
4. Rebuilds every KPI, chart, cascading filter, drill-down and outlet ranking in the browser.

If the dataset refreshes while the same Power BI Publish-to-web publication remains active, the website reads the updated values without another GitHub upload. The header shows the model refresh time in Bangladesh Standard Time.

The site contains no copied business-data snapshot, username or password. It uses only the existing public Power BI publication and public Zone Distribution snapshot that already support the source dashboards.

## Using the cascading searches

1. Type in the RHO, Zonal, outlet or article field.
2. Choose a shortened suggestion, or enter an exact code/name and press **Enter**.
3. The dashboard refreshes and limits the other available filters to the resulting data scope.
4. Clear a search field to remove that filter, or choose **Reset all** to return to Packed Commodity / latest 30 days.

Click any table heading to sort by that field. Click it again to reverse the order. Click any underlined number to open live article-level details for that KPI, division or outlet.

## Metric definitions and QA notes

- **Received units:** sum of `qty_in_unit_of_entry` from the receiving query.
- **Sold units:** sum of `ActualInvoicedQuantity` from the invoiced-sales query.
- **Receipt balance:** received units minus sold units. Positive means inventory build; negative means inventory drawdown.
- **Inventory, stock day, incident rates, over-receiving value and active outlets:** existing measures from the Power BI model.
- The dashboard keeps the published report's Packed Commodity and movement-type scope.
- Unit and incident totals reconcile across the headline, category and division views.
- **Over-receiving value is a non-additive Power BI measure.** It is recalculated in each category, division and outlet context; do not add its row values. Use the headline card for the overall total.
- A period means the latest complete days ending at the report's current end date.
- Zonal/RHO labels are joined by outlet code. If a Power BI outlet has no row in the newest Zone Distribution file, the dashboard keeps the outlet available and clearly displays **Not mapped** instead of inventing a hierarchy.
- During the 8 September 2026 validation, 990 of 1,225 Power BI outlet options matched the current organization snapshot. This coverage changes automatically when either public source is refreshed.

## If the live Power BI publication changes

Normal dataset refreshes require no code change. A code update is required only if the owner creates a different Publish-to-web link, disables public access, moves the model to a different Power BI API cluster, or renames/removes model fields.

The connection settings are at the top of `powerbi.js`:

- `API_ROOT`
- `RESOURCE_KEY`
- `POWER_BI_URL`

## Troubleshooting

- **GitHub 404:** verify the repository name, confirm `index.html` is in the root, and check that Pages deploys from `main` / `/(root)`.
- **Live data unavailable:** first open the **Source report** button. If that report also fails, restore or republish it in Power BI. If only this dashboard fails, the publication key, API cluster or model schema likely changed.
- **Old design still appears:** wait for the Pages workflow to finish and hard-refresh the browser.
- **Filters return no exposure:** the selected combination may have no incidents; reset the filters or widen the period.
- **RHO/Zonal shows Not mapped:** the outlet is present in Power BI but its code is absent from the newest Zone Distribution snapshot. Correct the outlet-code mapping in the source file; the dashboard will use it after the public snapshot updates.

## Optional local preview

From the extracted folder, run:

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000`. Opening `index.html` directly as a `file://` URL will not reliably load JavaScript modules.
