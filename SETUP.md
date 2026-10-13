# Receiving dashboard: RHO / Zonal mapping fix

## Upload

1. Extract `Receiving_RHO_Zonal_Drive_Mapping_Fix_GitHub_Upload.zip`.
2. Open your existing GitHub repository **aftabz-lab/receiving-dashboard-shwapno**.
3. Upload these four files to the repository's top level, replacing existing copies where applicable:
   - `app.js`
   - `index.html`
   - `organization.js`
   - `organization-fallback.json` (new)
4. Commit the upload to the branch used by GitHub Pages. Wait for its deployment to complete.
5. Open https://aftabz-lab.github.io/receiving-dashboard-shwapno/ and refresh once with Ctrl+Shift+R.

No changes to your Power BI watcher, snapshot workflow, Google Apps Script watcher, dashboard styles or export calculation rules are required by this patch.

## Source and automatic updates

The hierarchy source remains the Zone Distribution workbook from this existing Google Drive folder:
https://drive.google.com/drive/u/1/folders/16HTr8nfPz4P2PMr4QB0bjgwiD110Qd-0

The dashboard reads its published `zone-distribution` snapshot and the existing Zone Distribution dashboard data. It accepts the compact fields (`code`, `name`, `leader`, `zonal`) as well as the original Excel headers (`CODE`, `Outlet Name`, `Regional Head HR Name`, `Zonal HR Name`). Full HR names are used when available for the current assignment.

The mapping check now runs independently on the existing one-minute dashboard timer and when returning to the tab, including while dashboard filters are active. Once your existing Drive publisher publishes an updated Zone Distribution file, the next mapping check updates the outlet names, RHO, Zonal, relevant filter choices and cached management details. The newer published source takes precedence when the two published copies have different timestamps. The local hierarchy-only file provides a backup if both published sources are unavailable; it does not replace fresher remote assignments.

The Drive publisher must remain active. This patch does not introduce a new Drive upload trigger or change its schedule. Publication time plus the existing one-minute browser check determines when a Drive change appears; this is not a zero-delay upload webhook. Power BI snapshot generation, timing, date rules, measures and numeric calculations are unchanged.

## Checked against the current source

Workbook: **Zone Distribution Oct 2026 w location type.xlsx**. All **1,032** source outlet codes have RHO and Zonal assignments.

**F109 / Hazaribagh:**
- RHO: **Ahmed Yusuf**
- Zonal: **MD. Afzal Hossain**

Of the **1,056** outlet codes in the current receiving snapshot, **1,029** are present in this source. These **27** codes are absent from the current source workbook and therefore have no RHO / Zonal assignment to use:

`9000, D000, D009, D089, DK01, DK02, DK04, DK07, DK11, DK14, F013, F039, F112, F151, F412, F445, F515, F591, F618, F686, F691, F692, F704, F824, F841, F844, SY01`

Add their actual assignments to the Zone Distribution workbook if those codes should be mapped. They will be picked up through the same published-source refresh. Historical outlet search options may also contain codes that are absent from the current Zone Distribution workbook.

Verified: all 1,032 source joins, F109 full HR names, compact and original-header formats, updated manager assignments, newer published data, invalid / empty publication handling, remote-source failures, automatic refresh with filters and an open detail dialog, and cached rows for all six management tables. Receiving financial values and the snapshot time were unchanged during the mapping refresh check.
