# TopLine Fieldglass Extractor

Built by Otraz Technologies for TopLine Middle East Pipelines Contracting L.L.C.

HR uploads an employee's documents. The app reads them with Claude, checks them against each other, lets HR review and correct the values, and writes the employee's row into the Fieldglass sheet.

## How it works

1. **Sign in.** The sign-in page uses the split-screen layout.
2. **Upload.** Drop files, a folder or a `.zip`. A zip or folder with one sub-folder per employee (e.g. `31.AYUB KHAN YASIN KHAN/`) queues each employee separately.
3. **Read.** Each file is prepared in the browser. PDFs with a text layer send their text, scans and photos go as compressed images, and short PDFs (ID cards, certificates) always include an image as well. Files go to `/api/extract` one at a time, 3 in parallel. Claude identifies the document type and returns its fields.
4. **Merge and check.** Each field comes from the most reliable document (see the table below). The app then checks that:
   - the Emirates ID, passport number, DOB and name match across documents
   - the Emirates ID format is valid
   - the age is plausible
   - H2S, Life-Saving and Medical certificates are not expired or expiring within 30 days
   - no required documents or fields are missing
5. **Review.** Every field shows which document it came from. Colours mark fields to check, problems and defaults. HR can correct any value or change a file's detected document type.
6. **Write.**
   - **Add to workbook** writes the employee into the open workbook. If the Emirates ID is already in the sheet, that row is updated; otherwise the employee goes into the first empty row.
   - **Load master sheet** lets HR work on their current Fieldglass file instead of the blank template.
   - **Download workbook** saves the result.
   - **Excel for this employee** downloads a copy of the template containing just that employee.

The template's formatting, column widths, filters and validation are kept.

## Field sources

| Column | Field | Source (priority order) |
|---|---|---|
| D / E | Name EN / AR | Emirates ID → Residence visa |
| G | Emirates ID | Emirates ID → Residence → Medical |
| I, J, L | Nationality, DOB, Gender | Emirates ID → Passport → Medical |
| K | Blood group | ADNOC medical (lab "Blood typing ABO" + Rh) |
| M | Place of issue (residence) | Residence visa |
| N | Passport no. | Passport → Residence → Medical |
| S | Job title | Residence visa → Emirates ID (occupation) |
| V | H2S validity | H2S certificate – Certificate Validity |
| W | ADNOC Life Saving validity | HSE Induction & Life Saving cert – Certificate Validity |
| X | ADNOC Medical validity | Medical fitness certificate – Expiry |
| Y | HSE Induction date | HSE Induction cert – Issue Date |
| Z | Certificate No. | HSE Induction cert – Certificate No. |
| B, C, F, H, O, P, T, U | Contract, company, email, marital status, education, phones | Defaults from the current sheet (HR confirms) |
| Q, R, AA–AE | Religion, sect, Fieldglass IDs, location, remarks | Entered manually (not on the documents) |

## Deploy on Netlify

1. Push this folder to a GitHub repo (or drag-and-drop the folder in Netlify → *Add new site → Deploy manually*).
2. Netlify reads `netlify.toml`, which publishes `public/` and sets up the functions in `netlify/functions/`. There is no build step.
3. In **Site configuration → Environment variables**, add:

| Variable | Value |
|---|---|
| `ANTHROPIC_API_KEY` | Your Claude API key (stays on the server, never sent to the browser) |
| `APP_USERS` | Logins, comma-separated: `hr@toplineme.com:StrongPass1,syed@otraz.ai:StrongPass2` |
| `AUTH_SECRET` | Any long random string (signs the 12-hour session tokens) |
| `ANTHROPIC_MODEL` | Optional. Defaults to `claude-sonnet-4-5` |

4. Redeploy after adding the variables.

**Limits.** Netlify accepts request bodies up to about 6 MB. The browser keeps each file under about 4.3 MB by compressing images and capping each PDF at 8 rendered pages.

## Files

```
netlify.toml
netlify/functions/auth.mjs     sign-in → signed session token
netlify/functions/extract.mjs  one document → Claude (structured tool output)
netlify/lib/auth-utils.mjs     token sign/verify
public/index.html, styles.css, app.js   UI
public/merge.js                merge rules + checks (also runs in Node for tests)
public/excel.js                writes the row into the workbook (ExcelJS)
public/template.xlsx           blank Fieldglass template (from "New for Fieldglass.xlsx")
public/vendor/                 pdf.js, JSZip, ExcelJS (self-hosted)
```

## Local test

```
npm i -g netlify-cli
netlify dev        # needs the env vars above in a .env file
```
