# CR-Projects — CR Workdesk v3 source

**Status:** Source-only, GitHub-ready. Not deployed. Microsoft Graph live synchronization is **not activated or verified**.

This repository contains **application code, synthetic test fixtures, and a ZERO-record safe bootstrap**. It intentionally excludes the real `Progress CR TSEL.xlsx`, all workbook-derived snapshots, any company records, credentials, and export files. It should be **PRIVATE** and limited to authorized collaborators.

## Layout

- `dist/` — static website, UI, data model and Microsoft Graph client. Run `npm run prepare:safe` to create `dist/data.js` from the zero-record template.
- `tools/build_data.py` — optional local-only snapshot creator. DO NOT commit its generated `dist/data.js`; the default deployment intentionally starts empty until the authorized user signs into Microsoft.
- `tests/` — pure synthetic model and mocked Graph tests; no company data needed.
- `dist/config.js` — fill in non-secret, approved SharePoint site details in your **private environment**. Never add secrets here. Browser receives public client ID only.

## Quick start

```bash
npm test
npm run serve
```

Then open `http://localhost:8080`. It initially displays **no data** because the corporate workbook is not shipped with source. Local drafts remain local to browser; live read/write requires an authorized Microsoft Entra SPA registration, delegated consent, a valid workbook path, and a protected HTTPS deployment.

## Existing production site

**Do not create an unprotected public deployment.** CR Workdesk remains on the current Sites-hosted URL. This repo is a source-control preparation and does not migrate, connect, or update the existing production site. The original site's version and permissions must be reviewed before publishing any replacement.

## Integration limitations

- OAuth 2.0 PKCE implemented, but **not tested against the company's live tenant**.
- Read refresh every ~60 seconds while app is open and session valid; no background scheduled syncing.
- Excel cell edits are individually written and verified after a pre-write conflict check. This is not atomic, and a race still exists between check and write.
- Manual reconciliation confirmation, manual CR year overrides and document URLs are **local-only annotations**, not guaranteed synchronized.
- Microsoft Graph may require administrator consent. No browser login by itself guarantees API entitlement.
- No Excel original files or site production changed by the creation of this code bundle.

## GitHub setup

Create an **empty PRIVATE** repository named `CR-Projects` on the authorized GitHub account. Do not choose Initialize with README. Extract `CR-Projects_GitHub_Ready.zip`, then run:

```bash
git init -b main
git add .
git commit -m "Bootstrap CR Workdesk v3 source without business data"
git remote add origin https://github.com/<OWNER>/CR-Projects.git
git push -u origin main
```

`.gitignore` explicitly excludes `dist/data.js`, `.xlsx`, `.csv`, exports, screenshots, environment settings and other data. Run `git status` and inspect staged changes before each push; validate corporate policy for source hosting even in private repos.

**Gate for completion:** browser QA, live Microsoft read/write QA, original-site feature-parity comparison, access review, release authorization, and deployment/rollback test must all pass.

## GitHub Pages (deployment requested 30 October 2026)

This is a **static** HTML/CSS/JavaScript application. GitHub Actions builds an **empty-data** Pages artifact every time `main` changes. The artifact contains neither workbook records nor credentials. The app must still obtain browser-side Microsoft Entra delegated Graph authorization to show live data.

- Repository: `https://github.com/rzqllh/cr-projects` (**private**).
- Intended Pages URL, only after Pages is enabled and a deployment succeeds: `https://rzqllh.github.io/cr-projects/`.
- In GitHub open **Settings → Pages → Build and deployment → Source: GitHub Actions**; if a private-repo Pages site is unavailable under the current GitHub plan, **do not make this repository public without a separate source-code/security review and express approval**.
- Entra admin must register an SPA redirect URI **exactly** `https://rzqllh.github.io/cr-projects/` with appropriate delegated permissions and consent; put its application ID and tenant identifier into the on-page connection dialog. No client secrets. Do not use a real corporate workbook for testing until authorization and security review are complete.
- Note: GitHub Pages hosting is generally publicly reachable even with a private repository. The app code/branding are publicly served; the live CR data is not bundled. Verify corporate policy before connecting production data to a third-party-hosted frontend.
- GitHub Pages does **not** provide background synchronization: 60-second refresh works only while the authenticated page stays open.

Release gates remain: authorized Microsoft Graph workbook read, controlled test write to the proper cell, concurrency handling, authenticated access controls, and production acceptance. The old Sites deployment is not modified.
