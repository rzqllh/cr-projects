# Security notes

- PRIVATE repo only; review corporate source-control policy before using GitHub.
- No actual workbook bytes, customer or business records, screenshots, CSV exports or personal access tokens in Git.
- Authorization-code + PKCE requires an approved SPA redirect URI and Microsoft Graph delegated permissions; never use a client secret in a frontend app.
- This app is not an enterprise security boundary: for audit logging and atomic writes, implement a server-side IT-approved API.
- Do not publish an empty-source development repo to a public URL with corporate workbook data.
