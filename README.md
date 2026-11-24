# ResumeForge — Free ATS Resume Builder

First working version for freshers and experienced professionals.

## Included
- Personal details, summary, work experience, education, skills, projects and certifications.
- Live single-column resume preview.
- Rule-based job-description keyword comparison and section checklist.
- Editable Word (`.docx`) and PDF downloads.
- Automatic save in the current browser using localStorage.
- Responsive layout for desktop and mobile.
- No login, backend, database or paid AI API required.

## Important ATS-score note
The displayed score is a transparent heuristic based on keyword overlap and presence of resume sections. It is **not** a score from an actual ATS vendor, cannot predict hiring outcomes, and cannot guarantee that a resume will pass an employer's screening system. Users should only include skills and experience they genuinely have.

## Requirements
- Windows 10/11
- Node.js 20 LTS or newer (npm is included)
- Internet connection for first dependency installation

## Run locally on Windows
1. Install Node.js LTS from https://nodejs.org/.
2. Extract this ZIP to a folder, for example `C:\Projects\ResumeForge`.
3. Open the folder in VS Code.
4. Select **Terminal → New Terminal**.
5. Run:

   ```bash
   npm install
   npm run dev
   ```

6. Open the local URL printed by Vite, usually `http://localhost:5173`.

## Production build

```bash
npm run build
npm run preview
```

Production files are created in `dist`.

## Free deployment with Cloudflare Pages
1. Create a GitHub account and a repository named `resumeforge`.
2. Upload the extracted project files to the repository. Do not upload `node_modules`.
3. Sign in to Cloudflare and open **Workers & Pages → Create → Pages → Connect to Git**.
4. Select the `resumeforge` repository.
5. Set framework preset to `Vite`, build command to `npm run build`, and build output directory to `dist`.
6. Deploy. Cloudflare provides a public `*.pages.dev` address.

Check current hosting plan limits and terms before public launch. A custom domain is optional and may cost money.

## Privacy
Resume content is stored in the user's browser localStorage in this version; it is not uploaded to a ResumeForge server. Users should avoid shared/public computers and clear browser data when appropriate.

## Current limitations / next version
- Keyword matching is simple; it does not understand context, synonyms, skill equivalence or seniority.
- PDF export is text-based; verify page breaks before applying.
- No account, cloud sync, import-existing-resume parser, AI rewrite, template gallery or admin analytics yet.
- Add automated tests, accessibility review, security review and cross-browser validation before broad promotion.
