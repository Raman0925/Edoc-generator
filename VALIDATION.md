# Validation record — 7 October 2026

Completed in the development environment:

- Python compilation and import of the application modules.
- 43 portable automated tests passed, using the build's pinned pypdf 6.18.1.
- The opt-in real Excel smoke test was skipped because this environment is Linux and has no desktop Microsoft Excel.
- Real PDF merging was exercised: approval pages precede certificate pages, output page counts match, and encrypted/corrupt PDFs are rejected.
- Five-report batches, original-file byte preservation, missing-field failures, optional blanks, zero values, non-finite numeric rejection, saved profile round trips, duplicate destinations, merged-cell collisions, cancellation, filename sanitization and same-name reports with different extensions were exercised.
- `git diff --cached --check` passed.

Not completed here:

- Visual/functional operation of the GUI on Windows.
- Excel COM operation or PDF-layout comparison with the user's company template; no sample template/report has been supplied.
- Creation or verification of a Windows `.exe`. The included local build script and GitHub Actions workflow provide the build path.
- The company-template PDF comparison and real Office test remain pending. Source publication and build execution are now handled through the signed-in GitHub browser; consult the repository's Actions tab for the final Windows build and packaged-app startup-check result.

Before using generated documents for company approval, run the real Excel smoke test and compare one generated document with a known manually prepared result from the actual report/template. Confirm all mapped fields, units, dates, pagination and certificate pairing. The application prepares documents; it does not validate calibration measurements or grant approval.
