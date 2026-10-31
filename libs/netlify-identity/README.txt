netlify-identity-widget 1.9.2 (MIT, Netlify) - vendored 2026-10-31 from the npm package
(build/netlify-identity-widget.js, the same 1.x line the identity.netlify.com/v1 URL serves).
Loaded only on *.netlify.app / *.netlify.com hosts by js/identity.js; never on file:// or GitHub Pages.
SHA-256 466954df58db08e3c508d1442f9b986738f2fafa52fa28a58cdfa855fa5f54c1 (npm original, LF line endings; the repo checks it out with CRLF like every other lib)
Why: the app no longer loads a third-party script at runtime (audit 2026-10-31 M4) and can run
under a strict Content-Security-Policy (script-src 'self').
