# Proposed next improvements

Prioritized follow-up proposals; these are not implemented in this release.

1. Resolve page order, titles and dimensions from VSDX relationships and pages.xml; verify with reordered and 10+ page fixtures.
2. Replace the service worker's single pending-file slot with per-tab ownership; test two local files opened simultaneously and worker restarts.
3. Offer SVG export and a page picker with thumbnails; keep all conversion local.
4. Add a real Chrome extension test for navigation, signed download URLs and file access; retain downloads when automatic viewing fails.
