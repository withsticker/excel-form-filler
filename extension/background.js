// Open the side panel when the toolbar icon is clicked.
// Side panel stays open until the user closes it (X icon) — clicking outside does not dismiss it.
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((err) => console.error(err));
