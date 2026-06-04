/* global Office */

Office.onReady(() => {
  // Function commands can be registered here in future
});

/**
 * DocumentOpen event handler — auto-shows the task pane when a document
 * containing the embedded add-in reference is opened.
 * Registered via the manifest VersionOverridesV1_1 Events extension point.
 */
function onDocumentOpen(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

// Must be attached to the global scope for the manifest to call it
(globalThis as Record<string, unknown>)["onDocumentOpen"] = onDocumentOpen;
