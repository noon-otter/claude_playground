/* global Office */

Office.onReady(() => {
  // Commands will be registered here
});

function action(event: Office.AddinCommands.Event) {
  // Handle command actions
  event.completed();
}

// Register functions
(Office as any).actions = {
  action: action
};
