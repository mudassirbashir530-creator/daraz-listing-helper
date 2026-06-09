/**
 * LIFEINNO Daraz Helper - Background Service Worker (Manifest V3)
 */

// On extension installation, configure side panel defaults
chrome.runtime.onInstalled.addListener(() => {
  console.log('LIFEINNO Daraz Helper Extension successfully installed!');
  
  // Set default behaviors for side panel
  if (chrome.sidePanel && chrome.sidePanel.setPanelBehavior) {
    chrome.sidePanel.setPanelBehavior({
      openPanelOnActionClick: false // Click opens the popup by default, they can open sidebar from there
    }).catch(err => {
      console.warn('Could not set side panel behavior:', err);
    });
  }
});

// Listen for messages from popup or other scripts
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'open_sidebar') {
    const targetTabId = msg.tabId || (sender.tab ? sender.tab.id : null);
    
    if (targetTabId) {
      if (chrome.sidePanel && chrome.sidePanel.open) {
        chrome.sidePanel.open({ tabId: targetTabId })
          .then(() => {
            sendResponse({ success: true });
          })
          .catch((err) => {
            console.error('Error opening side panel:', err);
            sendResponse({ success: false, error: err.message });
          });
      } else {
        sendResponse({ success: false, error: 'chrome.sidePanel API not available' });
      }
    } else {
      sendResponse({ success: false, error: 'No active tab ID available' });
    }
    return true; // Keep message channel open for response
  }
});
