// SPX Auto Router — Background Service Worker (relay)
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.target === 'background') {
    // Relay to content script
    chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
      if (tab) chrome.tabs.sendMessage(tab.id, msg, sendResponse);
    });
    return true;
  }
});
