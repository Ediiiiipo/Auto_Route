// SPX Auto Router — Background Service Worker
// Abre popup como janela independente ao clicar no ícone
chrome.action.onClicked.addListener(() => {
  chrome.windows.create({
    url: chrome.runtime.getURL('popup.html'),
    type: 'popup',
    width: 560,
    height: 700,
    focused: true
  });
});

// Relay de mensagens para o content-script
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.target === 'background') {
    chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
      if (tab) chrome.tabs.sendMessage(tab.id, msg, sendResponse);
    });
    return true;
  }
});
