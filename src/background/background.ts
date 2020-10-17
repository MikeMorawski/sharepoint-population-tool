var bg_context = {
    spLegacyCtx: null,
};

chrome.tabs.onUpdated.addListener(function (tabId, changeInfo, tab) {
    chrome.pageAction.show(tabId);
    if (changeInfo.status === "complete") {
        chrome.tabs.executeScript(null, { file: "contextloader.js" }, (callback) => { });
    }
});

chrome.pageAction.onClicked.addListener(function (activeTab) {
    chrome.tabs.executeScript(null, { file: "contextloader.js" }, (callback) => {
        chrome.tabs.executeScript(null, { file: "contentscript.js" });
        chrome.tabs.insertCSS({
            file: "popup/css/main.css",
        });
        chrome.tabs.executeScript(null, { file: "popup/js/bundle.js" }, () => {
            chrome.tabs.sendMessage(activeTab.id, bg_context);
        });
    });
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (sender.id == chrome.runtime.id) {
        const ctx = msg._spPageContextInfo as _spPageContextInfo;
        if (ctx && ctx.pageItemId == -1 && ctx.pageListId && ctx.pageListId.length > 0) {
            chrome.pageAction.setIcon({
                tabId: sender.tab.id,
                path: "popup/page_32.png"
            }, () => { })
            bg_context.spLegacyCtx = ctx;
        }
    }
})

// if(typeof moduleLoaderPromise !== 'undefined'){
//     (moduleLoaderPromise as any).then((context)=>{});
// }