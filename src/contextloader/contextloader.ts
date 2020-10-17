var currentContext = null;

addEventListener(chrome.runtime.id + '-config', function (e: CustomEvent) {
    chrome.runtime.sendMessage({ _spPageContextInfo: e.detail });
}, false);

runInPage(emitConfigExtractor);

function runInPage(fn) {
    const script = document.createElement('script');
    document.head.appendChild(script).text = '(' + fn + ')("' + chrome.runtime.id + '")';
    script.remove();
}

//Sends context to background.js script
function emitConfigExtractor(extensionId, loop) {
    var loopinfo = loop ? loop : 0;

    if (typeof _spPageContextInfo === 'undefined') {
        if (loopinfo == 25) {
            console.log("Context not found, stopping...");
        } else {
            setTimeout(() => {
                emitConfigExtractor(extensionId, loopinfo + 1);
            }, 250);
        }
    } else {
        console.log("Context found: " + _spPageContextInfo.listTitle);
        window.dispatchEvent(new CustomEvent(extensionId + '-config', {
            detail: _spPageContextInfo,
        }));
    }

}