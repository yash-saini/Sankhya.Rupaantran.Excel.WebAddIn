var dotNetHelper = null;
console.log("googleSheetsInterop.js loaded.");

window.googleSheetInterop = {
    registerDotNetHelper: function (helper) {
        dotNetHelper = helper;
        console.log("DotNet helper registered.");
    },

    // C# calls this AFTER converting
    setRange: function (a1Notation, values) {
        console.log("setRange called from C#. Posting 'setRange' to parent.");
        window.parent.postMessage({
            action: 'setRange',
            a1Notation: a1Notation, // Pass the address
            values: values
        }, "*");
    },

    // C# calls this right BEFORE converting
    getRange: function () {
        console.log("getRange called from C#. Posting 'getRange' to parent.");
        window.parent.postMessage({ action: 'getRange' }, "*");
    }
};

// Listen for messages FROM the parent (Sidebar.html)
window.addEventListener("message", (event) => {
    const data = event.data;

    // Only listen for messages from the sidebar itself
    if (event.origin.includes("googleusercontent.com")) {
        console.log("googleSheetsInterop.js received message:", data);

        if (data.action === 'rangeDataReceived') {
            if (dotNetHelper) {
                console.log("Received 'rangeDataReceived'. Invoking C# 'ProcessValues'");
                dotNetHelper.invokeMethodAsync('ProcessValues', data.data);
            }
        }
        else if (data.action === 'setRangeSuccess') {
            if (dotNetHelper) {
                console.log("Received 'setRangeSuccess'. Invoking C# 'OnConversionComplete'");
                dotNetHelper.invokeMethodAsync('OnConversionComplete');
            }
        }
        else if (data.action === 'scriptError') {
            if (dotNetHelper) {
                console.log("Received 'scriptError'. Invoking C# 'OnScriptError'");
                dotNetHelper.invokeMethodAsync('OnScriptError', data.error);
            }
        }
    }
});