import { isEmptyObject } from "jquery";

//Needed for button to work
$("#Setup").on("click", () => tryCatch(run));
$("#register-handler").on("click", () => tryCatch(registerChangeEventHandler));

async function registerChangeEventHandler() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        //OnColumnChanged
        sheet.onChanged.add(onColumnChanged);
        //onSheetChanged
        //sheet.onChanged.add(onSheetChanged);
        await context.sync();
        console.log("Event handler registered for sheet.");
});
}


//Colunm Function
async function onColumnChanged(eventArgs: Excel.WorksheetChangedEventArgs) {
    await Excel.run(async (context) => {
        const details = eventArgs.details;
        const address = eventArgs.address;

        Office.context.ui.displayDialogAsync('http://127.0.0.1:3000/Clearfont/dialog.html', 
                    { height: 30, width: 20 });
        //Debug: console.log(`Address: ${address}`);

        //Remove the numbers from the address
        var column = address.replace(/[0-9]/g, '');
        console.log(`Column: ${column}`);

        if(column == 'B') {
            console.log(`Column B changed`);
            var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
            if(details.valueAfter == "RR") {
                console.log(`Value after: ${details.valueAfter}`);
                range.hyperlink = {address: "https://www.royalroad.com", textToDisplay: "RR"};
            } else if(details.valueAfter == "WN") {
                console.log(`Value after: ${details.valueAfter}`);
                range.hyperlink = {address: "https://www.webnovoel.com", textToDisplay: "WN"};
            } else if(details.valueAfter == "Ranobes") {
                console.log(`Value after: ${details.valueAfter}`);
                range.hyperlink = {address: "https://www.Ranobes.top", textToDisplay: "Ranobes"};
            }else if (details.valueAfter!== "RR" ||
                details.valueAfter!== "WN" || details.valueAfter!== "Ranobes") {
                    console.log(`Value after: ${details.valueAfter}`);
                    let dialog
                    Office.context.ui.displayDialogAsync('http://127.0.0.1:3000/Clearfont/dialog.html',
                    { height: 20, width: 30, displayInIframe: true },
                    function (asyncResult) {
                        dialog = asyncResult.value;
                        // callbacks from the parent
                        dialog.addEventHandler(Office.EventType.DialogEventReceived);
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived);
                    });
            }
            
        
            

        } 
    
    });
}        
//Sheet function
async function onSheetChanged(eventArgs: Excel.WorksheetChangedEventArgs) {
    await Excel.run(async (context) => {
    const details = eventArgs.details;
    const address = eventArgs.address;

    console.log(
        `Change at: ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),` +
            `now is ${details.valueAfter}(${details.valueTypeAfter})`
    );
    if (details.valueBefore != null && details.valueAfter === "") {
        console.log(
            `Value before: ${details.valueBefore}` + `"Value after: ${details.valueAfter}`
            
        );
        var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
        
        range.format.fill.clear();
        range.format.font.color = "black";
    
        console.log(`Cleared the format at: ${address}`);
    }
    });
}

function run() {
    return Excel.run(function (context) {
     //   var range = context.workbook.getSelectedRange();
     //   range.format.fill.clear();
     //   
     //   range.load("address");

     Office.context.ui.displayDialogAsync('http://127.0.0.1:3000/Clearfont/dialog.html', {height: 30, width: 20}, function (result) {
        // Handle the dialog result here
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Dialog closed successfully");
        } else {
            console.error("Dialog failed: " + result.error.message);
        }
    });

        return context.sync().then(function () {
            //` for string interpolation
           //console.log(`Selcected range is now cleared. Range was: "${range.address}" ` );
           
            console.log("Hello, world!");
        });
    }).catch(function (error) {
        console.log(error);
    });    

}

// tryCatch function to handle errors
function tryCatch(callback) {
    Promise.resolve()
    .then(callback)
    .catch(function (error) {
         // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    });
}

