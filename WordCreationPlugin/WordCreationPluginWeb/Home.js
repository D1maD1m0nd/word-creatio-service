
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");

                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");
            $("#insert-table-button").text("Insert Table");
            $("#insert-paragraph-button").text("Insert Paragraph");
            $("#send-request-button").text("Send simple request");
            $("#auth-button").text("Auth");
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $("#insert-table-button").click(insertTable);
            $("#insert-paragraph-button").click(insertParagraph);
            $("#auth-button").click(authCreatio);
            $("#send-request-button").click(getDocumentAsCompressed);
        });
    };
    function insertTable() {
        Word.run((context) => {

            const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

            const tableData = [
                ["Name", "ID", "Birth City"],
                ["Bob", "434", "Chicago"],
                ["Sue", "719", "Havana"],
            ];
            secondParagraph.insertTable(3, 3, "After", tableData);
            return context.sync();
        })
            .catch(errorHandler);
    }

    function insertParagraph() {
        Word.run((context) => {

            const docBody = context.document.body;
            docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                "Start");
            return context.sync();
        }).catch(errorHandler);
    }
    // The following example gets the document in Office Open XML ("compressed") format in 65536 bytes (64 KB) slices.
    // Note: The implementation of app.showNotification in this example is from the Visual Studio template for Office Add-ins.
    function getDocumentAsCompressed() {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 100000 },
            function (result) {
                if (result.status == "succeeded") {

                    // If the getFileAsync call succeeded, then
                    // result.value will return a valid File Object.
                    const myFile = result.value;
                    const sliceCount = myFile.sliceCount;
                    const docdataSlices = [];
                    let slicesReceived = 0, gotAllSlices = true;

                    //app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
                    showNotification("Request Info", "File size:" + myFile.size + " #Slices: " + sliceCount);
                    // Get the file slices.
                    getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                }
                else {
                    showNotification("Error:", result.error.message);
                }
            });
    }

    function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status == "succeeded") {
                if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                    return;
                }

                // Got one slice, store it in a temporary array.
                // (Or you can do something else, such as
                // send it to a third-party server.)
                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                if (++slicesReceived == sliceCount) {
                    // All slices have been received.
                    file.closeAsync();
                    onGotAllSlices(docdataSlices);
                }
                else {
                    getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                }
            }
            else {
                gotAllSlices = false;
                file.closeAsync();
                errorHandler("getSliceAsync Error:" + sliceResult.error.message)
            }
        });
    }

    function onGotAllSlices(docdataSlices) {
        let docdata = [];
        for (let i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);
        }

        //let fileContent = new String();
        //for (let j = 0; j < docdata.length; j++) {
       //     fileContent += String.fromCharCode(docdata[j]);
       // }
        sendFileData(docdata);
        // Now all the file content is stored in 'fileContent' variable,
        // you can do something with it, such as print, fax...
    }

    async function authCreatio() {
        try {
            const body = JSON.stringify({
                "UserName": "Supervisor",
                "UserPassword": "Y{YE3~oMpPr*"
            });

            const requestData = {
                method: "post",
                headers: {
                    'Content-Type': 'application/json'
                    },
                url: "https://fotodom.site:488/ServiceModel/AuthService.svc/Login",
                data: body
            };
            const response = await axios(requestData);
            showNotification("Response" + response);
        } catch (e) {
            errorHandler('Ошибка ' + e.name + ":" + e.message + "\n" + e.stack)
        }
    }

    async function sendFileDataV2(docdata) {
        try {
            const fileName = 'test.docx';
            const file = new Blob(
                [new Uint8Array(docdata)],
                fileName,
                { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
            );
           

            const params = new URLSearchParams(getFileApiConfig(file, fileName))
            const headers = getFileApiHeader(file, fileName)
            const requestData = {
                method: "post",
                params: params,
                headers: headers,
                //url: "http://bakdev.lexiasoft.ru/0/rest/FileApiService/UploadFile",
                url: "https://fotodom.site:488/0/rest/FileApiService/UploadFile",
                data: file
            }
            const response = await axios(requestData);
            showNotification("Response" + response);
        } catch (e) {
            errorHandler('Ошибка ' + e.name + ":" + e.message + "\n" + e.stack)
        }
    }

    function getFileApiConfig(file, fileName) {
        return {
            "fileapi16721259223571": null,
            "totalFileLength": file.size,
            "fileId": "c58cbd8d-015c-4597-a43e-00a313eb3033",
            "mimeType": 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            "columnName": "Data",
            "fileName": fileName,
            "parentColumnName": "Contact",
            "parentColumnValue": "172226d9-5afc-4f53-a6cf-03ff781d5753",
            "entitySchemaName": "ContactFile"
        }
    }

    function getFileApiHeader(file, fileName) {
        return {
            "Content-Disposition": `attachment; filename=${fileName}`,
            "Content-Length": file.size,
            "BPMCSRF": "IY8IglJC93oTo8n1QOryCu",
            "Content-Type": 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            "Content-Range": `bytes 0-${file.size}/${file.size + 1}`,
            "fileName": fileName
        }
    }

    async function sendFileData(docdata) {
        try {
            const aFile = new Blob(
                [new Uint8Array(docdata)],
                { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
            );
            const formData = new FormData();
            formData.append('file', aFile, 'testfile.docx');

            const response = await axios({
                port:5000,
                method: 'post',
                url: 'http://127.0.0.1:5000/file',
                data: formData,
                headers: {
                    'Content-Type': `multipart/form-data; boundary=${formData._boundary}`,
                },
            });

            showNotification("Reponse", response);
        } catch (e) {
            errorHandler('Ошибка ' + e.name + ":" + e.message + "\n" + e.stack)
        }
    }


   
    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // This variable will keep the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
