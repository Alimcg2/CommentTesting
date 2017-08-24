// The initialize function must be run each time a new page is loaded
(function () {
    var currentCellText;
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("#clickMe").click(function (event) {
                clickMe();
            });

            $("#createNew").click(function (event) {
                createNew();
            });
            $("#inactiveConvos").click(function (event) {
                showInactive();
            });
            $("#cancelButton").click(function (event) {
                bindNamedItem({row: 4, column: 3}, "#2f4260");
            });

        });
    };
    // Reads data from current document selection and displays a notification
    function clickMe() {
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#individualView")[0].classList.remove("hidden");
    }

    function bindNamedItem(row, color) {
        Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myBinding'}, function (result) {
            if (result.status == 'succeeded'){
                console.log('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
                }
            else
                console.log('Error: ' + result.error.message);
        });

        Office.select("bindings#myBinding").setFormatsAsync(
        [
            {cells: row, format: {fontColor: color}}], 
        function (asyncResult){});
        getDataWithContext();
    }

    function getDataWithContext() {
        var format = "Your data: ";
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
    }

     function showDataWithContext(asyncResult) {
        console.log(asyncResult.value);
        console.log(asyncResult.asyncContext);
    }

    function createNew() {
        bindNamedItem({row: 3, column: 3}, "#dd9b4b");
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#newView")[0].classList.remove("hidden");
        document.getElementById("updateCell").innerHTML = getText();
    }
    function backToAll() {
        $("#allComments")[0].classList.remove("hidden");
        $("#createNew")[0].classList.remove("hidden");
        $("#back")[0].classList.add("hidden");
        $("#individualView")[0].classList.add("hidden");
        $("#newView")[0].classList.add("hidden");
    }

    function getText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            { valueFormat: "unformatted", filterType: "all" },
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(error.name + ": " + error.message);
                }
                else {
                    // Get selected data.
                    var dataValue = asyncResult.value;
                    console.log('Selected data is ' + dataValue);
                    return dataValue;
                }
            });
    }
    function showInactive(){
        $("#inactive1")[0].classList.remove("hidden");
        $("#inactive2")[0].classList.remove("hidden");
        $("#inactiveConvos")[0].onclick = hideInactive;
    }
    function hideInactive(){
        $("#inactive1")[0].classList.add("hidden");
        $("#inactive2")[0].classList.add("hidden");
    }

})();