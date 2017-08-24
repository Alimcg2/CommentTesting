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
        });
    };
    // Reads data from current document selection and displays a notification
    function clickMe() {
        // Office.context.document.setSelectedDataAsync("testing", 
        //     function (asyncResult) {
        //         var error = asyncResult.error;
        //         if (asyncResult.status === "failed") {
        //             //show error. Upcoming displayDialog API will help here.
        //         }
        //         else {
        //             //show success.Upcoming displayDialog API will help here.
        //         }
        //     });
        bindNamedItem();
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#individualView")[0].classList.remove("hidden");
    }

    function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myBinding'}, function (result) {
        if (result.status == 'succeeded'){
            console.log('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            console.log('Error: ' + result.error.message);
    });

    Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fillColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "red", fontStyle: "bold"}}], 
    function (asyncResult){});
}

    function createNew() {
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#newView")[0].classList.remove("hidden");
        console.log(document.getElementById("updateCell"));
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

})();