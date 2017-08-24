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
         Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "red", fontStyle: "bold"}}], 
    function (asyncResult){});
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#individualView")[0].classList.remove("hidden");
        var mySelection = context.document.getSelection().parentTableCell;
            context.load(mySelection);
            console.log(context.sync());
            return context.sync()
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