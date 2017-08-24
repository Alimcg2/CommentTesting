// The initialize function must be run each time a new page is loaded
(function () {
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
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#individualView")[0].classList.remove("hidden");
    }
    function createNew() {
        $("#allComments")[0].classList.add("hidden");
        $("#createNew")[0].classList.add("hidden");
        $("#back")[0].classList.remove("hidden");
        $("#back")[0].onclick = backToAll;
        $("#newView")[0].classList.remove("hidden");
        getText;
    }
    function backToAll() {
        $("#allComments")[0].classList.remove("hidden");
        $("#createNew")[0].classList.remove("hidden");
        $("#back")[0].classList.add("hidden");
        $("#individualView")[0].classList.add("hidden");
        $("#newView")[0].classList.add("hidden");
    }

    function getText() {
        console.log("It ran");
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
                }
            });
    }

})();