
// The initialize function must be run each time a new page is loaded
(function () {
Office.initialize = function (reason) {
    $(document).ready(function () {
        $("#writeTextButton").click(function (event) {
            writeText(text);
        });
        //
    });
};

// Reads data from current document selection and displays a notification
function writeText() {
    Office.context.document.setSelectedDataAsync("Something",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                //show error. Upcoming displayDialog API will help here.
            }
            else {
                //show success.Upcoming displayDialog API will help here.
            }
        });
}

})();
