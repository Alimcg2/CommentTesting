
// The initialize function must be run each time a new page is loaded
(function () {
Office.initialize = function (reason) {
    $(document).ready(function () {
        $("#createNew").click(function (event) {
            createNew();
        });
        //
    });
};

// Reads data from current document selection and displays a notification
function createNew() {
    console.log("it worked!");
    $("#allComments")[0].classList.add("hidden");
    $("#createNew")[0].classList.add("hidden");
    $("#back")[0].classList.remove("hidden");
    $("#back")[0].onclick = backToAll;

    Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
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
    function backToAll(){
        $("#allComments")[0].classList.remove("hidden");
        $("#createNew")[0].classList.remove("hidden");
        $("#back")[0].classList.add("hidden");
    }

})();
