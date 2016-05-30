/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            displayItemDetails();
        });
    };

    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var from;
        from = Office.cast.item.toMessageRead(item).from;
        //var test = from.displayName;
        //var name = test.substr(0, test.indexOf(' '));
        //var name = test;
        //var company = name;
        //$("#company").attr("value", company);
        //$("#name").attr("value", name);
        //var surname = test.substr(test.indexOf(' ') + 1);
        //$("#surname").attr("value", surname);
        
          //if (from) {
         //$('#from').text(from.displayName);
         //$('#from').click(function () {
             // app.showNotification(from.displayName, from.emailAddress);
            //});
        //}
    }
})();