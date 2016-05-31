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

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails()
    {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var from = Office.cast.item.toMessageRead(item).from;
        var email = from.emailAddress;

        var name = from.displayName.substr(0, from.displayName.indexOf(' '));
        $("#name").val(name);

        var surname = from.displayName.substr(from.displayName.indexOf(' ') + 1);
        $("#surname").val(surname);

        var domain = email.substr(email.indexOf('@') + 1);
        var company = domain.substr(0, domain.indexOf('.'));
        $("#company").val(company);
    }
})();