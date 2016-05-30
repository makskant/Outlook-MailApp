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
        var email = from.emailAddress;
        var test = from.displayName;
        var name = test.substr(0, test.indexOf(' '));
        $("#name").attr("value", name);
        var surname = test.substr(test.indexOf(' ') + 1);
        $("#surname").attr("value", surname);
        var domain = email.substr(email.indexOf('@') + 1);
        var company = domain.substr(0, domain.indexOf('.'));
        $("#company").attr("value", company);
    }
})();