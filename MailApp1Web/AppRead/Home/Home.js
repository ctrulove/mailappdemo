/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
			item.body.getAsync("html", function (result){
				console.log("result.value",result);
				var reKQT = /KQT: -\d{10}/;
				var reKI = /KI: \d{7}/;
				var KQT = result.value.match(reKQT)[0];
				var KI = result.value.match(reKI)[0];
				if (KQT) {
					KQT = KQT.slice(5);
					$('#kqt').text(KQT);
				}
				if (KI) {
					KI = KI.slice(4);
					$('#ki').text(KI);
				}
				
				console.log("mathes",result.value.match(re));
			});
            console.log("Office.cast.item.toMessageRead(item)", Office.cast.item.toMessageRead(item));
			console.log("item", item);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }
})();