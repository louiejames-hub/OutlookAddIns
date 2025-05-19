Office.onReady(() => {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, checkRecipients);
});

function checkRecipients(eventArgs) {
    let item = Office.context.mailbox.item;
    item.getRecipientsAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            let recipients = asyncResult.value;
            let organizationDomain = "yourcompany.com"; // Change this to your actual domain

            let externalRecipients = recipients.filter(recipient => 
                !recipient.emailAddress.endsWith("@" + organizationDomain)
            );

            if (externalRecipients.length > 0) {
                let confirmation = confirm("Some recipients are outside your organization. Do you still want to send?");
                if (!confirmation) {
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            }
        }
        eventArgs.completed({ allowEvent: true });
    });
}
