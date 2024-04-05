let mailboxItem;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    console.log("Office.initialize");
    mailboxItem = Office.context.mailbox.item;
}

function validateMessage(event)
{
    console.log("validateMessage");

    event.completed({ allowEvent: false });
    return;
}
