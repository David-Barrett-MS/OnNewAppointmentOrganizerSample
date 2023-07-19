Office.onReady(function (info) {
    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
    UpdateTaskPaneUI(Office.context.mailbox.item);
    document.getElementById("app-body").style.display = "block";
    console.log('Finished initialisation');
});

function itemChanged(eventArgs) {
    // Update UI based on the new current item
    UpdateTaskPaneUI(Office.context.mailbox.item);
}

// Example implementation
function UpdateTaskPaneUI(item) {
    // Assuming that item is always a read item (instead of a compose item).
    if (item == null) {
        console.log('Item is null, can\'t read subject');
        return;
    }

    console.log(item.subject);
    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

