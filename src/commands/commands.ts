/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

function Mayuscula_Click(event) {
  // 1. Ejecutas tu lógica en la hoja
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();

    // Modificar valores a mayúsculas
    const uppercaseValues = range.values.map(row => 
      row.map(cell => typeof cell === "string" ? cell.toUpperCase() : cell)
    );
    range.values = uppercaseValues;
    await context.sync();

    // 2. Abres un diálogo pequeñito de aviso
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/dialog.html',
        { height: 15, width: 20, displayInIframe: true },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const dialog = asyncResult.value;
                // Cerrar automáticamente después de 1.5 segundos
                setTimeout(() => dialog.close(), 1500);
            }
            event.completed();
        }
    );
  }).catch((error) => {
    console.error(error);
    event.completed();
  });
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("Mayuscula_Click", Mayuscula_Click);
