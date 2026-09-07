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

export async function Mayuscula_Click(event: Office.AddinCommands.Event) {
  try{
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const uppercaseValues = range.values.map(row => 
        row.map(cell => typeof cell === "string" ? cell.toUpperCase() : cell)
      );

      range.values = uppercaseValues;
      await context.sync();
    });
  } catch (error) {
    console.error("Error al ejecutar Mayuscula_click",error);
  };

  event.completed();
}

export async function Minuscula_Click(event: Office.AddinCommands.Event) {
  try{
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const uppercaseValues = range.values.map(row => 
        row.map(cell => typeof cell === "string" ? cell.toLowerCase() : cell)
      );

      range.values = uppercaseValues;
      await context.sync();
    });
  } catch (error) {
    console.error("Error al ejecutar Minuscula_Click",error);
  };

  event.completed();
}

export async function Oracion_Click(event: Office.AddinCommands.Event): Promise<void> {
  try{
    await Excel.run(async (context:any) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const uppercaseValues = range.values.map((row : any[]) => 
        row.map((cellValue: any) => {
          if (typeof cellValue === "string" && cellValue.trim() !== "") {
            const texto = cellValue.toLowerCase();

            return texto.charAt(0).toUpperCase() + texto.slice(1)
          }
          return cellValue;
        })
      );

      range.values = uppercaseValues;
      await context.sync();
    });
  } catch (error) {
    console.error("Error al ejecutar Oracion_Click",error);
  };

  event.completed();
}

export async function Primera_Click(event: Office.AddinCommands.Event): Promise<void> {
  try{
    await Excel.run(async (context:any) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const uppercaseValues = range.values.map((row : any[]) => 
        row.map((cellValue: any) => {
          if (typeof cellValue === "string" && cellValue.trim() !== "") {

            return cellValue
              .toLowerCase()
              .replace(/\b\w/g, (char) => char.toUpperCase());
          }
          return cellValue;
        })
      );

      range.values = uppercaseValues;
      await context.sync();
    });
  } catch (error) {
    console.error("Error al ejecutar Oracion_Click",error);
  };

  event.completed();
}

export async function Agruparceldas_Click(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      // 1. Obtener la selección actual
      const selection = context.workbook.getSelectedRange();
      selection.load(["cellCount", "values", "rowCount", "columnCount"]);
      await context.sync();

      // Validación de mínimo 2 celdas
      if (selection.cellCount < 2) {
        // En Office.js la notificación nativa no bloqueante es recomendada o se puede usar un diálogo
        console.warn("Selecciona al menos 2 celdas");
        return;
      }

      // Procesamos por cada columna dentro de la selección
      for (let col = 0; col < selection.columnCount; col++) {
        let filaInicio = 0;
        const rowEnd = selection.rowCount - 1;

        while (filaInicio <= rowEnd) {
          const valorInicio = selection.values[filaInicio][col];
          let filaFin = filaInicio + 1;

          // Buscar hasta dónde se repite el valor consecutivo
          while (filaFin <= rowEnd) {
            const valorActual = selection.values[filaFin][col];

            if (
              (valorInicio === null && valorActual === null) ||
              (valorInicio !== null && valorInicio === valorActual)
            ) {
              filaFin++;
            } else {
              break;
            }
          }

          // Si hay 2 o más celdas consecutivas con el mismo valor, las combinamos
          if (filaFin - filaInicio > 1) {
            const rangoCombinar = selection.getCell(filaInicio, col).getAbsoluteResizedRange(
              filaFin - filaInicio,
              1
            );

            // Se borran los valores de las celdas inferiores para evitar la advertencia de Excel
            for (let i = filaInicio + 1; i < filaFin; i++) {
              selection.getCell(i, col).clear();
            }

            // Combinar y centrar contenido
            rangoCombinar.merge(true); // true = combina de forma independiente por fila/columna
            rangoCombinar.format.horizontalAlignment = Excel.HorizontalAlignment.center;
            rangoCombinar.format.verticalAlignment = Excel.VerticalAlignment.center;
          }

          filaInicio = filaFin;
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error al agrupar celdas:", error);
  }

  // Notificar fin del comando a Office
  event.completed();
}

export async function pdf_Click(event: Office.AddinCommands.Event): Promise<void> {
  try {
    // 1. Solicitar a Office que genere el documento completo en formato PDF
    Office.context.document.getFileAsync(
      Office.FileType.Pdf,
      { sliceSize: 65536 }, // Tamaño de fragmento: 64KB
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const file = result.value;
          const sliceCount = file.sliceCount;
          const slicesData: Uint8Array[] = [];
          let slicesReceived = 0;

          // Función recursiva para obtener todos los fragmentos del archivo
          const getSlice = (index: number) => {
            file.getSliceAsync(index, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                slicesData.push(new Uint8Array(sliceResult.value.data));
                slicesReceived++;

                if (slicesReceived === sliceCount) {
                  // 2. Cerrar el archivo en la API de Office
                  file.closeAsync();

                  // 3. Unir los fragmentos y disparar la descarga en el navegador
                  const blob = new Blob(slicesData, { type: "application/pdf" });
                  const downloadUrl = URL.createObjectURL(blob);
                  
                  const link = document.createElement("a");
                  link.href = downloadUrl;
                  link.download = "Documento.pdf";
                  document.body.appendChild(link);
                  link.click();
                  document.body.removeChild(link);
                  URL.revokeObjectURL(downloadUrl);
                } else {
                  getSlice(slicesReceived);
                }
              } else {
                file.closeAsync();
                console.error("Error al obtener fragmento del PDF:", sliceResult.error.message);
              }
            });
          };

          // Iniciar la lectura desde el primer fragmento (índice 0)
          getSlice(0);
        } else {
          console.error("Error al generar PDF:", result.error.message);
        }
      }
    );
  } catch (error) {
    console.error("Error inesperado al exportar PDF:", error);
  }

  // Notificar a Office que el comando finalizó
  event.completed();
}

async function guardarRutaEnLibro(ruta: string) {
  if (!ruta || ruta.trim() === "") return;

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 1. Verificar si ya existe el NamedItem "Ruta" y eliminarlo
    const existingName = workbook.names.getItemOrNullObject("Ruta");
    await context.sync();

    if (!existingName.isNullObject) {
      existingName.delete();
    }

    // 2. Crear o reemplazar el NamedItem guardando el valor de la ruta como fórmula de texto
    workbook.names.add("Ruta", `="${ruta.replace(/"/g, '""')}"`);

    await context.sync();
    console.log("Ruta guardada en el libro correctamente.");
  });
}

if (typeof window !== "undefined") {
  (window as any).Mayuscula_Click = Mayuscula_Click;
  (window as any).Minuscula_Click = Minuscula_Click;
  (window as any).Oracion_Click = Oracion_Click;
  (window as any).Primera_Click = Primera_Click;
  (window as any).Agruparceldas_Click = Agruparceldas_Click;
  (window as any).pdf_Click = pdf_Click;
}

Office.onReady(() => {
  Office.actions.associate("Mayuscula_Click", Mayuscula_Click);
  Office.actions.associate("Minuscula_Click", Minuscula_Click);
  Office.actions.associate("Oracion_Click", Oracion_Click);
  Office.actions.associate("Primera_Click", Primera_Click);
  Office.actions.associate("Agruparceldas_Click", Agruparceldas_Click);
  Office.actions.associate("pdf_Click", pdf_Click);
});
Office.actions.associate("action", action);

