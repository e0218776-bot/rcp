function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Guardar fila limpia con un solo campo de teléfono unificado
    sheet.appendRow([
      new Date(),
      data.nombre,
      data.email,
      data.telefono,
      data.activo
    ]);

    // Alerta inmediata por correo al broker
    var destinatario = "jorge.olvera.sp@gmail.com";
    var asunto = "🚨 NUEVO LEAD DISCOVERY: " + data.activo + " - " + data.nombre;
    var cuerpo = "Se ha recibido un nuevo registro confidencial en Selling Playa / Royal Coast Properties:\n\n" +
                 "• Nombre: " + data.nombre + "\n" +
                 "• Activo de Interés: " + data.activo + "\n" +
                 "• WhatsApp: " + data.telefono + "\n" +
                 "• Correo: " + data.email + "\n" +
                 "• Fecha: " + new Date().toLocaleString() + "\n\n" +
                 "Iniciar contacto para perfilamiento discovery.";

    MailApp.sendEmail(destinatario, asunto, cuerpo);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
