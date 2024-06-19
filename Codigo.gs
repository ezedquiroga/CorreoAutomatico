function enviarCorreosPersonalizados() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Envio de correos automaticos');
    if (!sheet) {
        Logger.log("La hoja 'Envio de correos automaticos' no existe.");
        return;
    }
    
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

    var direccionesEnviadas = obtenerDireccionesEnviadas(); // Se obtiene la lista de direcciones enviadas

    for (var i = 0; i < data.length; i++) {
        var destinatario = data[i][8]; // Columna del correo electrónico
        var asunto = "Curso de Manipulación Segura de Alimentos"; // asunto del mail
        var mensajeHtml = `<p>Se registró su inscripción al curso. Dos días antes de la fecha de inicio, le llegará un correo con la clave para ingresar a la clase.<br>Si tiene alguna consulta puede escribir a capacitaciones.bromatologia@vicentelopez.gov.ar</p>`; // Columna del mensaje HTML

        if (destinatario && !direccionesEnviadas.includes(destinatario)) {
            try {
                // Envía el correo
                MailApp.sendEmail({
                    to: destinatario,
                    subject: asunto,
                    htmlBody: mensajeHtml
                });

                // Registro de depuración
                Logger.log("Correo enviado a: " + destinatario);

                // Agrega la dirección al registro de direcciones enviadas
                agregarDireccionEnviada(destinatario, direccionesEnviadas);

            } catch (error) {
                Logger.log("Error al enviar correo a " + destinatario + ": " + error.message);
            }
        } else {
            // Registro de depuración si la dirección ya se ha enviado
            Logger.log("Correo omitido para: " + destinatario);
        }
    }
}

function agregarDireccionEnviada(destinatario, direccionesEnviadas) {
    // Agrega la dirección solo si no existe en la lista
    if (!direccionesEnviadas.includes(destinatario)) {
        direccionesEnviadas.push(destinatario);
        guardarDireccionesEnviadas(direccionesEnviadas);
    }
}

function obtenerDireccionesEnviadas() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var direccionesEnviadas = scriptProperties.getProperty('direccionesEnviadas');
    return direccionesEnviadas ? JSON.parse(direccionesEnviadas) : [];
}

function guardarDireccionesEnviadas(direccionesEnviadas) {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('direccionesEnviadas', JSON.stringify(direccionesEnviadas));
}

function obtenerDireccionesEnviadasDesdeHojaDeCalculo() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Envio de correos automaticos');
    if (!sheet) {
        Logger.log("La hoja 'Envio de correos automaticos' no existe.");
        return [];
    }
    
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

    var direccionesEnviadas = [];

    for (var i = 0; i < data.length; i++) {
        var destinatario = data[i][8]; // Columna del correo electrónico
        if (destinatario) {
            direccionesEnviadas.push(destinatario);
        }
    }

    Logger.log("Direcciones enviadas desde hoja de cálculo: " + direccionesEnviadas.join(", "));
    return direccionesEnviadas;
}

function eliminarFilasTrigger() {
    var direccionesEnviadasHoja = obtenerDireccionesEnviadasDesdeHojaDeCalculo();
    var direccionesEnviadas = obtenerDireccionesEnviadas();
    
    // Filtrar las direcciones enviadas que ya no están en la hoja de cálculo
    direccionesEnviadas = direccionesEnviadas.filter(function(destinatario) {
        return direccionesEnviadasHoja.includes(destinatario);
    });
    
    guardarDireccionesEnviadas(direccionesEnviadas);
    Logger.log("Direcciones enviadas actualizadas después de la eliminación: " + direccionesEnviadas.join(", "));
}
