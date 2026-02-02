const SHEET_NAME = 'Hoja 1';
const FOLDER_NAME = 'EXPEDIENTES_VISA_UPLOADS';
const EMAIL_DESTINO = 'barajaspamela010@gmail.com';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Formulario VISA - Completo')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Función auxiliar para obtener o crear carpeta
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

// Función para guardar archivo en Drive y devolver URL y Blob
function saveFileToDrive(base64Data, fileName, folder) {
  try {
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteString = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteString, type, fileName);
    const file = folder.createFile(blob);
    return { url: file.getUrl(), blob: blob };
  } catch (e) {
    return { url: "ERROR: " + e.toString(), blob: null };
  }
}

// --- FUNCIÓN PRINCIPAL ---
function processForm(formData) {
  const lock = LockService.getScriptLock();
  lock.tryLock(40000); // 40 seg espera

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("No se encontró la hoja: " + SHEET_NAME);

    // 1. GESTIÓN DE ENCABEZADOS (Si es hoja nueva)
    if (sheet.getLastRow() === 0) {
      const headers = [
        "FECHA REGISTRO",
        "APELLIDOS", "NOMBRE(S)", "NOMBRE NATIVO", "OTROS NOMBRES", 
        "GENERO", "ESTADO CIVIL", "FECHA NACIMIENTO", "CIUDAD NAC", "ESTADO NAC", "PAIS NAC", "NACIONALIDAD", "OTRA NACIONALIDAD", 
        "INE", "SSN", "TAX ID", 
        "DOMICILIO", "COLONIA", "CIUDAD", "ESTADO", "CP", 
        "TEL CASA", "TEL TRABAJO", "FAX", "CELULAR", "EMAIL", 
        "TIENE REDES", "REDES DETALLE",
        "PASAPORTE NUM", "PAIS EXP", "CIUDAD EXP", "ESTADO EXP", "FECHA EXP", "FECHA CAD", "ROBO PASAPORTE", "DETALLE ROBO",
        "TIPO VISA", "FECHA VIAJE", "ESTADIA", 
        "DIR USA", "CIUDAD USA", "ESTADO USA", "ZIP USA",
        "QUIEN PAGA", "DETALLE PAGADOR", "GRUPO", "NOMBRE GRUPO", "ACOMPAÑANTES", "LISTA ACOMPAÑANTES",
        "HA ESTADO USA", "ULTIMO VIAJE", "ESTADIA ULTIMO", 
        "LICENCIA USA", "NUM LICENCIA",
        "VISA ANTERIOR", "TIPO VISA ANT", "FECHA VISA ANT", "NUM VISA ANT",
        "MISMO TIPO", "MISMO PAIS", "HUELLAS", "VISA PERDIDA", "ANIO PERDIDA", "VISA CANCELADA", "MOTIVO CANCEL", "NEGACION", "MOTIVO NEGACION",
        "TIENE FAMILIAR USA", "CONTACTO USA", "COMPAÑIA USA",
        "PADRE INFO", "MADRE INFO", "TIENE FAM DIRECTOS", "LISTA FAMILIARES",
        "CONYUGE", "CONYUGE INFO",
        "EMPLEO ACTUAL", "EMPLEO INFO",
        "EMPLEO ANTERIOR", "EMPLEO ANT INFO", "EDUCACION",
        "TRIBU", "VIAJES 5 AÑOS", "ORG SOCIAL", "ARMAS", "MILITAR",
        "SALUD", "TRASTORNO", "DROGAS", "PENALES", "DELITO USA", "ILEGAL USA", "TRABAJO ILEGAL",
        "URL INE", "URL PASAPORTE"
      ];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    // 2. GUARDAR ARCHIVOS EN DRIVE
    let urlIne = "NO SUBIDO", urlPasaporte = "NO SUBIDO";
    let blobIne = null, blobPass = null; // Para el PDF
    
    if (formData.fileIneData || formData.filePassData) {
       const folder = getOrCreateFolder(FOLDER_NAME);
       const timestamp = new Date().getTime();
       
       if(formData.fileIneData) {
         const res = saveFileToDrive(formData.fileIneData, `INE_${formData.apellidos}_${timestamp}`, folder);
         urlIne = res.url;
         blobIne = res.blob;
       }
       if(formData.filePassData) {
         const res = saveFileToDrive(formData.filePassData, `PASS_${formData.apellidos}_${timestamp}`, folder);
         urlPasaporte = res.url;
         blobPass = res.blob;
       }
    }

    // 3. GUARDAR EN HOJA DE CÁLCULO
    const rowData = [
      new Date(),
      formData.apellidos, formData.nombres, formData.nombreNativo, formData.otrosNombres,
      formData.genero, formData.estadoCivil, formData.fechaNacimiento, formData.ciudadNac, formData.estadoNac, formData.paisNac, formData.nacionalidad, formData.otraNacionalidad,
      formData.ine, formData.ssn, formData.taxId,
      formData.calle, formData.colonia, formData.ciudadDom, formData.estadoDom, formData.cp,
      formData.ladaCasa + formData.telCasa, formData.ladaTrabajo + formData.telTrabajo, formData.ladaFax + formData.fax, formData.ladaCelular + formData.celular, formData.email,
      formData.tieneRedes, `FB:${formData.facebook} IG:${formData.instagram} TW:${formData.twitter} IN:${formData.linkedin}`,
      formData.pasaporteNum, formData.pasaportePais, formData.pasaporteCiudad, formData.pasaporteEstado, formData.pasaporteFechaExp, formData.pasaporteFechaCad, formData.roboPasaporte, formData.detalleRobo,
      formData.tipoVisa, formData.fechaViaje, formData.tiempoEstadia + " " + formData.unidadEstadia,
      formData.calleUsa, formData.ciudadUsa, formData.estadoUsa, formData.zipUsa,
      formData.quienPaga, formData.detallePagador, formData.viajaGrupo, formData.nombreGrupo, formData.viajaAcompanado, formData.listaAcompanantes,
      formData.haEstadoUsa, formData.fechaUltimoViaje, formData.estadiaUltimoViaje + " " + formData.unidadUltimoViaje,
      formData.licenciaUsa, formData.numLicencia,
      formData.visaAnterior, formData.tipoVisaAnt, formData.fechaVisaAnt, formData.numVisaAnt,
      formData.mismoTipoVisa, formData.mismoPaisVisa, formData.huellas, formData.visaPerdida, formData.anioVisaPerdida, formData.visaCancelada, formData.motivoCancelacion, formData.negacionEntrada, formData.motivoNegacion,
      formData.tieneFamiliarUSA, 
      `CONT:${formData.contUsNombre} ${formData.contUsApellido} REL:${formData.contUsRelacion}`, 
      `COMP:${formData.compNombre} TEL:${formData.compTel}`,
      `P:${formData.padreNombre} ${formData.padreApellidos} USA:${formData.padreEnUsa}`, 
      `M:${formData.madreNombre} ${formData.madreApellidos} USA:${formData.madreEnUsa}`,
      formData.tieneFamDirectos, formData.listaFamiliares,
      formData.espApellidos + " " + formData.espNombre, `NAC:${formData.espNac} VIVE:${formData.espViveUsted}`,
      formData.empNombre, `PUESTO:${formData.empPuesto} SAL:${formData.empSalario} ING:${formData.empFechaIngreso}`,
      formData.antEmpNombre, `PUESTO:${formData.antPuesto} JEFE:${formData.antJefeApe}`, formData.listaEducacion,
      formData.tribu, formData.viajes5, formData.orgSocial, formData.armas, formData.militar,
      formData.enfermedad, formData.trastorno, formData.adicto, formData.delito, formData.delitoUsa, formData.ilegal, formData.trabajoIlegal,
      urlIne, urlPasaporte
    ];
    sheet.appendRow(rowData);

    // 4. GENERAR Y ENVIAR PDF
    sendPdfEmail(formData, blobIne, blobPass);

    return { status: 'success', message: '¡SOLICITUD GUARDADA Y PDF ENVIADO!' };

  } catch (e) {
    console.error(e);
    return { status: 'error', message: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// --- FUNCIÓN PARA GENERAR EL PDF Y ENVIARLO ---
function sendPdfEmail(data, blobIne, blobPass) {
  
  // Construcción del HTML para el PDF
  let html = `
    <html>
      <head>
        <style>
          body { font-family: 'Helvetica', sans-serif; font-size: 10px; color: #333; }
          h1 { color: #003366; border-bottom: 2px solid #003366; padding-bottom: 5px; }
          h2 { background-color: #eee; padding: 5px; font-size: 12px; margin-top: 15px; border-left: 5px solid #003366; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
          td { border-bottom: 1px solid #ddd; padding: 4px; vertical-align: top; }
          .label { font-weight: bold; width: 35%; color: #555; }
          .val { width: 65%; color: #000; }
          .img-container { text-align: center; margin-top: 20px; page-break-inside: avoid; }
          img { max-width: 90%; max-height: 300px; border: 1px solid #ccc; }
        </style>
      </head>
      <body>
        <h1>EXPEDIENTE DE SOLICITUD DE VISA</h1>
        <p><strong>Solicitante:</strong> ${data.apellidos} ${data.nombres}</p>
        <p><strong>Fecha Registro:</strong> ${new Date().toLocaleString()}</p>

        <h2>1. DATOS PERSONALES</h2>
        <table>
          <tr><td class="label">Nombres Completos:</td><td class="val">${data.apellidos} ${data.nombres}</td></tr>
          <tr><td class="label">Otros Nombres / Nativo:</td><td class="val">${data.otrosNombres} / ${data.nombreNativo}</td></tr>
          <tr><td class="label">Género / Estado Civil:</td><td class="val">${data.genero} / ${data.estadoCivil}</td></tr>
          <tr><td class="label">Nacimiento:</td><td class="val">${data.fechaNacimiento} (${data.ciudadNac}, ${data.estadoNac}, ${data.paisNac})</td></tr>
          <tr><td class="label">Nacionalidad:</td><td class="val">${data.nacionalidad} ${data.otraNacionalidad ? '/ ' + data.otraNacionalidad : ''}</td></tr>
          <tr><td class="label">Identificaciones:</td><td class="val">INE: ${data.ine} | SSN: ${data.ssn} | TAX: ${data.taxId}</td></tr>
          <tr><td class="label">Domicilio:</td><td class="val">${data.calle}, ${data.colonia}, ${data.ciudadDom}, ${data.estadoDom} CP: ${data.cp}</td></tr>
          <tr><td class="label">Teléfonos:</td><td class="val">Cel: ${data.ladaCelular}-${data.celular} | Casa: ${data.ladaCasa}-${data.telCasa}</td></tr>
          <tr><td class="label">Email / Redes:</td><td class="val">${data.email} | FB:${data.facebook} IG:${data.instagram}</td></tr>
        </table>

        <h2>2. PASAPORTE Y VIAJE</h2>
        <table>
          <tr><td class="label">Pasaporte:</td><td class="val">${data.pasaporteNum} (${data.pasaportePais}) Exp: ${data.pasaporteFechaExp} Cad: ${data.pasaporteFechaCad}</td></tr>
          <tr><td class="label">Robo Pasaporte:</td><td class="val">${data.roboPasaporte} - ${data.detalleRobo}</td></tr>
          <tr><td class="label">Viaje Planeado:</td><td class="val">Visa: ${data.tipoVisa} | Fecha: ${data.fechaViaje} | Estadía: ${data.tiempoEstadia} ${data.unidadEstadia}</td></tr>
          <tr><td class="label">Destino USA:</td><td class="val">${data.calleUsa}, ${data.ciudadUsa}, ${data.estadoUsa}</td></tr>
          <tr><td class="label">Pagador / Grupo:</td><td class="val">${data.quienPaga} (${data.detallePagador}) | Grupo: ${data.viajaGrupo} (${data.nombreGrupo})</td></tr>
          <tr><td class="label">Acompañantes:</td><td class="val">${data.listaAcompanantes}</td></tr>
        </table>

        <h2>3. HISTORIAL</h2>
        <table>
          <tr><td class="label">Ha estado en USA:</td><td class="val">${data.haEstadoUsa} | Último: ${data.fechaUltimoViaje}</td></tr>
          <tr><td class="label">Visa Anterior:</td><td class="val">${data.visaAnterior} | Núm: ${data.numVisaAnt} | Fecha: ${data.fechaVisaAnt}</td></tr>
          <tr><td class="label">Problemas (Perdida/Cancel/Negada):</td><td class="val">Perdida: ${data.visaPerdida} | Cancel: ${data.visaCancelada} | Negada: ${data.negacionEntrada} (${data.motivoNegacion})</td></tr>
        </table>

        <h2>4. FAMILIA Y TRABAJO</h2>
        <table>
          <tr><td class="label">Contacto USA:</td><td class="val">${data.contUsNombre} ${data.contUsApellido} (${data.contUsRelacion}) / ${data.compNombre}</td></tr>
          <tr><td class="label">Padres:</td><td class="val">P: ${data.padreNombre} ${data.padreApellidos} | M: ${data.madreNombre} ${data.madreApellidos}</td></tr>
          <tr><td class="label">Familiares Directos USA:</td><td class="val">${data.listaFamiliares}</td></tr>
          <tr><td class="label">Cónyuge:</td><td class="val">${data.espNombre} ${data.espApellidos}</td></tr>
          <tr><td class="label">Empleo Actual:</td><td class="val">${data.empNombre} | Puesto: ${data.empPuesto} | Ingreso: ${data.empFechaIngreso}</td></tr>
          <tr><td class="label">Empleo Anterior / Edu:</td><td class="val">${data.antEmpNombre} / ${data.listaEducacion}</td></tr>
        </table>

        <h2>5. SEGURIDAD</h2>
        <table>
          <tr><td class="label">Info Varios:</td><td class="val">Tribu: ${data.tribu} | Viajes: ${data.viajes5} | Org: ${data.orgSocial} | Armas: ${data.armas} | Militar: ${data.militar}</td></tr>
          <tr><td class="label">Salud / Legal:</td><td class="val">Salud: ${data.enfermedad} | Trastorno: ${data.trastorno} | Drogas: ${data.adicto}</td></tr>
          <tr><td class="label">Penal / Migratorio:</td><td class="val">Delito: ${data.delito} | Delito USA: ${data.delitoUsa} | Ilegal: ${data.ilegal} | Trabajo Ilegal: ${data.trabajoIlegal}</td></tr>
        </table>

        <br><hr><br>
        <h3>DOCUMENTOS ADJUNTOS</h3>
  `;

  // Insertar imágenes en el HTML si existen (usando base64 directamente para el PDF)
  // NOTA: Usamos el string base64 original que viene en 'data'
  if(data.fileIneData) {
    html += `<div class="img-container"><h4>INE (Frente)</h4><img src="${data.fileIneData}" /></div>`;
  } else {
    html += `<p>No se adjuntó INE.</p>`;
  }

  if(data.filePassData) {
    html += `<div class="img-container"><h4>PASAPORTE</h4><img src="${data.filePassData}" /></div>`;
  } else {
    html += `<p>No se adjuntó Pasaporte.</p>`;
  }

  html += `</body></html>`;

  // Generar Blob PDF
  const pdfBlob = Utilities.newBlob(html, MimeType.HTML)
                  .getAs(MimeType.PDF)
                  .setName(`Visa_${data.apellidos}_${data.nombres}.pdf`);

  // Enviar Correo
  MailApp.sendEmail({
    to: EMAIL_DESTINO,
    subject: `NUEVA SOLICITUD VISA - ${data.apellidos} ${data.nombres}`,
    body: `Se ha recibido una nueva solicitud de visa.\n\nNombre: ${data.apellidos} ${data.nombres}\nFecha: ${new Date().toLocaleString()}\n\nSe adjunta el expediente completo en PDF con las fotos incluidas.`,
    attachments: [pdfBlob]
  });
}