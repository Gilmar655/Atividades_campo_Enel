/**
 * Backend do site Registro de Atividades de Campo.
 *
 * Opção recomendada:
 * 1. Abra a planilha Dashboard_acompanhamento em campo no Google Sheets.
 * 2. Acesse Extensões > Apps Script.
 * 3. Cole este conteúdo no arquivo Code.gs.
 * 4. Ajuste APP_TOKEN se desejar uma validação adicional.
 * 5. Publique como Aplicativo da Web.
 */

const CONFIG = {
  SPREADSHEET_ID: "", // Deixe vazio quando o script estiver vinculado à planilha.
  PHOTO_FOLDER_ID: "", // Deixe vazio para criar a pasta automaticamente.
  APP_TOKEN: "", // Opcional. Use o mesmo valor na tela Configuração do site.
  REPORT_FOLDER_NAME: "Relatorios_Atividades_Campo",
  PHOTO_FOLDER_NAME: "Fotos_Atividades_Campo",
};

const RECORD_HEADERS = [
  "ID_ATIVIDADE",
  "DATA_CRIACAO",
  "HORA_CRIACAO",
  "DATA_ATIVIDADE",
  "HORA_INICIO",
  "HORA_TERMINO",
  "DURACAO_ATIVIDADE",
  "EMAIL_USUARIO",
  "ID_TECNICO",
  "NOME_TECNICO_ENEL",
  "MATRICULA_TECNICO",
  "REGIAO",
  "POLO",
  "TIPO_ATIVIDADE",
  "OUTRO_TIPO_ATIVIDADE",
  "ID_PARCEIRA",
  "NOME_PARCEIRA",
  "NUMERO_CONTRATO",
  "NUMERO_PROJETO",
  "NUMERO_OS",
  "CIRCUITO",
  "TIPO_INTERVENCAO",
  "LOCALIZACAO_GPS",
  "LATITUDE",
  "LONGITUDE",
  "ENDERECO",
  "MUNICIPIO",
  "DESCRICAO_ATIVIDADE",
  "SITUACAO_ENCONTRADA",
  "ACAO_REALIZADA",
  "RESULTADO_ATIVIDADE",
  "EXISTE_NAO_CONFORMIDADE",
  "DESCRICAO_NAO_CONFORMIDADE",
  "EXISTE_PENDENCIA",
  "DESCRICAO_PENDENCIA",
  "RESPONSAVEL_PENDENCIA",
  "PRAZO_PENDENCIA",
  "PRIORIDADE",
  "STATUS_ATIVIDADE",
  "QUANTIDADE_FOTOS",
  "LINK_MAPA",
  "DATA_ULTIMA_ALTERACAO",
  "USUARIO_ULTIMA_ALTERACAO",
  "STATUS_SINCRONIZACAO",
  "OBSERVACOES",
];

const ACTION_HEADERS = [
  "ID_ACAO",
  "ID_ATIVIDADE",
  "TIPO_ACAO",
  "DESCRICAO_ACAO",
  "RESPONSAVEL",
  "EMPRESA_RESPONSAVEL",
  "DATA_ABERTURA",
  "PRAZO",
  "PRIORIDADE",
  "STATUS_ACAO",
  "DATA_CONCLUSAO",
  "COMPROVACAO_CONCLUSAO",
  "OBSERVACOES",
];

const PHOTO_HEADERS = [
  "ID_FOTO",
  "ID_ATIVIDADE",
  "FOTO",
  "LEGENDA_FOTO",
  "TIPO_EVIDENCIA",
  "DATA_FOTO",
  "HORA_FOTO",
  "LOCALIZACAO_FOTO",
  "USUARIO_RESPONSAVEL",
  "LINK_ARQUIVO",
];

function doGet(e) {
  try {
    validateToken_(e && e.parameter ? e.parameter.token : "");
    const action = (e && e.parameter && e.parameter.action) || "ping";

    if (action === "ping") {
      return json_({
        ok: true,
        message: "Conexão ativa",
        timestamp: new Date().toISOString(),
      });
    }

    if (action === "list") {
      return json_({
        ok: true,
        records: sheetToObjects_("REGISTROS_ATIVIDADES", RECORD_HEADERS),
        actions: sheetToObjects_("ACOES_PENDENCIAS", ACTION_HEADERS),
      });
    }

    return json_({ ok: false, error: "Ação GET não reconhecida." });
  } catch (error) {
    return json_({ ok: false, error: String(error.message || error) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    validateToken_(payload.token || "");

    if (payload.action !== "saveRecord" || !payload.record) {
      throw new Error("Payload inválido ou ação não reconhecida.");
    }

    const saved = saveRecord_(payload.record, payload.pendingAction || null);
    SpreadsheetApp.flush();
    return json_({ ok: true, id: saved.id, photoCount: saved.photoCount });
  } catch (error) {
    return json_({ ok: false, error: String(error.message || error) });
  } finally {
    lock.releaseLock();
  }
}

function saveRecord_(record, pendingAction) {
  if (!record.id) throw new Error("O registro não possui ID único.");

  const photos = Array.isArray(record.photos) ? record.photos : [];
  const photoLinks = savePhotos_(record, photos);
  const rowObject = recordToSheetObject_(record, photoLinks.length);
  upsertObject_("REGISTROS_ATIVIDADES", RECORD_HEADERS, "ID_ATIVIDADE", rowObject);

  if (pendingAction && pendingAction.id) {
    upsertObject_(
      "ACOES_PENDENCIAS",
      ACTION_HEADERS,
      "ID_ACAO",
      actionToSheetObject_(pendingAction),
    );
  }

  return { id: record.id, photoCount: photoLinks.length };
}

function savePhotos_(record, photos) {
  const links = [];
  const sheet = ensureSheet_("FOTOS_ATIVIDADES", PHOTO_HEADERS);
  const folder = getPhotoFolder_();

  photos.forEach(function (photo, index) {
    if (!photo || !photo.dataUrl) return;
    const match = String(photo.dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) return;

    const mimeType = match[1];
    const extension = mimeType.indexOf("png") >= 0 ? "png" : "jpg";
    const fileName =
      record.id +
      "_" +
      Utilities.formatString("%02d", index + 1) +
      "_" +
      new Date().getTime() +
      "." +
      extension;
    const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
    const file = folder.createFile(blob);
    const link = file.getUrl();
    links.push(link);

    const photoId = photo.id || Utilities.getUuid();
    const existingRow = findRowByKey_(sheet, "ID_FOTO", photoId);
    const values = [
      photoId,
      record.id,
      fileName,
      photo.legend || photo.name || "",
      photo.evidenceType || "Durante a atividade",
      record.activityDate || new Date(),
      record.startTime || "",
      record.location || "",
      record.technician || "",
      link,
    ];

    if (existingRow > 1) {
      sheet.getRange(existingRow, 1, 1, PHOTO_HEADERS.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  });

  return links;
}

function recordToSheetObject_(record, savedPhotoCount) {
  return {
    ID_ATIVIDADE: record.id,
    DATA_CRIACAO: record.createdDate,
    HORA_CRIACAO: record.createdTime,
    DATA_ATIVIDADE: record.activityDate,
    HORA_INICIO: record.startTime,
    HORA_TERMINO: record.endTime,
    DURACAO_ATIVIDADE: record.duration,
    EMAIL_USUARIO: record.userEmail,
    ID_TECNICO: record.technicianId,
    NOME_TECNICO_ENEL: record.technician,
    MATRICULA_TECNICO: record.technicianRegistration,
    REGIAO: record.region,
    POLO: record.pole,
    TIPO_ATIVIDADE: record.activityType,
    OUTRO_TIPO_ATIVIDADE: record.otherActivity,
    ID_PARCEIRA: record.partnerId,
    NOME_PARCEIRA: record.partner,
    NUMERO_CONTRATO: record.contractNumber,
    NUMERO_PROJETO: record.projectNumber,
    NUMERO_OS: record.workOrder,
    CIRCUITO: record.circuit,
    TIPO_INTERVENCAO: record.interventionType,
    LOCALIZACAO_GPS: record.location,
    LATITUDE: record.latitude,
    LONGITUDE: record.longitude,
    ENDERECO: record.address,
    MUNICIPIO: record.municipality,
    DESCRICAO_ATIVIDADE: record.description,
    SITUACAO_ENCONTRADA: record.situation,
    ACAO_REALIZADA: record.actionsTaken,
    RESULTADO_ATIVIDADE: record.activityResult,
    EXISTE_NAO_CONFORMIDADE: record.hasNonConformity ? "Sim" : "Não",
    DESCRICAO_NAO_CONFORMIDADE: record.nonConformity,
    EXISTE_PENDENCIA: record.hasPending ? "Sim" : "Não",
    DESCRICAO_PENDENCIA: record.pendingDescription,
    RESPONSAVEL_PENDENCIA: record.pendingOwner,
    PRAZO_PENDENCIA: record.pendingDeadline,
    PRIORIDADE: record.priority,
    STATUS_ATIVIDADE: record.status,
    QUANTIDADE_FOTOS: savedPhotoCount || record.photoCount || 0,
    LINK_MAPA: record.mapLink,
    DATA_ULTIMA_ALTERACAO: new Date(),
    USUARIO_ULTIMA_ALTERACAO: record.updatedBy || record.technician,
    STATUS_SINCRONIZACAO: "Sincronizado",
    OBSERVACOES: record.notes,
  };
}

function actionToSheetObject_(action) {
  return {
    ID_ACAO: action.id,
    ID_ATIVIDADE: action.activityId,
    TIPO_ACAO: action.type,
    DESCRICAO_ACAO: action.description,
    RESPONSAVEL: action.owner,
    EMPRESA_RESPONSAVEL: action.company,
    DATA_ABERTURA: action.openedAt,
    PRAZO: action.deadline,
    PRIORIDADE: action.priority,
    STATUS_ACAO: action.status,
    DATA_CONCLUSAO: action.completedAt,
    COMPROVACAO_CONCLUSAO: action.proof,
    OBSERVACOES: action.notes,
  };
}

function upsertObject_(sheetName, headers, keyHeader, object) {
  const sheet = ensureSheet_(sheetName, headers);
  const keyValue = object[keyHeader];
  const row = headers.map(function (header) {
    return object[header] === undefined || object[header] === null ? "" : object[header];
  });
  const existingRow = findRowByKey_(sheet, keyHeader, keyValue);
  if (existingRow > 1) {
    sheet.getRange(existingRow, 1, 1, headers.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function findRowByKey_(sheet, keyHeader, keyValue) {
  if (!keyValue || sheet.getLastRow() < 2) return -1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyColumn = headers.indexOf(keyHeader) + 1;
  if (!keyColumn) return -1;
  const finder = sheet
    .getRange(2, keyColumn, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(keyValue))
    .matchEntireCell(true)
    .findNext();
  return finder ? finder.getRow() : -1;
}

function sheetToObjects_(sheetName, expectedHeaders) {
  const sheet = ensureSheet_(sheetName, expectedHeaders);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, expectedHeaders.length).getDisplayValues();
  const headers = values[0];
  return values.slice(1).filter(function (row) {
    return row.some(String);
  }).map(function (row) {
    const result = {};
    headers.forEach(function (header, index) {
      result[header] = row[index];
    });
    return result;
  });
}

function ensureSheet_(name, headers) {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet
      .getRange(1, 1, 1, headers.length)
      .setBackground("#154F82")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
  }
  return sheet;
}

function getSpreadsheet_() {
  return CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getPhotoFolder_() {
  if (CONFIG.PHOTO_FOLDER_ID) return DriveApp.getFolderById(CONFIG.PHOTO_FOLDER_ID);
  const folders = DriveApp.getFoldersByName(CONFIG.PHOTO_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.PHOTO_FOLDER_NAME);
}

function validateToken_(receivedToken) {
  if (CONFIG.APP_TOKEN && receivedToken !== CONFIG.APP_TOKEN) {
    throw new Error("Token de integração inválido.");
  }
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

/**
 * Cria os gatilhos de consolidação:
 * - dia 16: primeira quinzena;
 * - dia 1: mês anterior.
 * Execute esta função uma única vez no editor do Apps Script.
 */
function criarGatilhosDeRelatorio() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (
      trigger.getHandlerFunction() === "gerarRelatorioPrimeiraQuinzena" ||
      trigger.getHandlerFunction() === "gerarRelatorioMensal"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("gerarRelatorioPrimeiraQuinzena")
    .timeBased()
    .onMonthDay(16)
    .atHour(7)
    .create();

  ScriptApp.newTrigger("gerarRelatorioMensal").timeBased().onMonthDay(1).atHour(7).create();
}

function gerarRelatorioPrimeiraQuinzena() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth(), 15, 23, 59, 59);
  gerarResumoPdf_(start, end, "Relatório da 1ª quinzena");
}

function gerarRelatorioMensal() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const end = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  gerarResumoPdf_(start, end, "Relatório mensal consolidado");
}

function gerarResumoPdf_(start, end, title) {
  const records = sheetToObjects_("REGISTROS_ATIVIDADES", RECORD_HEADERS).filter(function (record) {
    const parts = String(record.DATA_ATIVIDADE || "").split("/");
    const date =
      parts.length === 3
        ? new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]))
        : new Date(record.DATA_ATIVIDADE);
    return !isNaN(date.getTime()) && date >= start && date <= end;
  });

  const document = DocumentApp.create(
    title + " - " + Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM"),
  );
  const body = document.getBody();
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(
    "Período: " +
      Utilities.formatDate(start, Session.getScriptTimeZone(), "dd/MM/yyyy") +
      " a " +
      Utilities.formatDate(end, Session.getScriptTimeZone(), "dd/MM/yyyy"),
  );
  body.appendParagraph("Total de atividades: " + records.length);
  body.appendParagraph(
    "Atividades concluídas: " +
      records.filter(function (record) {
        return record.STATUS_ATIVIDADE === "Concluída";
      }).length,
  );
  body.appendParagraph(
    "Atividades com pendência: " +
      records.filter(function (record) {
        return record.EXISTE_PENDENCIA === "Sim";
      }).length,
  );
  body.appendParagraph("Relação consolidada").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  if (records.length) {
    const rows = [["ID", "Data", "Técnico", "Atividade", "Parceira", "Status"]];
    records.forEach(function (record) {
      rows.push([
        record.ID_ATIVIDADE,
        record.DATA_ATIVIDADE,
        record.NOME_TECNICO_ENEL,
        record.TIPO_ATIVIDADE,
        record.NOME_PARCEIRA,
        record.STATUS_ATIVIDADE,
      ]);
    });
    body.appendTable(rows);
  }

  document.saveAndClose();
  const pdf = DriveApp.getFileById(document.getId()).getAs(MimeType.PDF);
  const reportFolders = DriveApp.getFoldersByName(CONFIG.REPORT_FOLDER_NAME);
  const folder = reportFolders.hasNext()
    ? reportFolders.next()
    : DriveApp.createFolder(CONFIG.REPORT_FOLDER_NAME);
  const pdfFile = folder.createFile(pdf).setName(document.getName() + ".pdf");
  DriveApp.getFileById(document.getId()).setTrashed(true);
  return pdfFile.getUrl();
}
