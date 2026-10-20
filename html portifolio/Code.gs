// Script vinculado à sua Google Planilha (Extensões > Apps Script)
var ABA = "Agendamentos";
var HORARIOS = { manha: ["08:00", "12:00"], tarde: ["14:00", "18:00"] };
var CABECALHO = ["Recebido em", "Status", "Nome", "Telefone", "E-mail", "Pacote", "Plano", "Adicionais", "Total (R$)", "Data", "Horário"];

function aba_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(CABECALHO);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// Horários ocupados (a página usa para bloquear)
function ocupados_(sh) {
  var tz = Session.getScriptTimeZone();
  var linhas = sh.getDataRange().getValues().slice(1), lista = [];
  linhas.forEach(function (r) {
    if (r[1] === "Cancelado") return;
    var d = r[9] instanceof Date ? Utilities.formatDate(r[9], tz, "yyyy-MM-dd") : String(r[9]);
    var slot = String(r[10]).indexOf("08h") === 0 ? "manha" : "tarde";
    lista.push(d + "|" + slot);
  });
  return lista;
}

function doGet() {
  return saida_({ busy: ocupados_(aba_()) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents), sh = aba_();
    if (ocupados_(sh).indexOf(d.data + "|" + d.slot) > -1)
      return saida_({ ok: false, erro: "Esse horário acabou de ser reservado. Escolha outro." });

    sh.appendRow([new Date(), "Pendente", d.nome, d.tel, d.email, d.pacote, d.plano, d.adicionais, d.total, d.data, d.horario]);

    var h = HORARIOS[d.slot];
    CalendarApp.getDefaultCalendar().createEvent(
      "[PENDENTE] Gravação - " + d.nome + " (" + d.pacote + ")",
      new Date(d.data + "T" + h[0] + ":00-03:00"),
      new Date(d.data + "T" + h[1] + ":00-03:00"),
      { description: "Tel: " + d.tel + "\nE-mail: " + d.email + "\nAdicionais: " + (d.adicionais || "-") + "\nTotal: R$ " + d.total }
    );
    return saida_({ ok: true });
  } catch (err) {
    return saida_({ ok: false, erro: "Erro ao salvar. Tente novamente." });
  } finally {
    lock.releaseLock();
  }
}

function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
