/**
 * QUIZ DOS TRÊS PODERES — Planilha de resultados (Google Drive)
 *
 * COMO INSTALAR
 * 1. No Google Drive, crie uma Planilha Google (ex.: "Quiz Três Poderes - Resultados").
 * 2. Na planilha: Extensões > Apps Script.
 * 3. Apague o código que aparece e cole TODO este arquivo. Salve (ícone de disquete).
 * 4. Clique em Implantar > Nova implantação.
 *      - Tipo: App da Web
 *      - Executar como: Eu
 *      - Quem pode acessar: Qualquer pessoa
 *    Clique em Implantar e autorize o acesso com sua conta Google.
 *    (Se aparecer "O Google não verificou este app": Avançado > Acessar projeto.)
 * 5. Copie a "URL do app da Web" (termina em /exec) e cole em URL_PLANILHA no index.html.
 *
 * Se alterar este código depois: Implantar > Gerenciar implantações > editar (lápis) >
 * Versão: "Nova versão" > Implantar. A URL continua a mesma.
 */

const ABA = "Resultados";
const CABECALHO = [
  "Data", "ID", "Nome", "Sobrenome", "Idade", "Curso", "Bonequinho",
  "Acertos", "Total", "Tempo (s)",
  "Legislativo", "Executivo", "Judiciário", "Geral"
];
const LIMITE_RANKING = 100;

/** Recebe um resultado do site e grava uma linha na planilha. */
function doPost(e) {
  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const porPoder = d.porPoder || {};

    aba_().appendRow([
      new Date(),
      texto_(d.id, 60),
      texto_(d.nome, 40),
      texto_(d.sobrenome, 60),
      numero_(d.idade, 1, 120),
      texto_(d.curso, 60),
      boneco_(d.avatar),
      numero_(d.acertos, 0, 100),
      numero_(d.total, 1, 100),
      numero_(d.tempo, 0, 86400),
      // apóstrofo para o Sheets não transformar "2/4" em data
      "'" + texto_(porPoder["Legislativo"], 10),
      "'" + texto_(porPoder["Executivo"], 10),
      "'" + texto_(porPoder["Judiciário"], 10),
      "'" + texto_(porPoder["Geral"], 10)
    ]);

    return json_({ ok: true });
  } catch (erro) {
    return json_({ ok: false, erro: String(erro) });
  } finally {
    trava.releaseLock();
  }
}

/** Devolve o ranking (sem a idade, que não aparece publicamente). */
function doGet() {
  try {
    const linhas = aba_().getDataRange().getValues().slice(1);

    const ranking = linhas
      .filter(r => r[1])
      .map(r => ({
        id:        String(r[1]),
        data:      new Date(r[0]).getTime(),
        nome:      String(r[2]),
        sobrenome: String(r[3]),
        curso:     String(r[5]),
        avatar:    boneco_(r[6]),
        acertos:   Number(r[7]) || 0,
        total:     Number(r[8]) || 0,
        tempo:     Number(r[9]) || 0
      }))
      .sort((a, b) => b.acertos - a.acertos || a.tempo - b.tempo || a.data - b.data)
      .slice(0, LIMITE_RANKING);

    return json_({ ok: true, ranking: ranking });
  } catch (erro) {
    return json_({ ok: false, erro: String(erro) });
  }
}

/* ---------- Auxiliares ---------- */

function aba_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA);
    aba.appendRow(CABECALHO);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold");
  }
  return aba;
}

// Limita o tamanho e impede que o texto seja lido como fórmula (=, +, -, @)
function texto_(valor, max) {
  let s = String(valor == null ? "" : valor).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// m = masculino, f = feminino, n = branco (não se identificar)
function boneco_(valor) {
  return valor === "m" || valor === "f" ? valor : "n";
}

function numero_(valor, min, max) {
  const n = Math.round(Number(valor));
  if (!isFinite(n)) return min;
  return Math.min(max, Math.max(min, n));
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
