/**
 * QUIZ DOS TRÊS PODERES — Dados no Google Drive
 *
 * A cada quiz terminado, este script:
 *   1. grava uma linha na aba "Resultados" (todos os dados da partida);
 *   2. refaz a aba "Ranking" (ordenada, com a posição de cada pessoa);
 *   3. refaz o arquivo "resultados.csv" na pasta "Quiz Três Poderes - Dados"
 *      do seu Google Drive (abre no Excel, Google Planilhas, Power BI etc.).
 *
 * COMO INSTALAR
 * 1. No Google Drive, crie uma Planilha Google (ex.: "Quiz Três Poderes - Resultados").
 * 2. Na planilha: Extensões > Apps Script.
 * 3. Apague o código que aparece e cole TODO este arquivo. Salve (ícone de disquete).
 * 4. Clique em Implantar > Nova implantação.
 *      - Tipo: App da Web
 *      - Executar como: Eu
 *      - Quem pode acessar: Qualquer pessoa
 *    Clique em Implantar e autorize o acesso com sua conta Google
 *    (ele pede acesso à planilha e ao Drive, por causa do arquivo CSV).
 *    Se aparecer "O Google não verificou este app": Avançado > Acessar projeto.
 * 5. Copie a "URL do app da Web" (termina em /exec) e cole em URL_PLANILHA no index.html.
 *
 * Se alterar este código depois: Implantar > Gerenciar implantações > editar (lápis) >
 * Versão: "Nova versão" > Implantar. A URL continua a mesma.
 *
 * Na planilha aparece o menu "Quiz" > "Atualizar ranking e CSV agora",
 * útil se você apagar ou corrigir linhas à mão.
 */

const ABA_RESULTADOS = "Resultados";
const ABA_RANKING    = "Ranking";
const NOME_PASTA     = "Quiz Três Poderes - Dados";
const NOME_CSV       = "resultados.csv";
const LIMITE_RANKING = 100;   // quantas posições o site mostra

// Colunas novas entram sempre NO FINAL, para não desalinhar linhas já gravadas
const CABECALHO = [
  "Data", "ID", "Nome", "Sobrenome", "Monstrinho", "Idade", "Escolaridade / Curso",
  "Pontos", "Acertos", "Total de questões", "Tempo (s)",
  "Acertos fáceis (de 3)", "Acertos médios (de 4)", "Acertos difíceis (de 3)",
  "Origem", "E-mail"
];

// Colunas da aba Resultados (posição começando em 0)
const C = {
  data: 0, id: 1, nome: 2, sobrenome: 3, monstrinho: 4, idade: 5, curso: 6,
  pontos: 7, acertos: 8, total: 9, tempo: 10, facil: 11, medio: 12, dificil: 13,
  origem: 14, email: 15
};

// Monstrinho escolhido no site (código) → nome gravado na planilha
const MONSTRINHOS = { m: "Azul", f: "Rosa", n: "Amarelo" };
// Linhas gravadas antes da troca guardavam "sexo"; eram as mesmas três opções
const VALORES_ANTIGOS = { "Masculino": "m", "Feminino": "f", "Não informado": "n" };


/* =========================================================
   SITE → PLANILHA
   ========================================================= */

/** Recebe um resultado do site, grava e atualiza ranking + CSV. */
function doPost(e) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const porNivel = d.porNivel || {};

    abaResultados_().appendRow([
      new Date(),
      texto_(d.id, 60),
      texto_(d.nome, 40),
      texto_(d.sobrenome, 60),
      MONSTRINHOS[d.avatar] || MONSTRINHOS.n,
      numero_(d.idade, 1, 120),
      texto_(d.curso, 100),
      numero_(d.pontos, 0, 1000),
      numero_(d.acertos, 0, 100),
      numero_(d.total, 1, 100),
      numero_(d.tempo, 0, 86400),
      acertosDoNivel_(porNivel.facil),
      acertosDoNivel_(porNivel.medio),
      acertosDoNivel_(porNivel.dificil),
      texto_(d.origem, 30),
      texto_(d.email, 120)
    ]);

    atualizarRankingECsv_();
    return json_({ ok: true });
  } catch (erro) {
    return json_({ ok: false, erro: String(erro) });
  } finally {
    trava.releaseLock();
  }
}

/** Devolve o ranking para o site (sem idade, origem e e-mail, que não aparecem publicamente). */
function doGet() {
  try {
    const ranking = lerResultados_()
      .slice(0, LIMITE_RANKING)
      .map(r => ({
        id:        String(r[C.id]),
        data:      new Date(r[C.data]).getTime(),
        nome:      String(r[C.nome]),
        sobrenome: String(r[C.sobrenome]),
        curso:     String(r[C.curso]),
        avatar:    codigoMonstrinho_(r[C.monstrinho]),
        pontos:    Number(r[C.pontos])  || 0,
        acertos:   Number(r[C.acertos]) || 0,
        total:     Number(r[C.total])   || 0,
        tempo:     Number(r[C.tempo])   || 0
      }));
    return json_({ ok: true, ranking: ranking });
  } catch (erro) {
    return json_({ ok: false, erro: String(erro) });
  }
}


/* =========================================================
   RANKING E CSV
   ========================================================= */

/** Refaz a aba Ranking e o arquivo CSV a partir da aba Resultados. */
function atualizarRankingECsv_() {
  const linhas = lerResultados_();   // já ordenado
  const fuso = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  const tabela = [[
    "Posição", "Data", "Nome", "Sobrenome", "Monstrinho", "Idade", "Escolaridade / Curso",
    "Pontos", "Acertos", "Total de questões", "Tempo (s)", "Tempo (min:seg)",
    "Acertos fáceis (de 3)", "Acertos médios (de 4)", "Acertos difíceis (de 3)",
    "Origem", "E-mail"
  ]].concat(linhas.map((r, i) => [
    i + 1,
    Utilities.formatDate(new Date(r[C.data]), fuso, "dd/MM/yyyy HH:mm:ss"),
    r[C.nome], r[C.sobrenome], MONSTRINHOS[codigoMonstrinho_(r[C.monstrinho])], r[C.idade], r[C.curso],
    r[C.pontos], r[C.acertos], r[C.total], r[C.tempo], minSeg_(r[C.tempo]),
    r[C.facil], r[C.medio], r[C.dificil], r[C.origem], r[C.email]
  ].map(semFormula_)));

  // Aba Ranking
  const aba = abaRanking_();
  aba.clearContents();
  aba.getRange(1, 1, tabela.length, tabela[0].length).setValues(tabela);
  aba.getRange(1, 1, 1, tabela[0].length).setFontWeight("bold");
  aba.setFrozenRows(1);

  // Arquivo CSV no Drive
  arquivoCsv_().setContent(montarCsv_(tabela));
}

/**
 * CSV no padrão brasileiro: separado por ponto e vírgula e com BOM UTF-8,
 * para o Excel abrir com acentos certos e colunas separadas.
 */
function montarCsv_(tabela) {
  const celula = v => {
    const s = String(semFormula_(v == null ? "" : v));
    return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return "﻿" + tabela.map(l => l.map(celula).join(";")).join("\r\n") + "\r\n";
}

/** Linhas da aba Resultados, ordenadas: mais pontos, depois menor tempo, depois quem jogou antes. */
function lerResultados_() {
  const aba = abaResultados_();
  if (aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHO.length).getValues()
    .filter(r => r[C.id])
    .sort((a, b) =>
      (Number(b[C.pontos]) || 0) - (Number(a[C.pontos]) || 0) ||
      (Number(a[C.tempo])  || 0) - (Number(b[C.tempo])  || 0) ||
      new Date(a[C.data]) - new Date(b[C.data])
    );
}


/* =========================================================
   MENU NA PLANILHA
   ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Quiz")
    .addItem("Atualizar ranking e CSV agora", "atualizarManual")
    .addToUi();
}

function atualizarManual() {
  atualizarRankingECsv_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Ranking e "' + NOME_CSV + '" atualizados na pasta "' + NOME_PASTA + '".', "Quiz", 5
  );
}


/* =========================================================
   AUXILIARES
   ========================================================= */

function abaResultados_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA_RESULTADOS);
  if (!aba) {
    aba = planilha.insertSheet(ABA_RESULTADOS);
    aba.appendRow(CABECALHO);
    aba.setFrozenRows(1);
  }
  // Mantém o cabeçalho em dia numa planilha antiga (ex.: "Sexo" → "Monstrinho", + "E-mail")
  const titulo = aba.getRange(1, 1, 1, CABECALHO.length);
  if (titulo.getValues()[0].join("|") !== CABECALHO.join("|")) {
    titulo.setValues([CABECALHO]).setFontWeight("bold");
  }
  return aba;
}

function abaRanking_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  return planilha.getSheetByName(ABA_RANKING) || planilha.insertSheet(ABA_RANKING);
}

// Pasta e arquivo são criados uma vez; os IDs ficam guardados para não duplicar
function arquivoCsv_() {
  const props = PropertiesService.getScriptProperties();

  let pasta = null;
  try { pasta = DriveApp.getFolderById(props.getProperty("PASTA_ID")); } catch (e) {}
  if (!pasta || pasta.isTrashed()) {
    pasta = DriveApp.createFolder(NOME_PASTA);
    props.setProperty("PASTA_ID", pasta.getId());
    props.deleteProperty("CSV_ID");
  }

  let arquivo = null;
  try { arquivo = DriveApp.getFileById(props.getProperty("CSV_ID")); } catch (e) {}
  if (!arquivo || arquivo.isTrashed()) {
    arquivo = pasta.createFile(NOME_CSV, "", MimeType.CSV);
    props.setProperty("CSV_ID", arquivo.getId());
  }
  return arquivo;
}

// "2/3" → 2
function acertosDoNivel_(valor) {
  const n = parseInt(String(valor || "0").split("/")[0], 10);
  return isFinite(n) ? n : 0;
}

// "Rosa" (ou o valor antigo "Feminino") → "f" (o site usa a letra para desenhar o monstrinho)
function codigoMonstrinho_(valor) {
  for (const k in MONSTRINHOS) if (MONSTRINHOS[k] === valor) return k;
  return VALORES_ANTIGOS[valor] || "n";
}

function minSeg_(segundos) {
  const s = Math.max(0, Math.round(Number(segundos) || 0));
  return Math.floor(s / 60) + ":" + ("0" + (s % 60)).slice(-2);
}

// Texto começando com = + - @ viraria fórmula na planilha ou no Excel
function semFormula_(v) {
  return typeof v === "string" && /^[=+\-@]/.test(v) ? "'" + v : v;
}

// Limita o tamanho e impede que o texto seja lido como fórmula
function texto_(valor, max) {
  return semFormula_(String(valor == null ? "" : valor).trim().slice(0, max));
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
