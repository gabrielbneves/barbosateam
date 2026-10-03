/**
 * =========================================================================================
 * BARBOSA TEAM - GOOGLE APPS SCRIPT DE SINCRONIZAÇÃO EM NUVEM
 * Sistema Integrado de Avaliação Física, Prescrição & Banco de Dados de Exercícios
 * =========================================================================================
 * 
 * INSTRUÇÕES DE INSTALAÇÃO NO GOOGLE PLANILHAS:
 * 1. Abra sua planilha do Google Planilhas.
 * 2. Clique no menu superior em: "Extensões" > "Apps Script".
 * 3. Apague todo o código existente no editor e cole este código completo.
 * 4. Clique no ícone de disquete "Salvar projeto".
 * 5. Clique no botão azul "Implantar" (canto superior direito) > "Nova implantação".
 * 6. Selecione o tipo "App da Web":
 *    - Descrição: "API Barbosa Team - Alunos e Exercícios"
 *    - Executar como: "Eu (seu e-mail)"
 *    - Quem tem acesso: "Qualquer pessoa" (necessário para que o aplicativo possa sincronizar)
 * 7. Clique em "Implantar", autorize as permissões da sua conta Google e copie a URL da Web gerada.
 * 8. A aba "Exercicios" e a aba "Alunos" serão criadas e formatadas automaticamente!
 * =========================================================================================
 */

function obterOuCriarAbaExercicios(ss) {
  var aba = ss.getSheetByName('Exercicios');
  if (!aba) {
    aba = ss.insertSheet('Exercicios');
    var cabecalhos = ['ID', 'Nome do Exercício', 'Tipo', 'Tipo de Vídeo', 'URL do Vídeo / ID', 'Grupamentos Musculares', 'Data de Criação'];
    aba.appendRow(cabecalhos);
    var headerRange = aba.getRange(1, 1, 1, cabecalhos.length);
    headerRange.setBackground('#1F4E78');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 160); // ID
    aba.setColumnWidth(2, 260); // Nome
    aba.setColumnWidth(3, 110); // Tipo
    aba.setColumnWidth(4, 110); // Tipo de Video
    aba.setColumnWidth(5, 300); // URL
    aba.setColumnWidth(6, 260); // Grupos
    aba.setColumnWidth(7, 160); // Data
  }
  return aba;
}

function obterOuCriarAbaAlunos(ss) {
  var aba = ss.getSheetByName('Alunos');
  if (!aba) {
    aba = ss.insertSheet('Alunos');
    var cabecalhos = ['ID', 'Nome', 'Idade', 'Sexo', 'Telefone', 'Objetivo', 'Dados_JSON', 'Ultima_Atualizacao'];
    aba.appendRow(cabecalhos);
    var headerRange = aba.getRange(1, 1, 1, cabecalhos.length);
    headerRange.setBackground('#1F4E78');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

// -----------------------------------------------------------------------------
// GET: LER DADOS DA PLANILHA (ALUNOS E EXERCÍCIOS)
// -----------------------------------------------------------------------------
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ler Alunos
    var abaAlunos = obterOuCriarAbaAlunos(ss);
    var dadosAlunos = abaAlunos.getDataRange().getValues();
    var alunos = [];
    if (dadosAlunos.length > 1) {
      for (var i = 1; i < dadosAlunos.length; i++) {
        var rowA = dadosAlunos[i];
        if (rowA[0]) {
          try {
            var objAluno = rowA[6] ? JSON.parse(rowA[6]) : null;
            if (objAluno) {
              alunos.push(objAluno);
            } else {
              alunos.push({
                id: String(rowA[0]),
                nome: String(rowA[1]),
                idade: rowA[2],
                sexo: String(rowA[3]),
                telefone: String(rowA[4]),
                objetivo: String(rowA[5])
              });
            }
          } catch(errA) {
            alunos.push({ id: String(rowA[0]), nome: String(rowA[1]) });
          }
        }
      }
    }

    // 2. Ler Banco de Exercícios
    var abaExercicios = obterOuCriarAbaExercicios(ss);
    var dadosEx = abaExercicios.getDataRange().getValues();
    var exercicios = [];
    if (dadosEx.length > 1) {
      for (var j = 1; j < dadosEx.length; j++) {
        var rowE = dadosEx[j];
        if (rowE[0] && rowE[1]) {
          var gruposArr = [];
          if (rowE[5]) {
            gruposArr = String(rowE[5]).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
          }
          exercicios.push({
            id: String(rowE[0]),
            nome: String(rowE[1]),
            tipo: String(rowE[2] || 'forca'),
            tipoVideo: String(rowE[3] || 'youtube'),
            videoUrl: String(rowE[4] || ''),
            grupos: gruposArr,
            grupo: gruposArr[0] || 'Geral',
            dataCriacao: String(rowE[6] || '')
          });
        }
      }
    }

    // 3. Ler Excluídos se houver aba
    var excluidos = [];
    var abaExcluidos = ss.getSheetByName('Excluidos');
    if (abaExcluidos) {
      var dadosExc = abaExcluidos.getDataRange().getValues();
      for (var k = 1; k < dadosExc.length; k++) {
        if (dadosExc[k][0]) excluidos.push(String(dadosExc[k][0]));
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'sucesso',
      alunos: alunos,
      exercicios: exercicios,
      excluidos: excluidos
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'erro',
      mensagem: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------------------------------------
// POST: SALVAR OU SINCRONIZAR DADOS (ALUNOS E EXERCÍCIOS)
// -----------------------------------------------------------------------------
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);
    var acao = payload.acao;

    // -------------------------------------------------------------------------
    // AÇÃO 1: SALVAR OU ATUALIZAR EXERCÍCIO INDIVIDUAL
    // -------------------------------------------------------------------------
    if (acao === 'salvar_exercicio') {
      var ex = payload.exercicio;
      if (!ex || !ex.nome) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Dados do exercício incompletos' })).setMimeType(ContentService.MimeType.JSON);
      }
      var abaEx = obterOuCriarAbaExercicios(ss);
      var dadosT = abaEx.getDataRange().getValues();
      var idEx = ex.id || ('ex_' + new Date().getTime());
      var gruposStr = Array.isArray(ex.grupos) ? ex.grupos.join(', ') : (ex.grupo || 'Geral');
      var dataCriacaoStr = ex.dataCriacao || new Date().toISOString();

      var linhaExistente = -1;
      for (var l = 1; l < dadosT.length; l++) {
        if (String(dadosT[l][0]) === String(idEx) || String(dadosT[l][1]).toLowerCase() === String(ex.nome).toLowerCase()) {
          linhaExistente = l + 1; // 1-based index
          break;
        }
      }

      if (linhaExistente > 0) {
        abaEx.getRange(linhaExistente, 1, 1, 7).setValues([[
          idEx,
          ex.nome,
          ex.tipo || 'forca',
          ex.tipoVideo || 'youtube',
          ex.videoUrl || '',
          gruposStr,
          dataCriacaoStr
        ]]);
      } else {
        abaEx.appendRow([
          idEx,
          ex.nome,
          ex.tipo || 'forca',
          ex.tipoVideo || 'youtube',
          ex.videoUrl || '',
          gruposStr,
          dataCriacaoStr
        ]);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: 'sucesso', id: idEx, mensagem: 'Exercício salvo com sucesso na aba Exercicios' })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------------------
    // AÇÃO 2: EXCLUIR EXERCÍCIO DA PLANILHA
    // -------------------------------------------------------------------------
    if (acao === 'excluir_exercicio') {
      var idDel = payload.id;
      var nomeDel = payload.nome;
      var abaExDel = obterOuCriarAbaExercicios(ss);
      var dadosDel = abaExDel.getDataRange().getValues();
      for (var d = 1; d < dadosDel.length; d++) {
        if ((idDel && String(dadosDel[d][0]) === String(idDel)) || (nomeDel && String(dadosDel[d][1]).toLowerCase() === String(nomeDel).toLowerCase())) {
          abaExDel.deleteRow(d + 1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'sucesso', mensagem: 'Exercício excluído da planilha' })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------------------
    // AÇÃO 3: SINCRONIZAR BANCO COMPLETO DE EXERCÍCIOS
    // -------------------------------------------------------------------------
    if (acao === 'sincronizar_exercicios') {
      var listaEx = Array.isArray(payload.exercicios) ? payload.exercicios : [];
      var abaExSync = obterOuCriarAbaExercicios(ss);
      
      // Limpar linhas anteriores mantendo cabeçalho
      var ultimaLinha = abaExSync.getLastRow();
      if (ultimaLinha > 1) {
        abaExSync.deleteRows(2, ultimaLinha - 1);
      }

      var linhasParaInserir = [];
      for (var m = 0; m < listaEx.length; m++) {
        var it = listaEx[m];
        var gStr = Array.isArray(it.grupos) ? it.grupos.join(', ') : (it.grupo || 'Geral');
        linhasParaInserir.push([
          it.id || ('ex_' + (new Date().getTime() + m)),
          it.nome || 'Exercício',
          it.tipo || 'forca',
          it.tipoVideo || 'youtube',
          it.videoUrl || '',
          gStr,
          it.dataCriacao || new Date().toISOString()
        ]);
      }

      if (linhasParaInserir.length > 0) {
        abaExSync.getRange(2, 1, linhasParaInserir.length, 7).setValues(linhasParaInserir);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: 'sucesso', total: linhasParaInserir.length, mensagem: 'Aba Exercicios atualizada com sucesso' })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------------------
    // AÇÕES DE ALUNOS (SALVAR ALUNO / SINCRONIZAR BANCO COMPLETO)
    // -------------------------------------------------------------------------
    if (acao === 'salvar_aluno') {
      var aluno = payload.aluno;
      if (!aluno || !aluno.id) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Aluno inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      var abaAl = obterOuCriarAbaAlunos(ss);
      var dadosAl = abaAl.getDataRange().getValues();
      var lAl = -1;
      for (var a = 1; a < dadosAl.length; a++) {
        if (String(dadosAl[a][0]) === String(aluno.id)) {
          lAl = a + 1;
          break;
        }
      }
      var jsonStr = JSON.stringify(aluno);
      var dataHora = new Date().toISOString();
      if (lAl > 0) {
        abaAl.getRange(lAl, 1, 1, 8).setValues([[
          aluno.id, aluno.nome, aluno.idade || '', aluno.sexo || '', aluno.telefone || '', aluno.objetivo || '', jsonStr, dataHora
        ]]);
      } else {
        abaAl.appendRow([
          aluno.id, aluno.nome, aluno.idade || '', aluno.sexo || '', aluno.telefone || '', aluno.objetivo || '', jsonStr, dataHora
        ]);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'sucesso', mensagem: 'Aluno salvo na nuvem' })).setMimeType(ContentService.MimeType.JSON);
    }

    if (acao === 'sincronizar_banco') {
      var alunosBanco = payload.banco && payload.banco.alunos ? payload.banco.alunos : [];
      var abaAlSync = obterOuCriarAbaAlunos(ss);
      var ultL = abaAlSync.getLastRow();
      if (ultL > 1) {
        abaAlSync.deleteRows(2, ultL - 1);
      }
      var rowsAl = [];
      for (var b = 0; b < alunosBanco.length; b++) {
        var alItem = alunosBanco[b];
        rowsAl.push([
          alItem.id, alItem.nome, alItem.idade || '', alItem.sexo || '', alItem.telefone || '', alItem.objetivo || '', JSON.stringify(alItem), new Date().toISOString()
        ]);
      }
      if (rowsAl.length > 0) {
        abaAlSync.getRange(2, 1, rowsAl.length, 8).setValues(rowsAl);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'sucesso', mensagem: 'Banco sincronizado' })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Ação não reconhecida: ' + acao })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
