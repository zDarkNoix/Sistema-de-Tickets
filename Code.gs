const NOME_DA_GUIA = "2025.06.24(BJT)";
const COLUNA_CLIENTE = 'L'; 
const NOME_DA_GUIA_FAMILIARES = "DADOSFAMILIARES";

const COLUNAS_DADOS = {
  B: "Contato 1", C: "Contato 2",
  D: "Contato 3", E: "Contato 4",
  F: "Contato 5", G: "Contato 6",
  L: "Nome da Loja", M: "Nome do responsável pela loja", N: "Número de telefone", O: "CPF",
  R: "CNPJ", S: "Data de registro do CNPJ", T: "Categoria do CNPJ", U: "Endereço do CNPJ",
  V: { pergunta: "Contato 1 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  W: "Contato 1 (Data selecionada automaticamente1)",
  X: { pergunta: "Contato 2 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  Y: "Contato 2 (Data selecionada automaticamente2)",
  Z: { pergunta: "Contato 3 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  AA: "Contato 3 (Data selecionada automaticamente3)",
  AB: { pergunta: "Contato 4 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  AC: "Contato 4 (Data selecionada automaticamente4)",
  AD: { pergunta: "Contato 5 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  AE: "Contato 5 (Data selecionada automaticamente5)",
  AF: { pergunta: "Contato 6 (Status)", opcoes: ['Sucesso', 'Não atende', 'Número incorreto', 'Retorno agendado']},
  AG: "Contato 6 (Data selecionada automaticamente6)",
  AH: "AGENDAMENTO [RETORNO] Data: / Horário: /Nome da pessoa:",
  AI: "ANOTAÇÕES SOBRE TENTATIVAS DE LIGAÇÕES/ RETORNO",
  AX: "INFORMAÇÕES COMPLEMENTARES (DADOS FORNECIDOS/ PERCEPÇÃO)",
  AJ: "O comerciante está ciente do processo de repasse com o 99Food?", 
  AK: "O nome da loja está correto?",
  AL: "O nome do responsável legal está correto?",
  AM: "O telefone do responsável legal está correto?",
  AN: "O CPF está correto?",
  AO: "O CPF do familiar está correto?",
  AP: "O endereço do CPF está correto?",
  AQ: "O CNPJ está correto?",
  AR: "O tempo de registro do CNPJ está correto?",
  AS: "A categoria do CNPJ está correta?",
  AT: "O endereço registrado no CNPJ está correto?",
  AU: "O proprietário está disposto a cooperar enviando uma foto sua segurando um documento de identificação válido? SE NÃO, ANOTAR MOTIVO", 
  AV: { pergunta: "Classificação FINAL do comerciante", opcoes: ['Número incorreto/Não atende', 'Uso fraudulento das informações comerciais', 'Comerciante com informações irregulares', 'Comerciante com informações regulares']},
  AW: { pergunta: "Agente", opcoes: ["Juliana", "Fagner", "Pedro", "Debora", "Mateus"] },
  AZ: "Log Última Alteração"
};

const COLUNA_STATUS_TICKET = 'AY'; 
const COLUNA_INFO_ULTIMA_ALTERACAO = 'AZ';


function getFamilyData(rowNumber) {
  Logger.log(`Iniciando getFamilyData para o ticket ID (linha): ${rowNumber}`);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_DA_GUIA_FAMILIARES);
    
    if (!sheet) {
      Logger.log(`ERRO: A planilha '${NOME_DA_GUIA_FAMILIARES}' não foi encontrada.`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("AVISO: A planilha de familiares está vazia ou contém apenas o cabeçalho.");
      return [];
    }

    const allData = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); 
    Logger.log(`Processando ${allData.length} linhas da planilha de familiares.`);

    const matchingRows = allData.filter(row => {

      return String(row[0]).trim() == String(rowNumber);
    });

    Logger.log(`Encontradas ${matchingRows.length} pessoas para o ticket ID ${rowNumber}.`);

    if (matchingRows.length === 0) {
      return [];
    }


    const familyMembers = matchingRows.map(row => {
      return {

        parentesco: row[8] || "Não informado", 

        nome: row[6] || "Não informado",      

        documento: row[5] || "Não informado"  
      };
    });
    

    Logger.log("Dados dos familiares a serem enviados: " + JSON.stringify(familyMembers));
    return familyMembers;

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em getFamilyData para a linha ${rowNumber}: ${e.message}`);
    return [];
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sistema de Tickets')
    .addItem('Abrir Painel de Tickets', 'abrirPainelDeTickets')
    .addSeparator()
    .addItem('Ver Relatório de Agentes', 'abrirPainelRelatorio')
    .addToUi();
}

function extrairPessoasHorizontal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const origem = ss.getSheetByName('2025.06.24(BJT)');
  const destinoNome = 'DADOSFAMILIARES';

  const campos = ['pontuacao', 'idade', 'aposentado', 'documento', 'nome', 'obito', 'campo', 'tipo_beneficio', 'idade_informacao'];


  let destino = ss.getSheetByName(destinoNome);
  if (!destino) {
    destino = ss.insertSheet(destinoNome);
  } else {
    destino.clear();
  }


  const cabecalho = ['Linha Original', 'Pessoa #', ...campos];
  destino.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);

  const ultimaLinha = origem.getLastRow();
  const dados = origem.getRange(6, 16, ultimaLinha - 5, 16).getValues();

  const linhasSaida = [];

  for (let i = 0; i < dados.length; i++) {
    const linhaOrigem = i + 6;
    let texto = dados[i][0];
    if (!texto || typeof texto !== 'string') continue;

    texto = texto.trim();


    const matches = texto.match(/\{[^{}]+\}/g);
    if (!matches) continue;

    let pessoaIndex = 1;
    for (const jsonStr of matches) {
      try {
        const pessoa = JSON.parse(jsonStr);

        const linhaPessoa = campos.map(campo => pessoa[campo] || '');
        linhasSaida.push([linhaOrigem, pessoaIndex, ...linhaPessoa]);
        pessoaIndex++;
      } catch (e) {
        Logger.log(`Erro ao analisar JSON na linha ${linhaOrigem}: ${e.message}`);
      }
    }
  }

  if (linhasSaida.length > 0) {
    destino.getRange(2, 1, linhasSaida.length, cabecalho.length).setValues(linhasSaida);
  }
}



function abrirPainelDeTickets() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1250) 
    .setHeight(750); 
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sistema de Tickets - Fagner Gomes de Lima');
}

function abrirPainelRelatorio() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('relatorio.html')
    .setWidth(600) 
    .setHeight(500); 
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Relatório de Sucessos por Dia');
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_DA_GUIA);
}

function getColumnLetter(colIndex) {
  let letter = ''; let temp;
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

function getColumnIndex(letter) {
  let index = 0; letter = letter.toUpperCase();
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

function normalizeCheckboxValueGS(value) {
    if (value === "" || value === undefined || value === null) { return false; }
    return String(value).toLowerCase() === 'true' || value === true;
}

function getColumnDefinitions() { 
  return COLUNAS_DADOS; 
}

function getTicketDetails(rowNumber) {
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error(`Guia "${NOME_DA_GUIA}" não encontrada.`);
    if (rowNumber < 5) throw new Error("Número da linha inválido. Deve ser >= 5.");

    const identificadorPrincipalTicket = sheet.getRange(`${COLUNA_CLIENTE}${rowNumber}`).getValue();
    let ticketData = {
      id: rowNumber,
      clienteNome: identificadorPrincipalTicket || "Identificador não encontrado", 
      status: sheet.getRange(`${COLUNA_STATUS_TICKET}${rowNumber}`).getValue() || "Novo",
      campos: {}
    };

    for (const col in COLUNAS_DADOS) {
      const colIdx = getColumnIndex(col);
      let valor = sheet.getRange(rowNumber, colIdx).getValue();
      if (typeof valor === 'boolean') { /* Mantém booleano */ }
      else if (valor instanceof Date) { valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy"); }
      else { valor = valor === null || valor === undefined ? "" : String(valor); }
      ticketData.campos[col] = valor;
    }

    ticketData.familiares = getFamilyData(rowNumber); 
    
    return ticketData;
  } catch (error) { Logger.log(`Erro em getTicketDetails R${rowNumber}: ${error.message}`); return { error: `Erro: ${error.message}` }; }
}

function createOrUpdateTicket(rowNumber) {
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error(`Guia "${NOME_DA_GUIA}" não encontrada.`);
    if (rowNumber < 5) throw new Error("Número da linha inválido. Deve ser >= 5.");
    const statusRange = sheet.getRange(`${COLUNA_STATUS_TICKET}${rowNumber}`);
    if (statusRange.getValue() !== "Resolvido") { statusRange.setValue("Em Andamento"); }
    return getTicketDetails(rowNumber); 
  } catch (error) { Logger.log(`Erro em createOrUpdateTicket R${rowNumber}: ${error.message}`); return { error: `Erro: ${error.message}` }; }
}

function saveTicketData(ticketData) {
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error(`Guia "${NOME_DA_GUIA}" não encontrada.`);
    const rowNumber = ticketData.id;
    if (rowNumber < 5) throw new Error("ID de ticket inválido.");

    for (const col in ticketData.campos) {
      if (COLUNAS_DADOS.hasOwnProperty(col)) {
        if (col === COLUNA_INFO_ULTIMA_ALTERACAO || col === 'AW') continue; 
        const colIdx = getColumnIndex(col);
        let valor = ticketData.campos[col];
        if (['B','C','D','E','F','G'].includes(col)) { 
          valor = (String(valor).toLowerCase() === 'true' || valor === true);
        } 
        else if (['W','Y','AA','AC','AE','AG', 'S'].includes(col)) { 
            if (typeof valor === 'string' && valor.match(/^\d{4}-\d{2}-\d{2}$/) && valor.length === 10) {
                const parts = valor.split('-');
                if (parts.length === 3 && parts[0].length === 4) {
                    valor = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                } else { valor = null; }
            } else if (!valor) { valor = null; }
        }
        sheet.getRange(rowNumber, colIdx).setValue(valor);
      }
    }
    
    const userEmail = Session.getActiveUser().getEmail() || "Usuário";
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm 'do dia' dd/MM/yyyy");
    let logInfo = '';

    const statusRange = sheet.getRange(`${COLUNA_STATUS_TICKET}${rowNumber}`);
    if (ticketData.finalizar) {
        logInfo = `Ticket fechado às ${timestamp} por ${userEmail}`;
      

        const contactAttemptDropdowns = ['V', 'X', 'Z', 'AB', 'AD', 'AF'];
        let houveSucesso = false;
        for (const col of contactAttemptDropdowns) {
          if (ticketData.campos[col] === 'Sucesso') {
            houveSucesso = true;
            break;
          }
        }
        
        if(houveSucesso) {
          statusRange.setValue("Resolvido");
        } else {
          statusRange.setValue("Resolvido S/ Sucesso");
        }



      let agentName = '';
      switch(userEmail.toLowerCase()) {
        case 'mailagent1@99app.com': agentName = 'agent1'; break;
        case 'mailagent2@99app.com': agentName = 'agent2'; break;
        case 'mailagent3@99app.com': agentName = 'agent3'; break;
        default: agentName = 'Não identificado';
      }
      sheet.getRange(rowNumber, getColumnIndex('AW')).setValue(agentName);

    } else {
        logInfo = `Última alteração às ${timestamp} por ${userEmail}`;
        if (statusRange.getValue() !== "Resolvido") { 
             statusRange.setValue("Em Andamento");
        }
    }
    sheet.getRange(`${COLUNA_INFO_ULTIMA_ALTERACAO}${rowNumber}`).setValue(logInfo);
    
    return getTicketDetails(rowNumber); 
  } catch (error) { 
    Logger.log(`Erro em saveTicketData T${ticketData.id}: ${error.message}`); 
    return { error: `Erro: ${error.message}`, saved: false }; 
  }
}

function getAllTickets() { 
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error(`Guia "${NOME_DA_GUIA}" não encontrada.`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const contactAttemptDropdowns = ['V', 'X', 'Z', 'AB', 'AD', 'AF'];
    const colsToFetch = ['B','C','D','E','F','G', COLUNA_CLIENTE, COLUNA_STATUS_TICKET, COLUNA_INFO_ULTIMA_ALTERACAO, ...contactAttemptDropdowns];
    
    let minColIdx = Infinity, maxColIdx = 0;
    colsToFetch.forEach(c => {
        if (c) { const idx = getColumnIndex(c); if (idx < minColIdx) minColIdx = idx; if (idx > maxColIdx) maxColIdx = idx; }
    });
    if (minColIdx === Infinity) return [];
    
    const rangeString = `${getColumnLetter(minColIdx)}2:${getColumnLetter(maxColIdx)}${lastRow}`;
    const allData = sheet.getRange(rangeString).getValues();
    const ticketsResumidos = [];
    const mapaDeIndices = {};
    colsToFetch.forEach(c => { if (c) mapaDeIndices[c] = getColumnIndex(c) - minColIdx; });

    for (let i = 0; i < allData.length; i++) {
      const linhaAtualValores = allData[i];
      const numeroLinhaPlanilha = i + 2; 
      const status = linhaAtualValores[mapaDeIndices[COLUNA_STATUS_TICKET]];


      if (status === "Em Andamento" || status === "Resolvido" || status === "Resolvido S/ Sucesso") {
        let ticketResumo = {
          id: numeroLinhaPlanilha,
          nomeLoja: linhaAtualValores[mapaDeIndices[COLUNA_CLIENTE]] || "Loja não definida", 
          status: status,
          logUltimaAlteracao: linhaAtualValores[mapaDeIndices[COLUNA_INFO_ULTIMA_ALTERACAO]] || "",
          campos: {} 
        };
        ['B','C','D','E','F','G'].forEach(cbCol => {
            if (mapaDeIndices[cbCol] !== undefined) {
                 ticketResumo.campos[cbCol] = normalizeCheckboxValueGS(linhaAtualValores[mapaDeIndices[cbCol]]);
            } else { ticketResumo.campos[cbCol] = false; }
        });
        contactAttemptDropdowns.forEach(ddCol => {
            if (mapaDeIndices[ddCol] !== undefined) {
                ticketResumo.campos[ddCol] = linhaAtualValores[mapaDeIndices[ddCol]] || "";
            }
        });

        ticketsResumidos.push(ticketResumo);
      }
    }
    return ticketsResumidos;
  } catch (error) { Logger.log(`Erro em getAllTickets: ${error.message}`); return { error: `Erro: ${error.message}` }; }
}

function getReportData() {
  const callStatusCounts = getAgentStatusCountsByDate(); 
  const resolutionCounts = getResolutionCountsByAgent();
  const dailyFinalStatusCounts = getDailyFinalStatusCounts();

  const agentOptions = (COLUNAS_DADOS && COLUNAS_DADOS['AW'] && COLUNAS_DADOS['AW'].opcoes) ? COLUNAS_DADOS['AW'].opcoes : [];

  const errorMsg = (callStatusCounts.error || "") + " " + (resolutionCounts.error || "") + " " + (dailyFinalStatusCounts.error || "");
  if (errorMsg.trim()) {
    return { error: errorMsg };
  }

  return {
    counts: callStatusCounts,
    resolutionCounts: resolutionCounts,
    dailyFinalStatusCounts: dailyFinalStatusCounts,
    agents: agentOptions
  };
}


function getAgentStatusCountsByDate() {
  try {
    const sheet = getSheet();
    if (!sheet) {
      throw new Error(`Planilha "${NOME_DA_GUIA}" não encontrada.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};


    const statusesToTrack = {
      'Sucesso': 'sucesso',
      'Não atende': 'naoAtende',
      'Número incorreto': 'numeroIncorreto',
      'Retorno agendado': 'retornoAgendado'
    };
    const statusValues = Object.keys(statusesToTrack);


    const statusDateMap = {
      'V': 'W', 'X': 'Y', 'Z': 'AA',
      'AB': 'AC', 'AD': 'AE', 'AF': 'AG'
    };
    
    const agentColLetter = 'AW';
    const statusCols = Object.keys(statusDateMap);
    const dateCols = Object.values(statusDateMap);
    const colsToFetch = [agentColLetter, ...statusCols, ...dateCols];

    let minColIdx = Infinity, maxColIdx = 0;
    colsToFetch.forEach(colLetter => {
      const idx = getColumnIndex(colLetter);
      if (idx < minColIdx) minColIdx = idx;
      if (idx > maxColIdx) maxColIdx = idx;
    });

    const range = sheet.getRange(2, minColIdx, lastRow - 1, (maxColIdx - minColIdx) + 1);
    const values = range.getValues();

    const agentStatusData = {};

    const agentOptions = COLUNAS_DADOS[agentColLetter].opcoes || [];
    agentOptions.forEach(agent => {
      agentStatusData[agent] = {};
    });

    for (const row of values) {
      const agentColInArray = getColumnIndex(agentColLetter) - minColIdx;
      const agentName = row[agentColInArray];

      if (agentName && agentStatusData.hasOwnProperty(agentName)) {
        
        for (const statusCol of statusCols) {
          const statusColInArray = getColumnIndex(statusCol) - minColIdx;
          const statusValue = row[statusColInArray];
          

          if (statusValues.includes(statusValue)) {
            const dateColLetter = statusDateMap[statusCol];
            const dateColInArray = getColumnIndex(dateColLetter) - minColIdx;
            const dateValue = row[dateColInArray];
            
            if (dateValue instanceof Date) {
              const formattedDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
              

              if (!agentStatusData[agentName][formattedDate]) {
                agentStatusData[agentName][formattedDate] = {
                    sucesso: 0,
                    naoAtende: 0,
                    numeroIncorreto: 0,
                    retornoAgendado: 0
                };
              }


                const statusKey = statusesToTrack[statusValue];
              agentStatusData[agentName][formattedDate][statusKey]++;

            }
          }
        }
      }
    }
    
    return agentStatusData;

  } catch (e) {
    Logger.log(`Erro em getAgentStatusCountsByDate: ${e.message}`);
    return { error: `Erro ao processar contagem de status: ${e.message}` };
  }
}


function getResolutionCountsByAgent() {
  try {
    const sheet = getSheet();
    if (!sheet) {
      throw new Error(`Planilha "${NOME_DA_GUIA}" não encontrada.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const agentColLetter = 'AW';
    const statusColLetter = 'AY';
    

    const agentRange = sheet.getRange(`${agentColLetter}2:${agentColLetter}${lastRow}`).getValues();
    const statusRange = sheet.getRange(`${statusColLetter}2:${statusColLetter}${lastRow}`).getValues();
    
    const resolutionCounts = {};
    const agentOptions = (COLUNAS_DADOS[agentColLetter] && COLUNAS_DADOS[agentColLetter].opcoes) || [];


    agentOptions.forEach(agent => {
      resolutionCounts[agent] = {
        resolvido: 0,
        resolvidoSemSucesso: 0
      };
    });


    for (let i = 0; i < agentRange.length; i++) {
      const agentName = agentRange[i][0];
      const status = statusRange[i][0];

      if (agentName && resolutionCounts.hasOwnProperty(agentName)) {
        if (status === 'Resolvido') {
          resolutionCounts[agentName].resolvido++;
        } else if (status === 'Resolvido S/ Sucesso') {
          resolutionCounts[agentName].resolvidoSemSucesso++;
        }
      }
    }
    
    return resolutionCounts;

  } catch (e) {
    Logger.log(`Erro em getResolutionCountsByAgent: ${e.message}`);
    return { error: `Erro ao processar contagem de resoluções: ${e.message}` };
  }
}


function getDailyFinalStatusCounts() {
  try {
    const sheet = getSheet();
    if (!sheet) throw new Error("Planilha não encontrada.");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const statuses = sheet.getRange(`AY2:AY${lastRow}`).getValues();
    const logs = sheet.getRange(`AZ2:AZ${lastRow}`).getValues();
    const dailyCounts = {};
    const dateRegex = /(\d{2}\/\d{2}\/\d{4})/;

    for (let i = 0; i < statuses.length; i++) {
      const status = statuses[i][0];
      const log = logs[i][0];

      if ((status === 'Resolvido' || status === 'Resolvido S/ Sucesso') && log) {
        const match = log.match(dateRegex);
        if (match) {
          const date = match[0];
          if (!dailyCounts[date]) {
            dailyCounts[date] = { resolvido: 0, resolvidoSemSucesso: 0 };
          }
          if (status === 'Resolvido') {
            dailyCounts[date].resolvido++;
          } else {
            dailyCounts[date].resolvidoSemSucesso++;
          }
        }
      }
    }
    return dailyCounts;
  } catch(e) {
    Logger.log(`Erro em getDailyFinalStatusCounts: ${e.message}`);
    return { error: e.message };
  }
}
