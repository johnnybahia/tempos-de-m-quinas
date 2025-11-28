// ==========================================================
// 1. ROTEADOR E LOGIN
// ==========================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.maquina) {
    return salvarDadosESP32(e);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Monitoramento Fabril')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function salvarDadosESP32(e) {
  var maquina = e.parameter.maquina;
  var evento = e.parameter.evento;
  var duracao = e.parameter.duracao;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Página1");
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var dataAtual = new Date();
  var timezone = Session.getScriptTimeZone();
  var dataStr = Utilities.formatDate(dataAtual, timezone, "dd/MM/yyyy");
  var horaStr = Utilities.formatDate(dataAtual, timezone, "HH:mm:ss");
  
  sheet.appendRow([dataStr, horaStr, maquina, evento, duracao]);
  return ContentService.createTextOutput("OK");
}

function verificarLogin(usuario, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LOGIN");
  if (!sheet) return { erro: "Aba LOGIN não encontrada." };
  
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === String(usuario).toLowerCase() && String(dados[i][1]) == senha) {
      return { sucesso: true, usuario: dados[i][0] };
    }
  }
  return { erro: "Acesso negado." };
}

// ==========================================================
// 2. FUNÇÕES DE DADOS (LEITURA)
// ==========================================================

function buscarDadosTempoReal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDados = ss.getSheetByName("Página1");
  const dados = sheetDados.getDataRange().getValues();
  
  const sheetTurnos = ss.getSheetByName("TURNOS");
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  const configTurnos = {};
  for (let i = 1; i < dadosTurnos.length; i++) {
    configTurnos[String(dadosTurnos[i][0]).trim()] = [
       { nome: "Turno 1", inicio: dadosTurnos[i][1], fim: dadosTurnos[i][2] },
       { nome: "Turno 2", inicio: dadosTurnos[i][3], fim: dadosTurnos[i][4] },
       { nome: "Turno 3", inicio: dadosTurnos[i][5], fim: dadosTurnos[i][6] }
    ];
  }

  const sheetDadosConfig = ss.getSheetByName("DADOS");
  const mapaFamilias = {};
  if (sheetDadosConfig) {
    const dConfig = sheetDadosConfig.getDataRange().getValues();
    if (dConfig.length > 0) {
      const h = dConfig[0].map(c => String(c).toUpperCase().trim());
      const idxM = h.indexOf("MÁQUINAS");
      const idxF = h.findIndex(x => x.includes("FAMÍLIA") || x.includes("FAMILIA"));
      if (idxM > -1 && idxF > -1) {
        for (let i = 1; i < dConfig.length; i++) {
          let m = String(dConfig[i][idxM]).trim();
          mapaFamilias[m] = dConfig[i][idxF] || "GERAL";
        }
      }
    }
  }

  const statusMaquinas = {};
  const agora = new Date(); 

  for (let i = dados.length - 1; i > 0; i--) {
    let linha = dados[i];
    let maquina = String(linha[2]).trim();
    if (!maquina) continue;

    if (!statusMaquinas[maquina]) {
      let infoTurnoAtual = descobrirTurnoCompleto(agora, maquina, configTurnos);
      let nomeTurnoAtual = "Fora de Turno";
      let dataProducaoAtual = null;

      if (infoTurnoAtual) {
        nomeTurnoAtual = infoTurnoAtual.nome;
        dataProducaoAtual = new Date(agora);
        if (infoTurnoAtual.cruzaMeiaNoite) {
           let horaAgora = agora.getHours();
           let horaInicio = Math.floor(infoTurnoAtual.minInicio / 60);
           if (horaAgora < horaInicio) {
             dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
           }
        }
        dataProducaoAtual.setHours(0,0,0,0);
      }

      let dataReg = lerDataBR(linha[0]);
      let timestampFinal;
      let timePart = new Date(linha[1]); 
      if (!isNaN(dataReg.getTime()) && !isNaN(timePart.getTime())) {
         let d = new Date(dataReg);
         d.setHours(timePart.getHours(), timePart.getMinutes(), timePart.getSeconds(), 0);
         timestampFinal = d.getTime();
      } else {
         timestampFinal = new Date().getTime(); 
      }

      statusMaquinas[maquina] = {
        ultimoEvento: linha[3],
        timestamp: timestampFinal,
        turnoAtual: nomeTurnoAtual,
        familia: mapaFamilias[maquina] || "OUTROS",
        totalProduzindo: 0,
        totalParada: 0,
        refNomeTurno: nomeTurnoAtual,
        refDataProducao: dataProducaoAtual ? dataProducaoAtual.getTime() : null 
      };
    }

    let ref = statusMaquinas[maquina];
    if (ref.refNomeTurno !== "Fora de Turno" && ref.refDataProducao !== null) {
      let dataReg = lerDataBR(linha[0]);
      let horaRegObj = new Date(linha[1]);
      
      if (!isNaN(dataReg.getTime()) && !isNaN(horaRegObj.getTime())) {
        let fullDateReg = new Date(dataReg);
        fullDateReg.setHours(horaRegObj.getHours(), horaRegObj.getMinutes(), horaRegObj.getSeconds());
        
        let infoTurnoReg = descobrirTurnoCompleto(fullDateReg, maquina, configTurnos);
        
        if (infoTurnoReg && infoTurnoReg.nome === ref.refNomeTurno) {
          let dataProdReg = new Date(dataReg);
          if (infoTurnoReg.cruzaMeiaNoite) {
             let h = fullDateReg.getHours();
             let hIni = Math.floor(infoTurnoReg.minInicio / 60);
             if (h < hIni) dataProdReg.setDate(dataProdReg.getDate() - 1);
          }
          dataProdReg.setHours(0,0,0,0);

          if (dataProdReg.getTime() === ref.refDataProducao) {
             let duracao = parseDuration(linha[4]);
             if (linha[3] === "TEMPO PRODUZINDO") ref.totalProduzindo += duracao;
             else if (linha[3] === "TEMPO PARADA") ref.totalParada += duracao;
          }
        }
      }
    }
  }
  return statusMaquinas;
}

// === BUSCA PARADAS DETALHADAS DO TURNO ATUAL ===
// Busca as paradas da aba PAINEL (dados consolidados)
function buscarParadasTurnoAtual(maquina, turnoNome, dataProducao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPainel = ss.getSheetByName("PAINEL") || ss.getSheetByName("Painel");

  if (!sheetPainel) {
    Logger.log("ERRO: Aba PAINEL não encontrada");
    return [];
  }

  const dados = sheetPainel.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  // Converter dataProducao para formato comparável
  const dataProdBusca = new Date(dataProducao);
  dataProdBusca.setHours(0, 0, 0, 0);
  const dataProdStr = Utilities.formatDate(dataProdBusca, timezone, "dd/MM/yyyy");

  Logger.log("DEBUG buscarParadasTurnoAtual:");
  Logger.log("  Máquina: " + maquina);
  Logger.log("  Turno: " + turnoNome);
  Logger.log("  Data: " + dataProdStr);
  Logger.log("  Total linhas PAINEL: " + dados.length);

  // Buscar a linha correspondente
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqPainel = String(linha[0]).trim();
    let turnoPainel = String(linha[1]).trim();

    // Converter data da planilha
    let dataPainel = lerDataBR(linha[2]);
    let dataPainelStr = Utilities.formatDate(dataPainel, timezone, "dd/MM/yyyy");

    // Verificar se é a linha correta
    if (maqPainel === maquina && turnoPainel === turnoNome && dataPainelStr === dataProdStr) {
      Logger.log("  ✓ Linha encontrada! Linha " + (i+1));
      Logger.log("  Dados da linha:");
      Logger.log("    Col 5 (>3min): " + linha[5]);
      Logger.log("    Col 6 (>10min): " + linha[6]);
      Logger.log("    Col 7 (>20min): " + linha[7]);
      Logger.log("    Col 8 (>30min): " + linha[8]);

      // Colunas: 5=TEMPOS>3min, 6=TEMPOS>10min, 7=TEMPOS>20min, 8=TEMPOS>30min
      const paradas = [];

      // Parsear as colunas de tempos (assumindo formato: "00:10:30, 00:15:20")
      const categorias = [
        { coluna: 5, nome: "> 3 min", minDuracao: 180 },
        { coluna: 6, nome: "> 10 min", minDuracao: 600 },
        { coluna: 7, nome: "> 20 min", minDuracao: 1200 },
        { coluna: 8, nome: "> 30 min", minDuracao: 1800 }
      ];

      categorias.forEach(cat => {
        const valorColuna = String(linha[cat.coluna] || "").trim();

        Logger.log("    Processando " + cat.nome + ": '" + valorColuna + "'");

        if (valorColuna && valorColuna !== "-" && valorColuna !== "") {
          // Parsear múltiplos tempos separados por vírgula
          const tempos = valorColuna.split(",");

          tempos.forEach(tempo => {
            tempo = tempo.trim();
            if (tempo && tempo.includes(":")) {
              const partes = tempo.split(":");
              if (partes.length >= 2) {
                const h = parseInt(partes[0]) || 0;
                const m = parseInt(partes[1]) || 0;
                const s = partes.length > 2 ? (parseInt(partes[2]) || 0) : 0;
                const duracaoSeg = h * 3600 + m * 60 + s;

                // NÃO filtrar por duração - os valores JÁ vêm filtrados do gerarRelatorioTurnos()
                // Adicionar TODOS os tempos da categoria
                if (duracaoSeg > 0) {
                  paradas.push({
                    categoria: cat.nome,
                    duracao: duracaoSeg,
                    duracaoFmt: tempo,
                    horario: "-" // Não temos horário específico no PAINEL
                  });
                  Logger.log("      ✓ Adicionada: " + tempo + " (" + duracaoSeg + "s)");
                }
              }
            }
          });
        }
      });

      // Ordenar por duração (maior primeiro)
      paradas.sort((a, b) => b.duracao - a.duracao);

      Logger.log("  Total paradas encontradas: " + paradas.length);

      return paradas;
    }
  }

  return [];
}

function formatarSegundosParaHora(segundos) {
  if (typeof segundos !== 'number' || isNaN(segundos)) return "00:00:00";
  segundos = Math.round(segundos);
  const h = Math.floor(segundos/3600).toString().padStart(2,'0');
  const m = Math.floor((segundos%3600)/60).toString().padStart(2,'0');
  const sec = (segundos%60).toString().padStart(2,'0');
  return `${h}:${m}:${sec}`;
}

// === BUSCA HISTÓRICO ATUALIZADA (FILTRO POR PERÍODO) ===
// Busca da aba "Página" que contém o histórico completo com motivos, serviços, etc.
function buscarHistorico(maquinaFiltro, dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("DEBUG buscarHistorico:");
  Logger.log("  Máquina: " + maquinaFiltro);
  Logger.log("  Data Início: " + dataInicio);
  Logger.log("  Data Fim: " + dataFim);

  // Buscar da aba "Pagina" (sem acento) ou "Página" (com acento) - não PAINEL
  const sheetPagina = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");

  if (!sheetPagina) {
    Logger.log("  ⚠ Aba 'Página' não encontrada, usando PAINEL");
    // Fallback: tentar PAINEL se Página não existir
    return buscarHistoricoPainel(maquinaFiltro, dataInicio, dataFim);
  }

  Logger.log("  ✓ Usando aba: " + sheetPagina.getName());
  const dados = sheetPagina.getDataRange().getValues();
  Logger.log("  Total linhas: " + dados.length);
  const resultados = [];
  
  // Prepara as datas de filtro (se existirem)
  let dInicio = null;
  let dFim = null;

  if (dataInicio) {
    let p = dataInicio.split('-');
    dInicio = new Date(p[0], p[1]-1, p[2]);
    dInicio.setHours(0,0,0,0);
  }
  
  if (dataFim) {
    let p = dataFim.split('-');
    dFim = new Date(p[0], p[1]-1, p[2]);
    dFim.setHours(23,59,59,999);
  }

  // Estrutura da aba "Página":
  // Col 0: MÁQUINAS
  // Col 1: CUSTO MÃO DE OBRA
  // Col 2: TURNO
  // Col 3: DATA
  // Col 4: LIGADA
  // Col 5: DESLIGADA
  // Col 6: TEMPOS > 30 min
  // Col 7: MOTIVO DA PARADA
  // Col 8: MOTIVO DA PARADA (duplicado)
  // Col 9: SERVIÇOS REALIZADOS
  // Col 10: PEÇAS TROCADAS
  // Col 11: CUSTO DE PEÇAS
  // Col 12: DATA DE FABRICAÇÃO OU COMPRA

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqPagina = String(linha[0]).trim();

    // Filtra pela máquina
    if (maqPagina !== maquinaFiltro) continue;

    // Filtra pela data (se houver filtro)
    if (dInicio || dFim) {
      let dataPagina = lerDataBR(linha[3]); // Coluna 3 na aba Página
      dataPagina.setHours(12,0,0,0);

      if (dInicio && dataPagina < dInicio) continue;
      if (dFim && dataPagina > dFim) continue;
    }

    // Se passou, adiciona
    resultados.push({
      turno: linha[2],
      data: Utilities.formatDate(lerDataBR(linha[3]), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"),
      ligada: formatarHoraExcel(linha[4]),
      desligada: formatarHoraExcel(linha[5]),
      paradas3min: "-", // Não existe na aba Página
      paradas10min: "-", // Não existe na aba Página
      paradas20min: "-", // Não existe na aba Página
      paradas30min: linha[6] || "-",
      custoMO: typeof linha[1] === 'number' ? linha[1] : 0, // Coluna 1
      motivo: linha[7] || linha[8] || "-", // Colunas 7 ou 8 (duplicadas)
      servico: linha[9] || "-",
      pecas: linha[10] || "-",
      custoPecas: typeof linha[11] === 'number' ? linha[11] : 0,
      obs: linha[12] || "-" // Data de fabricação como observação
    });
  }
  
  // Ordena por data decrescente (mais novo primeiro)
  resultados.sort((a, b) => {
    let da = lerDataBR(a.data);
    let db = lerDataBR(b.data);
    return db - da;
  });

  Logger.log("  Total registros encontrados: " + resultados.length);

  return resultados;
}

// === BUSCA HISTÓRICO DA ABA PAINEL (FALLBACK) ===
function buscarHistoricoPainel(maquinaFiltro, dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPainel = ss.getSheetByName("PAINEL") || ss.getSheetByName("Painel");

  Logger.log("DEBUG buscarHistoricoPainel (fallback):");
  Logger.log("  Máquina: " + maquinaFiltro);

  if (!sheetPainel) {
    Logger.log("  ERRO: Aba PAINEL também não encontrada!");
    return [];
  }

  Logger.log("  ✓ Usando aba: " + sheetPainel.getName());
  const dados = sheetPainel.getDataRange().getValues();
  Logger.log("  Total linhas: " + dados.length);
  const resultados = [];

  let dInicio = null;
  let dFim = null;

  if (dataInicio) {
    let p = dataInicio.split('-');
    dInicio = new Date(p[0], p[1]-1, p[2]);
    dInicio.setHours(0,0,0,0);
  }

  if (dataFim) {
    let p = dataFim.split('-');
    dFim = new Date(p[0], p[1]-1, p[2]);
    dFim.setHours(23,59,59,999);
  }

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqPainel = String(linha[0]).trim();

    if (maqPainel !== maquinaFiltro) continue;

    if (dInicio || dFim) {
      let dataPainel = lerDataBR(linha[2]);
      dataPainel.setHours(12,0,0,0);

      if (dInicio && dataPainel < dInicio) continue;
      if (dFim && dataPainel > dFim) continue;
    }

    resultados.push({
      turno: linha[1],
      data: Utilities.formatDate(lerDataBR(linha[2]), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"),
      ligada: formatarHoraExcel(linha[3]),
      desligada: formatarHoraExcel(linha[4]),
      paradas3min: linha[5] || "-",
      paradas10min: linha[6] || "-",
      paradas20min: linha[7] || "-",
      paradas30min: linha[8] || "-",
      custoMO: typeof linha[9] === 'number' ? linha[9] : 0,
      motivo: "-",
      servico: "-",
      pecas: "-",
      custoPecas: 0,
      obs: "-"
    });
  }

  resultados.sort((a, b) => {
    let da = lerDataBR(a.data);
    let db = lerDataBR(b.data);
    return db - da;
  });

  Logger.log("  Total registros encontrados (PAINEL): " + resultados.length);

  return resultados;
}

function buscarListasDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DADOS");
  if (!sheet) return { motivos: [], servicos: [] };
  const dados = sheet.getDataRange().getValues();
  let motivos = [], servicos = [];
  let colMotivo = -1, colServico = -1;
  if (dados.length > 0) {
    for (let c = 0; c < dados[0].length; c++) {
      let head = String(dados[0][c]).toUpperCase().trim();
      if (head === "MOTIVO DA PARADA") colMotivo = c;
      if (head === "SERVIÇOS REALIZADOS") colServico = c; 
    }
  }
  for (let i = 1; i < dados.length; i++) {
    if (colMotivo > -1 && dados[i][colMotivo]) motivos.push(dados[i][colMotivo]);
    if (colServico > -1 && dados[i][colServico]) servicos.push(dados[i][colServico]);
  }
  return { motivos: motivos, servicos: servicos };
}

function salvarApontamento(dadosForm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PAINEL") || ss.getSheetByName("Painel");
  const dados = sheet.getDataRange().getValues();
  let linhaEncontrada = -1;
  
  const partes = dadosForm.data.split('-');
  const dataFiltroBR = `${partes[2]}/${partes[1]}/${partes[0]}`;

  for (let i = 1; i < dados.length; i++) {
    let maqPainel = String(dados[i][0]).trim();
    let turnoPainel = String(dados[i][1]).trim();
    let valData = dados[i][2];
    let dataPainelStr = "";
    if (valData instanceof Date) {
      dataPainelStr = Utilities.formatDate(valData, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    } else {
      dataPainelStr = String(valData).trim();
    }
    
    if (maqPainel === dadosForm.maquina && turnoPainel === dadosForm.turno && dataPainelStr === dataFiltroBR) {
      linhaEncontrada = i + 1; 
      break;
    }
  }
  if (linhaEncontrada > 0) {
    sheet.getRange(linhaEncontrada, 11).setValue(dadosForm.motivo);
    sheet.getRange(linhaEncontrada, 12).setValue(dadosForm.servico);
    sheet.getRange(linhaEncontrada, 13).setValue(dadosForm.pecas);
    sheet.getRange(linhaEncontrada, 14).setValue(dadosForm.custo);
    sheet.getRange(linhaEncontrada, 15).setValue(dadosForm.obs);
    return "✅ Dados salvos com sucesso!";
  } else {
    return "⚠️ Registro não encontrado. Rode o 'gerarRelatorioTurnos' para atualizar o painel.";
  }
}

// ==========================================================
// 3. GERADOR DE RELATÓRIO
// ==========================================================

function gerarRelatorioTurnos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDados = ss.getSheetByName("Página1");
  const sheetTurnos = ss.getSheetByName("TURNOS");
  const sheetPainel = ss.getSheetByName("PAINEL");
  const sheetCustos = ss.getSheetByName("DADOS");
  if (!sheetDados || !sheetTurnos || !sheetPainel) return;
  
  const dadosPainelAntigo = sheetPainel.getDataRange().getValues();
  const mapaPainelExistente = {}; 
  const agora = new Date();
  
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  const configTurnos = {}; 
  for (let i = 1; i < dadosTurnos.length; i++) {
    let maquina = dadosTurnos[i][0];
    if (maquina) {
      configTurnos[String(maquina).trim()] = [
         { nome: "Turno 1", inicio: dadosTurnos[i][1], fim: dadosTurnos[i][2] },
         { nome: "Turno 2", inicio: dadosTurnos[i][3], fim: dadosTurnos[i][4] },
         { nome: "Turno 3", inicio: dadosTurnos[i][5], fim: dadosTurnos[i][6] }
      ];
    }
  }

  for (let i = 1; i < dadosPainelAntigo.length; i++) {
    let row = dadosPainelAntigo[i];
    if (row.length > 2) { 
      let d = lerDataBR(row[2]); 
      let dataStr = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      let chave = String(row[0]).trim() + "|" + dataStr + "|" + String(row[1]).trim();
      mapaPainelExistente[chave] = row;
    }
  }
  
  const mapaCustos = {};
  if (sheetCustos) {
    const dadosCustos = sheetCustos.getDataRange().getValues();
    if (dadosCustos.length > 0) {
      const header = dadosCustos[0].map(c => String(c).toUpperCase().trim());
      const idxM = header.indexOf("MÁQUINAS");
      const idxC = header.indexOf("MÃO DE OBRA");
      if (idxM > -1 && idxC > -1) {
        for (let i = 1; i < dadosCustos.length; i++) {
          let m = String(dadosCustos[i][idxM]).trim();
          let c = dadosCustos[i][idxC];
          if (typeof c === 'string') c = parseFloat(c.replace(/[^\d,.-]/g, '').replace(',', '.')) || 0;
          if (m) mapaCustos[m] = c;
        }
      }
    }
  }
  
  const dadosBrutos = sheetDados.getDataRange().getValues();
  const resumo = {}; 
  
  for (let i = 1; i < dadosBrutos.length; i++) {
    let linha = dadosBrutos[i];
    let dataOriginal = lerDataBR(linha[0]); 
    let hora = linha[1];
    let maquinaRaw = linha[2];
    let evento = linha[3];
    let duracao = parseDuration(linha[4]); 
    
    if (!maquinaRaw || !hora) continue;
    let maquina = String(maquinaRaw).trim();
    let nomeEvento = evento ? String(evento).trim() : "";
    if (nomeEvento !== "TEMPO PRODUZINDO" && nomeEvento !== "TEMPO PARADA") continue;
    
    let infoTurno = descobrirTurnoCompleto(hora, maquina, configTurnos);
    if (infoTurno) {
      let dataFim = new Date(dataOriginal);
      if (isNaN(dataFim.getTime())) dataFim = new Date();
      let horaObj = new Date(hora);
      if (!isNaN(horaObj.getTime())) { dataFim.setHours(horaObj.getHours(), horaObj.getMinutes(), horaObj.getSeconds(), 0); }
      let dataProducao = new Date(dataFim);
      if (infoTurno.cruzaMeiaNoite) {
         let horaReg = dataFim.getHours();
         let horaInicio = Math.floor(infoTurno.minInicio / 60);
         if (horaReg < horaInicio) dataProducao.setDate(dataProducao.getDate() - 1);
      }
      processarRegistro(resumo, ss, maquina, dataProducao, infoTurno.nome, nomeEvento, duracao);
    }
  }
  
  const linhasSaida = [];
  const SEGUNDOS_DIA = 86400;
  
  for (let chave in resumo) {
    let item = resumo[chave];
    let rowFinal = [];
    
    let infoTurnoAgora = descobrirTurnoCompleto(agora, item.maquina, configTurnos);
    let ehTurnoAtual = false;
    
    if (infoTurnoAgora && infoTurnoAgora.nome === item.turno) {
      let dataProdAgora = new Date(agora);
      if (infoTurnoAgora.cruzaMeiaNoite) {
         let h = agora.getHours();
         let hIni = Math.floor(infoTurnoAgora.minInicio / 60);
         if (h < hIni) dataProdAgora.setDate(dataProdAgora.getDate() - 1);
      }
      dataProdAgora.setHours(0,0,0,0);
      let dataItem = new Date(item.data);
      dataItem.setHours(0,0,0,0);
      if (dataItem.getTime() === dataProdAgora.getTime()) {
        ehTurnoAtual = true;
      }
    }

    if (mapaPainelExistente[chave] && !ehTurnoAtual) {
       rowFinal = mapaPainelExistente[chave];
    } else {
       let manual = { motivo: "", servico: "", pecas: "", custoPecas: "", obs: "" };
       if (mapaPainelExistente[chave]) {
          let old = mapaPainelExistente[chave];
          manual = { motivo: old[10], servico: old[11], pecas: old[12], custoPecas: old[13], obs: old[14] };
       }

       let tempoParadoLiq = Math.max(0, item.desligada - 3600);
       let custoMO = (tempoParadoLiq / 3600) * (mapaCustos[item.maquina] || 0);
       let valLigada = Math.max(0, item.ligada) / SEGUNDOS_DIA;
       let valDesligada = Math.max(0, item.desligada) / SEGUNDOS_DIA;
       
       rowFinal = [
          item.maquina, item.turno, item.data, valLigada, valDesligada, 
          formatarListaTempos(item.listaStop3), formatarListaTempos(item.listaStop10), formatarListaTempos(item.listaStop20), formatarListaTempos(item.listaStop30), 
          custoMO, manual.motivo, manual.servico, manual.pecas, manual.custoPecas, manual.obs
       ];
    }
    
    linhasSaida.push(rowFinal);
    delete mapaPainelExistente[chave];
  }

  for (let chave in mapaPainelExistente) {
    linhasSaida.push(mapaPainelExistente[chave]);
  }
  
  sheetPainel.clearContents();
  const cabecalho = [[ "MÁQUINAS", "TURNO", "DATA", "LIGADA", "DESLIGADA", "TEMPOS > 3 min", "TEMPOS > 10 min", "TEMPOS > 20 min", "TEMPOS > 30 min", "CUSTO MÃO DE OBRA", "MOTIVO DA PARADA", "SERVIÇOS REALIZADOS", "PEÇAS TROCADAS", "CUSTO PEÇAS", "OBSERVAÇÃO" ]];
  sheetPainel.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho).setFontWeight("bold");
  
  if (linhasSaida.length > 0) {
    const ordemTurno = {"Turno 1": 1, "Turno 2": 2, "Turno 3": 3};
    linhasSaida.sort((a,b) => {
      let dA = lerDataBR(a[2]), dB = lerDataBR(b[2]);
      if (dA.getTime() !== dB.getTime()) return dA.getTime() - dB.getTime();
      let tA = ordemTurno[a[1]] || 9, tB = ordemTurno[b[1]] || 9;
      if (tA !== tB) return tA - tB;
      return String(a[0]).localeCompare(b[0]);
    });
    
    sheetPainel.getRange(2, 1, linhasSaida.length, linhasSaida[0].length).setValues(linhasSaida);
    sheetPainel.getRange(2, 3, linhasSaida.length, 1).setNumberFormat("dd/MM/yyyy");
    sheetPainel.getRange(2, 4, linhasSaida.length, 2).setNumberFormat("[h]:mm:ss");
    sheetPainel.getRange(2, 10, linhasSaida.length, 1).setNumberFormat("R$ #,##0.00");
    sheetPainel.getRange(2, 14, linhasSaida.length, 1).setNumberFormat("R$ #,##0.00");
    sheetPainel.getRange(2, 6, linhasSaida.length, 4).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}

// 4. LIMPEZA
function limparDadosAntigos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Página1");
  if (!sheet) return;
  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  const limite = new Date(hoje.getTime() - (15 * 24 * 60 * 60 * 1000));
  limite.setHours(0,0,0,0);
  let linhasParaDeletar = 0;
  for (let i = 0; i < dados.length; i++) {
    let dataLinha = lerDataBR(dados[i][0]); 
    if (dataLinha < limite) { linhasParaDeletar++; } else { break; }
  }
  if (linhasParaDeletar > 0) { sheet.deleteRows(1, linhasParaDeletar); }
}

// === AUXILIARES ===
function lerDataBR(valor) {
  if (!valor) return new Date();
  if (valor instanceof Date) return valor;
  if (typeof valor === 'string') {
    let partes = valor.split('/');
    if (partes.length === 3) return new Date(partes[2], partes[1]-1, partes[0]);
    partes = valor.split('-');
    if (partes.length === 3) return new Date(partes[0], partes[1]-1, partes[2]);
  }
  return new Date(); 
}
function parseDuration(raw) {
  if (typeof raw === 'number') return raw;
  if (raw instanceof Date) return raw.getHours() * 3600 + raw.getMinutes() * 60 + raw.getSeconds();
  if (typeof raw === 'string') { let s = parseFloat(raw.replace(',', '.').trim()); return isNaN(s) ? 0 : s; }
  return 0;
}
function horaParaMinutos(val) {
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  if (typeof val === 'string' && val.includes(':')) { let p = val.split(':'); return parseInt(p[0]) * 60 + parseInt(p[1]); }
  return 0;
}
function descobrirTurnoCompleto(hora, maq, config) { 
  let c = config[maq] || config[String(maq).trim()];
  if(!c) return null;
  let min = horaParaMinutos(new Date(hora));
  for(let t of c) {
     let i = horaParaMinutos(t.inicio);
     let f = horaParaMinutos(t.fim);
     let cruza = i > f;
     if (!cruza) { if (min >= i && min < f) return { nome: t.nome, minInicio: i, minFim: f, cruzaMeiaNoite: false }; }
     else { if (min >= i || min < f) return { nome: t.nome, minInicio: i, minFim: f, cruzaMeiaNoite: true }; }
  }
  return null;
}
function formatarHoraExcel(val) {
  if (typeof val !== 'number' || val < 0) return 0;
  let s = Math.round(val * 86400);
  let h = Math.floor(s/3600), m = Math.floor((s%3600)/60), sec = s%60;
  return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
}
function processarRegistro(resumo, ss, maquina, data, turno, evento, segundos) {
  if (segundos > 86400) return;
  if (segundos < 0) segundos = 0;
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina + "|" + dStr + "|" + turno;
  if (!resumo[chave]) resumo[chave] = { maquina: maquina, data: data, turno: turno, ligada: 0, desligada: 0, listaStop3: [], listaStop10: [], listaStop20: [], listaStop30: [] };
  if (evento === "TEMPO PRODUZINDO") resumo[chave].ligada += segundos;
  else {
    resumo[chave].desligada += segundos;
    if (segundos > 10) { 
        resumo[chave].qtdParadas++;
        if (segundos > 180) resumo[chave].listaStop3.push(segundos);
        if (segundos > 600) resumo[chave].listaStop10.push(segundos);
        if (segundos > 1200) resumo[chave].listaStop20.push(segundos);
        if (segundos > 1800) resumo[chave].listaStop30.push(segundos);
    }
  }
}
function formatarListaTempos(lista) {
  if (!lista || !lista.length) return "-";
  return lista.map(s => {
    s = Math.max(0, parseFloat(s) || 0);
    let h = Math.floor(s/3600), m = Math.floor((s%3600)/60), sec = Math.floor(s%60);
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
  }).join(", ");
}
function exibirMensagem(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

// === FUNÇÃO DE DIAGNÓSTICO ===
function diagnosticarMaquina(nomeMaquina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDados = ss.getSheetByName("Página1");

  if (!sheetDados) return { erro: "Aba Página1 não encontrada" };

  const dados = sheetDados.getDataRange().getValues();
  const registros = [];

  // Buscar últimos 10 registros da máquina
  for (let i = dados.length - 1; i > 0 && registros.length < 10; i--) {
    let linha = dados[i];
    let maqLinha = String(linha[2]).trim();

    if (maqLinha === nomeMaquina) {
      registros.push({
        linha: i + 1,
        data: linha[0],
        hora: linha[1],
        maquina: linha[2],
        evento: linha[3],
        duracao: linha[4]
      });
    }
  }

  return {
    maquina: nomeMaquina,
    totalRegistros: registros.length,
    ultimoEvento: registros.length > 0 ? registros[0].evento : "Nenhum",
    registros: registros
  };
}

// === FUNÇÃO DE TESTE COMPLETO ===
function testarFuncoes() {
  Logger.log("=== TESTE DE DIAGNÓSTICO ===");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Verificar abas
  Logger.log("\n1. ABAS DISPONÍVEIS:");
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    Logger.log("  - " + sheet.getName() + " (" + sheet.getLastRow() + " linhas)");
  });

  // 2. Verificar aba PAINEL
  Logger.log("\n2. ESTRUTURA ABA PAINEL:");
  const painel = ss.getSheetByName("PAINEL") || ss.getSheetByName("Painel");
  if (painel) {
    const cabecalho = painel.getRange(1, 1, 1, 10).getValues()[0];
    Logger.log("  Cabeçalhos:");
    cabecalho.forEach((col, idx) => {
      Logger.log("    Col " + idx + ": " + col);
    });

    // Primeira linha de dados
    if (painel.getLastRow() > 1) {
      Logger.log("  Primeira linha de dados:");
      const primeiraLinha = painel.getRange(2, 1, 1, 10).getValues()[0];
      primeiraLinha.forEach((val, idx) => {
        Logger.log("    Col " + idx + ": " + val);
      });
    }
  }

  // 3. Verificar aba Página/Pagina
  Logger.log("\n3. ESTRUTURA ABA PÁGINA/PAGINA:");
  const pagina = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");
  if (pagina) {
    Logger.log("  ✓ Encontrada: " + pagina.getName());
    const cabecalho = pagina.getRange(1, 1, 1, 13).getValues()[0];
    Logger.log("  Cabeçalhos:");
    cabecalho.forEach((col, idx) => {
      Logger.log("    Col " + idx + ": " + col);
    });

    // Primeira linha de dados
    if (pagina.getLastRow() > 1) {
      Logger.log("  Primeira linha de dados:");
      const primeiraLinha = pagina.getRange(2, 1, 1, 13).getValues()[0];
      primeiraLinha.forEach((val, idx) => {
        Logger.log("    Col " + idx + ": " + val);
      });
    }
  } else {
    Logger.log("  ⚠ Aba 'Página' ou 'Pagina' NÃO ENCONTRADA");
  }

  Logger.log("\n=== FIM DO TESTE ===");
}
