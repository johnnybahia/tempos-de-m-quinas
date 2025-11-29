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

  // Atualizar PAINEL automaticamente após receber dados do ESP32
  gerarRelatorioTurnos();

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

        // REGRA: Dia de produção vai de 07:00 às 06:59 do dia seguinte
        // Se estamos entre 00:00 e 06:59, pertence ao dia anterior
        let horaAgora = agora.getHours();
        if (horaAgora < 7) {
          // Madrugada (00:00-06:59) = dia de produção começou ontem
          dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
        } else if (infoTurnoAtual.cruzaMeiaNoite && horaAgora < Math.floor(infoTurnoAtual.minInicio / 60)) {
          // Turno individual que cruza meia-noite (ex: Turno 2: 16:49-01:40)
          dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
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
          let h = fullDateReg.getHours();

          // REGRA: Dia de produção vai de 07:00 às 06:59 do dia seguinte
          if (h < 7) {
            // Madrugada (00:00-06:59) = dia de produção começou ontem
            dataProdReg.setDate(dataProdReg.getDate() - 1);
          } else if (infoTurnoReg.cruzaMeiaNoite && h < Math.floor(infoTurnoReg.minInicio / 60)) {
            // Turno individual que cruza meia-noite
            dataProdReg.setDate(dataProdReg.getDate() - 1);
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
// SIMPLIFICADO: Busca diretamente na aba "Pagina"
function buscarParadasTurnoAtual(maquina, turnoNome, dataProducao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PAINEL");

  if (!sheet) {
    Logger.log("ERRO: Aba 'PAINEL' não encontrada");
    return [];
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  // Converter dataProducao para formato comparável
  const dataBusca = new Date(dataProducao);
  const dataBuscaStr = Utilities.formatDate(dataBusca, timezone, "dd/MM/yyyy");

  Logger.log("=== buscarParadasTurnoAtual ===");
  Logger.log("Máquina: " + maquina);
  Logger.log("Turno: " + turnoNome);
  Logger.log("Data: " + dataBuscaStr);

  // Estrutura da aba "PAINEL":
  // Col 0: MÁQUINAS | Col 1: TURNO | Col 2: DATA
  // Col 3: LIGADA | Col 4: DESLIGADA
  // Col 5: TEMPOS > 3 min ← APENAS ESTA COLUNA
  // Col 6: TEMPOS > 10 min | Col 7: TEMPOS > 20 min | Col 8: TEMPOS > 30 min
  // Col 9: CUSTO MÃO DE OBRA | Col 10: MOTIVO DA PARADA | ...

  // Procurar a linha que corresponde
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqLinha = String(linha[0]).trim();
    let turnoLinha = String(linha[1]).trim();
    let dataLinha = lerDataBR(linha[2]);
    let dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy");

    // Verificar se é a linha correta (máquina + turno + data)
    if (maqLinha === maquina && turnoLinha === turnoNome && dataLinhaStr === dataBuscaStr) {
      Logger.log("✓ Linha encontrada: " + (i + 1));

      const paradas = [];

      // Processar APENAS a coluna "> 3 min" (col 5)
      const temposStr = String(linha[5] || "").trim();

      if (temposStr && temposStr !== "-" && temposStr !== "") {
        Logger.log("> 3 min: " + temposStr);

        // Separar por vírgula
        const tempos = temposStr.split(",");

        tempos.forEach(tempo => {
          tempo = tempo.trim();
          if (tempo && tempo.includes(":")) {
            // Parsear HH:MM:SS ou HH:MM
            const partes = tempo.split(":");
            const h = parseInt(partes[0]) || 0;
            const m = parseInt(partes[1]) || 0;
            const s = partes.length > 2 ? (parseInt(partes[2]) || 0) : 0;
            const duracaoSeg = h * 3600 + m * 60 + s;

            if (duracaoSeg > 0) {
              paradas.push({
                duracao: duracaoSeg,
                duracaoFmt: tempo
              });
            }
          }
        });
      }

      Logger.log("Total paradas > 3min: " + paradas.length);
      return paradas;
    }
  }

  Logger.log("⚠ Nenhuma linha encontrada para: " + maquina + " | " + turnoNome + " | " + dataBuscaStr);
  return [];
}

// === BUSCA RELATÓRIO POR FAMÍLIA ===
function buscarRelatorioFamilia(familia, dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");

  if (!sheet) {
    Logger.log("ERRO: Aba 'Pagina' não encontrada");
    return { erro: "Aba 'Pagina' não encontrada" };
  }

  // Carregar mapeamento de máquinas -> famílias da aba DADOS
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
          mapaFamilias[m] = String(dConfig[i][idxF] || "GERAL").trim();
        }
      }
    }
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  Logger.log("=== buscarRelatorioFamilia ===");
  Logger.log("Família: " + familia);
  Logger.log("Período: " + dataInicio + " até " + dataFim);
  Logger.log("Total linhas: " + dados.length);
  Logger.log("Total máquinas no mapa: " + Object.keys(mapaFamilias).length);

  // Estrutura da aba "Pagina" (16 colunas):
  // Col 0: MÁQUINAS | Col 1: CUSTO MO | Col 2: TURNO | Col 3: DATA
  // Col 4: LIGADA | Col 5: DESLIGADA
  // Col 6: TEMPOS > 3 min | Col 7: TEMPOS > 10 min
  // Col 8: TEMPOS > 20 min | Col 9: TEMPOS > 30 min
  // Col 10: MOTIVO | Col 11: MOTIVO(dup) | Col 12: SERVIÇOS
  // Col 13: PEÇAS | Col 14: CUSTO PEÇAS | Col 15: DATA FAB

  const maquinasPorFamilia = {};
  let totalRodandoSeg = 0;
  let totalParadoSeg = 0;
  let totalParadasCriticas = 0;

  // Percorrer dados da aba Pagina
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const nomeMaquina = String(linha[0]).trim();

    // Verificar se pertence à família usando o mapeamento da aba DADOS
    const familiaMaquina = mapaFamilias[nomeMaquina] || "";
    if (familiaMaquina.toUpperCase() !== familia.toUpperCase()) {
      continue;
    }

    const turno = String(linha[2]).trim(); // Col 2: TURNO
    const dataLinha = lerDataBR(linha[3]); // Col 3: DATA
    const dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "yyyy-MM-dd");

    Logger.log(`Linha ${i}: ${nomeMaquina} (fam: ${familiaMaquina}) | ${turno} | ${dataLinhaStr}`);

    // Filtrar por data
    if (dataInicio && dataLinhaStr < dataInicio) {
      Logger.log(`  ⊗ Rejeitada (antes de ${dataInicio})`);
      continue;
    }
    if (dataFim && dataLinhaStr > dataFim) {
      Logger.log(`  ⊗ Rejeitada (depois de ${dataFim})`);
      continue;
    }

    Logger.log("  ✓ Aceita");

    // Extrair dados
    const ligada = linha[4]; // Col 4: LIGADA
    const desligada = linha[5]; // Col 5: DESLIGADA
    const paradasCriticas = String(linha[6] || "").trim(); // Col 6: TEMPOS > 3 min

    // Converter para segundos (pode ser Date object ou número)
    let ligadaSeg = 0;
    let desligadaSeg = 0;

    if (ligada instanceof Date) {
      ligadaSeg = ligada.getHours() * 3600 + ligada.getMinutes() * 60 + ligada.getSeconds();
    } else if (typeof ligada === 'number') {
      ligadaSeg = Math.round(ligada * 86400);
    }

    if (desligada instanceof Date) {
      desligadaSeg = desligada.getHours() * 3600 + desligada.getMinutes() * 60 + desligada.getSeconds();
    } else if (typeof desligada === 'number') {
      desligadaSeg = Math.round(desligada * 86400);
    }

    // Contar paradas críticas (separadas por vírgula)
    let qtdParadas = 0;
    if (paradasCriticas && paradasCriticas !== "-") {
      qtdParadas = paradasCriticas.split(",").filter(p => p.trim().length > 0).length;
    }

    // Acumular totais
    totalRodandoSeg += ligadaSeg;
    totalParadoSeg += desligadaSeg;
    totalParadasCriticas += qtdParadas;

    // Agrupar por máquina
    if (!maquinasPorFamilia[nomeMaquina]) {
      maquinasPorFamilia[nomeMaquina] = [];
    }

    maquinasPorFamilia[nomeMaquina].push({
      data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
      turno: turno,
      ligada: formatarHoraExcel(ligada),
      desligada: formatarHoraExcel(desligada),
      paradas3min: formatarCelulaParada(linha[6]),
      paradas10min: formatarCelulaParada(linha[7]),
      paradas20min: formatarCelulaParada(linha[8]),
      paradas30min: formatarCelulaParada(linha[9]),
      motivo: String(linha[10] || linha[11] || "-"),
      servico: String(linha[12] || "-"),
      pecas: String(linha[13] || "-"),
      custoMO: typeof linha[1] === 'number' ? linha[1] : 0,
      custoPecas: typeof linha[14] === 'number' ? linha[14] : 0,
      obs: String(linha[15] || "-")
    });
  }

  // Montar resultado
  const maquinas = [];
  for (let nomeMaq in maquinasPorFamilia) {
    maquinas.push({
      nome: nomeMaq,
      turnos: maquinasPorFamilia[nomeMaq]
    });
  }

  // Ordenar máquinas por nome
  maquinas.sort((a, b) => a.nome.localeCompare(b.nome));

  Logger.log("Total máquinas encontradas: " + maquinas.length);
  Logger.log("Total paradas críticas: " + totalParadasCriticas);

  // Debug: mostrar famílias disponíveis se não encontrou nada
  if (maquinas.length === 0) {
    const familiasUnicas = [...new Set(Object.values(mapaFamilias))];
    Logger.log("⚠️ NENHUMA MÁQUINA ENCONTRADA!");
    Logger.log("Famílias disponíveis no mapa: " + familiasUnicas.join(", "));
    Logger.log("Família buscada: '" + familia + "'");
  }

  return {
    familia: familia,
    dataInicio: dataInicio.split('-').reverse().join('/'),
    dataFim: dataFim.split('-').reverse().join('/'),
    totais: {
      rodando: formatarSegundosParaHora(totalRodandoSeg),
      parado: formatarSegundosParaHora(totalParadoSeg),
      paradasCriticas: totalParadasCriticas
    },
    maquinas: maquinas
  };
}

function formatarSegundosParaHora(segundos) {
  if (typeof segundos !== 'number' || isNaN(segundos)) return "00:00:00";
  segundos = Math.round(segundos);
  const h = Math.floor(segundos/3600).toString().padStart(2,'0');
  const m = Math.floor((segundos%3600)/60).toString().padStart(2,'0');
  const sec = (segundos%60).toString().padStart(2,'0');
  return `${h}:${m}:${sec}`;
}

// === BUSCA HISTÓRICO SIMPLIFICADO ===
// Busca na aba "Pagina": máquina + filtro de data (opcional)
function buscarHistorico(maquinaFiltro, dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");

  Logger.log("=== buscarHistorico ===");
  Logger.log("Máquina: " + maquinaFiltro);
  Logger.log("Data Início: " + dataInicio);
  Logger.log("Data Fim: " + dataFim);

  if (!sheet) {
    Logger.log("⚠ Aba 'Pagina' não encontrada");
    return [];
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();
  const resultados = [];

  Logger.log("Total linhas: " + dados.length);

  // Estrutura da aba "Pagina" (16 colunas):
  // Col 0: MÁQUINAS | Col 1: CUSTO MO | Col 2: TURNO | Col 3: DATA
  // Col 4: LIGADA | Col 5: DESLIGADA
  // Col 6: TEMPOS > 3 min | Col 7: TEMPOS > 10 min
  // Col 8: TEMPOS > 20 min | Col 9: TEMPOS > 30 min
  // Col 10: MOTIVO | Col 11: MOTIVO(dup) | Col 12: SERVIÇOS
  // Col 13: PEÇAS | Col 14: CUSTO PEÇAS | Col 15: DATA FAB

  // Percorrer todas as linhas
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqLinha = String(linha[0]).trim();

    // Filtro 1: Máquina
    if (maqLinha !== maquinaFiltro) continue;

    // Filtro 2: Data (se fornecido)
    let dataLinha = lerDataBR(linha[3]);
    let dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "yyyy-MM-dd");

    Logger.log("Linha " + i + " - Máq: " + maqLinha + " | Data: " + dataLinhaStr + " | Filtro: " + dataInicio + " a " + dataFim);

    if (dataInicio && dataLinhaStr < dataInicio) {
      Logger.log("  ⊗ Rejeitada (antes do início)");
      continue;
    }
    if (dataFim && dataLinhaStr > dataFim) {
      Logger.log("  ⊗ Rejeitada (depois do fim)");
      continue;
    }

    Logger.log("  ✓ Aceita");

    // Passou nos filtros - adicionar ao resultado
    try {
      // Debug: ver valores brutos das colunas de tempo
      if (i === 1) { // Apenas primeira linha para não encher os logs
        Logger.log("  DEBUG Col 4 (LIGADA): " + linha[4] + " (tipo: " + typeof linha[4] + ")");
        Logger.log("  DEBUG Col 5 (DESLIGADA): " + linha[5] + " (tipo: " + typeof linha[5] + ")");
        Logger.log("  DEBUG Col 6 (PARADAS >3): " + linha[6] + " (tipo: " + typeof linha[6] + ")");
      }

      const registro = {
        turno: String(linha[2] || "-"),
        data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
        ligada: formatarHoraExcel(linha[4]),
        desligada: formatarHoraExcel(linha[5]),
        paradas3min: formatarCelulaParada(linha[6]),
        paradas10min: formatarCelulaParada(linha[7]),
        paradas20min: formatarCelulaParada(linha[8]),
        paradas30min: formatarCelulaParada(linha[9]),
        motivo: String(linha[10] || linha[11] || "-"),
        servico: String(linha[12] || "-"),
        pecas: String(linha[13] || "-"),
        custoMO: typeof linha[1] === 'number' ? linha[1] : 0,
        custoPecas: typeof linha[14] === 'number' ? linha[14] : 0,
        obs: String(linha[15] || "-")
      };

      if (i === 1) {
        Logger.log("  DEBUG Após formatação:");
        Logger.log("    ligada: " + registro.ligada);
        Logger.log("    desligada: " + registro.desligada);
        Logger.log("    paradas3min: " + registro.paradas3min);
      }

      resultados.push(registro);
      Logger.log("  → Registro adicionado: " + registro.data + " | " + registro.turno);
    } catch (erro) {
      Logger.log("  ⚠ ERRO ao processar linha " + i + ": " + erro.message);
    }
  }

  Logger.log("Registros encontrados: " + resultados.length);
  Logger.log("Retornando array com " + resultados.length + " elementos");
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

  Logger.log("=== PROCESSANDO PÁGINA1 ===");
  Logger.log("Total de linhas: " + (dadosBrutos.length - 1));

  let contadores = { total: 0, processadas: 0, ignoradas: 0, semTurno: 0 };

  for (let i = 1; i < dadosBrutos.length; i++) {
    let linha = dadosBrutos[i];
    let dataOriginal = lerDataBR(linha[0]);
    let hora = linha[1];
    let maquinaRaw = linha[2];
    let evento = linha[3];
    let duracaoRaw = linha[4];
    let duracao = parseDuration(duracaoRaw);

    contadores.total++;

    if (!maquinaRaw || !hora) {
      contadores.ignoradas++;
      continue;
    }

    let maquina = String(maquinaRaw).trim();
    let nomeEvento = evento ? String(evento).trim() : "";

    if (nomeEvento !== "TEMPO PRODUZINDO" && nomeEvento !== "TEMPO PARADA") {
      contadores.ignoradas++;
      Logger.log(`Linha ${i+1}: Evento ignorado "${nomeEvento}"`);
      continue;
    }

    let infoTurno = descobrirTurnoCompleto(hora, maquina, configTurnos);
    if (infoTurno) {
      let dataFim = new Date(dataOriginal);
      if (isNaN(dataFim.getTime())) dataFim = new Date();
      let horaObj = new Date(hora);
      if (!isNaN(horaObj.getTime())) { dataFim.setHours(horaObj.getHours(), horaObj.getMinutes(), horaObj.getSeconds(), 0); }
      let dataProducao = new Date(dataFim);
      let horaReg = dataFim.getHours();

      // REGRA: Dia de produção vai de 07:00 às 06:59 do dia seguinte
      if (horaReg < 7) {
        // Madrugada (00:00-06:59) = dia de produção começou ontem
        dataProducao.setDate(dataProducao.getDate() - 1);
      } else if (infoTurno.cruzaMeiaNoite && horaReg < Math.floor(infoTurno.minInicio / 60)) {
        // Turno individual que cruza meia-noite
        dataProducao.setDate(dataProducao.getDate() - 1);
      }

      // Log detalhado para eventos TEMPO PARADA
      if (nomeEvento === "TEMPO PARADA" && duracao > 180) {
        Logger.log(`Linha ${i+1}: ${maquina} | ${nomeEvento} | Raw: ${duracaoRaw} | Segundos: ${duracao} | Turno: ${infoTurno.nome}`);
      }

      processarRegistro(resumo, ss, maquina, dataProducao, infoTurno.nome, nomeEvento, duracao);
      contadores.processadas++;
    } else {
      contadores.semTurno++;
      Logger.log(`Linha ${i+1}: Sem turno - ${maquina} às ${hora}`);
    }
  }

  Logger.log("=== RESUMO PROCESSAMENTO ===");
  Logger.log(`Total: ${contadores.total} | Processadas: ${contadores.processadas} | Ignoradas: ${contadores.ignoradas} | Sem turno: ${contadores.semTurno}`);

  // Log do resumo acumulado
  Logger.log("=== RESUMO PARADAS ACUMULADAS ===");
  for (let chave in resumo) {
    let item = resumo[chave];
    if (item.listaStop3 && item.listaStop3.length > 0) {
      Logger.log(`${chave}: ${item.listaStop3.length} paradas >3min | ${item.listaStop10.length} >10min | ${item.listaStop20.length} >20min | ${item.listaStop30.length} >30min`);
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
      let h = agora.getHours();

      // REGRA: Dia de produção vai de 07:00 às 06:59 do dia seguinte
      if (h < 7) {
        // Madrugada (00:00-06:59) = dia de produção começou ontem
        dataProdAgora.setDate(dataProdAgora.getDate() - 1);
      } else if (infoTurnoAgora.cruzaMeiaNoite && h < Math.floor(infoTurnoAgora.minInicio / 60)) {
        // Turno individual que cruza meia-noite
        dataProdAgora.setDate(dataProdAgora.getDate() - 1);
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

  // Se já é um objeto Date válido, retornar diretamente
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return new Date(valor);
  }

  // Se é string, tentar parsear
  if (typeof valor === 'string') {
    let partes = valor.split('/');
    if (partes.length === 3) {
      // Formato dd/MM/yyyy
      return new Date(parseInt(partes[2]), parseInt(partes[1])-1, parseInt(partes[0]));
    }
    partes = valor.split('-');
    if (partes.length === 3) {
      // Formato yyyy-MM-dd
      return new Date(parseInt(partes[0]), parseInt(partes[1])-1, parseInt(partes[2]));
    }
  }

  // Fallback: tentar converter diretamente
  try {
    const d = new Date(valor);
    if (!isNaN(d.getTime())) return d;
  } catch (e) {
    // Ignorar erro
  }

  return new Date();
}
function parseDuration(raw) {
  if (typeof raw === 'number') return raw;
  if (raw instanceof Date) return raw.getHours() * 3600 + raw.getMinutes() * 60 + raw.getSeconds();
  if (typeof raw === 'string') {
    let str = raw.trim();
    // Se está no formato HH:MM:SS ou MM:SS
    if (str.includes(':')) {
      let partes = str.split(':');
      if (partes.length === 3) {
        // HH:MM:SS
        return parseInt(partes[0]) * 3600 + parseInt(partes[1]) * 60 + parseInt(partes[2]);
      } else if (partes.length === 2) {
        // MM:SS
        return parseInt(partes[0]) * 60 + parseInt(partes[1]);
      }
    }
    // Se é um número puro como string
    let s = parseFloat(str.replace(',', '.'));
    return isNaN(s) ? 0 : s;
  }
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
  // Se é um Date object (Google Sheets armazena horários como Date)
  if (val instanceof Date && !isNaN(val.getTime())) {
    const h = val.getHours();
    const m = val.getMinutes();
    const s = val.getSeconds();
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
  }

  // Se já é uma string formatada (ex: "1:30:14"), retorna direto
  if (typeof val === 'string' && val.includes(':')) {
    return val;
  }

  // Se é um número (fração de dia), converte para HH:MM:SS
  if (typeof val === 'number' && val >= 0 && !isNaN(val)) {
    let s = Math.round(val * 86400);
    let h = Math.floor(s/3600), m = Math.floor((s%3600)/60), sec = s%60;
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
  }

  // Caso padrão
  return "00:00:00";
}

// Formata valores que podem ser strings ou Date objects
function formatarCelulaParada(val) {
  if (!val || val === "" || val === "-") return "-";

  // Se é Date object, formatar
  if (val instanceof Date && !isNaN(val.getTime())) {
    return formatarHoraExcel(val);
  }

  // Se já é string, retornar direto
  return String(val);
}
function processarRegistro(resumo, ss, maquina, data, turno, evento, segundos) {
  if (segundos > 86400) return;
  if (segundos < 0) segundos = 0;
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina + "|" + dStr + "|" + turno;
  if (!resumo[chave]) resumo[chave] = { maquina: maquina, data: data, turno: turno, ligada: 0, desligada: 0, listaStop3: [], listaStop10: [], listaStop20: [], listaStop30: [] };

  if (evento === "TEMPO PRODUZINDO") {
    resumo[chave].ligada += segundos;
  } else {
    resumo[chave].desligada += segundos;
    if (segundos > 10) {
      if (!resumo[chave].qtdParadas) resumo[chave].qtdParadas = 0;
      resumo[chave].qtdParadas++;

      let adicionado = [];
      if (segundos > 180) {
        resumo[chave].listaStop3.push(segundos);
        adicionado.push(">3min");
      }
      if (segundos > 600) {
        resumo[chave].listaStop10.push(segundos);
        adicionado.push(">10min");
      }
      if (segundos > 1200) {
        resumo[chave].listaStop20.push(segundos);
        adicionado.push(">20min");
      }
      if (segundos > 1800) {
        resumo[chave].listaStop30.push(segundos);
        adicionado.push(">30min");
      }

      if (adicionado.length > 0) {
        Logger.log(`  → Parada ${segundos}s adicionada em: ${adicionado.join(", ")} | Total >3min: ${resumo[chave].listaStop3.length}`);
      }
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

// Função para testar busca histórica com datas específicas
function testarBuscaHistorica() {
  Logger.log("=== TESTE DE BUSCA HISTÓRICA ===");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pagina = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");

  if (!pagina) {
    Logger.log("⚠ Aba 'Página' não encontrada!");
    return;
  }

  Logger.log("\n1. DADOS NA ABA PÁGINA:");
  const dados = pagina.getDataRange().getValues();
  Logger.log("  Total de linhas: " + dados.length);

  // Mostra as primeiras 5 linhas com foco na data
  Logger.log("\n  Primeiras 5 linhas (Máquina | Turno | Data):");
  for (let i = 1; i < Math.min(6, dados.length); i++) {
    let maq = dados[i][0];
    let turno = dados[i][2];
    let dataVal = dados[i][3];
    let dataConvertida = lerDataBR(dataVal);

    Logger.log("    Linha " + i + ":");
    Logger.log("      Máquina: " + maq);
    Logger.log("      Turno: " + turno);
    Logger.log("      Data original: " + dataVal + " (tipo: " + typeof dataVal + ")");
    Logger.log("      Data convertida: " + dataConvertida);
    Logger.log("      Data formatada: " + Utilities.formatDate(dataConvertida, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss"));
  }

  // Testa busca com data de hoje
  Logger.log("\n2. TESTE DE BUSCA COM DATA DE HOJE:");
  const hoje = new Date();
  const hojeStr = Utilities.formatDate(hoje, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  Logger.log("  Data de hoje: " + hojeStr);

  // Pega a primeira máquina da aba para teste
  const maquinaTeste = dados.length > 1 ? String(dados[1][0]).trim() : "";
  if (maquinaTeste) {
    Logger.log("  Testando com máquina: " + maquinaTeste);
    const resultado = buscarHistorico(maquinaTeste, hojeStr, hojeStr);
    Logger.log("  Resultados encontrados: " + resultado.length);
    if (resultado.length > 0) {
      Logger.log("  Primeiro resultado: " + JSON.stringify(resultado[0]));
    }
  }

  // Testa busca sem filtro de data
  Logger.log("\n3. TESTE DE BUSCA SEM FILTRO DE DATA:");
  if (maquinaTeste) {
    const resultadoTotal = buscarHistorico(maquinaTeste, null, null);
    Logger.log("  Resultados encontrados: " + resultadoTotal.length);
    if (resultadoTotal.length > 0) {
      Logger.log("  Primeira data: " + resultadoTotal[0].data);
      Logger.log("  Última data: " + resultadoTotal[resultadoTotal.length - 1].data);
    }
  }

  Logger.log("\n=== FIM DO TESTE ===");
}

// Função para testar busca de histórico com data de hoje
function testarHistoricoHoje() {
  Logger.log("=== TESTE HISTÓRICO HOJE ===");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pagina") || ss.getSheetByName("Página");

  if (!sheet) {
    Logger.log("⚠ Aba 'Pagina' não encontrada");
    return;
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  // Pegar a primeira máquina
  const maquinaTeste = dados.length > 1 ? String(dados[1][0]).trim() : "";

  if (!maquinaTeste) {
    Logger.log("⚠ Nenhuma máquina encontrada");
    return;
  }

  // Data de hoje no formato yyyy-MM-dd (igual ao frontend)
  const hoje = new Date();
  const hojeStr = Utilities.formatDate(hoje, timezone, "yyyy-MM-dd");

  Logger.log("Máquina: " + maquinaTeste);
  Logger.log("Data de hoje: " + hojeStr);
  Logger.log("");

  // Mostrar todas as datas disponíveis para essa máquina
  Logger.log("Datas disponíveis para " + maquinaTeste + ":");
  for (let i = 1; i < dados.length; i++) {
    let maq = String(dados[i][0]).trim();
    if (maq === maquinaTeste) {
      let dataOriginal = dados[i][3];
      let dataConvertida = lerDataBR(dataOriginal);
      let dataStr = Utilities.formatDate(dataConvertida, timezone, "yyyy-MM-dd");
      let turno = dados[i][2];

      Logger.log("  Linha " + (i+1) + ": " + dataStr + " (" + turno + ")");
    }
  }

  Logger.log("");
  Logger.log("Chamando buscarHistorico('" + maquinaTeste + "', '" + hojeStr + "', '" + hojeStr + "')...");

  const resultado = buscarHistorico(maquinaTeste, hojeStr, hojeStr);

  Logger.log("");
  Logger.log("Resultado: " + resultado.length + " registros");

  if (resultado.length > 0) {
    Logger.log("Primeiro registro:");
    Logger.log(JSON.stringify(resultado[0], null, 2));
  }

  Logger.log("\n=== FIM DO TESTE ===");
}
