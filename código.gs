// ==========================================================
// CONFIGURAÇÃO GLOBAL
// ==========================================================

// ID DA PLANILHA
const ID_PLANILHA = "1jApQbNfM7gUyIL9b3I0CuEFePlnr2DMKeuARCIjrq7g";

// Função auxiliar para pegar a planilha correta (pelo ID ou ativa)
function getSS() {
  try {
    if (ID_PLANILHA && ID_PLANILHA !== "") {
      return SpreadsheetApp.openById(ID_PLANILHA);
    }
  } catch (e) {
    console.error("Erro ao abrir pelo ID: " + e.message);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

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
  // Trava de segurança para evitar conflito de escrita (Alta Escala)
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return ContentService.createTextOutput("BUSY"); }

  try {
    var maquina = e.parameter.maquina;
    var evento = e.parameter.evento;
    var duracao = e.parameter.duracao;
    
    var ss = getSS();
    var sheet = ss.getSheetByName("Página1");
    if (!sheet) sheet = ss.getActiveSheet();
    
    var dataAtual = new Date();
    var timezone = Session.getScriptTimeZone();
    var dataStr = Utilities.formatDate(dataAtual, timezone, "dd/MM/yyyy");
    var horaStr = Utilities.formatDate(dataAtual, timezone, "HH:mm:ss");
    
    sheet.appendRow([dataStr, horaStr, maquina, evento, duracao]);
    
    // O cálculo pesado (gerarRelatorioTurnos) foi removido daqui para não travar o ESP32.
    // Ele deve rodar via TRIGGER (Acionador) a cada 1 minuto.
    
    return ContentService.createTextOutput("OK");
    
  } catch (error) {
    return ContentService.createTextOutput("ERROR");
  } finally {
    lock.releaseLock(); 
  }
}

function verificarLogin(usuario, senha) {
  const ss = getSS();
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
// 2. FUNÇÕES DE DADOS (LEITURA TEMPO REAL)
// ==========================================================

function buscarDadosTempoReal() {
  try {
    var ss = getSS();
    if (!ss) throw new Error("Não foi possível abrir a planilha.");

    var sheetDados = ss.getSheetByName("Página1");
    if (!sheetDados) throw new Error("Erro: Aba 'Página1' não encontrada");

    var dados = sheetDados.getDataRange().getValues();
    if (!dados || dados.length <= 1) return JSON.stringify({});

    var sheetTurnos = ss.getSheetByName("TURNOS");
    var dadosTurnos = sheetTurnos ? sheetTurnos.getDataRange().getValues() : [];
    var configTurnos = {};
    var mapaMetas = {};

    for (var i = 1; i < dadosTurnos.length; i++) {
      if (dadosTurnos[i][0]) {
        var maqNome = String(dadosTurnos[i][0]).trim();
        configTurnos[maqNome] = [
           { nome: "Turno 1", inicio: dadosTurnos[i][1], fim: dadosTurnos[i][2] },
           { nome: "Turno 2", inicio: dadosTurnos[i][3], fim: dadosTurnos[i][4] },
           { nome: "Turno 3", inicio: dadosTurnos[i][5], fim: dadosTurnos[i][6] }
        ];

        // Ler metas por turno da aba TURNOS (colunas H, I, J = índices 7, 8, 9)
        // Suporta tanto números decimais quanto formato de horário (HH:MM)
        mapaMetas[maqNome] = {
          "Turno 1": converterMetaParaHoras(dadosTurnos[i][7]),
          "Turno 2": converterMetaParaHoras(dadosTurnos[i][8]),
          "Turno 3": converterMetaParaHoras(dadosTurnos[i][9])
        };
      }
    }

    var sheetDadosConfig = ss.getSheetByName("DADOS");
    var mapaFamilias = {};

    if (sheetDadosConfig) {
      var dConfig = sheetDadosConfig.getDataRange().getValues();
      if (dConfig.length > 0) {
        var h = dConfig[0].map(function(c) { return String(c).toUpperCase().trim(); });
        var idxM = h.indexOf("MÁQUINAS");
        var idxF = h.findIndex(function(x) { return x.includes("FAMÍLIA") || x.includes("FAMILIA"); });

        if (idxM > -1 && idxF > -1) {
          for (var i = 1; i < dConfig.length; i++) {
            var m = String(dConfig[i][idxM]).trim();
            mapaFamilias[m] = dConfig[i][idxF] || "GERAL";
          }
        }
      }
    }

    var statusMaquinas = {};
    var agora = new Date();
    var timezone = ss.getSpreadsheetTimeZone();
    var sheetPainel = ss.getSheetByName("PAINEL");
    var dadosPainel = sheetPainel ? sheetPainel.getDataRange().getValues() : [];

    for (var i = dados.length - 1; i > 0; i--) {
      var linha = dados[i];
      var maquina = String(linha[2]).trim();
      if (!maquina) continue;

      if (!statusMaquinas[maquina]) {
        var infoTurnoAtual = descobrirTurnoCompleto(agora, maquina, configTurnos);
        var nomeTurnoAtual = infoTurnoAtual ? infoTurnoAtual.nome : "Fora de Turno";
        var dataProducaoAtual = null;

        if (infoTurnoAtual) {
          dataProducaoAtual = new Date(agora);
          var horaAgora = agora.getHours();
          var horaInicioPrimeiroTurno = obterHoraInicioPrimeiroTurno(maquina, configTurnos);
          if (horaAgora < horaInicioPrimeiroTurno) {
            dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
          } else if (infoTurnoAtual.cruzaMeiaNoite && horaAgora < Math.floor(infoTurnoAtual.minInicio / 60)) {
            dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
          }
          dataProducaoAtual.setHours(0,0,0,0);
        }

        var dataReg = lerDataBR(linha[0]);
        var timestampFinal;
        var timePart = new Date(linha[1]);

        if (!isNaN(dataReg.getTime()) && !isNaN(timePart.getTime())) {
           var d = new Date(dataReg);
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
          refDataProducao: dataProducaoAtual ? dataProducaoAtual.getTime() : null,
          horarioInicio: "",
          primeiraHora: null,
          primeiraDuracao: 0,
          metaTurno1: mapaMetas[maquina] ? mapaMetas[maquina]["Turno 1"] : 0,
          metaTurno2: mapaMetas[maquina] ? mapaMetas[maquina]["Turno 2"] : 0,
          metaTurno3: mapaMetas[maquina] ? mapaMetas[maquina]["Turno 3"] : 0,
          metaTurnoAtual: mapaMetas[maquina] && mapaMetas[maquina][nomeTurnoAtual] ? mapaMetas[maquina][nomeTurnoAtual] : 0
        };
      }
      
      var ref = statusMaquinas[maquina];
      if (ref.refNomeTurno !== "Fora de Turno" && ref.refDataProducao !== null) {
        var dataReg = lerDataBR(linha[0]);
        var horaRegObj = new Date(linha[1]);

        if (!isNaN(dataReg.getTime()) && !isNaN(horaRegObj.getTime())) {
          var fullDateReg = new Date(dataReg);
          fullDateReg.setHours(horaRegObj.getHours(), horaRegObj.getMinutes(), horaRegObj.getSeconds());

          var infoTurnoReg = descobrirTurnoCompleto(fullDateReg, maquina, configTurnos);

          if (infoTurnoReg && infoTurnoReg.nome === ref.refNomeTurno) {
            var dataProdReg = new Date(dataReg);
            var h = fullDateReg.getHours();
            var horaInicioPrimeiroTurno = obterHoraInicioPrimeiroTurno(maquina, configTurnos);

            if (h < horaInicioPrimeiroTurno) {
              dataProdReg.setDate(dataProdReg.getDate() - 1);
            } else if (infoTurnoReg.cruzaMeiaNoite && h < Math.floor(infoTurnoReg.minInicio / 60)) {
              dataProdReg.setDate(dataProdReg.getDate() - 1);
            }

            dataProdReg.setHours(0,0,0,0);

            if (dataProdReg.getTime() === ref.refDataProducao) {
               var duracao = parseDuration(linha[4]);

               // CLAMPAR duração para que só conte a parte do evento que está DENTRO do turno
               // Isso corrige o bug de eventos que cruzam turnos (ex: evento que começa no Turno 3 e termina no Turno 1)
               var fimEvento = fullDateReg.getTime();
               var inicioEvento = fimEvento - (duracao * 1000);

               // Calcular início e fim do turno atual
               var dataInicioTurnoEvento = new Date(dataProdReg);
               var horaInicioTurnoEvento = new Date(infoTurnoReg.inicio);
               dataInicioTurnoEvento.setHours(horaInicioTurnoEvento.getHours(), horaInicioTurnoEvento.getMinutes(), horaInicioTurnoEvento.getSeconds(), 0);

               var dataFimTurnoEvento = new Date(dataProdReg);
               var horaFimTurnoEvento = new Date(infoTurnoReg.fim);
               dataFimTurnoEvento.setHours(horaFimTurnoEvento.getHours(), horaFimTurnoEvento.getMinutes(), horaFimTurnoEvento.getSeconds(), 0);

               // Se turno cruza meia-noite, ajustar data de fim
               if (infoTurnoReg.cruzaMeiaNoite) {
                 dataFimTurnoEvento.setDate(dataFimTurnoEvento.getDate() + 1);
               }

               var inicioTurno = dataInicioTurnoEvento.getTime();
               var fimTurno = dataFimTurnoEvento.getTime();

               // Clampar o evento para os limites do turno
               var inicioEfetivo = Math.max(inicioEvento, inicioTurno);
               var fimEfetivo = Math.min(fimEvento, fimTurno);

               // Calcular duração clampada (em segundos)
               var duracaoClampada = Math.max(0, Math.floor((fimEfetivo - inicioEfetivo) / 1000));

               if (linha[3] === "TEMPO PRODUZINDO") ref.totalProduzindo += duracaoClampada;
               else if (linha[3] === "TEMPO PARADA") ref.totalParada += duracaoClampada;

               // Rastrear primeiro evento (loop vai de trás pra frente, então sempre sobrescreve)
               ref.primeiraHora = fullDateReg;
               ref.primeiraDuracao = duracaoClampada;
            }
          }
        }
      }
    }

    // Buscar horário de início do PAINEL
    if (dadosPainel.length > 0) {
       for (var maq in statusMaquinas) {
          var info = statusMaquinas[maq];
          if (info.refNomeTurno !== "Fora de Turno" && info.refDataProducao !== null) {
            var dataBusca = new Date(info.refDataProducao);
            var dataBuscaStr = Utilities.formatDate(dataBusca, timezone, "dd/MM/yyyy");

            for (var i = 1; i < dadosPainel.length; i++) {
              var linha = dadosPainel[i];
              if (!linha || linha.length < 16) continue;

              var maqPainel = String(linha[0] || "").trim();
              var turnoPainel = String(linha[1] || "").trim();
              var dataPainelStr = "";

              if (linha[2] instanceof Date) {
                dataPainelStr = Utilities.formatDate(linha[2], timezone, "dd/MM/yyyy");
              } else {
                dataPainelStr = String(linha[2] || "").trim();
              }

              if (maqPainel === maq && turnoPainel === info.refNomeTurno && dataPainelStr === dataBuscaStr) {
                var horarioInicio = linha[15];
                if (horarioInicio && horarioInicio !== "") {
                  if (horarioInicio instanceof Date) {
                    info.horarioInicio = Utilities.formatDate(horarioInicio, timezone, "HH:mm:ss");
                  } else {
                    info.horarioInicio = String(horarioInicio);
                  }
                }
                break;
              }
            }
          }
       }
    }

    // Calcular gap inicial (tempo entre início do turno e primeiro evento)
    for (var maq in statusMaquinas) {
      var info = statusMaquinas[maq];
      if (info.refNomeTurno !== "Fora de Turno" && info.refDataProducao !== null && info.primeiraHora) {
        var configMaq = configTurnos[maq];
        if (configMaq) {
          var turnoConfig = configMaq.find(function(t) { return t.nome === info.refNomeTurno; });
          if (turnoConfig && turnoConfig.inicio) {
            var dataInicioTurno = new Date(info.refDataProducao);
            var horaInicioTurno = new Date(turnoConfig.inicio);

            if (!isNaN(horaInicioTurno.getTime()) && !isNaN(dataInicioTurno.getTime())) {
              dataInicioTurno.setHours(horaInicioTurno.getHours(), horaInicioTurno.getMinutes(), horaInicioTurno.getSeconds(), 0);

              var horaPrimeiroRegistro = new Date(info.primeiraHora);

              if (!isNaN(horaPrimeiroRegistro.getTime())) {
                // Calcular o início efetivo do evento (subtraindo a duração)
                var inicioEfetivoEvento = new Date(horaPrimeiroRegistro.getTime() - (info.primeiraDuracao * 1000));

                // CLAMPAR: Se o evento começou antes do turno, considerar que começou no início do turno
                // Isso evita gap negativo quando eventos cruzam turnos
                var inicioEfetivoClampado = Math.max(inicioEfetivoEvento.getTime(), dataInicioTurno.getTime());

                var diferencaSegundos = Math.floor((inicioEfetivoClampado - dataInicioTurno.getTime()) / 1000);

                // Se há gap de mais de 60 segundos, adicionar como parada
                if (diferencaSegundos >= 60) {
                  info.totalParada += diferencaSegundos;
                }
              }
            }
          }
        }
      }
    }

    return JSON.stringify(statusMaquinas);

  } catch (error) {
    console.error("ERRO CRÍTICO: " + error.message);
    return JSON.stringify({});
  }
}

// ==========================================================
// 3. BUSCA DADOS PARA GRÁFICO POR TURNO (GRÁFICOS)
// ==========================================================

function buscarDadosGrafico(maquinaNome) {
  const ss = getSS();
  const sheet = ss.getSheetByName("PAINEL");

  if (!sheet) return { erro: "Aba PAINEL não encontrada" };

  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  const dataLimite = new Date();
  dataLimite.setDate(hoje.getDate() - 30);
  dataLimite.setHours(0,0,0,0);

  const timezone = ss.getSpreadsheetTimeZone();
  const labels = [];

  const output = { "Turno 1": [], "Turno 2": [], "Turno 3": [] };
  const mapIndices = {};

  // Buscar metas da máquina na aba TURNOS (colunas H, I, J = índices 7, 8, 9)
  const sheetTurnos = ss.getSheetByName("TURNOS");
  let metaTurno1 = 0, metaTurno2 = 0, metaTurno3 = 0;

  if (sheetTurnos) {
    const dadosTurnos = sheetTurnos.getDataRange().getValues();
    for (let i = 1; i < dadosTurnos.length; i++) {
      if (String(dadosTurnos[i][0]).trim() === maquinaNome) {
        metaTurno1 = converterMetaParaHoras(dadosTurnos[i][7]);
        metaTurno2 = converterMetaParaHoras(dadosTurnos[i][8]);
        metaTurno3 = converterMetaParaHoras(dadosTurnos[i][9]);
        break;
      }
    }
  }

  // Criar array de dias dos últimos 30 dias
  let idx = 0;
  for (let d = new Date(dataLimite); d <= hoje; d.setDate(d.getDate() + 1)) {
    let k = Utilities.formatDate(d, timezone, "dd/MM");
    labels.push(k);
    output["Turno 1"].push(0);
    output["Turno 2"].push(0);
    output["Turno 3"].push(0);
    mapIndices[k] = idx++;
  }

  // Processar dados do PAINEL
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqLinha = String(linha[0]).trim();

    if (maqLinha !== maquinaNome) continue;

    let turno = String(linha[1]).trim();
    let dataLinha = lerDataBR(linha[2]);

    if (dataLinha < dataLimite) continue;

    let dataStr = Utilities.formatDate(dataLinha, timezone, "dd/MM");
    let index = mapIndices[dataStr];

    if (index === undefined) continue;

    // Pegar tempo ligada (coluna 3) - já está em formato Excel (número decimal)
    let tempoLigada = linha[3];
    let segundos = converterParaSegundos(tempoLigada);
    let horas = parseFloat((segundos / 3600).toFixed(2));

    if (output[turno] && index !== undefined) {
      // Substituir ao invés de acumular (para evitar duplicações)
      output[turno][index] = horas;
    }
  }

  // Arredondar valores finais para 2 casas decimais
  ["Turno 1", "Turno 2", "Turno 3"].forEach(t => {
    output[t] = output[t].map(h => parseFloat(h.toFixed(2)));
  });

  return {
    labels: labels,
    t1: output["Turno 1"],
    t2: output["Turno 2"],
    t3: output["Turno 3"],
    maquina: maquinaNome,
    metaTurno1: metaTurno1,
    metaTurno2: metaTurno2,
    metaTurno3: metaTurno3
  };
}

// ==========================================================
// 4. FUNÇÕES AUXILIARES E LISTAS
// ==========================================================

function buscarParadasTurnoAtual(maquina, turnoNome, dataProducao) {
  const ss = getSS();
  const sheet = ss.getSheetByName("PAINEL");

  if (!sheet) {
    Logger.log("ERRO: Aba 'PAINEL' não encontrada");
    return [];
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  const dataBusca = new Date(dataProducao);
  const dataBuscaStr = Utilities.formatDate(dataBusca, timezone, "dd/MM/yyyy");

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqLinha = String(linha[0]).trim();
    let turnoLinha = String(linha[1]).trim();
    let dataLinha = lerDataBR(linha[2]);
    let dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy");

    if (maqLinha === maquina && turnoLinha === turnoNome && dataLinhaStr === dataBuscaStr) {
      const paradas = [];
      const temposStr = String(linha[5] || "").trim();

      if (temposStr && temposStr !== "-" && temposStr !== "") {
        const tempos = temposStr.split(",");
        tempos.forEach(tempo => {
          tempo = tempo.trim();
          let duracaoPura = tempo.split(" ")[0];
          
          if (duracaoPura && duracaoPura.includes(":")) {
            const partes = duracaoPura.split(":");
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
      return paradas;
    }
  }
  return [];
}

function buscarRelatorioFamilia(familia, dataInicio, dataFim) {
  const ss = getSS();
  const sheet = ss.getSheetByName("PAINEL"); 

  if (!sheet) {
    return { erro: "Aba 'PAINEL' não encontrada" };
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
          mapaFamilias[m] = String(dConfig[i][idxF] || "GERAL").trim();
        }
      }
    }
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();

  const maquinasPorFamilia = {};
  let totalRodandoSeg = 0;
  let totalParadoSeg = 0;
  let totalParadasCriticas = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const nomeMaquina = String(linha[0]).trim();
    const familiaMaquina = mapaFamilias[nomeMaquina] || "";
    
    if (familiaMaquina.toUpperCase() !== familia.toUpperCase()) {
      continue;
    }

    const turno = String(linha[1]).trim();
    const dataLinha = lerDataBR(linha[2]);
    const dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "yyyy-MM-dd");

    if (dataInicio && dataLinhaStr < dataInicio) continue;
    if (dataFim && dataLinhaStr > dataFim) continue;

    const ligada = linha[3];
    const desligada = linha[4];
    const paradasCriticas = String(linha[5] || "").trim();

    let ligadaSeg = converterParaSegundos(ligada);
    let desligadaSeg = converterParaSegundos(desligada);

    let qtdParadas = 0;
    if (paradasCriticas && paradasCriticas !== "-") {
      qtdParadas = paradasCriticas.split(",").filter(p => p.trim().length > 0).length;
    }

    totalRodandoSeg += ligadaSeg;
    totalParadoSeg += desligadaSeg;
    totalParadasCriticas += qtdParadas;

    if (!maquinasPorFamilia[nomeMaquina]) {
      maquinasPorFamilia[nomeMaquina] = [];
    }

    maquinasPorFamilia[nomeMaquina].push({
      data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
      turno: turno,
      ligada: formatarHoraExcel(ligada),
      desligada: formatarHoraExcel(desligada),
      paradas3min: formatarCelulaParada(linha[5]), 
      paradas10min: formatarCelulaParada(linha[6]),
      paradas20min: formatarCelulaParada(linha[7]), 
      paradas30min: formatarCelulaParada(linha[8])
    });
  }

  const maquinas = [];
  for (let nomeMaq in maquinasPorFamilia) {
    maquinas.push({
      nome: nomeMaq,
      turnos: maquinasPorFamilia[nomeMaq]
    });
  }

  maquinas.sort((a, b) => a.nome.localeCompare(b.nome));

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

function buscarHistorico(maquinaFiltro, dataInicio, dataFim) {
  const ss = getSS();
  const sheet = ss.getSheetByName("PAINEL");

  if (!sheet) {
    Logger.log("⚠ Aba 'PAINEL' não encontrada");
    return [];
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();
  const resultados = [];

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let maqLinha = String(linha[0]).trim();

    if (maqLinha !== maquinaFiltro) continue;

    let dataLinha = lerDataBR(linha[2]);
    let dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "yyyy-MM-dd");

    if (dataInicio && dataLinhaStr < dataInicio) continue;
    if (dataFim && dataLinhaStr > dataFim) continue;

    try {
      const registro = {
        turno: String(linha[1] || "-"),
        data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
        ligada: formatarHoraExcel(linha[3]),
        desligada: formatarHoraExcel(linha[4]),
        paradas3min: formatarCelulaParada(linha[5]),
        paradas10min: formatarCelulaParada(linha[6]),
        paradas20min: formatarCelulaParada(linha[7]),
        paradas30min: formatarCelulaParada(linha[8]),
        motivo: String(linha[10] || "-"),
        servico: String(linha[11] || "-"),
        pecas: String(linha[12] || "-"),
        custoMO: typeof linha[9] === 'number' ? linha[9] : 0,
        custoPecas: typeof linha[13] === 'number' ? linha[13] : 0,
        obs: String(linha[14] || "-")
      };
      resultados.push(registro);
    } catch (erro) {
      Logger.log(" ⚠ ERRO ao processar linha " + i + ": " + erro.message);
    }
  }

  return resultados;
}

function buscarListasDropdown() {
  const ss = getSS();
  const sheet = ss.getSheetByName("DADOS");
  if (!sheet) return { motivos: [], servicos: [], familias: [] };
  
  const dados = sheet.getDataRange().getValues();
  let motivos = [], servicos = [], familias = [];
  
  if (dados.length > 0) {
    for (let c = 0; c < dados[0].length; c++) {
      let head = String(dados[0][c]).toUpperCase().trim();
      
      if (head === "MOTIVO DA PARADA") {
         for(let i=1; i<dados.length; i++) if(dados[i][c]) motivos.push(dados[i][c]);
      }
      if (head === "SERVIÇOS REALIZADOS") {
         for(let i=1; i<dados.length; i++) if(dados[i][c]) servicos.push(dados[i][c]);
      }
      if (head.includes("FAMÍLIA") || head.includes("FAMILIA")) {
         for(let i=1; i<dados.length; i++) {
             let f = String(dados[i][c]).trim();
             if(f && !familias.includes(f)) familias.push(f);
         }
      }
    }
  }
  return { motivos: motivos, servicos: servicos, familias: familias.sort() };
}

function salvarApontamento(dadosForm) {
  const ss = getSS();
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
// 5. GERADOR DE RELATÓRIO
// ==========================================================

function gerarRelatorioTurnos() {
  const ss = getSS();
  const sheetDados = ss.getSheetByName("Página1");
  const sheetTurnos = ss.getSheetByName("TURNOS");
  const sheetPainel = ss.getSheetByName("PAINEL");
  const sheetCustos = ss.getSheetByName("DADOS");
  if (!sheetDados || !sheetTurnos || !sheetPainel) return;
  
  const dadosPainelAntigo = sheetPainel.getDataRange().getValues();
  const mapaPainelExistente = {}; 
  const agora = new Date();
  const dataHojeStr = Utilities.formatDate(agora, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  
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
  
  const dadosBrutos = sheetDados.getDataRange().getValues();
  const resumo = {};
  const mapUltimoFim = {};

  for (let i = 1; i < dadosBrutos.length; i++) {
    let linha = dadosBrutos[i];
    let dataOriginal = lerDataBR(linha[0]);
    let hora = linha[1];
    let maquinaRaw = linha[2];
    let evento = linha[3];
    let duracaoRaw = linha[4];
    let duracao = parseDuration(duracaoRaw);

    if (!maquinaRaw || !hora) continue;

    let maquina = String(maquinaRaw).trim();
    let nomeEvento = evento ? String(evento).trim() : "";

    if (nomeEvento !== "TEMPO PRODUZINDO" && nomeEvento !== "TEMPO PARADA") continue;

    let infoTurno = descobrirTurnoCompleto(hora, maquina, configTurnos);
    if (infoTurno) {
      let dataFim = new Date(dataOriginal);
      if (isNaN(dataFim.getTime())) dataFim = new Date();
      let horaObj = new Date(hora);
      if (!isNaN(horaObj.getTime())) { 
        dataFim.setHours(horaObj.getHours(), horaObj.getMinutes(), horaObj.getSeconds(), 0); 
      }
      
      let dataProducao = new Date(dataFim);
      let horaReg = dataFim.getHours();
      let horaInicioPrimeiroTurno = obterHoraInicioPrimeiroTurno(maquina, configTurnos);

      if (horaReg < horaInicioPrimeiroTurno) {
        dataProducao.setDate(dataProducao.getDate() - 1);
      } else if (infoTurno.cruzaMeiaNoite && horaReg < Math.floor(infoTurno.minInicio / 60)) {
        dataProducao.setDate(dataProducao.getDate() - 1);
      }
      
      // CLAMPING
      if (configTurnos[maquina]) {
        let turnoConfig = configTurnos[maquina].find(t => t.nome === infoTurno.nome);
        if (turnoConfig && turnoConfig.inicio) {
           let horaInicioConfig = new Date(turnoConfig.inicio);
           let dataInicioTurnoAbsoluto = new Date(dataProducao);
           dataInicioTurnoAbsoluto.setHours(horaInicioConfig.getHours(), horaInicioConfig.getMinutes(), 0, 0);
           
           let inicioRealEvento = new Date(dataFim.getTime() - (duracao * 1000));
           
           if (inicioRealEvento.getTime() < dataInicioTurnoAbsoluto.getTime()) {
              let novaDuracao = Math.floor((dataFim.getTime() - dataInicioTurnoAbsoluto.getTime()) / 1000);
              if (novaDuracao < 0) novaDuracao = 0; 
              duracao = novaDuracao;
           }
        }
      }

      let dProdStr = Utilities.formatDate(dataProducao, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      let chaveGap = maquina + "|" + dProdStr + "|" + infoTurno.nome;

      if (mapUltimoFim[chaveGap]) {
        let ultimoFim = mapUltimoFim[chaveGap]; 
        let inicioEventoAtual = new Date(dataFim.getTime() - (duracao * 1000));
        let gapIntermediario = Math.floor((inicioEventoAtual.getTime() - ultimoFim.getTime()) / 1000);
        
        if (gapIntermediario > 60) {
           registrarGapIntermediario(resumo, ss, maquina, dataProducao, infoTurno.nome, gapIntermediario, ultimoFim);
        }
      }
      mapUltimoFim[chaveGap] = dataFim;

      processarRegistro(resumo, ss, maquina, dataProducao, infoTurno.nome, nomeEvento, duracao, dataFim);
    }
  }

  // Adicionar turnos sem eventos como paradas completas
  const diasParaVerificar = 7; // Últimos 7 dias
  const dataLimiteVerificacao = new Date(agora);
  dataLimiteVerificacao.setDate(dataLimiteVerificacao.getDate() - diasParaVerificar);
  dataLimiteVerificacao.setHours(0, 0, 0, 0);

  for (let maquinaKey in configTurnos) {
    let turnosConfig = configTurnos[maquinaKey];

    // Para cada dia dos últimos X dias
    for (let d = new Date(dataLimiteVerificacao); d <= agora; d.setDate(d.getDate() + 1)) {
      let dataVerificacao = new Date(d);
      dataVerificacao.setHours(0, 0, 0, 0);
      let dataStr = Utilities.formatDate(dataVerificacao, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

      // Para cada turno configurado
      turnosConfig.forEach(turnoConfig => {
        let chave = maquinaKey + "|" + dataStr + "|" + turnoConfig.nome;

        // Se não existe entrada no resumo = máquina parada o turno inteiro
        if (!resumo[chave]) {
          // Calcular duração do turno em segundos
          let horaInicio = new Date(turnoConfig.inicio);
          let horaFim = new Date(turnoConfig.fim);

          let minInicio = horaInicio.getHours() * 60 + horaInicio.getMinutes();
          let minFim = horaFim.getHours() * 60 + horaFim.getMinutes();

          let duracaoTurnoSegundos;
          if (minFim > minInicio) {
            duracaoTurnoSegundos = (minFim - minInicio) * 60;
          } else {
            // Turno cruza meia-noite
            duracaoTurnoSegundos = ((1440 - minInicio) + minFim) * 60;
          }

          // Criar entrada com turno completo como parada
          let horaInicioTurno = new Date(dataVerificacao);
          horaInicioTurno.setHours(horaInicio.getHours(), horaInicio.getMinutes(), 0, 0);

          let horaFimTurno = new Date(dataVerificacao);
          horaFimTurno.setHours(horaFim.getHours(), horaFim.getMinutes(), 0, 0);
          if (minFim <= minInicio) {
            horaFimTurno.setDate(horaFimTurno.getDate() + 1);
          }

          // Formatar horários para a parada crítica
          let hIni = Utilities.formatDate(horaInicioTurno, ss.getSpreadsheetTimeZone(), "HH:mm");
          let hFim = Utilities.formatDate(horaFimTurno, ss.getSpreadsheetTimeZone(), "HH:mm");

          let objParada = { s: duracaoTurnoSegundos, h: hIni, f: hFim };

          resumo[chave] = {
            maquina: maquinaKey,
            data: dataVerificacao,
            turno: turnoConfig.nome,
            ligada: 0,
            desligada: duracaoTurnoSegundos,
            listaStop3: duracaoTurnoSegundos > 180 ? [objParada] : [],
            listaStop10: duracaoTurnoSegundos > 600 ? [objParada] : [],
            listaStop20: duracaoTurnoSegundos > 1200 ? [objParada] : [],
            listaStop30: duracaoTurnoSegundos > 1800 ? [objParada] : [],
            horarioInicio: null,
            primeiraHora: null,
            primeiraDuracao: 0,
            ultimaHora: null
          };
        }
      });
    }
  }

  const agoraTimestamp = agora.getTime();

  for (let chave in resumo) {
    let item = resumo[chave];
    
    try {
      if (item.ultimaHora) {
        let configMaq = configTurnos[item.maquina];
        if (configMaq) {
          let turnoConfig = configMaq.find(t => t.nome === item.turno);
          if (turnoConfig && turnoConfig.fim) {
            let dataFimTurno = new Date(item.data);
            let horaFimConfig = new Date(turnoConfig.fim);
            let horaInicioConfig = new Date(turnoConfig.inicio);
            
            let minI = horaInicioConfig.getHours() * 60 + horaInicioConfig.getMinutes();
            let minF = horaFimConfig.getHours() * 60 + horaFimConfig.getMinutes();
            
            if (minI > minF) {
               dataFimTurno.setDate(dataFimTurno.getDate() + 1);
            }
            
            dataFimTurno.setHours(horaFimConfig.getHours(), horaFimConfig.getMinutes(), 0, 0);
            
            if (agoraTimestamp > dataFimTurno.getTime()) {
               let ultimaHoraReg = new Date(item.ultimaHora);
               
               if (!isNaN(ultimaHoraReg.getTime())) {
                 let gapFinalSegundos = Math.floor((dataFimTurno.getTime() - ultimaHoraReg.getTime()) / 1000);
                 
                 if (gapFinalSegundos > 60) {
                    item.desligada += gapFinalSegundos;
                    
                    let hIni = Utilities.formatDate(ultimaHoraReg, ss.getSpreadsheetTimeZone(), "HH:mm");
                    let hFim = Utilities.formatDate(dataFimTurno, ss.getSpreadsheetTimeZone(), "HH:mm");
                    
                    let objGap = { s: gapFinalSegundos, h: hIni, f: hFim };

                    if (gapFinalSegundos > 180) item.listaStop3.push(objGap);
                    if (gapFinalSegundos > 600) item.listaStop10.push(objGap);
                    if (gapFinalSegundos > 1200) item.listaStop20.push(objGap);
                    if (gapFinalSegundos > 1800) item.listaStop30.push(objGap);
                 }
               }
            }
          }
        }
      }
    } catch(e) {
      Logger.log("Erro no pós-processamento de fim de turno: " + e.message);
    }
  }

  const linhasSaida = [];
  const SEGUNDOS_DIA = 86400;

  for (let chave in resumo) {
    let item = resumo[chave];
    let rowFinal = [];

    try {
      if (item.primeiraHora) {
        let configMaq = configTurnos[item.maquina];
        if (configMaq) {
          let turnoConfig = configMaq.find(t => t.nome === item.turno);
          if (turnoConfig && turnoConfig.inicio) {
            let dataInicioTurno = new Date(item.data);
            let horaInicioTurno = new Date(turnoConfig.inicio);
            if (!isNaN(horaInicioTurno.getTime()) && !isNaN(dataInicioTurno.getTime())) {
              dataInicioTurno.setHours(horaInicioTurno.getHours(), horaInicioTurno.getMinutes(), horaInicioTurno.getSeconds(), 0);

              let horaPrimeiroRegistro = new Date(item.primeiraHora);

              if (!isNaN(horaPrimeiroRegistro.getTime())) {
                let inicioEfetivoEvento = new Date(horaPrimeiroRegistro.getTime() - (item.primeiraDuracao * 1000));
                let diferencaSegundos = Math.floor((inicioEfetivoEvento.getTime() - dataInicioTurno.getTime()) / 1000);

                if (diferencaSegundos >= 60) {
                  item.desligada += diferencaSegundos;
                  let hIni = Utilities.formatDate(dataInicioTurno, ss.getSpreadsheetTimeZone(), "HH:mm");
                  let hFim = Utilities.formatDate(inicioEfetivoEvento, ss.getSpreadsheetTimeZone(), "HH:mm");
                  
                  let objGap = { s: diferencaSegundos, h: hIni, f: hFim };

                  if (diferencaSegundos > 180) item.listaStop3.unshift(objGap);
                  if (diferencaSegundos > 600) item.listaStop10.unshift(objGap);
                  if (diferencaSegundos > 1200) item.listaStop20.unshift(objGap);
                  if (diferencaSegundos > 1800) item.listaStop30.unshift(objGap);
                }
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log(`Erro ao calcular parada inicial para ${item.maquina} ${item.turno}: ${e.message}`);
    }

    // SEMPRE recalcular tempos (ligada, desligada, paradas) mesmo para dias antigos
    // MAS preservar apontamentos manuais (motivo, serviço, peças, custos, obs) se existirem
    let manual = { motivo: "", servico: "", pecas: "", custoPecas: "", obs: "" };
    let horarioInicioAnterior = "";
    if (mapaPainelExistente[chave]) {
      let old = mapaPainelExistente[chave];
      manual = { motivo: old[10], servico: old[11], pecas: old[12], custoPecas: old[13], obs: old[14] };
      horarioInicioAnterior = old[15] || "";
    }

    let tempoParadoLiq = Math.max(0, item.desligada - 3600);
    let custoMO = 0;
    let valLigada = Math.max(0, item.ligada) / SEGUNDOS_DIA;
    let valDesligada = Math.max(0, item.desligada) / SEGUNDOS_DIA;

    let horarioInicioStr = horarioInicioAnterior;
    if (item.horarioInicio) {
      let horaObj = new Date(item.horarioInicio);
      if (!isNaN(horaObj.getTime())) {
        horarioInicioStr = Utilities.formatDate(horaObj, ss.getSpreadsheetTimeZone(), "HH:mm:ss");
      }
    }

    if(item.listaStop3.length > 0) item.listaStop3.sort((a,b) => b.s - a.s);
    if(item.listaStop10.length > 0) item.listaStop10.sort((a,b) => b.s - a.s);
    if(item.listaStop20.length > 0) item.listaStop20.sort((a,b) => b.s - a.s);
    if(item.listaStop30.length > 0) item.listaStop30.sort((a,b) => b.s - a.s);

    rowFinal = [
       item.maquina, item.turno, item.data, valLigada, valDesligada,
       formatarListaTempos(item.listaStop3), formatarListaTempos(item.listaStop10), formatarListaTempos(item.listaStop20), formatarListaTempos(item.listaStop30),
       custoMO, manual.motivo, manual.servico, manual.pecas, manual.custoPecas, manual.obs, horarioInicioStr
    ];
    
    linhasSaida.push(rowFinal);
    delete mapaPainelExistente[chave];
  }

  for (let chave in mapaPainelExistente) {
    linhasSaida.push(mapaPainelExistente[chave]);
  }
  
  sheetPainel.clearContents();
  const cabecalho = [[ "MÁQUINAS", "TURNO", "DATA", "LIGADA", "DESLIGADA", "TEMPOS > 3 min", "TEMPOS > 10 min", "TEMPOS > 20 min", "TEMPOS > 30 min", "CUSTO MÃO DE OBRA", "MOTIVO DA PARADA", "SERVIÇOS REALIZADOS", "PEÇAS TROCADAS", "CUSTO PEÇAS", "OBSERVAÇÃO", "HORÁRIO INÍCIO" ]];
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
  const ss = getSS();
  const sheet = ss.getSheetByName("Página1");
  if (!sheet) return;
  const dados = sheet.getDataRange().getValues();
  const hoje = new Date();
  const limite = new Date(hoje.getTime() - (45 * 24 * 60 * 60 * 1000));
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
  try {
    if (typeof valor === 'string') {
      let partes = valor.split('/');
      if (partes.length === 3) return new Date(parseInt(partes[2]), parseInt(partes[1])-1, parseInt(partes[0]));
      partes = valor.split('-');
      if (partes.length === 3) return new Date(parseInt(partes[0]), parseInt(partes[1])-1, parseInt(partes[2]));
    }
  } catch (e) {}
  return new Date();
}
function parseDuration(raw) {
  if (typeof raw === 'number') return raw;
  if (raw instanceof Date) return raw.getHours() * 3600 + raw.getMinutes() * 60 + raw.getSeconds();
  if (typeof raw === 'string') {
    let str = raw.trim();
    if (str.includes(':')) {
      let partes = str.split(':');
      if (partes.length === 3) return parseInt(partes[0]) * 3600 + parseInt(partes[1]) * 60 + parseInt(partes[2]);
      else if (partes.length === 2) return parseInt(partes[0]) * 60 + parseInt(partes[1]);
    }
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

     // Validar se os valores são válidos
     if (i === 0 && f === 0) continue; // Turno não configurado

     let cruza = i > f;
     // Usar <= para incluir o último minuto do turno
     if (!cruza) { if (min >= i && min <= f) return { nome: t.nome, minInicio: i, minFim: f, cruzaMeiaNoite: false }; }
     else { if (min >= i || min <= f) return { nome: t.nome, minInicio: i, minFim: f, cruzaMeiaNoite: true }; }
  }
  return null;
}
function obterHoraInicioPrimeiroTurno(maq, config) {
  // Retorna a hora de início do primeiro turno configurado para a máquina
  // Usado para determinar quando muda o dia de produção (ao invés de hardcoded "< 7")
  let c = config[maq] || config[String(maq).trim()];
  if (!c) return 7; // Fallback para 7 se não houver config

  for (let t of c) {
    if (t.inicio) {
      let horaInicio = new Date(t.inicio);
      if (!isNaN(horaInicio.getTime())) {
        return horaInicio.getHours();
      }
    }
  }
  return 7; // Fallback padrão
}
function converterMetaParaHoras(valor) {
  // Converte valor de meta para horas decimais
  // Suporta: números decimais (9.8), Date objects (05:25:00), strings vazio

  if (!valor || valor === "") return 0;

  // Se for número, retorna direto
  if (typeof valor === 'number') {
    return parseFloat(valor) || 0;
  }

  // Se for Date object (formato de horário HH:MM do Google Sheets)
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    const horas = valor.getHours();
    const minutos = valor.getMinutes();
    const segundos = valor.getSeconds();
    // Converter para horas decimais: 5h25min = 5 + (25/60) = 5.4166...
    return parseFloat((horas + (minutos / 60) + (segundos / 3600)).toFixed(4));
  }

  // Se for string, tenta converter
  if (typeof valor === 'string') {
    const num = parseFloat(valor);
    if (!isNaN(num)) return num;
  }

  return 0;
}
function formatarHoraExcel(val) {
  if (val instanceof Date) return `${val.getHours().toString().padStart(2,'0')}:${val.getMinutes().toString().padStart(2,'0')}:${val.getSeconds().toString().padStart(2,'0')}`;
  if (typeof val === 'number') {
    let s = Math.round(val * 86400);
    let h = Math.floor(s/3600), m = Math.floor((s%3600)/60), sec = s%60;
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
  }
  return String(val);
}
function formatarCelulaParada(val) {
  if (!val || val === "" || val === "-") return "-";
  if (val instanceof Date && !isNaN(val.getTime())) return formatarHoraExcel(val);
  return String(val);
}
function converterParaSegundos(val) {
  if (typeof val === 'number') {
    return Math.round(val * 86400);
  }
  if (val instanceof Date) {
    return val.getHours() * 3600 + val.getMinutes() * 60 + val.getSeconds();
  }
  if (typeof val === 'string') {
    const partes = val.split(':');
    if (partes.length === 3) {
      const h = parseInt(partes[0]) || 0;
      const m = parseInt(partes[1]) || 0;
      const s = parseInt(partes[2]) || 0;
      return h * 3600 + m * 60 + s;
    }
  }
  return 0;
}
function formatarSegundosParaHora(s) {
  if (typeof s !== 'number' || isNaN(s)) return "00:00:00";
  s = Math.round(s);
  let h = Math.floor(s/3600).toString().padStart(2,'0');
  let m = Math.floor((s%3600)/60).toString().padStart(2,'0');
  let sec = (s%60).toString().padStart(2,'0');
  return `${h}:${m}:${sec}`;
}
function processarRegistro(resumo, ss, maquina, data, turno, evento, segundos, horarioEvento) {
  if (segundos > 86400) return;
  if (segundos < 0) segundos = 0;
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina + "|" + dStr + "|" + turno;
  if (!resumo[chave]) {
    resumo[chave] = { 
      maquina: maquina, data: data, turno: turno, 
      ligada: 0, desligada: 0, 
      listaStop3: [], listaStop10: [], listaStop20: [], listaStop30: [], 
      horarioInicio: null, primeiraHora: new Date(horarioEvento), primeiraDuracao: segundos, ultimaHora: new Date(horarioEvento) 
    };
  } else {
    resumo[chave].ultimaHora = new Date(horarioEvento);
  }

  if (evento === "TEMPO PARADA" && !resumo[chave].horarioInicio && horarioEvento) resumo[chave].horarioInicio = horarioEvento;
  if (evento === "TEMPO PRODUZINDO") resumo[chave].ligada += segundos;
  else {
    resumo[chave].desligada += segundos;
    if (segundos > 10) {
      if (!resumo[chave].qtdParadas) resumo[chave].qtdParadas = 0;
      resumo[chave].qtdParadas++;
      let hIni = "", hFim = "";
      if (horarioEvento instanceof Date) {
         let inicioParada = new Date(horarioEvento.getTime() - (segundos * 1000));
         hIni = Utilities.formatDate(inicioParada, ss.getSpreadsheetTimeZone(), "HH:mm");
         hFim = Utilities.formatDate(horarioEvento, ss.getSpreadsheetTimeZone(), "HH:mm");
      }
      let objStop = { s: segundos, h: hIni, f: hFim };
      if (segundos > 180) resumo[chave].listaStop3.push(objStop);
      if (segundos > 600) resumo[chave].listaStop10.push(objStop);
      if (segundos > 1200) resumo[chave].listaStop20.push(objStop);
      if (segundos > 1800) resumo[chave].listaStop30.push(objStop);
    }
  }
}
function registrarGapIntermediario(resumo, ss, maquina, data, turno, segundos, horarioInicio) {
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina + "|" + dStr + "|" + turno;
  if (!resumo[chave]) return; 

  resumo[chave].desligada += segundos;
  if (segundos > 10) {
    if (!resumo[chave].qtdParadas) resumo[chave].qtdParadas = 0;
    resumo[chave].qtdParadas++;
    let hIni = "", hFim = "";
    if (horarioInicio instanceof Date) {
       hIni = Utilities.formatDate(horarioInicio, ss.getSpreadsheetTimeZone(), "HH:mm");
       let fimParada = new Date(horarioInicio.getTime() + (segundos * 1000));
       hFim = Utilities.formatDate(fimParada, ss.getSpreadsheetTimeZone(), "HH:mm");
    }
    let objStop = { s: segundos, h: hIni, f: hFim };
    if (segundos > 180) resumo[chave].listaStop3.push(objStop);
    if (segundos > 600) resumo[chave].listaStop10.push(objStop);
    if (segundos > 1200) resumo[chave].listaStop20.push(objStop);
    if (segundos > 1800) resumo[chave].listaStop30.push(objStop);
  }
}
function formatarListaTempos(lista) {
  if (!lista || !lista.length) return "-";
  return lista.map(item => {
    let s = 0, h = "", f = "";
    if (typeof item === 'object' && item.s) { s = parseFloat(item.s) || 0; h = item.h || ""; f = item.f || ""; } else { s = parseFloat(item) || 0; }
    s = Math.max(0, s);
    let hh = Math.floor(s/3600), mm = Math.floor((s%3600)/60), sec = Math.floor(s%60);
    let duracaoStr = `${hh.toString().padStart(2,'0')}:${mm.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
    if (h && f) return `${duracaoStr} (${h} - ${f})`;
    if (h) return `${duracaoStr} (${h})`;
    return duracaoStr;
  }).join(", ");
}

// ==========================================================
// 7. DISPARO DE E-MAIL (AUTOMÁTICO E TESTE)
// ==========================================================

// Função para o Trigger (Ontem)
function enviarRelatorioDiario() {
  const hoje = new Date();
  const ontem = new Date(hoje);
  ontem.setDate(hoje.getDate() - 1);
  enviarRelatorioBase(ontem);
}

// Função para Teste Manual (Hoje)
function testarEnvioEmailHoje() {
  const hoje = new Date();
  // const ontem = new Date(hoje);
  // ontem.setDate(hoje.getDate() - 1);
  // Logger.log("Iniciando relatório de teste (Ontem)...");
  // enviarRelatorioBase(ontem);
  
  // MANTENDO HOJE PARA TESTE INSTANTANEO DE DADOS RECENTES
  Logger.log("Iniciando relatório de teste (Hoje)...");
  enviarRelatorioBase(hoje);
}

// Lógica central
function enviarRelatorioBase(dataAlvo) {
  const ss = getSS();
  const timezone = ss.getSpreadsheetTimeZone();
  
  const dataStr = Utilities.formatDate(dataAlvo, timezone, "dd/MM/yyyy");
  
  // 1. Pegar E-mails
  const sheetEmail = ss.getSheetByName("EMAIL");
  let destinatarios = "";
  if (sheetEmail) {
    const lista = sheetEmail.getDataRange().getValues();
    destinatarios = lista.map(r => r[0]).filter(e => String(e).includes("@")).join(",");
  }
  
  if (!destinatarios) {
    Logger.log("❌ Nenhum e-mail na aba EMAIL.");
    return;
  }

  // 2. Pegar Máquinas da aba TURNOS (garante que todas apareçam)
  const sheetTurnos = ss.getSheetByName("TURNOS");
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  const maquinasStats = {}; 
  
  for (let i = 1; i < dadosTurnos.length; i++) {
    let maq = String(dadosTurnos[i][0]).trim();
    if (maq) maquinasStats[maq] = { "Turno 1": 0, "Turno 2": 0, "Turno 3": 0 };
  }

  // 3. Ler Dados do PAINEL (Dados processados e corretos)
  const sheetPainel = ss.getSheetByName("PAINEL");
  if (!sheetPainel) return;

  const dadosPainel = sheetPainel.getDataRange().getValues();
  // Colunas PAINEL: 0=Maq, 1=Turno, 2=Data, 3=Ligada
  
  for (let i = 1; i < dadosPainel.length; i++) {
    let linha = dadosPainel[i];
    
    // Verificar Data
    let dataLinha = linha[2];
    let dataLinhaStr = "";
    if (dataLinha instanceof Date) {
       dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy");
    } else {
       // Tenta converter se for string
       let d = lerDataBR(dataLinha);
       dataLinhaStr = Utilities.formatDate(d, timezone, "dd/MM/yyyy");
    }

    if (dataLinhaStr !== dataStr) continue;

    let maquina = String(linha[0]).trim();
    let turno = String(linha[1]).trim();
    let tempoLigada = linha[3]; // Valor numérico (fração de dia) ou Date

    // Converter para segundos para o formatador do email
    let segundos = converterParaSegundos(tempoLigada);

    if (maquinasStats[maquina] && maquinasStats[maquina][turno] !== undefined) {
       maquinasStats[maquina][turno] = segundos;
    }
  }

  // 4. HTML
  let html = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      <p>Bom dia!</p>
      <p>Segue produção em horas das máquinas monitoradas na data <strong>${dataStr}</strong>.</p>
      
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #ddd;">
        <tr style="background-color: #0056b3; color: white;">
          <th style="text-align: left;">MÁQUINA</th>
          <th style="text-align: center;">TURNO 1</th>
          <th style="text-align: center;">TURNO 2</th>
          <th style="text-align: center;">TURNO 3</th>
        </tr>
  `;
  
  const nomesMaquinas = Object.keys(maquinasStats).sort();
  
  nomesMaquinas.forEach(maq => {
    html += `<tr><td style="font-weight:bold;">${maq}</td>`;
    
    ["Turno 1", "Turno 2", "Turno 3"].forEach(t => {
      let segundos = maquinasStats[maq][t];
      let celula = "";
      if (segundos > 0) {
        celula = `<span style="color: green; font-weight: bold; font-size: 14px;">${formatarSegundosParaEmail(segundos)}</span>`;
      } else {
        celula = `<span style="color: #dc3545; font-size: 10px; font-weight: bold;">SEM OPERAÇÃO NESTE PERÍODO</span>`;
      }
      html += `<td style="text-align: center;">${celula}</td>`;
    });
    
    html += `</tr>`;
  });
  
  html += `
      </table>
      <br>
      <p>Atenciosamente,<br><strong>Controle de Rotinas e Prazos Marfim.</strong></p>
    </div>
  `;

  // 5. Enviar
  try {
    MailApp.sendEmail({
      to: destinatarios,
      subject: "Controle de Produtividade por Máquina - " + dataStr,
      htmlBody: html
    });
    Logger.log("✅ E-mail enviado para: " + destinatarios);
  } catch (e) {
    Logger.log("❌ Erro ao enviar: " + e.message);
  }
}

function formatarSegundosParaEmail(segundos) {
  if (typeof segundos !== 'number' || isNaN(segundos)) return "00:00:00";
  segundos = Math.round(segundos);
  const h = Math.floor(segundos/3600).toString().padStart(2,'0');
  const m = Math.floor((segundos%3600)/60).toString().padStart(2,'0');
  const s = (segundos%60).toString().padStart(2,'0');
  return `${h}:${m}:${s}`;
}

function FORCAR_AUTORIZACAO() {
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log("Cota restante: " + quota);
}

// FUNÇÃO DE TESTE - Execute esta função para verificar se as metas estão sendo lidas
function testarLeituraMetas() {
  const ss = getSS();
  const sheetTurnos = ss.getSheetByName("TURNOS");

  if (!sheetTurnos) {
    Logger.log("❌ ERRO: Aba TURNOS não encontrada!");
    return;
  }

  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  Logger.log("📋 Total de linhas na aba TURNOS: " + dadosTurnos.length);

  // Mostrar cabeçalho
  Logger.log("\n🔤 CABEÇALHO (linha 1):");
  Logger.log("Coluna A: " + dadosTurnos[0][0]);
  Logger.log("Coluna B: " + dadosTurnos[0][1]);
  Logger.log("Coluna C: " + dadosTurnos[0][2]);
  Logger.log("Coluna H: " + dadosTurnos[0][7]);
  Logger.log("Coluna I: " + dadosTurnos[0][8]);
  Logger.log("Coluna J: " + dadosTurnos[0][9]);

  // Ler metas de cada máquina
  Logger.log("\n🎯 METAS LIDAS:");
  for (let i = 1; i < dadosTurnos.length && i < 5; i++) {
    const maquina = String(dadosTurnos[i][0]).trim();
    const metaT1 = dadosTurnos[i][7];
    const metaT2 = dadosTurnos[i][8];
    const metaT3 = dadosTurnos[i][9];

    Logger.log("\n--- Máquina: " + maquina);
    Logger.log("    Coluna H (META T1) - Valor bruto: " + metaT1 + " | Tipo: " + typeof metaT1);
    Logger.log("    Coluna I (META T2) - Valor bruto: " + metaT2 + " | Tipo: " + typeof metaT2);
    Logger.log("    Coluna J (META T3) - Valor bruto: " + metaT3 + " | Tipo: " + typeof metaT3);
    Logger.log("    ✅ Após converterMetaParaHoras:");
    Logger.log("      META T1: " + converterMetaParaHoras(metaT1) + " horas");
    Logger.log("      META T2: " + converterMetaParaHoras(metaT2) + " horas");
    Logger.log("      META T3: " + converterMetaParaHoras(metaT3) + " horas");
  }

  Logger.log("\n✅ Teste concluído! Verifique os logs acima.");
}

// ==========================================================
// FUNÇÃO DE DEBUG: Rastrear cálculos de tempo de uma máquina
// ==========================================================

// Wrapper para executar facilmente pelo dropdown do Apps Script
function DEBUG_espuladeira_torre_4_bocas() {
  debugarCalculosMaquina("ESPULADEIRA TORRE 4 BOCAS");
}

// Função para listar todas as máquinas e suas configurações de turno
function DEBUG_listar_maquinas_e_turnos() {
  try {
    const ss = getSS();
    const timezone = ss.getSpreadsheetTimeZone();

    Logger.log("📋 ========================================");
    Logger.log("📋 MÁQUINAS CONFIGURADAS NA ABA TURNOS");
    Logger.log("📋 ========================================\n");

    const sheetTurnos = ss.getSheetByName("TURNOS");
    if (!sheetTurnos) {
      Logger.log("❌ Aba TURNOS não encontrada!");
      return;
    }

    const dadosTurnos = sheetTurnos.getDataRange().getValues();
    Logger.log("Total de linhas: " + dadosTurnos.length + "\n");

    // Listar todas as máquinas
    for (let i = 1; i < dadosTurnos.length; i++) {
      const maquina = String(dadosTurnos[i][0]).trim();
      if (!maquina) continue;

      Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
      Logger.log("🔧 Máquina: [" + maquina + "]");
      Logger.log("   Comprimento do nome: " + maquina.length + " caracteres");
      Logger.log("   Nome em código: " + JSON.stringify(maquina));
      Logger.log("");

      // Turno 1
      if (dadosTurnos[i][1] && dadosTurnos[i][2]) {
        const t1Inicio = new Date(dadosTurnos[i][1]);
        const t1Fim = new Date(dadosTurnos[i][2]);
        Logger.log("   Turno 1: " +
                   Utilities.formatDate(t1Inicio, timezone, "HH:mm") + " - " +
                   Utilities.formatDate(t1Fim, timezone, "HH:mm"));
      } else {
        Logger.log("   Turno 1: NÃO CONFIGURADO");
      }

      // Turno 2
      if (dadosTurnos[i][3] && dadosTurnos[i][4]) {
        const t2Inicio = new Date(dadosTurnos[i][3]);
        const t2Fim = new Date(dadosTurnos[i][4]);
        Logger.log("   Turno 2: " +
                   Utilities.formatDate(t2Inicio, timezone, "HH:mm") + " - " +
                   Utilities.formatDate(t2Fim, timezone, "HH:mm"));
      } else {
        Logger.log("   Turno 2: NÃO CONFIGURADO");
      }

      // Turno 3
      if (dadosTurnos[i][5] && dadosTurnos[i][6]) {
        const t3Inicio = new Date(dadosTurnos[i][5]);
        const t3Fim = new Date(dadosTurnos[i][6]);
        Logger.log("   Turno 3: " +
                   Utilities.formatDate(t3Inicio, timezone, "HH:mm") + " - " +
                   Utilities.formatDate(t3Fim, timezone, "HH:mm"));
      } else {
        Logger.log("   Turno 3: NÃO CONFIGURADO");
      }

      Logger.log("");
    }

    Logger.log("\n📋 ========================================");
    Logger.log("📋 MÁQUINAS NA PÁGINA1 (ÚLTIMOS REGISTROS)");
    Logger.log("📋 ========================================\n");

    const sheetPagina1 = ss.getSheetByName("Página1");
    if (!sheetPagina1) {
      Logger.log("❌ Aba Página1 não encontrada!");
      return;
    }

    const dadosPagina1 = sheetPagina1.getDataRange().getValues();
    const maquinasEncontradas = new Set();

    // Pegar últimas 100 linhas
    const inicio = Math.max(1, dadosPagina1.length - 100);
    for (let i = inicio; i < dadosPagina1.length; i++) {
      const maquina = String(dadosPagina1[i][2]).trim();
      if (maquina) maquinasEncontradas.add(maquina);
    }

    Array.from(maquinasEncontradas).sort().forEach(maq => {
      Logger.log("🔧 [" + maq + "]");
      Logger.log("   Comprimento: " + maq.length + " caracteres");
      Logger.log("   Código: " + JSON.stringify(maq));
      Logger.log("");
    });

    Logger.log("\n🔍 ========================================");
    Logger.log("🔍 BUSCAR MÁQUINA ESPECÍFICA");
    Logger.log("🔍 ========================================\n");

    const buscar = "ESPULADEIRA TORRE 4 BOCAS";
    Logger.log("Buscando por: [" + buscar + "]");
    Logger.log("");

    // Buscar na TURNOS
    let encontrouTurnos = false;
    for (let i = 1; i < dadosTurnos.length; i++) {
      const maquina = String(dadosTurnos[i][0]).trim();
      if (maquina.toLowerCase().includes(buscar.toLowerCase()) ||
          buscar.toLowerCase().includes(maquina.toLowerCase())) {
        Logger.log("✅ Encontrado na TURNOS (linha " + (i+1) + "): [" + maquina + "]");
        encontrouTurnos = true;
      }
    }
    if (!encontrouTurnos) {
      Logger.log("❌ NÃO encontrado na aba TURNOS");
    }

    // Buscar na Página1
    Logger.log("");
    let encontrouPagina1 = false;
    for (let i = inicio; i < dadosPagina1.length; i++) {
      const maquina = String(dadosPagina1[i][2]).trim();
      if (maquina.toLowerCase().includes(buscar.toLowerCase()) ||
          buscar.toLowerCase().includes(maquina.toLowerCase())) {
        if (!encontrouPagina1) {
          Logger.log("✅ Encontrado na Página1:");
          encontrouPagina1 = true;
        }
        Logger.log("   Linha " + (i+1) + ": [" + maquina + "]");
        if (encontrouPagina1) break; // Mostrar só o primeiro
      }
    }
    if (!encontrouPagina1) {
      Logger.log("❌ NÃO encontrado na Página1 (últimos 100 registros)");
    }

    Logger.log("\n✅ Listagem concluída!");

  } catch (error) {
    Logger.log("❌ ERRO: " + error.message);
    Logger.log(error.stack);
  }
}

function debugarCalculosMaquina(nomeMaquina) {
  try {
    Logger.log("🔍 ========================================");
    Logger.log("🔍 DEBUG DE CÁLCULOS DE TEMPO");
    Logger.log("🔍 Máquina: " + nomeMaquina);
    Logger.log("🔍 ========================================\n");

    const ss = getSS();
    const sheet = ss.getSheetByName("Página1");
    const dados = sheet ? sheet.getDataRange().getValues() : [];

    if (dados.length <= 1) {
      Logger.log("❌ Nenhum dado encontrado na Página1!");
      return;
    }

    const agora = new Date();
    const timezone = ss.getSpreadsheetTimeZone();

    Logger.log("⏰ Hora atual: " + Utilities.formatDate(agora, timezone, "dd/MM/yyyy HH:mm:ss"));

    // Ler configuração de turnos
    const sheetTurnos = ss.getSheetByName("TURNOS");
    const dadosTurnos = sheetTurnos ? sheetTurnos.getDataRange().getValues() : [];
    const configTurnos = {};

    for (let i = 1; i < dadosTurnos.length; i++) {
      if (dadosTurnos[i][0]) {
        const maqNome = String(dadosTurnos[i][0]).trim();
        configTurnos[maqNome] = [
           { nome: "Turno 1", inicio: dadosTurnos[i][1], fim: dadosTurnos[i][2] },
           { nome: "Turno 2", inicio: dadosTurnos[i][3], fim: dadosTurnos[i][4] },
           { nome: "Turno 3", inicio: dadosTurnos[i][5], fim: dadosTurnos[i][6] }
        ];
      }
    }

    // Descobrir turno atual
    const infoTurnoAtual = descobrirTurnoCompleto(agora, nomeMaquina, configTurnos);
    if (!infoTurnoAtual) {
      Logger.log("❌ Máquina fora de turno no momento!");
      return;
    }

    Logger.log("📋 Turno atual: " + infoTurnoAtual.nome);
    Logger.log("   Horário do turno: " + Utilities.formatDate(new Date(infoTurnoAtual.inicio), timezone, "HH:mm") +
               " até " + Utilities.formatDate(new Date(infoTurnoAtual.fim), timezone, "HH:mm"));
    Logger.log("   Cruza meia-noite? " + (infoTurnoAtual.cruzaMeiaNoite ? "SIM" : "NÃO"));

    // Calcular data de produção
    let dataProducaoAtual = new Date(agora);
    const horaAgora = agora.getHours();
    const horaInicioPrimeiroTurno = obterHoraInicioPrimeiroTurno(nomeMaquina, configTurnos);

    if (horaAgora < horaInicioPrimeiroTurno) {
      dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
    } else if (infoTurnoAtual.cruzaMeiaNoite && horaAgora < Math.floor(infoTurnoAtual.minInicio / 60)) {
      dataProducaoAtual.setDate(dataProducaoAtual.getDate() - 1);
    }
    dataProducaoAtual.setHours(0,0,0,0);

    Logger.log("📅 Data de produção calculada: " + Utilities.formatDate(dataProducaoAtual, timezone, "dd/MM/yyyy"));

    // Calcular horário de início do turno
    const turnoConfig = configTurnos[nomeMaquina].find(t => t.nome === infoTurnoAtual.nome);
    const dataInicioTurno = new Date(dataProducaoAtual);
    const horaInicioTurno = new Date(turnoConfig.inicio);
    dataInicioTurno.setHours(horaInicioTurno.getHours(), horaInicioTurno.getMinutes(), horaInicioTurno.getSeconds(), 0);

    Logger.log("🕐 Horário de início do turno: " + Utilities.formatDate(dataInicioTurno, timezone, "dd/MM/yyyy HH:mm:ss"));

    const tempoDecorrido = Math.floor((agora.getTime() - dataInicioTurno.getTime()) / 1000);
    const hDecorrido = Math.floor(tempoDecorrido / 3600);
    const mDecorrido = Math.floor((tempoDecorrido % 3600) / 60);
    const sDecorrido = tempoDecorrido % 60;

    Logger.log("⏱️  Tempo decorrido desde início do turno: " +
               hDecorrido.toString().padStart(2, '0') + ":" +
               mDecorrido.toString().padStart(2, '0') + ":" +
               sDecorrido.toString().padStart(2, '0'));

    Logger.log("\n📊 ========================================");
    Logger.log("📊 EVENTOS ENCONTRADOS NA PÁGINA1");
    Logger.log("📊 ========================================\n");

    let totalProduzindo = 0;
    let totalParada = 0;
    let primeiraHora = null;
    let primeiraDuracao = 0;
    let eventosProduzindo = [];
    let eventosParada = [];
    let eventosForaDoFiltro = [];

    // Loop de trás pra frente (igual ao código original)
    for (let i = dados.length - 1; i > 0; i--) {
      const linha = dados[i];
      const maquina = String(linha[2]).trim();

      if (maquina !== nomeMaquina) continue;

      const dataReg = lerDataBR(linha[0]);
      const horaRegObj = new Date(linha[1]);

      if (isNaN(dataReg.getTime()) || isNaN(horaRegObj.getTime())) {
        Logger.log("⚠️  Linha " + (i+1) + ": Data/hora inválida - IGNORADO");
        continue;
      }

      const fullDateReg = new Date(dataReg);
      fullDateReg.setHours(horaRegObj.getHours(), horaRegObj.getMinutes(), horaRegObj.getSeconds());

      const infoTurnoReg = descobrirTurnoCompleto(fullDateReg, nomeMaquina, configTurnos);

      // Verificar se pertence ao turno atual
      const pertenceTurnoAtual = infoTurnoReg && infoTurnoReg.nome === infoTurnoAtual.nome;

      // Calcular data de produção do evento
      let dataProdReg = new Date(dataReg);
      const h = fullDateReg.getHours();

      if (h < horaInicioPrimeiroTurno) {
        dataProdReg.setDate(dataProdReg.getDate() - 1);
      } else if (infoTurnoReg && infoTurnoReg.cruzaMeiaNoite && h < Math.floor(infoTurnoReg.minInicio / 60)) {
        dataProdReg.setDate(dataProdReg.getDate() - 1);
      }

      dataProdReg.setHours(0,0,0,0);

      const pertenceDataProducao = dataProdReg.getTime() === dataProducaoAtual.getTime();

      const duracao = parseDuration(linha[4]);
      const evento = linha[3];

      const eventoInfo = {
        linha: i + 1,
        dataHora: Utilities.formatDate(fullDateReg, timezone, "dd/MM/yyyy HH:mm:ss"),
        evento: evento,
        duracao: linha[4],
        duracaoSegundos: duracao,
        turno: infoTurnoReg ? infoTurnoReg.nome : "SEM TURNO",
        dataProducao: Utilities.formatDate(dataProdReg, timezone, "dd/MM/yyyy"),
        pertenceTurnoAtual: pertenceTurnoAtual,
        pertenceDataProducao: pertenceDataProducao
      };

      // Se pertence ao turno e data corretos, somar
      if (pertenceTurnoAtual && pertenceDataProducao) {
        if (evento === "TEMPO PRODUZINDO") {
          totalProduzindo += duracao;
          eventosProduzindo.push(eventoInfo);
        } else if (evento === "TEMPO PARADA") {
          totalParada += duracao;
          eventosParada.push(eventoInfo);
        }

        // Rastrear primeiro evento (loop vai de trás pra frente)
        primeiraHora = fullDateReg;
        primeiraDuracao = duracao;
      } else {
        eventosForaDoFiltro.push(eventoInfo);
      }
    }

    // Mostrar eventos PRODUZINDO
    Logger.log("🟢 EVENTOS PRODUZINDO (" + eventosProduzindo.length + " eventos):");
    if (eventosProduzindo.length === 0) {
      Logger.log("   (nenhum)");
    } else {
      eventosProduzindo.reverse().forEach((evt, idx) => {
        Logger.log("   " + (idx+1) + ". [Linha " + evt.linha + "] " + evt.dataHora + " | Duração: " + evt.duracao + " (" + evt.duracaoSegundos + "s)");
      });
      const hP = Math.floor(totalProduzindo / 3600);
      const mP = Math.floor((totalProduzindo % 3600) / 60);
      const sP = totalProduzindo % 60;
      Logger.log("   ➡️  TOTAL PRODUZINDO: " + hP.toString().padStart(2, '0') + ":" + mP.toString().padStart(2, '0') + ":" + sP.toString().padStart(2, '0') + " (" + totalProduzindo + "s)");
    }

    Logger.log("\n🔴 EVENTOS PARADA (" + eventosParada.length + " eventos):");
    if (eventosParada.length === 0) {
      Logger.log("   (nenhum)");
    } else {
      eventosParada.reverse().forEach((evt, idx) => {
        Logger.log("   " + (idx+1) + ". [Linha " + evt.linha + "] " + evt.dataHora + " | Duração: " + evt.duracao + " (" + evt.duracaoSegundos + "s)");
      });
      const hPa = Math.floor(totalParada / 3600);
      const mPa = Math.floor((totalParada % 3600) / 60);
      const sPa = totalParada % 60;
      Logger.log("   ➡️  TOTAL PARADA (antes do gap): " + hPa.toString().padStart(2, '0') + ":" + mPa.toString().padStart(2, '0') + ":" + sPa.toString().padStart(2, '0') + " (" + totalParada + "s)");
    }

    // Calcular gap inicial
    Logger.log("\n📐 ========================================");
    Logger.log("📐 CÁLCULO DO GAP INICIAL");
    Logger.log("📐 ========================================\n");

    if (primeiraHora) {
      Logger.log("🕐 Primeiro evento registrado: " + Utilities.formatDate(primeiraHora, timezone, "dd/MM/yyyy HH:mm:ss"));
      Logger.log("⏱️  Duração do primeiro evento: " + primeiraDuracao + "s");

      const inicioEfetivoEvento = new Date(primeiraHora.getTime() - (primeiraDuracao * 1000));
      Logger.log("🔙 Início efetivo do evento: " + Utilities.formatDate(inicioEfetivoEvento, timezone, "dd/MM/yyyy HH:mm:ss"));
      Logger.log("🕐 Início do turno: " + Utilities.formatDate(dataInicioTurno, timezone, "dd/MM/yyyy HH:mm:ss"));

      const diferencaSegundos = Math.floor((inicioEfetivoEvento.getTime() - dataInicioTurno.getTime()) / 1000);

      Logger.log("📊 Gap calculado: " + diferencaSegundos + "s");

      if (diferencaSegundos >= 60) {
        totalParada += diferencaSegundos;
        const hGap = Math.floor(diferencaSegundos / 3600);
        const mGap = Math.floor((diferencaSegundos % 3600) / 60);
        const sGap = diferencaSegundos % 60;
        Logger.log("✅ Gap >= 60s: ADICIONADO ao tempo parado");
        Logger.log("   Gap: " + hGap.toString().padStart(2, '0') + ":" + mGap.toString().padStart(2, '0') + ":" + sGap.toString().padStart(2, '0'));
      } else {
        Logger.log("⚠️  Gap < 60s: NÃO adicionado");
      }
    } else {
      Logger.log("⚠️  Nenhum evento encontrado para calcular gap!");
    }

    // TOTAIS FINAIS
    Logger.log("\n🏁 ========================================");
    Logger.log("🏁 TOTAIS FINAIS");
    Logger.log("🏁 ========================================\n");

    const hProd = Math.floor(totalProduzindo / 3600);
    const mProd = Math.floor((totalProduzindo % 3600) / 60);
    const sProd = totalProduzindo % 60;

    const hPar = Math.floor(totalParada / 3600);
    const mPar = Math.floor((totalParada % 3600) / 60);
    const sPar = totalParada % 60;

    const totalGeral = totalProduzindo + totalParada;
    const hGeral = Math.floor(totalGeral / 3600);
    const mGeral = Math.floor((totalGeral % 3600) / 60);
    const sGeral = totalGeral % 60;

    Logger.log("🟢 PRODUZINDO: " + hProd.toString().padStart(2, '0') + ":" + mProd.toString().padStart(2, '0') + ":" + sProd.toString().padStart(2, '0') + " (" + totalProduzindo + "s)");
    Logger.log("🔴 PARADO:     " + hPar.toString().padStart(2, '0') + ":" + mPar.toString().padStart(2, '0') + ":" + sPar.toString().padStart(2, '0') + " (" + totalParada + "s)");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("📊 TOTAL:      " + hGeral.toString().padStart(2, '0') + ":" + mGeral.toString().padStart(2, '0') + ":" + sGeral.toString().padStart(2, '0') + " (" + totalGeral + "s)");
    Logger.log("");
    Logger.log("⏱️  ESPERADO:   " + hDecorrido.toString().padStart(2, '0') + ":" + mDecorrido.toString().padStart(2, '0') + ":" + sDecorrido.toString().padStart(2, '0') + " (" + tempoDecorrido + "s)");

    const diferenca = totalGeral - tempoDecorrido;
    const difAbs = Math.abs(diferenca);
    const hDif = Math.floor(difAbs / 3600);
    const mDif = Math.floor((difAbs % 3600) / 60);
    const sDif = difAbs % 60;

    if (Math.abs(diferenca) > 60) {
      Logger.log("");
      Logger.log("❌ DIFERENÇA:  " + (diferenca > 0 ? "+" : "-") + hDif.toString().padStart(2, '0') + ":" + mDif.toString().padStart(2, '0') + ":" + sDif.toString().padStart(2, '0') + " (" + diferenca + "s)");
      Logger.log("❌ ERRO DETECTADO! Total não bate com tempo decorrido!");
    } else {
      Logger.log("✅ Total correto!");
    }

    // Mostrar eventos fora do filtro (se houver)
    if (eventosForaDoFiltro.length > 0) {
      Logger.log("\n⚠️  ========================================");
      Logger.log("⚠️  EVENTOS IGNORADOS (fora do filtro)");
      Logger.log("⚠️  Total: " + eventosForaDoFiltro.length + " eventos");
      Logger.log("⚠️  ========================================\n");

      eventosForaDoFiltro.slice(0, 10).forEach((evt, idx) => {
        Logger.log("   " + (idx+1) + ". [Linha " + evt.linha + "] " + evt.dataHora);
        Logger.log("      Evento: " + evt.evento + " | Duração: " + evt.duracao);
        Logger.log("      Turno do evento: " + evt.turno + " | Data produção: " + evt.dataProducao);
        Logger.log("      Pertence turno atual? " + (evt.pertenceTurnoAtual ? "SIM" : "NÃO"));
        Logger.log("      Pertence data produção? " + (evt.pertenceDataProducao ? "SIM" : "NÃO"));
        Logger.log("");
      });

      if (eventosForaDoFiltro.length > 10) {
        Logger.log("   ... e mais " + (eventosForaDoFiltro.length - 10) + " eventos ignorados.");
      }
    }

    Logger.log("\n✅ Debug concluído! Verifique os logs acima.");

  } catch (error) {
    Logger.log("❌ ERRO no debug: " + error.message);
    Logger.log(error.stack);
  }
}
