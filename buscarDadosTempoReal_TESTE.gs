// VERSÃO DE TESTE - COPIE ESTA FUNÇÃO PARA O GOOGLE APPS SCRIPT
// Substitua a função buscarDadosTempoReal() existente por esta

function buscarDadosTempoReal() {
  // Envolver TUDO em try-catch com retorno garantido
  var resultado = null;

  try {
    Logger.log("=== INÍCIO buscarDadosTempoReal ===");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error("Erro: getActiveSpreadsheet() retornou null");
      return {};
    }

    var sheetDados = ss.getSheetByName("Página1");
    if (!sheetDados) {
      console.error("Erro: Aba 'Página1' não encontrada");
      return {};
    }

    var dados = sheetDados.getDataRange().getValues();
    if (!dados || dados.length <= 1) {
      console.error("Aviso: Aba 'Página1' está vazia");
      return {};
    }

    var sheetTurnos = ss.getSheetByName("TURNOS");
    if (!sheetTurnos) {
      console.error("Erro: Aba 'TURNOS' não encontrada");
      return {};
    }

    var dadosTurnos = sheetTurnos.getDataRange().getValues();
    var configTurnos = {};

    for (var i = 1; i < dadosTurnos.length; i++) {
      if (dadosTurnos[i][0]) {
        configTurnos[String(dadosTurnos[i][0]).trim()] = [
           { nome: "Turno 1", inicio: dadosTurnos[i][1], fim: dadosTurnos[i][2] },
           { nome: "Turno 2", inicio: dadosTurnos[i][3], fim: dadosTurnos[i][4] },
           { nome: "Turno 3", inicio: dadosTurnos[i][5], fim: dadosTurnos[i][6] }
        ];
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

    for (var i = dados.length - 1; i > 0; i--) {
      var linha = dados[i];
      var maquina = String(linha[2]).trim();
      if (!maquina) continue;

      if (!statusMaquinas[maquina]) {
        var infoTurnoAtual = descobrirTurnoCompleto(agora, maquina, configTurnos);
        var nomeTurnoAtual = "Fora de Turno";
        var dataProducaoAtual = null;

        if (infoTurnoAtual) {
          nomeTurnoAtual = infoTurnoAtual.nome;
          dataProducaoAtual = new Date(agora);

          var horaAgora = agora.getHours();
          if (horaAgora < 7) {
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
          refDataProducao: dataProducaoAtual ? dataProducaoAtual.getTime() : null
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

            if (h < 7) {
              dataProdReg.setDate(dataProdReg.getDate() - 1);
            } else if (infoTurnoReg.cruzaMeiaNoite && h < Math.floor(infoTurnoReg.minInicio / 60)) {
              dataProdReg.setDate(dataProdReg.getDate() - 1);
            }

            dataProdReg.setHours(0,0,0,0);

            if (dataProdReg.getTime() === ref.refDataProducao) {
               var duracao = parseDuration(linha[4]);
               if (linha[3] === "TEMPO PRODUZINDO") ref.totalProduzindo += duracao;
               else if (linha[3] === "TEMPO PARADA") ref.totalParada += duracao;
            }
          }
        }
      }
    }

    // Buscar horário de início do PAINEL
    try {
      var sheetPainel = ss.getSheetByName("PAINEL");
      if (sheetPainel) {
        var dadosPainel = sheetPainel.getDataRange().getValues();
        var timezone = ss.getSpreadsheetTimeZone();

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
                  info.horarioInicio = horarioInicio;
                }
                break;
              }
            }
          }
        }
      }
    } catch (e) {
      console.error("Erro ao buscar horário de início: " + e.message);
    }

    resultado = statusMaquinas;
    Logger.log("=== FIM buscarDadosTempoReal - Total máquinas: " + Object.keys(resultado).length + " ===");

  } catch (error) {
    console.error("ERRO CRÍTICO em buscarDadosTempoReal: " + error.message);
    console.error("Stack: " + error.stack);
    resultado = {};
  }

  // Garantir que SEMPRE retorna um objeto
  return resultado || {};
}
