// ==========================================================
// CONFIGURAÇÃO GLOBAL
// ==========================================================

const ID_PLANILHA = "1jApQbNfM7gUyIL9b3I0CuEFePlnr2DMKeuARCIjrq7g";

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
  if (e && e.parameter && e.parameter.evento === "HEARTBEAT") {
    return salvarHeartbeat(e);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Monitoramento Fabril')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================================
// RECEPÇÃO DE DADOS DO ESP32
// ==========================================================

function salvarDadosESP32(e) {
  try {
    var maquina = e.parameter.maquina || "";
    var evento  = e.parameter.evento  || "";
    var duracao = e.parameter.duracao || "0";

    var ss       = getSS();
    var timezone = Session.getScriptTimeZone();
    var dataAtual = new Date();
    var dataStr  = Utilities.formatDate(dataAtual, timezone, "dd/MM/yyyy");
    var horaStr  = Utilities.formatDate(dataAtual, timezone, "HH:mm:ss");

    // appendRow é thread-safe — sem lock necessário
    if (evento === "TEMPO PRODUZINDO" || evento === "TEMPO PARADA") {
      var sheet = ss.getSheetByName("Página1");
      if (!sheet) sheet = ss.getActiveSheet();
      sheet.appendRow([dataStr, horaStr, maquina, evento, duracao]);
      return ContentService.createTextOutput("OK");
    }

    if (evento === "ESTADO INICIAL PRODUZINDO" || evento === "ESTADO INICIAL PARADA") {
      registrarDiagnostico(ss, dataStr, horaStr, maquina, evento, "Boot detectado");
      return ContentService.createTextOutput("OK");
    }

    // Lock apenas para operações que precisam: SENSOR TRAVADO e HEARTBEAT
    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (e) { return ContentService.createTextOutput("BUSY"); }

    try {
      if (evento === "SENSOR TRAVADO") {
        registrarAlerta(ss, dataStr, horaStr, maquina, evento, duracao);
        var sheet = ss.getSheetByName("Página1");
        if (!sheet) sheet = ss.getActiveSheet();
        sheet.appendRow([dataStr, horaStr, maquina, "TEMPO PARADA", duracao]);
        return ContentService.createTextOutput("OK");
      }
      if (evento === "HEARTBEAT") {
        return salvarHeartbeat(e);
      }
      return ContentService.createTextOutput("OK");
    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    console.error("Erro em salvarDadosESP32: " + error.message);
    return ContentService.createTextOutput("ERROR");
  }
}

// ==========================================================
// HEARTBEAT — confirma que o ESP32 está vivo
// Grava na aba ESP32_STATUS com RSSI e reconexões
// ==========================================================

function salvarHeartbeat(e) {
  try {
    var ss       = getSS();
    var timezone = Session.getScriptTimeZone();
    var agora    = new Date();
    var dataStr  = Utilities.formatDate(agora, timezone, "dd/MM/yyyy");
    var horaStr  = Utilities.formatDate(agora, timezone, "HH:mm:ss");

    var enviados   = e.parameter.enviados   || "?";
    var perdidos   = e.parameter.perdidos   || "0";
    var errosHTTP  = e.parameter.errosHTTP  || "0";
    var rssi       = e.parameter.rssi       || "?";
    var reconexoes = e.parameter.reconexoes || "0";

    var statusTxt = "✅ OK";
    var rssiNum = parseInt(rssi);
    if (!isNaN(rssiNum)) {
      if (rssiNum < -80)      statusTxt = "🔴 SINAL CRÍTICO";
      else if (rssiNum < -70) statusTxt = "⚠️ SINAL FRACO";
      else if (Number(perdidos) > 0) statusTxt = "⚠️ PERDAS";
    } else if (Number(perdidos) > 0) {
      statusTxt = "⚠️ PERDAS";
    }

    var sheet = garantirAba(ss, "ESP32_STATUS",
      ["DATA", "HORA", "EVENTOS ENV.", "EVENTOS PERD.", "ERROS HTTP", "RSSI (dBm)", "RECONEXÕES", "STATUS"]);

    var dados = sheet.getDataRange().getValues();
    var ultimaLinha = dados.length;
    var atualizado = false;

    if (ultimaLinha > 1) {
      var ultimaData = String(dados[ultimaLinha - 1][0]).trim();
      if (ultimaData === dataStr) {
        sheet.getRange(ultimaLinha, 1, 1, 8).setValues([
          [dataStr, horaStr, enviados, perdidos, errosHTTP, rssi, reconexoes, statusTxt]
        ]);
        atualizado = true;
      }
    }

    if (!atualizado) {
      sheet.appendRow([dataStr, horaStr, enviados, perdidos, errosHTTP, rssi, reconexoes, statusTxt]);
    }

    // Coloração por qualidade do sinal
    var ultimaLinhaAtual = sheet.getLastRow();
    var celulaRSSI   = sheet.getRange(ultimaLinhaAtual, 6);
    var celulaStatus = sheet.getRange(ultimaLinhaAtual, 8);
    if (!isNaN(rssiNum)) {
      if (rssiNum < -80) {
        celulaRSSI.setBackground("#fce8e6");
        celulaStatus.setBackground("#fce8e6");
      } else if (rssiNum < -70) {
        celulaRSSI.setBackground("#fff3e0");
        celulaStatus.setBackground("#fff3e0");
      } else {
        celulaRSSI.setBackground("#e8f5e9");
        celulaStatus.setBackground("#e8f5e9");
      }
    }

    limparAbaAntiga(sheet, 90);
    return ContentService.createTextOutput("OK");

  } catch (error) {
    console.error("Erro em salvarHeartbeat: " + error.message);
    return ContentService.createTextOutput("ERROR");
  }
}

// ==========================================================
// ALERTAS — sensor travado, anomalias
// ==========================================================

function registrarAlerta(ss, dataStr, horaStr, maquina, evento, duracao) {
  try {
    var sheet = garantirAba(ss, "ESP32_ALERTAS",
      ["DATA", "HORA", "MÁQUINA", "EVENTO", "DURAÇÃO (s)", "OBSERVAÇÃO"]);

    sheet.appendRow([dataStr, horaStr, maquina, evento, duracao,
      "Verificar sensor/pino físico da máquina"]);

    var ultima = sheet.getLastRow();
    sheet.getRange(ultima, 1, 1, 6).setBackground("#fce8e6");

    // Envio de e-mail de alerta de sensor travado desativado
    /*
    try {
      var sheetEmail = ss.getSheetByName("EMAIL");
      if (sheetEmail) {
        var lista = sheetEmail.getDataRange().getValues();
        var destinatarios = lista.map(r => r[0])
          .filter(e => String(e).includes("@")).join(",");

        if (destinatarios) {
          var duracaoHoras = (parseFloat(duracao) / 3600).toFixed(1);
          MailApp.sendEmail({
            to: destinatarios,
            subject: "⚠ ALERTA — Sensor possivelmente travado: " + maquina,
            htmlBody: `
              <div style="font-family:Arial,sans-serif;color:#333;">
                <h2 style="color:#c0392b;">⚠ Alerta de Sensor Travado</h2>
                <p>O sistema detectou que a máquina abaixo ficou sem mudança de estado por tempo excessivo.</p>
                <table border="1" cellpadding="8" cellspacing="0"
                  style="border-collapse:collapse;border:1px solid #ddd;">
                  <tr style="background:#c0392b;color:white;">
                    <th>Máquina</th><th>Data</th><th>Hora</th><th>Duração sem mudança</th>
                  </tr>
                  <tr>
                    <td><strong>${maquina}</strong></td>
                    <td>${dataStr}</td>
                    <td>${horaStr}</td>
                    <td style="color:#c0392b;font-weight:bold;">${duracaoHoras} horas</td>
                  </tr>
                </table>
                <br>
                <p>Verifique o sensor e o pino GPIO conectado a esta máquina.</p>
                <p>Atenciosamente,<br><strong>Controle de Rotinas e Prazos Marfim.</strong></p>
              </div>`
          });
        }
      }
    } catch (emailErr) {
      console.error("Erro ao enviar e-mail de alerta: " + emailErr.message);
    }
    */

  } catch (error) {
    console.error("Erro em registrarAlerta: " + error.message);
  }
}

// ==========================================================
// DIAGNÓSTICO — boots e eventos administrativos
// ==========================================================

function registrarDiagnostico(ss, dataStr, horaStr, maquina, evento, obs) {
  try {
    var sheet = garantirAba(ss, "ESP32_DIAGNOSTICO",
      ["DATA", "HORA", "MÁQUINA", "EVENTO", "OBSERVAÇÃO"]);
    sheet.appendRow([dataStr, horaStr, maquina, evento, obs]);
    limparAbaAntiga(sheet, 30);
  } catch (error) {
    console.error("Erro em registrarDiagnostico: " + error.message);
  }
}

// ==========================================================
// UTILITÁRIOS DE ABAS
// ==========================================================

function garantirAba(ss, nomeAba, cabecalho) {
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
    sheet.getRange(1, 1, 1, cabecalho.length)
         .setValues([cabecalho])
         .setFontWeight("bold")
         .setBackground("#0056b3")
         .setFontColor("white");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function limparAbaAntiga(sheet, dias) {
  try {
    var dados = sheet.getDataRange().getValues();
    if (dados.length <= 2) return;
    var corte = new Date();
    corte.setDate(corte.getDate() - dias);
    corte.setHours(0, 0, 0, 0);
    var linhasAnteriores = 0;
    for (var i = 1; i < dados.length; i++) {
      var d = lerDataBR(dados[i][0]);
      if (d < corte) linhasAnteriores++;
      else break;
    }
    if (linhasAnteriores > 0) sheet.deleteRows(2, linhasAnteriores);
  } catch (e) { /* silencioso */ }
}

// ==========================================================
// LOGIN
// ==========================================================

function verificarLogin(usuario, senha) {
  const ss = getSS();
  const sheet = ss.getSheetByName("LOGIN");
  if (!sheet) return { erro: "Aba LOGIN não encontrada." };
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === String(usuario).toLowerCase()
        && String(dados[i][1]) == senha) {
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
            var m = String(dConfig[i][idxM]).trim().toUpperCase();
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

      var eventoLinha = String(linha[3]).trim();
      if (eventoLinha !== "TEMPO PRODUZINDO" && eventoLinha !== "TEMPO PARADA") continue;

      // CORREÇÃO: ignora registros com duração 0 (boots) para definir
      // o estado atual — evita o bug do relógio marcando horas erradas
      var duracaoLinha = parseDuration(linha[4]);
      if (!statusMaquinas[maquina] && duracaoLinha === 0) continue;

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
          familia: mapaFamilias[maquina.toUpperCase()] || "OUTROS",
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
          metaTurnoAtual: mapaMetas[maquina] && mapaMetas[maquina][nomeTurnoAtual]
                          ? mapaMetas[maquina][nomeTurnoAtual] : 0
        };
      }

      // Se a máquina ainda não tem estado válido, pula
      if (!statusMaquinas[maquina]) continue;

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

               // Não soma registros de boot (duração 0) nos totais do turno
               if (duracao === 0) continue;

               var fimEvento = fullDateReg.getTime();
               var inicioEvento = fimEvento - (duracao * 1000);

               var configMaq = configTurnos[maquina];
               var turnoConfig = null;
               if (configMaq) {
                 turnoConfig = configMaq.find(function(t) { return t.nome === infoTurnoReg.nome; });
               }

               var duracaoClampada = duracao;

               if (turnoConfig && turnoConfig.inicio && turnoConfig.fim) {
                 var dataInicioTurnoEvento = new Date(dataProdReg);
                 var horaInicioTurnoEvento = new Date(turnoConfig.inicio);
                 dataInicioTurnoEvento.setHours(horaInicioTurnoEvento.getHours(), horaInicioTurnoEvento.getMinutes(), horaInicioTurnoEvento.getSeconds(), 0);

                 var dataFimTurnoEvento = new Date(dataProdReg);
                 var horaFimTurnoEvento = new Date(turnoConfig.fim);
                 dataFimTurnoEvento.setHours(horaFimTurnoEvento.getHours(), horaFimTurnoEvento.getMinutes(), horaFimTurnoEvento.getSeconds(), 0);

                 if (infoTurnoReg.cruzaMeiaNoite) {
                   dataFimTurnoEvento.setDate(dataFimTurnoEvento.getDate() + 1);
                 }

                 var inicioTurno = dataInicioTurnoEvento.getTime();
                 var fimTurno = dataFimTurnoEvento.getTime();
                 var inicioEfetivo = Math.max(inicioEvento, inicioTurno);
                 var fimEfetivo = Math.min(fimEvento, fimTurno);
                 duracaoClampada = Math.max(0, Math.floor((fimEfetivo - inicioEfetivo) / 1000));
               }

               if (linha[3] === "TEMPO PRODUZINDO") ref.totalProduzindo += duracaoClampada;
               else if (linha[3] === "TEMPO PARADA") ref.totalParada += duracaoClampada;

               ref.primeiraHora = fullDateReg;
               ref.primeiraDuracao = duracaoClampada;
            }
          }
        }
      }
    }

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
                var inicioEfetivoEvento = new Date(horaPrimeiroRegistro.getTime() - (info.primeiraDuracao * 1000));
                var inicioEfetivoClampado = Math.max(inicioEfetivoEvento.getTime(), dataInicioTurno.getTime());
                var diferencaSegundos = Math.floor((inicioEfetivoClampado - dataInicioTurno.getTime()) / 1000);

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
// 3. BUSCA DADOS PARA GRÁFICO POR TURNO
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

  let idx = 0;
  for (let d = new Date(dataLimite); d <= hoje; d.setDate(d.getDate() + 1)) {
    let k = Utilities.formatDate(d, timezone, "dd/MM");
    labels.push(k);
    output["Turno 1"].push(0);
    output["Turno 2"].push(0);
    output["Turno 3"].push(0);
    mapIndices[k] = idx++;
  }

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

    let tempoLigada = linha[3];
    let segundos = converterParaSegundos(tempoLigada);
    let horas = parseFloat((segundos / 3600).toFixed(2));

    if (output[turno] && index !== undefined) {
      output[turno][index] = horas;
    }
  }

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
// 4. FUNÇÕES DE DADOS
// ==========================================================

function buscarParadasTurnoAtual(maquina, turnoNome, dataProducao) {
  const ss = getSS();
  const sheet = ss.getSheetByName("PAINEL");
  if (!sheet) { Logger.log("ERRO: Aba 'PAINEL' não encontrada"); return []; }

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
              paradas.push({ duracao: duracaoSeg, duracaoFmt: tempo });
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
  if (!sheet) return { erro: "Aba 'PAINEL' não encontrada" };

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
          let m = String(dConfig[i][idxM]).trim().toUpperCase();
          mapaFamilias[m] = String(dConfig[i][idxF] || "GERAL").trim();
        }
      }
    }
  }

  const dados = sheet.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();
  const maquinasPorFamilia = {};
  let totalRodandoSeg = 0, totalParadoSeg = 0, totalParadasCriticas = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const nomeMaquina = String(linha[0]).trim();
    const familiaMaquina = mapaFamilias[nomeMaquina.toUpperCase()] || "";
    if (familiaMaquina.toUpperCase() !== familia.toUpperCase()) continue;

    const turno = String(linha[1]).trim();
    const dataLinha = lerDataBR(linha[2]);
    const dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "yyyy-MM-dd");

    if (dataInicio && dataLinhaStr < dataInicio) continue;
    if (dataFim && dataLinhaStr > dataFim) continue;

    let ligadaSeg = converterParaSegundos(linha[3]);
    let desligadaSeg = converterParaSegundos(linha[4]);
    let paradasCriticas = String(linha[5] || "").trim();
    let qtdParadas = 0;
    if (paradasCriticas && paradasCriticas !== "-") {
      qtdParadas = paradasCriticas.split(",").filter(p => p.trim().length > 0).length;
    }

    totalRodandoSeg += ligadaSeg;
    totalParadoSeg += desligadaSeg;
    totalParadasCriticas += qtdParadas;

    if (!maquinasPorFamilia[nomeMaquina]) maquinasPorFamilia[nomeMaquina] = [];
    maquinasPorFamilia[nomeMaquina].push({
      data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
      turno: turno,
      ligada: formatarHoraExcel(linha[3]),
      desligada: formatarHoraExcel(linha[4]),
      paradas3min:  formatarCelulaParada(linha[5]),
      paradas10min: formatarCelulaParada(linha[6]),
      paradas20min: formatarCelulaParada(linha[7]),
      paradas30min: formatarCelulaParada(linha[8])
    });
  }

  const maquinas = [];
  for (let nomeMaq in maquinasPorFamilia) {
    maquinas.push({ nome: nomeMaq, turnos: maquinasPorFamilia[nomeMaq] });
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
  if (!sheet) { Logger.log("⚠ Aba 'PAINEL' não encontrada"); return []; }

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
      resultados.push({
        turno: String(linha[1] || "-"),
        data: Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy"),
        ligada: formatarHoraExcel(linha[3]),
        desligada: formatarHoraExcel(linha[4]),
        paradas3min:  formatarCelulaParada(linha[5]),
        paradas10min: formatarCelulaParada(linha[6]),
        paradas20min: formatarCelulaParada(linha[7]),
        paradas30min: formatarCelulaParada(linha[8]),
        motivo: String(linha[10] || "-"),
        servico: String(linha[11] || "-"),
        pecas: String(linha[12] || "-"),
        custoMO: typeof linha[9] === 'number' ? linha[9] : 0,
        custoPecas: typeof linha[13] === 'number' ? linha[13] : 0,
        obs: String(linha[14] || "-")
      });
    } catch (erro) {
      Logger.log("⚠ ERRO ao processar linha " + i + ": " + erro.message);
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
      if (head === "MOTIVO DA PARADA")
         for(let i=1;i<dados.length;i++) if(dados[i][c]) motivos.push(dados[i][c]);
      if (head === "SERVIÇOS REALIZADOS")
         for(let i=1;i<dados.length;i++) if(dados[i][c]) servicos.push(dados[i][c]);
      if (head.includes("FAMÍLIA") || head.includes("FAMILIA"))
         for(let i=1;i<dados.length;i++) {
           let f=String(dados[i][c]).trim();
           if(f && !familias.includes(f)) familias.push(f);
         }
    }
  }
  return { motivos, servicos, familias: familias.sort() };
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

    if (maqPainel === dadosForm.maquina && turnoPainel === dadosForm.turno
        && dataPainelStr === dataFiltroBR) {
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

  const diasParaVerificar = 7;
  const dataLimiteVerificacao = new Date(agora);
  dataLimiteVerificacao.setDate(dataLimiteVerificacao.getDate() - diasParaVerificar);
  dataLimiteVerificacao.setHours(0, 0, 0, 0);

  for (let maquinaKey in configTurnos) {
    let turnosConfig = configTurnos[maquinaKey];
    for (let d = new Date(dataLimiteVerificacao); d <= agora; d.setDate(d.getDate() + 1)) {
      let dataVerificacao = new Date(d);
      dataVerificacao.setHours(0, 0, 0, 0);
      let dataStr = Utilities.formatDate(dataVerificacao, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

      turnosConfig.forEach(turnoConfig => {
        let chave = maquinaKey + "|" + dataStr + "|" + turnoConfig.nome;
        if (!resumo[chave]) {
          let horaInicio = new Date(turnoConfig.inicio);
          let horaFim = new Date(turnoConfig.fim);
          let minInicio = horaInicio.getHours() * 60 + horaInicio.getMinutes();
          let minFim = horaFim.getHours() * 60 + horaFim.getMinutes();
          let duracaoTurnoSegundos = (minFim > minInicio)
            ? (minFim - minInicio) * 60
            : ((1440 - minInicio) + minFim) * 60;

          let horaInicioTurno = new Date(dataVerificacao);
          horaInicioTurno.setHours(horaInicio.getHours(), horaInicio.getMinutes(), 0, 0);
          let horaFimTurno = new Date(dataVerificacao);
          horaFimTurno.setHours(horaFim.getHours(), horaFim.getMinutes(), 0, 0);
          if (minFim <= minInicio) horaFimTurno.setDate(horaFimTurno.getDate() + 1);

          let hIni = Utilities.formatDate(horaInicioTurno, ss.getSpreadsheetTimeZone(), "HH:mm");
          let hFim = Utilities.formatDate(horaFimTurno, ss.getSpreadsheetTimeZone(), "HH:mm");
          let objParada = { s: duracaoTurnoSegundos, h: hIni, f: hFim };

          resumo[chave] = {
            maquina: maquinaKey, data: dataVerificacao, turno: turnoConfig.nome,
            ligada: 0, desligada: duracaoTurnoSegundos,
            listaStop3:  duracaoTurnoSegundos > 180  ? [objParada] : [],
            listaStop10: duracaoTurnoSegundos > 600  ? [objParada] : [],
            listaStop20: duracaoTurnoSegundos > 1200 ? [objParada] : [],
            listaStop30: duracaoTurnoSegundos > 1800 ? [objParada] : [],
            horarioInicio: null, primeiraHora: null, primeiraDuracao: 0, ultimaHora: null
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
            if (minI > minF) dataFimTurno.setDate(dataFimTurno.getDate() + 1);
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
                    if (gapFinalSegundos > 180)  item.listaStop3.push(objGap);
                    if (gapFinalSegundos > 600)  item.listaStop10.push(objGap);
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
                  if (diferencaSegundos > 180)  item.listaStop3.unshift(objGap);
                  if (diferencaSegundos > 600)  item.listaStop10.unshift(objGap);
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

    let manual = { motivo: "", servico: "", pecas: "", custoPecas: "", obs: "" };
    let horarioInicioAnterior = "";
    if (mapaPainelExistente[chave]) {
      let old = mapaPainelExistente[chave];
      manual = { motivo: old[10], servico: old[11], pecas: old[12], custoPecas: old[13], obs: old[14] };
      horarioInicioAnterior = old[15] || "";
    }

    let valLigada    = Math.max(0, item.ligada)    / SEGUNDOS_DIA;
    let valDesligada = Math.max(0, item.desligada) / SEGUNDOS_DIA;
    let horarioInicioStr = horarioInicioAnterior;
    if (item.horarioInicio) {
      let horaObj = new Date(item.horarioInicio);
      if (!isNaN(horaObj.getTime())) {
        horarioInicioStr = Utilities.formatDate(horaObj, ss.getSpreadsheetTimeZone(), "HH:mm:ss");
      }
    }

    if(item.listaStop3.length  > 0) item.listaStop3.sort((a,b)  => b.s - a.s);
    if(item.listaStop10.length > 0) item.listaStop10.sort((a,b) => b.s - a.s);
    if(item.listaStop20.length > 0) item.listaStop20.sort((a,b) => b.s - a.s);
    if(item.listaStop30.length > 0) item.listaStop30.sort((a,b) => b.s - a.s);

    rowFinal = [
       item.maquina, item.turno, item.data, valLigada, valDesligada,
       formatarListaTempos(item.listaStop3), formatarListaTempos(item.listaStop10),
       formatarListaTempos(item.listaStop20), formatarListaTempos(item.listaStop30),
       0, manual.motivo, manual.servico, manual.pecas, manual.custoPecas,
       manual.obs, horarioInicioStr
    ];
    linhasSaida.push(rowFinal);
    delete mapaPainelExistente[chave];
  }

  for (let chave in mapaPainelExistente) {
    linhasSaida.push(mapaPainelExistente[chave]);
  }

  sheetPainel.clearContents();
  const cabecalho = [["MÁQUINAS","TURNO","DATA","LIGADA","DESLIGADA","TEMPOS > 3 min","TEMPOS > 10 min","TEMPOS > 20 min","TEMPOS > 30 min","CUSTO MÃO DE OBRA","MOTIVO DA PARADA","SERVIÇOS REALIZADOS","PEÇAS TROCADAS","CUSTO PEÇAS","OBSERVAÇÃO","HORÁRIO INÍCIO"]];
  sheetPainel.getRange(1,1,1,cabecalho[0].length).setValues(cabecalho).setFontWeight("bold");

  if (linhasSaida.length > 0) {
    const ordemTurno = {"Turno 1":1,"Turno 2":2,"Turno 3":3};
    linhasSaida.sort((a,b) => {
      let dA = lerDataBR(a[2]), dB = lerDataBR(b[2]);
      if (dA.getTime() !== dB.getTime()) return dA.getTime() - dB.getTime();
      let tA = ordemTurno[a[1]] || 9, tB = ordemTurno[b[1]] || 9;
      if (tA !== tB) return tA - tB;
      return String(a[0]).localeCompare(b[0]);
    });
    sheetPainel.getRange(2,1,linhasSaida.length,linhasSaida[0].length).setValues(linhasSaida);
    sheetPainel.getRange(2,3,linhasSaida.length,1).setNumberFormat("dd/MM/yyyy");
    sheetPainel.getRange(2,4,linhasSaida.length,2).setNumberFormat("[h]:mm:ss");
    sheetPainel.getRange(2,10,linhasSaida.length,1).setNumberFormat("R$ #,##0.00");
    sheetPainel.getRange(2,14,linhasSaida.length,1).setNumberFormat("R$ #,##0.00");
    sheetPainel.getRange(2,6,linhasSaida.length,4).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}

// ==========================================================
// 6. LIMPEZA DE DADOS ANTIGOS
// ==========================================================

function limparPagina1() {
  var ss = getSS();
  var sheet = ss.getSheetByName("Página1");
  if (!sheet) { Logger.log("Aba Página1 não encontrada."); return; }

  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return;

  var timezone = ss.getSpreadsheetTimeZone();
  var corte = new Date();
  corte.setDate(corte.getDate() - 2);
  corte.setHours(0, 0, 0, 0);

  var ultimaLinhaAntiga = 0;
  for (var i = 1; i < dados.length; i++) {
    var celula = dados[i][0];
    var dataLinha;
    if (celula instanceof Date) {
      dataLinha = celula;
    } else {
      var partes = String(celula).split("/");
      if (partes.length !== 3) continue;
      dataLinha = new Date(partes[2], partes[1] - 1, partes[0]);
    }
    if (dataLinha < corte) { ultimaLinhaAntiga = i + 1; }
    else { break; }
  }

  if (ultimaLinhaAntiga < 2) { Logger.log("Nenhuma linha antiga para apagar."); return; }

  var qtdLinhas = ultimaLinhaAntiga - 1;
  sheet.deleteRows(2, qtdLinhas);
  Logger.log("✅ Limpeza Página1: " + qtdLinhas + " linha(s) apagadas (anteriores a "
    + Utilities.formatDate(corte, timezone, "dd/MM/yyyy") + ").");
}

function instalarTriggerLimpeza() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "limparPagina1") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("limparPagina1")
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(0).create();
  Logger.log("✅ Trigger semanal de limpeza instalado (domingo à meia-noite).");
}

// ==========================================================
// 7. DISPARO DE E-MAIL
// ==========================================================

function enviarRelatorioDiario() {
  const hoje = new Date();
  const ontem = new Date(hoje);
  ontem.setDate(hoje.getDate() - 1);
  enviarRelatorioBase(ontem);
}

function testarEnvioEmailHoje() {
  Logger.log("Iniciando relatório de teste (Hoje)...");
  enviarRelatorioBase(new Date());
}

function enviarRelatorioBase(dataAlvo) {
  const ss = getSS();
  const timezone = ss.getSpreadsheetTimeZone();
  const dataStr = Utilities.formatDate(dataAlvo, timezone, "dd/MM/yyyy");

  const sheetEmail = ss.getSheetByName("EMAIL");
  let destinatarios = "";
  if (sheetEmail) {
    const lista = sheetEmail.getDataRange().getValues();
    destinatarios = lista.map(r => r[0]).filter(e => String(e).includes("@")).join(",");
  }
  if (!destinatarios) { Logger.log("❌ Nenhum e-mail na aba EMAIL."); return; }

  const sheetTurnos = ss.getSheetByName("TURNOS");
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  const maquinasStats = {};
  for (let i = 1; i < dadosTurnos.length; i++) {
    let maq = String(dadosTurnos[i][0]).trim();
    if (maq) maquinasStats[maq] = { "Turno 1": 0, "Turno 2": 0, "Turno 3": 0 };
  }

  const sheetDadosEmail = ss.getSheetByName("DADOS");
  const mapaFamiliasEmail = {};
  if (sheetDadosEmail) {
    const dConfig = sheetDadosEmail.getDataRange().getValues();
    if (dConfig.length > 0) {
      const h = dConfig[0].map(c => String(c).toUpperCase().trim());
      const idxM = h.findIndex(x => x.includes("MÁQUINAS") || x.includes("MAQUINAS"));
      const idxF = h.findIndex(x => x.includes("FAMÍLIA") || x.includes("FAMILIA"));
      if (idxM > -1 && idxF > -1) {
        for (let i = 1; i < dConfig.length; i++) {
          let m = String(dConfig[i][idxM]).trim().toUpperCase();
          mapaFamiliasEmail[m] = String(dConfig[i][idxF] || "OUTROS").trim();
        }
      }
    }
  }

  const sheetPainel = ss.getSheetByName("PAINEL");
  if (!sheetPainel) return;
  const dadosPainel = sheetPainel.getDataRange().getValues();

  for (let i = 1; i < dadosPainel.length; i++) {
    let linha = dadosPainel[i];
    let dataLinha = linha[2];
    let dataLinhaStr = "";
    if (dataLinha instanceof Date) {
       dataLinhaStr = Utilities.formatDate(dataLinha, timezone, "dd/MM/yyyy");
    } else {
       dataLinhaStr = Utilities.formatDate(lerDataBR(dataLinha), timezone, "dd/MM/yyyy");
    }
    if (dataLinhaStr !== dataStr) continue;

    let maquina = String(linha[0]).trim();
    let turno = String(linha[1]).trim();
    let segundos = converterParaSegundos(linha[3]);

    if (maquinasStats[maquina] && maquinasStats[maquina][turno] !== undefined) {
       maquinasStats[maquina][turno] = segundos;
    }
  }

  const familiasMaquinasEmail = {};
  Object.keys(maquinasStats).sort().forEach(maq => {
    const fam = mapaFamiliasEmail[maq.toUpperCase()] || "OUTROS";
    if (!familiasMaquinasEmail[fam]) familiasMaquinasEmail[fam] = [];
    familiasMaquinasEmail[fam].push(maq);
  });

  let html = `
    <div style="font-family:Arial,sans-serif;color:#333;">
      <p>Bom dia!</p>
      <p>Segue produção em horas das máquinas monitoradas na data <strong>${dataStr}</strong>.</p>
      <table border="1" cellpadding="8" cellspacing="0"
        style="border-collapse:collapse;width:100%;border:1px solid #ddd;">
        <tr style="background-color:#0056b3;color:white;">
          <th style="text-align:left;">MÁQUINA</th>
          <th style="text-align:center;">TURNO 1</th>
          <th style="text-align:center;">TURNO 2</th>
          <th style="text-align:center;">TURNO 3</th>
        </tr>`;

  Object.keys(familiasMaquinasEmail).sort().forEach(fam => {
    html += `<tr style="background-color:#d0e4f7;">
      <td colspan="4" style="font-weight:bold;color:#0056b3;padding:6px 10px;
        font-size:13px;letter-spacing:0.5px;">SETOR &mdash; ${fam}</td></tr>`;

    familiasMaquinasEmail[fam].forEach(maq => {
      html += `<tr><td style="font-weight:bold;padding-left:18px;">${maq}</td>`;
      ["Turno 1","Turno 2","Turno 3"].forEach(t => {
        let seg = maquinasStats[maq][t];
        let celula = seg > 0
          ? `<span style="color:green;font-weight:bold;font-size:14px;">${formatarSegundosParaEmail(seg)}</span>`
          : `<span style="color:#dc3545;font-size:10px;font-weight:bold;">SEM OPERAÇÃO NESTE PERÍODO</span>`;
        html += `<td style="text-align:center;">${celula}</td>`;
      });
      html += `</tr>`;
    });
  });

  html += `</table><br>
    <p>Atenciosamente,<br><strong>Controle de Rotinas e Prazos Marfim.</strong></p></div>`;

  try {
    MailApp.sendEmail({
      to: destinatarios,
      subject: "Controle de Produtividade por Máquina CEARÁ - " + dataStr,
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
  return Math.floor(segundos/3600).toString().padStart(2,'0') + ":" +
         Math.floor((segundos%3600)/60).toString().padStart(2,'0') + ":" +
         (segundos%60).toString().padStart(2,'0');
}

function FORCAR_AUTORIZACAO() {
  Logger.log("Cota restante: " + MailApp.getRemainingDailyQuota());
}

// ==========================================================
// 8. FUNÇÕES AUXILIARES
// ==========================================================

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
  if (raw instanceof Date) return raw.getHours()*3600+raw.getMinutes()*60+raw.getSeconds();
  if (typeof raw === 'string') {
    let str = raw.trim();
    if (str.includes(':')) {
      let p = str.split(':');
      if (p.length===3) return parseInt(p[0])*3600+parseInt(p[1])*60+parseInt(p[2]);
      else if (p.length===2) return parseInt(p[0])*60+parseInt(p[1]);
    }
    let s = parseFloat(str.replace(',','.'));
    return isNaN(s) ? 0 : s;
  }
  return 0;
}
function horaParaMinutos(val) {
  if (val instanceof Date) return val.getHours()*60+val.getMinutes();
  if (typeof val==='string' && val.includes(':')) { let p=val.split(':'); return parseInt(p[0])*60+parseInt(p[1]); }
  return 0;
}
function descobrirTurnoCompleto(hora, maq, config) {
  let c = config[maq] || config[String(maq).trim()];
  if(!c) return null;
  let min = horaParaMinutos(new Date(hora));
  for(let t of c) {
     let i = horaParaMinutos(t.inicio);
     let f = horaParaMinutos(t.fim);
     if (i===0 && f===0) continue;
     let cruza = i > f;
     if (!cruza) { if (min>=i && min<=f) return {nome:t.nome,minInicio:i,minFim:f,cruzaMeiaNoite:false}; }
     else { if (min>=i || min<=f) return {nome:t.nome,minInicio:i,minFim:f,cruzaMeiaNoite:true}; }
  }
  return null;
}
function obterHoraInicioPrimeiroTurno(maq, config) {
  let c = config[maq] || config[String(maq).trim()];
  if (!c) return 7;
  for (let t of c) {
    if (t.inicio) {
      let h = new Date(t.inicio);
      if (!isNaN(h.getTime())) return h.getHours();
    }
  }
  return 7;
}
function converterMetaParaHoras(valor) {
  if (!valor || valor==="") return 0;
  if (typeof valor==='number') return parseFloat(valor)||0;
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return parseFloat((valor.getHours()+(valor.getMinutes()/60)+(valor.getSeconds()/3600)).toFixed(4));
  }
  if (typeof valor==='string') { let n=parseFloat(valor); if(!isNaN(n)) return n; }
  return 0;
}
function formatarHoraExcel(val) {
  if (val instanceof Date) return `${val.getHours().toString().padStart(2,'0')}:${val.getMinutes().toString().padStart(2,'0')}:${val.getSeconds().toString().padStart(2,'0')}`;
  if (typeof val==='number') {
    let s=Math.round(val*86400);
    let h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=s%60;
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
  }
  return String(val);
}
function formatarCelulaParada(val) {
  if (!val || val==="" || val==="-") return "-";
  if (val instanceof Date && !isNaN(val.getTime())) return formatarHoraExcel(val);
  return String(val);
}
function converterParaSegundos(val) {
  if (typeof val==='number') return Math.round(val*86400);
  if (val instanceof Date) return val.getHours()*3600+val.getMinutes()*60+val.getSeconds();
  if (typeof val==='string') {
    const p=val.split(':');
    if (p.length===3) return (parseInt(p[0])||0)*3600+(parseInt(p[1])||0)*60+(parseInt(p[2])||0);
  }
  return 0;
}
function formatarSegundosParaHora(s) {
  if (typeof s!=='number'||isNaN(s)) return "00:00:00";
  s=Math.round(s);
  return Math.floor(s/3600).toString().padStart(2,'0')+":"+
         Math.floor((s%3600)/60).toString().padStart(2,'0')+":"+
         (s%60).toString().padStart(2,'0');
}
function processarRegistro(resumo, ss, maquina, data, turno, evento, segundos, horarioEvento) {
  if (segundos > 86400) return;
  if (segundos < 0) segundos = 0;
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina+"|"+dStr+"|"+turno;
  if (!resumo[chave]) {
    resumo[chave] = {
      maquina, data, turno, ligada:0, desligada:0,
      listaStop3:[], listaStop10:[], listaStop20:[], listaStop30:[],
      horarioInicio:null, primeiraHora:new Date(horarioEvento),
      primeiraDuracao:segundos, ultimaHora:new Date(horarioEvento)
    };
  } else {
    resumo[chave].ultimaHora = new Date(horarioEvento);
  }
  if (evento==="TEMPO PARADA" && !resumo[chave].horarioInicio && horarioEvento) resumo[chave].horarioInicio = horarioEvento;
  if (evento==="TEMPO PRODUZINDO") resumo[chave].ligada += segundos;
  else {
    resumo[chave].desligada += segundos;
    if (segundos > 10) {
      if (!resumo[chave].qtdParadas) resumo[chave].qtdParadas = 0;
      resumo[chave].qtdParadas++;
      let hIni="", hFim="";
      if (horarioEvento instanceof Date) {
         let ini = new Date(horarioEvento.getTime()-(segundos*1000));
         hIni = Utilities.formatDate(ini, ss.getSpreadsheetTimeZone(), "HH:mm");
         hFim = Utilities.formatDate(horarioEvento, ss.getSpreadsheetTimeZone(), "HH:mm");
      }
      let objStop={s:segundos,h:hIni,f:hFim};
      if (segundos>180)  resumo[chave].listaStop3.push(objStop);
      if (segundos>600)  resumo[chave].listaStop10.push(objStop);
      if (segundos>1200) resumo[chave].listaStop20.push(objStop);
      if (segundos>1800) resumo[chave].listaStop30.push(objStop);
    }
  }
}
function registrarGapIntermediario(resumo, ss, maquina, data, turno, segundos, horarioInicio) {
  let dStr = Utilities.formatDate(data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  let chave = maquina+"|"+dStr+"|"+turno;
  if (!resumo[chave]) return;
  resumo[chave].desligada += segundos;
  if (segundos > 10) {
    if (!resumo[chave].qtdParadas) resumo[chave].qtdParadas=0;
    resumo[chave].qtdParadas++;
    let hIni="", hFim="";
    if (horarioInicio instanceof Date) {
       hIni = Utilities.formatDate(horarioInicio, ss.getSpreadsheetTimeZone(), "HH:mm");
       let fim = new Date(horarioInicio.getTime()+(segundos*1000));
       hFim = Utilities.formatDate(fim, ss.getSpreadsheetTimeZone(), "HH:mm");
    }
    let objStop={s:segundos,h:hIni,f:hFim};
    if (segundos>180)  resumo[chave].listaStop3.push(objStop);
    if (segundos>600)  resumo[chave].listaStop10.push(objStop);
    if (segundos>1200) resumo[chave].listaStop20.push(objStop);
    if (segundos>1800) resumo[chave].listaStop30.push(objStop);
  }
}
function formatarListaTempos(lista) {
  if (!lista||!lista.length) return "-";
  return lista.map(item => {
    let s=0,h="",f="";
    if (typeof item==='object'&&item.s){s=parseFloat(item.s)||0;h=item.h||"";f=item.f||"";}
    else{s=parseFloat(item)||0;}
    s=Math.max(0,s);
    let hh=Math.floor(s/3600),mm=Math.floor((s%3600)/60),sec=Math.floor(s%60);
    let dur=`${hh.toString().padStart(2,'0')}:${mm.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
    if (h&&f) return `${dur} (${h} - ${f})`;
    if (h) return `${dur} (${h})`;
    return dur;
  }).join(", ");
}

// ==========================================================
// 9. FUNÇÕES DE DEBUG
// ==========================================================

function testarLeituraMetas() {
  const ss = getSS();
  const sheetTurnos = ss.getSheetByName("TURNOS");
  if (!sheetTurnos) { Logger.log("❌ ERRO: Aba TURNOS não encontrada!"); return; }
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  Logger.log("📋 Total de linhas na aba TURNOS: " + dadosTurnos.length);
  for (let i = 1; i < dadosTurnos.length && i < 5; i++) {
    const maquina = String(dadosTurnos[i][0]).trim();
    Logger.log("\n--- Máquina: " + maquina);
    Logger.log("    META T1: " + converterMetaParaHoras(dadosTurnos[i][7]) + " horas");
    Logger.log("    META T2: " + converterMetaParaHoras(dadosTurnos[i][8]) + " horas");
    Logger.log("    META T3: " + converterMetaParaHoras(dadosTurnos[i][9]) + " horas");
  }
  Logger.log("\n✅ Teste concluído!");
}

function DEBUG_listar_maquinas_e_turnos() {
  const ss = getSS();
  const sheetTurnos = ss.getSheetByName("TURNOS");
  if (!sheetTurnos) { Logger.log("❌ Aba TURNOS não encontrada!"); return; }
  const dadosTurnos = sheetTurnos.getDataRange().getValues();
  const timezone = ss.getSpreadsheetTimeZone();
  Logger.log("📋 MÁQUINAS CONFIGURADAS NA ABA TURNOS\n");
  for (let i = 1; i < dadosTurnos.length; i++) {
    const maquina = String(dadosTurnos[i][0]).trim();
    if (!maquina) continue;
    Logger.log("🔧 Máquina: [" + maquina + "]");
    ["Turno 1","Turno 2","Turno 3"].forEach((nome, idx) => {
      const col = idx * 2 + 1;
      if (dadosTurnos[i][col] && dadosTurnos[i][col+1]) {
        Logger.log("   " + nome + ": " +
          Utilities.formatDate(new Date(dadosTurnos[i][col]),   timezone, "HH:mm") + " - " +
          Utilities.formatDate(new Date(dadosTurnos[i][col+1]), timezone, "HH:mm"));
      } else {
        Logger.log("   " + nome + ": NÃO CONFIGURADO");
      }
    });
    Logger.log("");
  }
  Logger.log("✅ Listagem concluída!");
}
