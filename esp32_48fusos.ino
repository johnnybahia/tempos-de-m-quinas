// ============================================================
// MONITOR DE MÁQUINAS — MARFIM
// Versão 4.4 — 48 Fusos (Unimat e Chinesa)
// Correções: timeout HTTP real, follow-redirects, descarte de
// eventos envenenados, heartbeat nunca persistido, guarda de heap,
// checkpoint horário (máquina que roda sem parar não fica sem
// registrar produção), removido o conceito de "sensor travado"
// (sinal vem de relé — biestável, não trava fisicamente).
// ============================================================

#include <WiFi.h>
#include <HTTPClient.h>
#include <WiFiClientSecure.h>
#include <esp_task_wdt.h>
#include <Preferences.h>
#include <LittleFS.h>
#include <ArduinoJson.h>

// ============================================================
// BLOCO 1 — CONFIGURAÇÕES DE REDE
// ============================================================
#define WIFI_SSID     "MARFIM_PRODUCAO"
#define WIFI_PASSWORD "Marfimm0403"

// ============================================================
// BLOCO 2 — URL DO GOOGLE APPS SCRIPT
// ============================================================
#define GOOGLE_SCRIPT_URL \
  "https://script.google.com/macros/s/AKfycbyoMZd9g_A7IIgcCajQm71OZTEf6R4kMWQvJXy79C6W8MR24OH_Q2sQ9-uko1q9uvH8fg/exec"

// ============================================================
// BLOCO 3 — CADASTRO DE MÁQUINAS
// ============================================================
#define QTD_MAQUINAS 8

struct Maquina {
  const char* nome;
  int pino;
};

// Utilizando pinos seguros para INPUT_PULLUP no ESP32
const Maquina MAQUINAS[QTD_MAQUINAS] = {
  { "48 FUSOS UNIMAT 1",  13 },
  { "48 FUSOS UNIMAT 2",  14 },
  { "48 FUSOS UNIMAT 3",  25 },
  { "48 FUSOS UNIMAT 4",  26 },
  { "48 FUSOS UNIMAT 5",  27 },
  { "48 FUSOS CHINESA 1", 32 },
  { "48 FUSOS CHINESA 2", 33 },
  { "48 FUSOS CHINESA 3", 22 },
};

// ============================================================
// BLOCO 4 — PARÂMETROS OPERACIONAIS
// ============================================================
#define DEBOUNCE_MS              50
#define DURACAO_MINIMA_MS        1000
#define CHECKPOINT_INTERVALO_MS  3600000UL  // 1 hora — ver taskEnviarGoogle/loop()
#define HEARTBEAT_INTERVALO_MS   600000UL
#define FILA_TAMANHO             50
#define WATCHDOG_TIMEOUT_S       120
#define DELAY_ALEATORIO_MAX_MS   8000
#define STACK_WIFI_TASK          20000

// HTTP
#define HTTP_MAX_TENTATIVAS      3
#define HTTP_TIMEOUT_MS          20000UL   // read timeout REAL (HTTPClient, ms)
#define HTTP_TIMEOUT_S           20        // WiFiClientSecure usa SEGUNDOS
#define HTTP_HANDSHAKE_TIMEOUT_S 15
#define HTTP_BACKOFF_BASE_MS     2000UL
#define HEAP_MINIMO_TLS          55000UL   // abaixo disso o handshake TLS falha

// LittleFS
#define LITTLEFS_EVENTOS_PATH   "/eventos.log"
#define LITTLEFS_TMP_PATH       "/eventos_tmp.log"
#define LITTLEFS_MAX_BYTES      1400000UL
#define FS_MAX_FALHAS_EVENTO    15        // após isso o evento é descartado
#define FS_BACKOFF_BASE_MS      5000UL
#define FS_BACKOFF_MAX_MS       60000UL

// ============================================================
// ESTRUTURAS E VARIÁVEIS
// ============================================================
struct EventoData {
  int maquinaIndex;
  bool produzindo;
  unsigned long duracaoMs;
  bool ehHeartbeat;
  bool ehEstadoInicial;
};

int estadoAtual[QTD_MAQUINAS];
unsigned long tempoInicio[QTD_MAQUINAS];
unsigned long lastDebounceTime[QTD_MAQUINAS];
int leituraDebounce[QTD_MAQUINAS];
unsigned long ultimoCheckpoint[QTD_MAQUINAS];

volatile uint32_t eventosEnviados    = 0;
volatile uint32_t eventosPerdidos    = 0;
volatile uint32_t errosHTTP          = 0;
volatile uint32_t reconexoesWiFi     = 0;
volatile uint32_t eventosPersistidos = 0;
volatile uint32_t eventosDescartados = 0;

uint32_t falhasEventoAtualFS = 0;   // falhas consecutivas do evento no topo do arquivo

unsigned long ultimoHeartbeat = 0;
QueueHandle_t filaEnvios;
Preferences prefs;
SemaphoreHandle_t littlefsMutex = NULL;

// ============================================================
// UTILIDADES
// ============================================================

void delayComWdt(uint32_t ms) {
  uint32_t restante = ms;
  while (restante > 0) {
    uint32_t passo = (restante > 1000) ? 1000 : restante;
    vTaskDelay(pdMS_TO_TICKS(passo));
    esp_task_wdt_reset();
    restante -= passo;
  }
}

// ============================================================
// NVS — SALVAR E RECUPERAR ESTADO NA FLASH
// ============================================================

void nvsSalvarEstado(int idx, int estado, unsigned long msInicio) {
  char keyEst[12], keyMs[12];
  snprintf(keyEst, sizeof(keyEst), "est_%d", idx);
  snprintf(keyMs,  sizeof(keyMs),  "ms_%d",  idx);
  prefs.begin("marfim", false);
  prefs.putInt(keyEst, estado);
  prefs.putULong(keyMs, msInicio);
  prefs.end();
}

void nvsSalvarMomentoDoBootAtual() {
  unsigned long agora = millis();
  prefs.begin("marfim", false);
  for (int i = 0; i < QTD_MAQUINAS; i++) {
    char keyMsb[12];
    snprintf(keyMsb, sizeof(keyMsb), "msb_%d", i);
    prefs.putULong(keyMsb, agora);
  }
  prefs.end();
  Serial.println("  [NVS] Momento do boot salvo.");
}

bool nvsLerEstado(int idx, int* estado, unsigned long* msInicio, unsigned long* msBootAnterior) {
  char keyEst[12], keyMs[12], keyMsb[12];
  snprintf(keyEst,  sizeof(keyEst),  "est_%d", idx);
  snprintf(keyMs,   sizeof(keyMs),   "ms_%d",  idx);
  snprintf(keyMsb,  sizeof(keyMsb),  "msb_%d", idx);
  prefs.begin("marfim", true);
  bool temDado = prefs.isKey(keyEst);
  if (temDado) {
    *estado         = prefs.getInt(keyEst, HIGH);
    *msInicio       = prefs.getULong(keyMs, 0);
    *msBootAnterior = prefs.getULong(keyMsb, 0);
  }
  prefs.end();
  return temDado;
}

// ============================================================
// LITTLEFS — PERSISTÊNCIA OFFLINE DE EVENTOS
// ============================================================

String serializarEvento(const EventoData& ev) {
  StaticJsonDocument<128> doc;
  doc["m"] = ev.maquinaIndex;
  doc["p"] = ev.produzindo ? 1 : 0;
  doc["d"] = ev.duracaoMs;
  if (ev.ehEstadoInicial) doc["t"] = "EI";
  else                    doc["t"] = "EV";
  String linha;
  serializeJson(doc, linha);
  linha += "\n";
  return linha;
}

bool deserializarEvento(const String& linha, EventoData& ev) {
  StaticJsonDocument<128> doc;
  if (deserializeJson(doc, linha) != DeserializationError::Ok) return false;
  if (!doc.containsKey("m") || !doc.containsKey("t")) return false;

  int idx = doc["m"] | -1;
  if (idx < 0 || idx >= QTD_MAQUINAS) return false;

  ev.maquinaIndex    = idx;
  ev.produzindo      = (doc["p"].as<int>() == 1);
  ev.duracaoMs       = doc["d"].as<unsigned long>();
  const char* tipo   = doc["t"] | "";
  ev.ehHeartbeat     = (strcmp(tipo, "HB") == 0);
  ev.ehEstadoInicial = (strcmp(tipo, "EI") == 0);
  return true;
}

void persistirEventoLittleFS(const EventoData& ev) {
  if (ev.ehHeartbeat) { return; }
  if (littlefsMutex == NULL) return;
  if (xSemaphoreTake(littlefsMutex, pdMS_TO_TICKS(3000)) != pdTRUE) return;

  size_t tamanhoAtual = 0;
  File fCheck = LittleFS.open(LITTLEFS_EVENTOS_PATH, "r");
  if (fCheck) { tamanhoAtual = fCheck.size(); fCheck.close(); }

  if (tamanhoAtual >= LITTLEFS_MAX_BYTES) {
    File fOld = LittleFS.open(LITTLEFS_EVENTOS_PATH, "r");
    File fNew = LittleFS.open(LITTLEFS_TMP_PATH, "w");
    if (fOld && fNew) {
      int descartadas = 0;
      while (fOld.available()) {
        String linha = fOld.readStringUntil('\n');
        linha.trim();
        if (descartadas < 50) { descartadas++; continue; }
        if (linha.length() > 2) { fNew.print(linha); fNew.print('\n'); }
      }
      fOld.close();
      fNew.close();
      LittleFS.remove(LITTLEFS_EVENTOS_PATH);
      LittleFS.rename(LITTLEFS_TMP_PATH, LITTLEFS_EVENTOS_PATH);
      eventosDescartados += descartadas;
      falhasEventoAtualFS = 0;
      Serial.printf("  [FS] Rotacao: %d linhas descartadas.\n", descartadas);
    } else {
      if (fOld) fOld.close();
      if (fNew) fNew.close();
    }
  }

  File f = LittleFS.open(LITTLEFS_EVENTOS_PATH, "a");
  if (f) {
    f.print(serializarEvento(ev));
    f.close();
    eventosPersistidos++;
    Serial.printf("  [FS] Evento persistido. Total: %u\n", eventosPersistidos);
  } else {
    Serial.println("xx [FS] Falha ao abrir arquivo para escrita.");
  }

  xSemaphoreGive(littlefsMutex);
}

uint32_t contarEventosPersistidos() {
  if (littlefsMutex == NULL) return 0;
  if (xSemaphoreTake(littlefsMutex, pdMS_TO_TICKS(3000)) != pdTRUE) return 0;

  uint32_t count = 0;
  if (LittleFS.exists(LITTLEFS_EVENTOS_PATH)) {
    File f = LittleFS.open(LITTLEFS_EVENTOS_PATH, "r");
    if (f) {
      while (f.available()) {
        String l = f.readStringUntil('\n');
        l.trim();
        if (l.length() > 2) count++;
      }
      f.close();
    }
  }

  xSemaphoreGive(littlefsMutex);
  return count;
}

void removerPrimeiraLinhaLittleFS() {
  if (littlefsMutex == NULL) return;
  if (xSemaphoreTake(littlefsMutex, pdMS_TO_TICKS(3000)) != pdTRUE) return;

  File fOld = LittleFS.open(LITTLEFS_EVENTOS_PATH, "r");
  File fNew = LittleFS.open(LITTLEFS_TMP_PATH, "w");

  if (fOld && fNew) {
    bool primeiraPulada = false;
    while (fOld.available()) {
      String linha = fOld.readStringUntil('\n');
      linha.trim();
      if (linha.length() <= 2) continue;
      if (!primeiraPulada) { primeiraPulada = true; continue; }
      fNew.print(linha);
      fNew.print('\n');
    }
    fOld.close();
    fNew.close();
    LittleFS.remove(LITTLEFS_EVENTOS_PATH);
    LittleFS.rename(LITTLEFS_TMP_PATH, LITTLEFS_EVENTOS_PATH);
  } else {
    if (fOld) fOld.close();
    if (fNew) fNew.close();
  }

  falhasEventoAtualFS = 0;
  xSemaphoreGive(littlefsMutex);
}

int lerPrimeiroEventoLittleFS(EventoData& ev) {
  if (littlefsMutex == NULL) return 0;
  if (xSemaphoreTake(littlefsMutex, pdMS_TO_TICKS(3000)) != pdTRUE) return 0;

  int resultado = 0;
  if (LittleFS.exists(LITTLEFS_EVENTOS_PATH)) {
    File f = LittleFS.open(LITTLEFS_EVENTOS_PATH, "r");
    if (f) {
      String linha = "";
      while (f.available()) {
        linha = f.readStringUntil('\n');
        linha.trim();
        if (linha.length() > 2) break;
      }
      f.close();
      if (linha.length() > 2) {
        resultado = deserializarEvento(linha, ev) ? 1 : -1;
      }
    }
  }

  xSemaphoreGive(littlefsMutex);
  return resultado;
}

// ============================================================
// FUNÇÕES AUXILIARES
// ============================================================

String urlEncode(const String& s) {
  String enc = "";
  for (int i = 0; i < s.length(); i++) {
    char c = s.charAt(i);
    if (c == ' ') enc += "%20";
    else if (isAlphaNumeric(c) || c == '-' || c == '_' || c == '.' || c == '/') enc += c;
    else {
      char buf[4];
      snprintf(buf, sizeof(buf), "%%%02X", (unsigned char)c);
      enc += buf;
    }
  }
  return enc;
}

String montarURL(int idx, const char* evento, unsigned long duracaoMs) {
  String url = GOOGLE_SCRIPT_URL;
  url += "?maquina=" + urlEncode(String(MAQUINAS[idx].nome));
  url += "&evento="  + urlEncode(String(evento));
  url += "&duracao=" + String(duracaoMs / 1000.0, 2);
  return url;
}

bool enviarHTTP(const String& url) {
  for (int tentativa = 1; tentativa <= HTTP_MAX_TENTATIVAS; tentativa++) {
    esp_task_wdt_reset();

    uint32_t heap = ESP.getFreeHeap();
    if (heap < HEAP_MINIMO_TLS) {
      Serial.printf("xx [HTTP] Heap insuficiente para TLS (%u bytes). Abortando envio.\n", heap);
      errosHTTP++;
      return false;
    }

    WiFiClientSecure client;
    client.setInsecure();
    client.setTimeout(HTTP_TIMEOUT_S);
    client.setHandshakeTimeout(HTTP_HANDSHAKE_TIMEOUT_S);

    HTTPClient http;
    if (!http.begin(client, url)) {
      Serial.println("xx [HTTP] http.begin() falhou.");
      errosHTTP++;
      delayComWdt(HTTP_BACKOFF_BASE_MS);
      continue;
    }

    http.setTimeout(HTTP_TIMEOUT_MS);
    http.setConnectTimeout(HTTP_TIMEOUT_MS);
    http.setReuse(false);
    http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);

    int httpCode = http.GET();
    String body = "";
    if (httpCode > 0) {
      body = http.getString();
      body.trim();
      if (body.length() > 60) body = body.substring(0, 60);
    }
    http.end();

    if (httpCode == HTTP_CODE_OK && body != "BUSY") {
      return true;
    }

    Serial.printf("xx [HTTP] Tentativa %d/%d falhou (codigo %d | heap %u | body: %s)\n",
                  tentativa, HTTP_MAX_TENTATIVAS, httpCode, heap, body.c_str());
    errosHTTP++;

    if (tentativa < HTTP_MAX_TENTATIVAS) {
      delayComWdt(HTTP_BACKOFF_BASE_MS * tentativa);
    }
  }
  return false;
}

bool reconectarWiFi() {
  if (WiFi.status() == WL_CONNECTED) return true;
  Serial.println("!! [WiFi] Desconectado. Reconectando...");
  WiFi.disconnect();
  vTaskDelay(pdMS_TO_TICKS(1000));
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);
  int espera = 0;
  while (WiFi.status() != WL_CONNECTED && espera < 50) {
    vTaskDelay(pdMS_TO_TICKS(1000));
    espera++;
    Serial.printf("  [WiFi] Aguardando... (%ds)\n", espera);
    esp_task_wdt_reset();
  }
  if (WiFi.status() == WL_CONNECTED) {
    Serial.println(">> [WiFi] Reconectado! IP: " + WiFi.localIP().toString());
    reconexoesWiFi++;
    return true;
  }
  nvsSalvarMomentoDoBootAtual();
  Serial.println("!! [WiFi] Falha apos 50s. Reiniciando...");
  vTaskDelay(pdMS_TO_TICKS(500));
  ESP.restart();
  return false;
}

void enfileirarOuPersistir(const EventoData& ev) {
  if (xQueueSend(filaEnvios, &ev, 0) != pdTRUE) {
    eventosPerdidos++;
    persistirEventoLittleFS(ev);
  }
}

String montarURLHeartbeat() {
  int32_t rssi = WiFi.RSSI();
  String url  = String(GOOGLE_SCRIPT_URL);
  url += "?evento=HEARTBEAT";
  url += "&enviados="    + String(eventosEnviados);
  url += "&perdidos="    + String(eventosPerdidos);
  url += "&errosHTTP="   + String(errosHTTP);
  url += "&rssi="        + String(rssi);
  url += "&reconexoes="  + String(reconexoesWiFi);
  url += "&persistidos=" + String(eventosPersistidos);
  url += "&descartados=" + String(eventosDescartados);
  url += "&fsPendentes=" + String(contarEventosPersistidos());
  url += "&heap="        + String(ESP.getFreeHeap());
  return url;
}

// ============================================================
// TASK CORE 0 — ENVIO WiFi
// ============================================================

void taskEnviarGoogle(void* parameter) {
  esp_task_wdt_add(NULL);
  EventoData pacote;
  randomSeed(ESP.getEfuseMac());

  while (true) {
    esp_task_wdt_reset();

    // ---------- Drenagem do LittleFS ----------
    EventoData evFS;
    int statusFS = lerPrimeiroEventoLittleFS(evFS);

    if (statusFS == -1) {
      Serial.println("xx [FS] Linha corrompida detectada. Removendo.");
      removerPrimeiraLinhaLittleFS();
      eventosDescartados++;
      continue;
    }

    if (statusFS == 1) {
      if (evFS.ehHeartbeat) {
        Serial.println("  [FS] Heartbeat legado descartado.");
        removerPrimeiraLinhaLittleFS();
        eventosDescartados++;
        continue;
      }

      if (reconectarWiFi()) {
        String url, nomeLog;
        if (evFS.ehEstadoInicial) {
          const char* ev2 = evFS.produzindo ? "ESTADO INICIAL PRODUZINDO" : "ESTADO INICIAL PARADA";
          url     = montarURL(evFS.maquinaIndex, ev2, 0);
          nomeLog = String("[FS-BOOT] ") + MAQUINAS[evFS.maquinaIndex].nome;
        } else {
          const char* ev2 = evFS.produzindo ? "TEMPO PRODUZINDO" : "TEMPO PARADA";
          url     = montarURL(evFS.maquinaIndex, ev2, evFS.duracaoMs);
          nomeLog = String("[FS-EV] ") + MAQUINAS[evFS.maquinaIndex].nome +
                    " - " + String(evFS.duracaoMs / 1000.0, 1) + "s";
        }

        Serial.println(">> " + nomeLog);

        if (enviarHTTP(url)) {
          eventosEnviados++;
          removerPrimeiraLinhaLittleFS();
          Serial.println("<< FS Sucesso.");
        } else {
          falhasEventoAtualFS++;
          if (falhasEventoAtualFS >= FS_MAX_FALHAS_EVENTO) {
            Serial.printf("xx [FS] Evento descartado apos %u falhas (envenenado).\n",
                          falhasEventoAtualFS);
            removerPrimeiraLinhaLittleFS();
            eventosDescartados++;
          } else {
            uint32_t backoff = FS_BACKOFF_BASE_MS * falhasEventoAtualFS;
            if (backoff > FS_BACKOFF_MAX_MS) backoff = FS_BACKOFF_MAX_MS;
            Serial.printf("xx [FS] Falha %u/%u - aguardando %ums.\n",
                          falhasEventoAtualFS, (uint32_t)FS_MAX_FALHAS_EVENTO, backoff);
            delayComWdt(backoff);
          }
        }
      }
      esp_task_wdt_reset();
      continue;
    }

    // ---------- Fila RAM ----------
    if (xQueueReceive(filaEnvios, &pacote, pdMS_TO_TICKS(1000))) {
      if (!reconectarWiFi()) {
        persistirEventoLittleFS(pacote);
        continue;
      }

      if (!pacote.ehHeartbeat) {
        delayComWdt(random(0, DELAY_ALEATORIO_MAX_MS));
      }

      String url, nomeLog;

      if (pacote.ehHeartbeat) {
        url     = montarURLHeartbeat();
        nomeLog = "[HEARTBEAT] RSSI:" + String(WiFi.RSSI()) +
                  "dBm | Heap:" + String(ESP.getFreeHeap());
      }
      else if (pacote.ehEstadoInicial) {
        const char* ev = pacote.produzindo ? "ESTADO INICIAL PRODUZINDO" : "ESTADO INICIAL PARADA";
        url     = montarURL(pacote.maquinaIndex, ev, 0);
        nomeLog = String("[BOOT] ") + MAQUINAS[pacote.maquinaIndex].nome;
      }
      else {
        const char* ev = pacote.produzindo ? "TEMPO PRODUZINDO" : "TEMPO PARADA";
        url      = montarURL(pacote.maquinaIndex, ev, pacote.duracaoMs);
        nomeLog  = String(pacote.produzindo ? "[PROD] " : "[PARADA] ");
        nomeLog += MAQUINAS[pacote.maquinaIndex].nome;
        nomeLog += " - " + String(pacote.duracaoMs / 1000.0, 1) + "s";
      }

      Serial.println(">> " + nomeLog);

      if (enviarHTTP(url)) {
        eventosEnviados++;
        Serial.println("<< Sucesso.");
      } else if (pacote.ehHeartbeat) {
        Serial.println("xx Heartbeat perdido (nao persistido).");
      } else {
        persistirEventoLittleFS(pacote);
        Serial.println("xx HTTP falhou - persistido no FS.");
      }
    }

    esp_task_wdt_reset();
  }
}

// ============================================================
// SETUP
// ============================================================

void setup() {
  Serial.begin(115200);
  Serial.println("\n===== MONITOR MARFIM v4.4 - 48 FUSOS =====");

  littlefsMutex = xSemaphoreCreateMutex();

  if (!LittleFS.begin(true)) {
    Serial.println("xx [FS] LittleFS falhou! Operando sem persistencia offline.");
  } else {
    Serial.printf("  [FS] LittleFS OK. Eventos pendentes: %u\n", contarEventosPersistidos());
  }

  for (int i = 0; i < QTD_MAQUINAS; i++) {
    pinMode(MAQUINAS[i].pino, INPUT_PULLUP);
    estadoAtual[i]      = digitalRead(MAQUINAS[i].pino);
    leituraDebounce[i]  = estadoAtual[i];
    tempoInicio[i]      = millis();
    lastDebounceTime[i] = millis();
    ultimoCheckpoint[i] = millis();
  }

  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.config(INADDR_NONE, INADDR_NONE, INADDR_NONE, IPAddress(8, 8, 8, 8));
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);

  Serial.print("Conectando WiFi");
  int t = 0;
  while (WiFi.status() != WL_CONNECTED && t < 30) {
    delay(500);
    Serial.print(".");
    t++;
  }

  if (WiFi.status() != WL_CONNECTED) {
    Serial.println("\n!! WiFi falhou no boot. Salvando NVS e continuando offline...");
    nvsSalvarMomentoDoBootAtual();
  } else {
    Serial.println("\nConectado! IP: " + WiFi.localIP().toString());
  }

  Serial.printf("  [MEM] Heap livre apos WiFi: %u bytes\n", ESP.getFreeHeap());

  filaEnvios = xQueueCreate(FILA_TAMANHO, sizeof(EventoData));

#if ESP_ARDUINO_VERSION_MAJOR >= 3
  esp_task_wdt_config_t wdt_config = {
    .timeout_ms     = WATCHDOG_TIMEOUT_S * 1000,
    .idle_core_mask = 0,
    .trigger_panic  = true
  };
  esp_task_wdt_reconfigure(&wdt_config);
#else
  esp_task_wdt_init(WATCHDOG_TIMEOUT_S, true);
#endif
  esp_task_wdt_add(NULL);

  xTaskCreatePinnedToCore(taskEnviarGoogle, "WiFiTask", STACK_WIFI_TASK, NULL, 1, NULL, 0);

  Serial.println("Verificando NVS...");
  for (int i = 0; i < QTD_MAQUINAS; i++) {
    int eS = HIGH;
    unsigned long mI = 0, mB = 0;

    if (nvsLerEstado(i, &eS, &mI, &mB) && mB > mI) {
      unsigned long dur = mB - mI;
      if (dur >= DURACAO_MINIMA_MS && dur < 86400000UL) {
        Serial.printf("  [NVS] %s - recuperado %.1fs %s\n",
                      MAQUINAS[i].nome, dur / 1000.0,
                      (eS == LOW) ? "PRODUZINDO" : "PARADA");
        EventoData ev = { i, (eS == LOW), dur, false, false };
        enfileirarOuPersistir(ev);
      } else {
        Serial.printf("  [NVS] %s - duracao invalida (%.1fs). Ignorado.\n",
                      MAQUINAS[i].nome, dur / 1000.0);
      }
    } else {
      Serial.printf("  [NVS] %s - sem recuperacao.\n", MAQUINAS[i].nome);
    }

    nvsSalvarEstado(i, estadoAtual[i], millis());

    EventoData bootEv = { i, (estadoAtual[i] == LOW), 0, false, true };
    enfileirarOuPersistir(bootEv);
  }

  ultimoHeartbeat = millis();

  Serial.println("Sistema iniciado! Monitorando " + String(QTD_MAQUINAS) + " maquinas.");
  for (int i = 0; i < QTD_MAQUINAS; i++) {
    Serial.printf("  GPIO %2d -> %s [%s]\n",
                  MAQUINAS[i].pino,
                  MAQUINAS[i].nome,
                  estadoAtual[i] == LOW ? "PRODUZINDO" : "PARADA");
  }
}

// ============================================================
// LOOP PRINCIPAL (CORE 1) — LEITURA GPIO
// ============================================================

void loop() {
  esp_task_wdt_reset();
  unsigned long agora = millis();

  if (agora - ultimoHeartbeat >= HEARTBEAT_INTERVALO_MS) {
    ultimoHeartbeat = agora;
    EventoData hb = { 0, false, 0, true, false };
    if (xQueueSend(filaEnvios, &hb, 0) != pdTRUE) eventosPerdidos++;
  }

  for (int i = 0; i < QTD_MAQUINAS; i++) {
    int leitura = digitalRead(MAQUINAS[i].pino);

    if (leitura != leituraDebounce[i]) {
      lastDebounceTime[i] = agora;
      leituraDebounce[i]  = leitura;
    }

    if ((agora - lastDebounceTime[i]) >= DEBOUNCE_MS) {
      if (leitura != estadoAtual[i]) {
        unsigned long duracaoMs = lastDebounceTime[i] - tempoInicio[i];

        if (duracaoMs >= DURACAO_MINIMA_MS) {
          EventoData ev = { i, (estadoAtual[i] == LOW), duracaoMs, false, false };
          enfileirarOuPersistir(ev);
          Serial.printf(">> Evento: %s | %s | %.1fs\n",
                        MAQUINAS[i].nome,
                        estadoAtual[i] == LOW ? "PRODUZINDO" : "PARADA",
                        duracaoMs / 1000.0);
        } else {
          Serial.printf("  [RUIDO] %s | %.0fms\n", MAQUINAS[i].nome, (float)duracaoMs);
        }

        estadoAtual[i]      = leitura;
        tempoInicio[i]      = lastDebounceTime[i];
        ultimoCheckpoint[i] = lastDebounceTime[i];
        nvsSalvarEstado(i, estadoAtual[i], millis());
      }
    }

    // Checkpoint: uma trançadeira pode rodar horas seguidas sem nunca mudar
    // de estado — sem isso, um evento só é enviado quando ela finalmente
    // para, e um turno inteiro passado no meio de uma corrida longa fica
    // sem nenhum dado. A cada CHECKPOINT_INTERVALO_MS, reporta o tempo
    // decorrido desde o último relato como um evento normal (produzindo ou
    // parada, o que estiver valendo) e reinicia a contagem a partir de
    // agora — sem tocar em estadoAtual[i], porque o estado físico não mudou.
    if ((agora - ultimoCheckpoint[i]) >= CHECKPOINT_INTERVALO_MS) {
      unsigned long duracaoMs = agora - tempoInicio[i];
      if (duracaoMs >= DURACAO_MINIMA_MS) {
        EventoData cp = { i, (estadoAtual[i] == LOW), duracaoMs, false, false };
        enfileirarOuPersistir(cp);
        Serial.printf(">> Checkpoint: %s | %s | %.1fs\n",
                      MAQUINAS[i].nome,
                      estadoAtual[i] == LOW ? "PRODUZINDO" : "PARADA",
                      duracaoMs / 1000.0);
      }
      tempoInicio[i]      = agora;
      ultimoCheckpoint[i] = agora;
      nvsSalvarEstado(i, estadoAtual[i], millis());
    }
  }

  vTaskDelay(pdMS_TO_TICKS(1));
}
