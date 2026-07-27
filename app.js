/* Registro de Atividades de Campo — aplicação estática para GitHub Pages */
(() => {
  "use strict";

  const DB_NAME = "enel-field-activities";
  const DB_VERSION = 1;
  const STORE_RECORDS = "records";
  const STORE_ACTIONS = "actions";
  const STORE_META = "meta";
  const SETTINGS_KEY = "enel-field-settings";
  const MODEL_URL = "./data/model.json";

  const ACTIVITY_STATUSES = [
    "Rascunho",
    "Em andamento",
    "Concluída",
    "Concluída com pendência",
    "Parcial",
    "Cancelada",
    "Aguardando retorno",
    "Reprogramada",
  ];

  const ACTION_STATUSES = [
    "Aberta",
    "Em andamento",
    "Aguardando parceira",
    "Aguardando Enel",
    "Concluída",
    "Cancelada",
  ];

  const DEFAULT_ACTIVITY_TYPES = [
    "Acompanhamento de Momento Enel na parceira",
    "Viabilidade de projeto",
    "Viabilidade para redução de clientes interrompidos",
    "Viabilidade de gerador de média tensão",
    "Fiscalização de segurança em equipe de parceira",
    "Realização de mega blitz no pátio da parceira",
    "Acompanhamento de projeto com utilização de gerador de média tensão",
    "Acompanhamento de projeto prioritário",
    "Visita técnica",
    "Inspeção de campo",
    "Acompanhamento de execução de projeto",
    "Reunião ou alinhamento em campo",
    "Verificação de materiais",
    "Comissionamento",
    "Acompanhamento de linha viva",
    "Acompanhamento de desligamento programado",
    "Outras",
  ];

  const DEFAULT_PARTNERS = [
    "Start Engenharia",
    "Essencial Energia",
    "PSC Energy",
    "Conecta Empreendimentos",
    "Engelmig",
    "Outras",
  ];

  const state = {
    db: null,
    model: {},
    records: [],
    actions: [],
    activityTypes: [],
    partners: [],
    regions: [],
    poles: [],
    interventions: [],
    priorities: [],
    selectedPhotos: [],
    selectedRecordId: null,
    statusChart: null,
    typeChart: null,
    settings: { endpoint: "", token: "" },
  };

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function normalize(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  }

  function generateId(prefix) {
    const date = new Date();
    const year = date.getFullYear();
    const random =
      typeof crypto !== "undefined" && crypto.randomUUID
        ? crypto.randomUUID().replaceAll("-", "").slice(0, 8).toUpperCase()
        : Math.random().toString(36).slice(2, 10).toUpperCase();
    return `${prefix}-${year}-${random}`;
  }

  function todayISO() {
    const date = new Date();
    return [
      date.getFullYear(),
      String(date.getMonth() + 1).padStart(2, "0"),
      String(date.getDate()).padStart(2, "0"),
    ].join("-");
  }

  function currentTime() {
    const date = new Date();
    return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
  }

  function excelSerialToDate(value) {
    if (value === null || value === undefined || value === "") return "";
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
    }
    if (typeof value === "number" && value > 20000) {
      const utc = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
      return utc.toISOString().slice(0, 10);
    }
    const text = String(value).trim();
    const br = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (br) return `${br[3]}-${br[2].padStart(2, "0")}-${br[1].padStart(2, "0")}`;
    const parsed = new Date(text);
    return Number.isNaN(parsed.valueOf()) ? text : parsed.toISOString().slice(0, 10);
  }

  function excelTimeToText(value) {
    if (value === null || value === undefined || value === "") return "";
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
    }
    if (typeof value === "number") {
      const totalMinutes = Math.round((value % 1) * 24 * 60);
      return `${String(Math.floor(totalMinutes / 60) % 24).padStart(2, "0")}:${String(totalMinutes % 60).padStart(2, "0")}`;
    }
    return String(value).slice(0, 5);
  }

  function formatDate(value) {
    const iso = excelSerialToDate(value);
    if (!iso) return "—";
    const [year, month, day] = iso.split("-");
    return year && month && day ? `${day}/${month}/${year}` : iso;
  }

  function formatDateTime(date, time) {
    return `${formatDate(date)}${time ? ` • ${time}` : ""}`;
  }

  function getMonthLabel(date = new Date()) {
    return new Intl.DateTimeFormat("pt-BR", { month: "long", year: "numeric" }).format(date);
  }

  function toObjects(rows) {
    if (!Array.isArray(rows) || rows.length < 2) return [];
    const headers = rows[0].map((header) => String(header ?? "").trim());
    return rows
      .slice(1)
      .filter((row) => Array.isArray(row) && row.some((cell) => cell !== null && cell !== ""))
      .map((row) =>
        Object.fromEntries(headers.map((header, index) => [header || `COL_${index + 1}`, row[index] ?? null])),
      );
  }

  function parseConfiguration(rows) {
    const output = { regions: [], poles: [], interventions: [], priorities: [] };
    if (!Array.isArray(rows)) return output;
    for (const row of rows.slice(2)) {
      if (row?.[1]) output.regions.push(String(row[1]));
      if (row?.[4]) output.poles.push(String(row[4]));
      if (row?.[7]) output.interventions.push(String(row[7]));
      if (row?.[10]) output.priorities.push(String(row[10]));
    }
    return output;
  }

  function unique(values) {
    return [...new Set(values.filter(Boolean))];
  }

  function statusClass(status) {
    const value = normalize(status);
    if (value.includes("concluida") && !value.includes("pendencia")) return "success";
    if (value.includes("cancel") || value.includes("vencid")) return "danger";
    if (value.includes("pendencia") || value.includes("parcial") || value.includes("aguard")) return "warning";
    if (value.includes("andamento") || value.includes("reprogram")) return "info";
    return "neutral";
  }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE_RECORDS)) {
          db.createObjectStore(STORE_RECORDS, { keyPath: "id" });
        }
        if (!db.objectStoreNames.contains(STORE_ACTIONS)) {
          db.createObjectStore(STORE_ACTIONS, { keyPath: "id" });
        }
        if (!db.objectStoreNames.contains(STORE_META)) {
          db.createObjectStore(STORE_META, { keyPath: "key" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function dbRequest(storeName, mode, operation) {
    return new Promise((resolve, reject) => {
      const transaction = state.db.transaction(storeName, mode);
      const store = transaction.objectStore(storeName);
      const request = operation(store);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  const dbGetAll = (store) => dbRequest(store, "readonly", (target) => target.getAll());
  const dbGet = (store, key) => dbRequest(store, "readonly", (target) => target.get(key));
  const dbPut = (store, value) => dbRequest(store, "readwrite", (target) => target.put(value));
  const dbDelete = (store, key) => dbRequest(store, "readwrite", (target) => target.delete(key));
  const dbClear = (store) => dbRequest(store, "readwrite", (target) => target.clear());

  function normalizeRecord(row) {
    const photos = Array.isArray(row.photos) ? row.photos : [];
    return {
      id: String(row.ID_ATIVIDADE ?? row.id ?? generateId("ATV")),
      createdDate: excelSerialToDate(row.DATA_CRIACAO ?? row.createdDate ?? todayISO()),
      createdTime: excelTimeToText(row.HORA_CRIACAO ?? row.createdTime ?? currentTime()),
      activityDate: excelSerialToDate(row.DATA_ATIVIDADE ?? row.activityDate ?? todayISO()),
      startTime: excelTimeToText(row.HORA_INICIO ?? row.startTime ?? ""),
      endTime: excelTimeToText(row.HORA_TERMINO ?? row.endTime ?? ""),
      duration: row.DURACAO_ATIVIDADE ?? row.duration ?? "",
      userEmail: row.EMAIL_USUARIO ?? row.userEmail ?? "",
      technicianId: row.ID_TECNICO ?? row.technicianId ?? "",
      technician: row.NOME_TECNICO_ENEL ?? row.technician ?? "",
      technicianRegistration: row.MATRICULA_TECNICO ?? row.technicianRegistration ?? "",
      region: row.REGIAO ?? row.region ?? "",
      pole: row.POLO ?? row.pole ?? "",
      activityType: row.TIPO_ATIVIDADE ?? row.activityType ?? "",
      otherActivity: row.OUTRO_TIPO_ATIVIDADE ?? row.otherActivity ?? "",
      partnerId: row.ID_PARCEIRA ?? row.partnerId ?? "",
      partner: row.NOME_PARCEIRA ?? row.partner ?? "",
      contractNumber: row.NUMERO_CONTRATO ?? row.contractNumber ?? "",
      projectNumber: row.NUMERO_PROJETO ?? row.projectNumber ?? "",
      workOrder: row.NUMERO_OS ?? row.workOrder ?? "",
      circuit: row.CIRCUITO ?? row.circuit ?? "",
      interventionType: row.TIPO_INTERVENCAO ?? row.interventionType ?? "",
      location: row.LOCALIZACAO_GPS ?? row.location ?? "",
      latitude: row.LATITUDE ?? row.latitude ?? "",
      longitude: row.LONGITUDE ?? row.longitude ?? "",
      address: row.ENDERECO ?? row.address ?? "",
      municipality: row.MUNICIPIO ?? row.municipality ?? "",
      description: row.DESCRICAO_ATIVIDADE ?? row.description ?? "",
      situation: row.SITUACAO_ENCONTRADA ?? row.situation ?? "",
      actionsTaken: row.ACAO_REALIZADA ?? row.actionsTaken ?? "",
      activityResult: row.RESULTADO_ATIVIDADE ?? row.activityResult ?? "",
      hasNonConformity:
        row.EXISTE_NAO_CONFORMIDADE === true ||
        normalize(row.EXISTE_NAO_CONFORMIDADE ?? row.hasNonConformity) === "sim",
      nonConformity: row.DESCRICAO_NAO_CONFORMIDADE ?? row.nonConformity ?? "",
      hasPending:
        row.EXISTE_PENDENCIA === true || normalize(row.EXISTE_PENDENCIA ?? row.hasPending) === "sim",
      pendingDescription: row.DESCRICAO_PENDENCIA ?? row.pendingDescription ?? "",
      pendingOwner: row.RESPONSAVEL_PENDENCIA ?? row.pendingOwner ?? "",
      pendingDeadline: excelSerialToDate(row.PRAZO_PENDENCIA ?? row.pendingDeadline ?? ""),
      priority: row.PRIORIDADE ?? row.priority ?? "Normal",
      status: row.STATUS_ATIVIDADE ?? row.status ?? "Em andamento",
      photoCount: Number(row.QUANTIDADE_FOTOS ?? row.photoCount ?? photos.length ?? 0),
      mapLink: row.LINK_MAPA ?? row.mapLink ?? "",
      updatedAt: row.DATA_ULTIMA_ALTERACAO ?? row.updatedAt ?? new Date().toISOString(),
      updatedBy: row.USUARIO_ULTIMA_ALTERACAO ?? row.updatedBy ?? "",
      syncStatus: row.STATUS_SINCRONIZACAO ?? row.syncStatus ?? "Pendente",
      notes: row.OBSERVACOES ?? row.notes ?? "",
      photos,
      sample: Boolean(row.sample),
    };
  }

  function normalizeAction(row) {
    return {
      id: String(row.ID_ACAO ?? row.id ?? generateId("ACAO")),
      activityId: String(row.ID_ATIVIDADE ?? row.activityId ?? ""),
      type: row.TIPO_ACAO ?? row.type ?? "Pendência",
      description: row.DESCRICAO_ACAO ?? row.description ?? "",
      owner: row.RESPONSAVEL ?? row.owner ?? "",
      company: row.EMPRESA_RESPONSAVEL ?? row.company ?? "",
      openedAt: excelSerialToDate(row.DATA_ABERTURA ?? row.openedAt ?? todayISO()),
      deadline: excelSerialToDate(row.PRAZO ?? row.deadline ?? ""),
      priority: row.PRIORIDADE ?? row.priority ?? "Normal",
      status: row.STATUS_ACAO ?? row.status ?? "Aberta",
      completedAt: excelSerialToDate(row.DATA_CONCLUSAO ?? row.completedAt ?? ""),
      proof: row.COMPROVACAO_CONCLUSAO ?? row.proof ?? "",
      notes: row.OBSERVACOES ?? row.notes ?? "",
    };
  }

  function loadSettings() {
    try {
      state.settings = { ...state.settings, ...JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}") };
    } catch {
      state.settings = { endpoint: "", token: "" };
    }
  }

  async function loadModel() {
    const response = await fetch(MODEL_URL);
    if (!response.ok) throw new Error("Não foi possível carregar o modelo de dados.");
    state.model = await response.json();

    const typeRows = toObjects(state.model.TIPOS_ATIVIDADE);
    state.activityTypes = typeRows
      .filter((item) => normalize(item.STATUS_TIPO) !== "inativo")
      .sort((a, b) => Number(a.ORDEM_EXIBICAO || 999) - Number(b.ORDEM_EXIBICAO || 999));

    if (!state.activityTypes.length) {
      state.activityTypes = DEFAULT_ACTIVITY_TYPES.map((name, index) => ({
        ID_TIPO: `TIPO${String(index + 1).padStart(3, "0")}`,
        NOME_TIPO_ATIVIDADE: name,
        EXIGE_FOTO: "Sim",
        EXIGE_LOCALIZACAO: "Sim",
      }));
    }

    state.partners = toObjects(state.model.PARCEIRAS).filter(
      (item) => !item.STATUS_PARCEIRA || normalize(item.STATUS_PARCEIRA).startsWith("ativ"),
    );
    if (!state.partners.length) {
      state.partners = DEFAULT_PARTNERS.map((name, index) => ({
        ID_PARCEIRA: `PARC${String(index + 1).padStart(3, "0")}`,
        NOME_PARCEIRA: name,
      }));
    }

    const config = parseConfiguration(state.model.CONFIGURACOES);
    state.regions = config.regions;
    state.poles = config.poles;
    state.interventions = config.interventions;
    state.priorities = config.priorities;
  }

  async function seedExamples() {
    const seeded = await dbGet(STORE_META, "seeded");
    if (seeded) return;

    const modelRecords = toObjects(state.model.REGISTROS_ATIVIDADES).map((row) =>
      normalizeRecord({ ...row, sample: true }),
    );
    const modelActions = toObjects(state.model.ACOES_PENDENCIAS).map(normalizeAction);
    const modelPhotos = toObjects(state.model.FOTOS_ATIVIDADES);

    for (const record of modelRecords) {
      const photoRows = modelPhotos.filter((photo) => String(photo.ID_ATIVIDADE) === record.id);
      record.photos = photoRows.map((photo) => ({
        id: String(photo.ID_FOTO ?? generateId("FOTO")),
        name: photo.LEGENDA_FOTO || "Evidência do modelo",
        legend: photo.LEGENDA_FOTO || "",
        evidenceType: photo.TIPO_EVIDENCIA || "",
        dataUrl: "",
        url: photo.LINK_ARQUIVO || "",
      }));
      record.photoCount = Math.max(record.photoCount, photoRows.length);
      await dbPut(STORE_RECORDS, record);
    }
    for (const action of modelActions) await dbPut(STORE_ACTIONS, action);
    await dbPut(STORE_META, { key: "seeded", value: true });
  }

  async function refreshState() {
    state.records = (await dbGetAll(STORE_RECORDS)).map(normalizeRecord);
    state.actions = (await dbGetAll(STORE_ACTIONS)).map(normalizeAction);
    state.records.sort((a, b) =>
      `${b.activityDate}${b.startTime}`.localeCompare(`${a.activityDate}${a.startTime}`),
    );
    state.actions.sort((a, b) => String(a.deadline || "9999").localeCompare(String(b.deadline || "9999")));
  }

  function populateSelect(select, values, placeholder = "Selecione") {
    if (!select) return;
    select.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>${values
      .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
      .join("")}`;
  }

  function populateLists() {
    const activityNames = state.activityTypes.map((item) => item.NOME_TIPO_ATIVIDADE);
    const partnerNames = state.partners.map((item) => item.NOME_PARCEIRA);
    populateSelect($("#activityType"), activityNames);
    populateSelect($("#partner"), partnerNames);
    populateSelect($("#pole"), state.poles);
    populateSelect($("#interventionType"), state.interventions);
    populateSelect($("#priority"), state.priorities.length ? state.priorities : ["Crítica", "Alta", "Média", "Normal", "Baixa"]);
    populateSelect($("#activityStatus"), ACTIVITY_STATUSES);

    for (const select of [$("#filterActivity")]) {
      select.innerHTML = `<option value="">Todos</option>${activityNames
        .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
        .join("")}`;
    }
    $("#filterPartner").innerHTML = `<option value="">Todas</option>${partnerNames
      .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
      .join("")}`;
    $("#filterStatus").innerHTML = `<option value="">Todos</option>${ACTIVITY_STATUSES.map(
      (status) => `<option value="${escapeHtml(status)}">${escapeHtml(status)}</option>`,
    ).join("")}`;
    $("#pendingCompanyFilter").innerHTML = `<option value="">Todas</option>${partnerNames
      .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
      .join("")}`;
  }

  function filteredRecords() {
    const search = normalize($("#filterSearch").value);
    const start = $("#filterDateStart").value;
    const end = $("#filterDateEnd").value;
    const partner = $("#filterPartner").value;
    const type = $("#filterActivity").value;
    const region = $("#filterRegion").value;
    const status = $("#filterStatus").value;

    return state.records.filter((record) => {
      const haystack = normalize(
        [
          record.id,
          record.technician,
          record.activityType,
          record.partner,
          record.projectNumber,
          record.workOrder,
          record.address,
          record.municipality,
          record.description,
        ].join(" "),
      );
      return (
        (!search || haystack.includes(search)) &&
        (!start || record.activityDate >= start) &&
        (!end || record.activityDate <= end) &&
        (!partner || record.partner === partner) &&
        (!type || record.activityType === type) &&
        (!region || record.region === region) &&
        (!status || record.status === status)
      );
    });
  }

  function filteredActions() {
    const search = normalize($("#pendingSearch").value);
    const deadline = $("#pendingDeadlineFilter").value;
    const company = $("#pendingCompanyFilter").value;
    const priority = $("#pendingPriorityFilter").value;
    const status = $("#pendingStatusFilter").value;
    return state.actions.filter((action) => {
      const haystack = normalize(
        [action.id, action.activityId, action.type, action.description, action.owner, action.company].join(" "),
      );
      return (
        (!search || haystack.includes(search)) &&
        (!deadline || !action.deadline || action.deadline <= deadline) &&
        (!company || action.company === company) &&
        (!priority || action.priority === priority) &&
        (!status || action.status === status)
      );
    });
  }

  function recordRows(records) {
    return records
      .map(
        (record) => `
          <tr>
            <td><span class="cell-main">${escapeHtml(record.id)}</span><span class="cell-sub">${record.sample ? "Exemplo do modelo" : escapeHtml(record.syncStatus)}</span></td>
            <td><span class="cell-main">${formatDate(record.activityDate)}</span><span class="cell-sub">${escapeHtml(record.startTime || "Sem horário")}</span></td>
            <td><span class="cell-main">${escapeHtml(record.technician || "Não informado")}</span><span class="cell-sub">${escapeHtml(record.technicianRegistration)}</span></td>
            <td><span class="cell-main">${escapeHtml(record.activityType)}</span><span class="cell-sub">${escapeHtml(record.interventionType)}</span></td>
            <td>${escapeHtml(record.partner || "—")}</td>
            <td><span class="cell-main">${escapeHtml(record.projectNumber || "—")}</span><span class="cell-sub">${escapeHtml(record.workOrder || "")}</span></td>
            <td>${escapeHtml(record.region || "—")}</td>
            <td><span class="status-pill ${statusClass(record.status)}">${escapeHtml(record.status)}</span></td>
            <td><span class="photo-count">▧ ${record.photoCount || record.photos.length || 0}</span></td>
            <td><button class="row-action" type="button" data-detail-id="${escapeHtml(record.id)}">Ver</button></td>
          </tr>`,
      )
      .join("");
  }

  function recordCards(records) {
    return records
      .map(
        (record) => `
          <article class="mobile-data-card">
            <header>
              <div><h3>${escapeHtml(record.activityType)}</h3><p>${escapeHtml(record.id)}</p></div>
              <span class="status-pill ${statusClass(record.status)}">${escapeHtml(record.status)}</span>
            </header>
            <p>${escapeHtml(record.description || "Sem descrição")}</p>
            <div class="mobile-meta">
              <div><span>Data</span><strong>${formatDate(record.activityDate)}</strong></div>
              <div><span>Técnico</span><strong>${escapeHtml(record.technician || "Não informado")}</strong></div>
              <div><span>Parceira</span><strong>${escapeHtml(record.partner || "—")}</strong></div>
              <div><span>Fotos</span><strong>${record.photoCount || record.photos.length || 0}</strong></div>
            </div>
            <button class="row-action" type="button" data-detail-id="${escapeHtml(record.id)}">Ver detalhes</button>
          </article>`,
      )
      .join("");
  }

  function renderRecords() {
    const records = filteredRecords();
    $("#recordsTableBody").innerHTML = recordRows(records);
    $("#recordsCardList").innerHTML = recordCards(records);
    $("#recordsEmpty").classList.toggle("show", records.length === 0);
    $("#recordResultCount").textContent = `${records.length} ${records.length === 1 ? "registro" : "registros"}`;
  }

  function renderActions() {
    const actions = filteredActions();
    $("#pendingTableBody").innerHTML = actions
      .map(
        (action) => `
          <tr>
            <td><span class="cell-main">${escapeHtml(action.id)}</span></td>
            <td><button class="text-button" type="button" data-detail-id="${escapeHtml(action.activityId)}">${escapeHtml(action.activityId)}</button></td>
            <td>${escapeHtml(action.type)}</td>
            <td>${escapeHtml(action.description)}</td>
            <td>${escapeHtml(action.owner || "—")}</td>
            <td>${escapeHtml(action.company || "—")}</td>
            <td>${formatDate(action.deadline)}</td>
            <td><span class="status-pill ${statusClass(action.priority)}">${escapeHtml(action.priority)}</span></td>
            <td><span class="status-pill ${statusClass(action.status)}">${escapeHtml(action.status)}</span></td>
          </tr>`,
      )
      .join("");
    $("#pendingCardList").innerHTML = actions
      .map(
        (action) => `
          <article class="mobile-data-card">
            <header>
              <div><h3>${escapeHtml(action.type)}</h3><p>${escapeHtml(action.id)}</p></div>
              <span class="status-pill ${statusClass(action.status)}">${escapeHtml(action.status)}</span>
            </header>
            <p>${escapeHtml(action.description)}</p>
            <div class="mobile-meta">
              <div><span>Atividade</span><strong>${escapeHtml(action.activityId)}</strong></div>
              <div><span>Prazo</span><strong>${formatDate(action.deadline)}</strong></div>
              <div><span>Responsável</span><strong>${escapeHtml(action.owner || "—")}</strong></div>
              <div><span>Prioridade</span><strong>${escapeHtml(action.priority)}</strong></div>
            </div>
          </article>`,
      )
      .join("");
    $("#pendingEmpty").classList.toggle("show", actions.length === 0);
    $("#pendingResultCount").textContent = `${actions.length} ${actions.length === 1 ? "pendência" : "pendências"}`;
  }

  function countBy(items, key) {
    return items.reduce((acc, item) => {
      const label = item[key] || "Não informado";
      acc[label] = (acc[label] || 0) + 1;
      return acc;
    }, {});
  }

  function renderCharts() {
    if (typeof Chart === "undefined") return;
    const statusCounts = countBy(state.records, "status");
    const typeCounts = Object.entries(countBy(state.records, "activityType"))
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6);

    state.statusChart?.destroy();
    state.typeChart?.destroy();

    state.statusChart = new Chart($("#statusChart"), {
      type: "doughnut",
      data: {
        labels: Object.keys(statusCounts),
        datasets: [
          {
            data: Object.values(statusCounts),
            backgroundColor: ["#0f5bff", "#14a66f", "#f4147d", "#ff5a1f", "#7c70ee", "#55bcd2"],
            borderWidth: 0,
            hoverOffset: 6,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "68%",
        plugins: {
          legend: {
            position: "bottom",
            labels: { usePointStyle: true, pointStyle: "circle", boxWidth: 8, color: "#667287", font: { size: 10 } },
          },
        },
      },
    });

    state.typeChart = new Chart($("#typeChart"), {
      type: "bar",
      data: {
        labels: typeCounts.map(([label]) => label.length > 32 ? `${label.slice(0, 29)}…` : label),
        datasets: [
          {
            label: "Atividades",
            data: typeCounts.map(([, value]) => value),
            backgroundColor: ["#0f5bff", "#3f7cf6", "#55bcd2", "#14a66f", "#f4147d", "#ff5a1f"],
            borderRadius: 7,
            borderSkipped: false,
          },
        ],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#edf0f5" } },
          y: { grid: { display: false }, ticks: { color: "#667287", font: { size: 9 } } },
        },
        plugins: { legend: { display: false } },
      },
    });
  }

  function renderRecent() {
    const recent = state.records.slice(0, 3);
    $("#recentRecords").innerHTML = recent.length
      ? recent
          .map(
            (record) => `
              <article class="recent-item" data-detail-id="${escapeHtml(record.id)}" tabindex="0">
                <header>
                  <strong>${escapeHtml(record.technician || "Técnico não informado")}</strong>
                  <span class="status-pill ${statusClass(record.status)}">${escapeHtml(record.status)}</span>
                </header>
                <p>${escapeHtml(record.description || record.activityType)}</p>
                <small>${formatDateTime(record.activityDate, record.startTime)} • ${escapeHtml(record.partner || "Sem parceira")}</small>
              </article>`,
          )
          .join("")
      : `<div class="empty-mini"><div><strong>Nenhum registro de campo</strong><p>Use o botão “Registrar atividade” para começar.</p></div><button class="btn btn-primary" type="button" data-open-record>Novo registro</button></div>`;
  }

  function renderDashboard() {
    const completed = state.records.filter(
      (record) => normalize(record.status) === "concluida",
    ).length;
    const openActions = state.actions.filter(
      (action) => !["concluida", "cancelada"].includes(normalize(action.status)),
    ).length;
    const nonConformities = state.records.filter((record) => record.hasNonConformity).length;
    const photos = state.records.reduce(
      (sum, record) => sum + Math.max(record.photoCount || 0, record.photos?.length || 0),
      0,
    );
    const completion = state.records.length ? Math.round((completed / state.records.length) * 100) : 0;

    $("#kpiRecords").textContent = state.records.length;
    $("#kpiRecordsNote").textContent = state.records.length ? "Base consolidada" : "Nenhum registro no período";
    $("#kpiCompleted").textContent = completed;
    $("#kpiCompletedNote").textContent = `${completion}% dos registros`;
    $("#kpiPending").textContent = openActions;
    $("#kpiPendingNote").textContent = openActions ? "Requer acompanhamento" : "Nenhuma pendência";
    $("#kpiNonConformities").textContent = nonConformities;
    $("#kpiNonConformitiesNote").textContent = nonConformities ? "Ocorrências registradas" : "Nenhuma ocorrência";
    $("#heroPhotoCount").textContent = photos;
    $("#recordNavCount").textContent = state.records.length;
    $("#pendingNavCount").textContent = openActions;
    $("#currentPeriodLabel").textContent = getMonthLabel();
    renderCharts();
    renderRecent();
  }

  function renderStorage() {
    $("#storageRecords").textContent = state.records.length;
    $("#storagePendings").textContent = state.actions.length;
    $("#storagePendingSync").textContent = state.records.filter(
      (record) => normalize(record.syncStatus) !== "sincronizado",
    ).length;
  }

  function renderReportPreview() {
    const start = $("#reportStart").value;
    const end = $("#reportEnd").value;
    const records = state.records.filter(
      (record) => (!start || record.activityDate >= start) && (!end || record.activityDate <= end),
    );
    const ids = new Set(records.map((record) => record.id));
    const actions = state.actions.filter((action) => ids.has(action.activityId));
    $("#reportPreviewTitle").textContent = $("#reportTitle").value || "Relatório consolidado";
    $("#reportPreviewPeriod").textContent =
      start || end ? `${start ? formatDate(start) : "Início"} a ${end ? formatDate(end) : "Hoje"}` : "Todos os registros";
    $("#reportTotal").textContent = records.length;
    $("#reportPhotos").textContent = records.reduce(
      (sum, record) => sum + Math.max(record.photoCount || 0, record.photos?.length || 0),
      0,
    );
    $("#reportPendings").textContent = actions.filter(
      (action) => !["concluida", "cancelada"].includes(normalize(action.status)),
    ).length;
    $("#reportTechnicians").textContent = unique(records.map((record) => record.technician)).length;
  }

  function renderAll() {
    renderDashboard();
    renderRecords();
    renderActions();
    renderStorage();
    renderReportPreview();
    updateConnectionBadge();
  }

  function switchView(viewName) {
    $$(".view").forEach((view) => view.classList.toggle("active", view.dataset.view === viewName));
    $$("[data-view-target]").forEach((button) =>
      button.classList.toggle("active", button.dataset.viewTarget === viewName),
    );
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (viewName === "dashboard") requestAnimationFrame(renderCharts);
  }

  function resetRecordForm() {
    $("#recordForm").reset();
    $("#technician").value = "";
    $("#activityDate").value = todayISO();
    $("#startTime").value = currentTime();
    $("#activityStatus").value = "Em andamento";
    $("#priority").value =
      state.priorities.find((priority) => normalize(priority) === "normal") || state.priorities[0] || "Normal";
    state.selectedPhotos = [];
    $("#photoPreviewGrid").innerHTML = "";
    $("#otherActivityField").classList.add("hidden");
    $("#otherPartnerField").classList.add("hidden");
    $("#nonConformityField").classList.add("hidden");
    $("#pendingFields").classList.add("hidden");
    $("#generatedIdPreview").textContent = "O ID será gerado automaticamente ao salvar.";
    $("#locationFeedback").textContent = "A permissão de localização será solicitada pelo navegador.";
    $("#locationFeedback").className = "";
  }

  function openRecordDialog() {
    resetRecordForm();
    $("#recordDialog").showModal();
  }

  function closeDialog(dialog) {
    if (dialog?.open) dialog.close();
  }

  function getActivityTypeConfig(name) {
    return state.activityTypes.find((item) => item.NOME_TIPO_ATIVIDADE === name) || {};
  }

  async function fileToPhoto(file) {
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });

    return new Promise((resolve) => {
      const image = new Image();
      image.onload = () => {
        const max = 1600;
        const scale = Math.min(1, max / Math.max(image.width, image.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(image.width * scale));
        canvas.height = Math.max(1, Math.round(image.height * scale));
        const context = canvas.getContext("2d");
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        resolve({
          id: generateId("FOTO"),
          name: file.name || "foto-campo.jpg",
          legend: "",
          evidenceType: "Durante a atividade",
          dataUrl: canvas.toDataURL("image/jpeg", 0.78),
          url: "",
        });
      };
      image.onerror = () =>
        resolve({
          id: generateId("FOTO"),
          name: file.name || "foto-campo",
          legend: "",
          evidenceType: "Durante a atividade",
          dataUrl,
          url: "",
        });
      image.src = dataUrl;
    });
  }

  function renderPhotoPreview() {
    $("#photoPreviewGrid").innerHTML = state.selectedPhotos
      .map(
        (photo, index) => `
          <div class="photo-preview">
            <img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" />
            <button type="button" data-remove-photo="${index}" aria-label="Remover foto">×</button>
          </div>`,
      )
      .join("");
  }

  async function handlePhotos(files) {
    const available = Math.max(0, 8 - state.selectedPhotos.length);
    const selected = [...files].slice(0, available);
    if (!selected.length) {
      showToast("O limite é de 8 fotos por registro.", "error");
      return;
    }
    showToast("Preparando as fotos...");
    const photos = await Promise.all(selected.map(fileToPhoto));
    state.selectedPhotos.push(...photos);
    renderPhotoPreview();
    showToast(`${photos.length} foto(s) adicionada(s).`, "success");
  }

  function readFormRecord(statusOverride) {
    const partnerName =
      $("#partner").value === "Outras" ? $("#otherPartner").value.trim() : $("#partner").value;
    const activityName =
      $("#activityType").value === "Outras"
        ? $("#otherActivity").value.trim() || "Outras"
        : $("#activityType").value;
    const partnerInfo = state.partners.find((item) => item.NOME_PARCEIRA === $("#partner").value) || {};
    const latitude = $("#latitude").value;
    const longitude = $("#longitude").value;
    const now = new Date();
    const record = normalizeRecord({
      id: generateId("ATV"),
      createdDate: todayISO(),
      createdTime: currentTime(),
      activityDate: $("#activityDate").value,
      startTime: $("#startTime").value,
      endTime: $("#endTime").value,
      technician: $("#technician").value.trim(),
      technicianRegistration: $("#technicianRegistration").value.trim(),
      region: $("#region").value,
      pole: $("#pole").value,
      activityType: activityName,
      otherActivity: $("#activityType").value === "Outras" ? $("#otherActivity").value.trim() : "",
      partnerId: partnerInfo.ID_PARCEIRA || "",
      partner: partnerName,
      contractNumber: $("#contractNumber").value.trim() || partnerInfo.NUMERO_CONTRATO || "",
      projectNumber: $("#projectNumber").value.trim(),
      workOrder: $("#workOrder").value.trim(),
      circuit: $("#circuit").value.trim(),
      interventionType: $("#interventionType").value,
      location: latitude && longitude ? `${latitude},${longitude}` : "",
      latitude,
      longitude,
      address: $("#address").value.trim(),
      municipality: $("#municipality").value.trim(),
      description: $("#description").value.trim(),
      situation: $("#situation").value.trim(),
      actionsTaken: $("#actionsTaken").value.trim(),
      activityResult: $("#activityResult").value.trim(),
      hasNonConformity: $("#hasNonConformity").checked,
      nonConformity: $("#nonConformity").value.trim(),
      hasPending: $("#hasPending").checked,
      pendingDescription: $("#pendingDescription").value.trim(),
      pendingOwner: $("#pendingOwner").value.trim(),
      pendingDeadline: $("#pendingDeadline").value,
      priority: $("#priority").value || "Normal",
      status: statusOverride || $("#activityStatus").value || "Em andamento",
      photoCount: state.selectedPhotos.length,
      mapLink: latitude && longitude ? `https://maps.google.com/?q=${latitude},${longitude}` : "",
      updatedAt: now.toISOString(),
      updatedBy: $("#technician").value.trim(),
      syncStatus: "Pendente",
      photos: state.selectedPhotos,
      sample: false,
    });
    return record;
  }

  function validateRecord(record, draft = false) {
    if (draft) return true;
    if (!record.technician || !record.activityDate || !record.startTime || !record.activityType || !record.partner) {
      showToast("Preencha os campos obrigatórios de identificação.", "error");
      return false;
    }
    if (!record.description) {
      showToast("Descreva a atividade realizada.", "error");
      return false;
    }
    const typeConfig = getActivityTypeConfig($("#activityType").value);
    if (normalize(typeConfig.EXIGE_FOTO) === "sim" && !record.photos.length) {
      showToast("Este tipo de atividade exige pelo menos uma foto.", "error");
      return false;
    }
    if (
      normalize(typeConfig.EXIGE_LOCALIZACAO) === "sim" &&
      (!record.latitude || !record.longitude)
    ) {
      showToast("Este tipo de atividade exige as coordenadas de localização.", "error");
      return false;
    }
    if (record.hasNonConformity && !record.nonConformity) {
      showToast("Descreva a não conformidade identificada.", "error");
      return false;
    }
    if (
      record.hasPending &&
      (!record.pendingDescription || !record.pendingOwner || !record.pendingDeadline)
    ) {
      showToast("Informe descrição, responsável e prazo da pendência.", "error");
      return false;
    }
    return true;
  }

  async function saveRecord(draft = false) {
    const record = readFormRecord(draft ? "Rascunho" : undefined);
    if (!validateRecord(record, draft)) return;

    await dbPut(STORE_RECORDS, record);
    if (record.hasPending) {
      const action = normalizeAction({
        id: generateId("ACAO"),
        activityId: record.id,
        type: "Pendência",
        description: record.pendingDescription,
        owner: record.pendingOwner,
        company: $("#pendingCompany").value.trim() || record.partner,
        openedAt: record.activityDate,
        deadline: record.pendingDeadline,
        priority: record.priority,
        status: $("#pendingStatus").value || "Aberta",
      });
      await dbPut(STORE_ACTIONS, action);
    }

    await refreshState();
    renderAll();
    closeDialog($("#recordDialog"));
    showToast(draft ? "Rascunho salvo." : `Registro ${record.id} salvo com sucesso.`, "success");

    if (state.settings.endpoint && !draft) {
      syncOneRecord(record).catch(() => {
        showToast("Registro salvo localmente. A sincronização ficará pendente.", "error");
      });
    }
  }

  function openDetail(recordId) {
    const record = state.records.find((item) => item.id === recordId);
    if (!record) return;
    state.selectedRecordId = record.id;
    $("#detailTitle").textContent = record.id;
    $("#detailSubtitle").textContent = `${formatDateTime(record.activityDate, record.startTime)} • ${record.technician || "Técnico não informado"}`;

    const details = [
      ["Técnico Enel", record.technician || "Não informado"],
      ["Matrícula", record.technicianRegistration || "—"],
      ["Status", record.status],
      ["Tipo de atividade", record.activityType],
      ["Parceira", record.partner || "—"],
      ["Contrato", record.contractNumber || "—"],
      ["Projeto / OS", `${record.projectNumber || "—"} / ${record.workOrder || "—"}`],
      ["Circuito", record.circuit || "—"],
      ["Região / Polo", `${record.region || "—"} / ${record.pole || "—"}`],
      ["Localização", record.location || "Não capturada"],
      ["Endereço", [record.address, record.municipality].filter(Boolean).join(" • ") || "—"],
      ["Descrição da atividade", record.description || "—", true],
      ["Situação encontrada", record.situation || "—", true],
      ["Ações realizadas", record.actionsTaken || "—", true],
      ["Resultado", record.activityResult || "—", true],
      ["Não conformidade", record.hasNonConformity ? record.nonConformity || "Sim" : "Não", true],
      ["Pendência", record.hasPending ? record.pendingDescription || "Sim" : "Não", true],
      ["Observações", record.notes || "—", true],
    ];

    const photoHtml = record.photos?.length
      ? `<div class="detail-item wide"><span>Evidências fotográficas</span><div class="detail-photo-grid">${record.photos
          .map((photo) =>
            photo.dataUrl
              ? `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name || "Evidência")}" />`
              : `<div class="detail-item"><strong>${escapeHtml(photo.legend || photo.name || "Evidência do modelo")}</strong><p>${escapeHtml(photo.evidenceType || "")}</p></div>`,
          )
          .join("")}</div></div>`
      : "";

    $("#detailContent").innerHTML = `<div class="detail-grid">${details
      .map(
        ([label, value, wide]) =>
          `<div class="detail-item${wide ? " wide" : ""}"><span>${escapeHtml(label)}</span><p>${escapeHtml(value)}</p></div>`,
      )
      .join("")}${photoHtml}</div>`;
    $("#detailDialog").showModal();
  }

  async function deleteSelectedRecord() {
    if (!state.selectedRecordId) return;
    const record = state.records.find((item) => item.id === state.selectedRecordId);
    if (!record || !window.confirm(`Excluir o registro ${record.id}? Esta ação removerá apenas a cópia local.`)) return;
    await dbDelete(STORE_RECORDS, record.id);
    for (const action of state.actions.filter((item) => item.activityId === record.id)) {
      await dbDelete(STORE_ACTIONS, action.id);
    }
    await refreshState();
    renderAll();
    closeDialog($("#detailDialog"));
    showToast("Registro local excluído.", "success");
  }

  function captureLocation() {
    const feedback = $("#locationFeedback");
    if (!navigator.geolocation) {
      feedback.textContent = "O navegador não oferece suporte à geolocalização.";
      feedback.className = "error";
      return;
    }
    feedback.textContent = "Capturando localização...";
    feedback.className = "";
    navigator.geolocation.getCurrentPosition(
      (position) => {
        $("#latitude").value = position.coords.latitude.toFixed(7);
        $("#longitude").value = position.coords.longitude.toFixed(7);
        feedback.textContent = `Localização capturada com precisão aproximada de ${Math.round(position.coords.accuracy)} metros.`;
        feedback.className = "success";
      },
      (error) => {
        feedback.textContent =
          error.code === error.PERMISSION_DENIED
            ? "Permissão de localização negada. Preencha as coordenadas manualmente."
            : "Não foi possível obter a localização. Tente novamente em área aberta.";
        feedback.className = "error";
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 },
    );
  }

  function mapRecordToSheet(record) {
    return [
      record.id,
      record.createdDate,
      record.createdTime,
      record.activityDate,
      record.startTime,
      record.endTime,
      record.duration,
      record.userEmail,
      record.technicianId,
      record.technician,
      record.technicianRegistration,
      record.region,
      record.pole,
      record.activityType,
      record.otherActivity,
      record.partnerId,
      record.partner,
      record.contractNumber,
      record.projectNumber,
      record.workOrder,
      record.circuit,
      record.interventionType,
      record.location,
      Number(record.latitude) || "",
      Number(record.longitude) || "",
      record.address,
      record.municipality,
      record.description,
      record.situation,
      record.actionsTaken,
      record.activityResult,
      record.hasNonConformity ? "Sim" : "Não",
      record.nonConformity,
      record.hasPending ? "Sim" : "Não",
      record.pendingDescription,
      record.pendingOwner,
      record.pendingDeadline,
      record.priority,
      record.status,
      Math.max(record.photoCount || 0, record.photos?.length || 0),
      record.mapLink,
      record.updatedAt,
      record.updatedBy,
      record.syncStatus,
      record.notes,
    ];
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function recordsInPeriod(start, end) {
    return state.records.filter(
      (record) => (!start || record.activityDate >= start) && (!end || record.activityDate <= end),
    );
  }

  async function exportExcel(records, title = "Relatório de atividades de campo") {
    if (typeof ExcelJS === "undefined") {
      showToast("O módulo de Excel não foi carregado.", "error");
      return;
    }
    showToast("Gerando Excel com fotos...");
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Registro de Atividades de Campo";
    workbook.created = new Date();

    const summary = workbook.addWorksheet("RESUMO", { views: [{ showGridLines: false }] });
    summary.columns = [{ width: 30 }, { width: 22 }, { width: 24 }, { width: 24 }];
    summary.mergeCells("A1:D2");
    summary.getCell("A1").value = title;
    summary.getCell("A1").font = { size: 20, bold: true, color: { argb: "FFFFFFFF" } };
    summary.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0D4FB7" } };
    summary.getCell("A1").alignment = { vertical: "middle", horizontal: "left" };
    const recordIds = new Set(records.map((record) => record.id));
    const relatedActions = state.actions.filter((action) => recordIds.has(action.activityId));
    const summaryData = [
      ["Total de atividades", records.length],
      ["Atividades concluídas", records.filter((record) => normalize(record.status) === "concluida").length],
      ["Atividades com pendência", records.filter((record) => record.hasPending).length],
      ["Não conformidades", records.filter((record) => record.hasNonConformity).length],
      ["Fotos registradas", records.reduce((sum, record) => sum + Math.max(record.photoCount || 0, record.photos?.length || 0), 0)],
      ["Ações e pendências", relatedActions.length],
      ["Técnicos envolvidos", unique(records.map((record) => record.technician)).length],
    ];
    summaryData.forEach(([label, value], index) => {
      const row = summary.getRow(index + 4);
      row.values = [label, value];
      row.getCell(1).font = { bold: true, color: { argb: "FF42516A" } };
      row.getCell(2).font = { bold: true, size: 14, color: { argb: "FF0F5BFF" } };
      row.height = 26;
    });
    summary.getCell("A13").value = "Arquivo gerado pelo site de Registro de Atividades de Campo.";
    summary.getCell("A13").font = { italic: true, color: { argb: "FF7A8597" }, size: 9 };

    const recordHeaders =
      state.model.REGISTROS_ATIVIDADES?.[0] ||
      [
        "ID_ATIVIDADE",
        "DATA_CRIACAO",
        "HORA_CRIACAO",
        "DATA_ATIVIDADE",
        "HORA_INICIO",
        "HORA_TERMINO",
        "DURACAO_ATIVIDADE",
        "EMAIL_USUARIO",
        "ID_TECNICO",
        "NOME_TECNICO_ENEL",
        "MATRICULA_TECNICO",
        "REGIAO",
        "POLO",
        "TIPO_ATIVIDADE",
        "OUTRO_TIPO_ATIVIDADE",
        "ID_PARCEIRA",
        "NOME_PARCEIRA",
        "NUMERO_CONTRATO",
        "NUMERO_PROJETO",
        "NUMERO_OS",
        "CIRCUITO",
        "TIPO_INTERVENCAO",
        "LOCALIZACAO_GPS",
        "LATITUDE",
        "LONGITUDE",
        "ENDERECO",
        "MUNICIPIO",
        "DESCRICAO_ATIVIDADE",
        "SITUACAO_ENCONTRADA",
        "ACAO_REALIZADA",
        "RESULTADO_ATIVIDADE",
        "EXISTE_NAO_CONFORMIDADE",
        "DESCRICAO_NAO_CONFORMIDADE",
        "EXISTE_PENDENCIA",
        "DESCRICAO_PENDENCIA",
        "RESPONSAVEL_PENDENCIA",
        "PRAZO_PENDENCIA",
        "PRIORIDADE",
        "STATUS_ATIVIDADE",
        "QUANTIDADE_FOTOS",
        "LINK_MAPA",
        "DATA_ULTIMA_ALTERACAO",
        "USUARIO_ULTIMA_ALTERACAO",
        "STATUS_SINCRONIZACAO",
        "OBSERVACOES",
      ];
    const recordSheet = workbook.addWorksheet("REGISTROS_ATIVIDADES", {
      views: [{ state: "frozen", ySplit: 1 }],
      properties: { defaultRowHeight: 22 },
    });
    recordSheet.addRow(recordHeaders);
    records.forEach((record) => recordSheet.addRow(mapRecordToSheet(record)));
    recordSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: recordHeaders.length } };
    recordSheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF154F82" } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    });
    recordSheet.getRow(1).height = 34;
    recordSheet.columns.forEach((column, index) => {
      const header = String(recordHeaders[index] || "");
      column.width = /DESCRICAO|SITUACAO|ACAO_|ENDERECO|OBSERVACOES/.test(header)
        ? 35
        : /ID_|NUMERO_|LINK|EMAIL|TIPO_/.test(header)
          ? 22
          : 16;
      column.alignment = { vertical: "top", wrapText: true };
    });
    recordSheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1 && rowNumber % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F8FC" } };
        });
      }
    });

    const actionSheet = workbook.addWorksheet("ACOES_PENDENCIAS", {
      views: [{ state: "frozen", ySplit: 1 }],
    });
    const actionHeaders =
      state.model.ACOES_PENDENCIAS?.[0] ||
      [
        "ID_ACAO",
        "ID_ATIVIDADE",
        "TIPO_ACAO",
        "DESCRICAO_ACAO",
        "RESPONSAVEL",
        "EMPRESA_RESPONSAVEL",
        "DATA_ABERTURA",
        "PRAZO",
        "PRIORIDADE",
        "STATUS_ACAO",
        "DATA_CONCLUSAO",
        "COMPROVACAO_CONCLUSAO",
        "OBSERVACOES",
      ];
    actionSheet.addRow(actionHeaders);
    relatedActions.forEach((action) =>
      actionSheet.addRow([
        action.id,
        action.activityId,
        action.type,
        action.description,
        action.owner,
        action.company,
        action.openedAt,
        action.deadline,
        action.priority,
        action.status,
        action.completedAt,
        action.proof,
        action.notes,
      ]),
    );
    actionSheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF154F82" } };
      cell.alignment = { wrapText: true, horizontal: "center" };
    });
    actionSheet.columns.forEach((column, index) => {
      column.width = [3, 11, 12].includes(index) ? 38 : 20;
      column.alignment = { vertical: "top", wrapText: true };
    });

    const photoSheet = workbook.addWorksheet("FOTOS_ATIVIDADES", {
      views: [{ state: "frozen", ySplit: 1 }],
    });
    photoSheet.columns = [
      { header: "ID_FOTO", width: 22 },
      { header: "ID_ATIVIDADE", width: 24 },
      { header: "LEGENDA_FOTO", width: 28 },
      { header: "TIPO_EVIDENCIA", width: 22 },
      { header: "DATA_FOTO", width: 14 },
      { header: "LOCALIZACAO_FOTO", width: 24 },
      { header: "LINK_ARQUIVO", width: 34 },
      { header: "IMAGEM", width: 24 },
    ];
    photoSheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF154F82" } };
      cell.alignment = { horizontal: "center" };
    });
    let photoIndex = 1;
    for (const record of records) {
      const photos = record.photos?.length
        ? record.photos
        : Array.from({ length: record.photoCount || 0 }, (_, index) => ({
            id: `FOTO-${record.id}-${index + 1}`,
            name: "Evidência registrada na base",
            legend: "Foto cadastrada no modelo",
            dataUrl: "",
            url: "",
          }));
      for (const photo of photos) {
        photoIndex += 1;
        photoSheet.addRow([
          photo.id || generateId("FOTO"),
          record.id,
          photo.legend || photo.name || "",
          photo.evidenceType || "",
          record.activityDate,
          record.location,
          photo.url || "",
          photo.dataUrl ? "" : "Imagem não armazenada localmente",
        ]);
        photoSheet.getRow(photoIndex).height = photo.dataUrl ? 88 : 28;
        if (photo.dataUrl) {
          const extension = photo.dataUrl.startsWith("data:image/png") ? "png" : "jpeg";
          const imageId = workbook.addImage({ base64: photo.dataUrl, extension });
          photoSheet.addImage(imageId, {
            tl: { col: 7.05, row: photoIndex - 0.95 },
            ext: { width: 120, height: 80 },
          });
        }
      }
    }
    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      `Relatorio_Atividades_Campo_${todayISO()}.xlsx`,
    );
    showToast("Excel gerado com sucesso.", "success");
  }

  async function exportPdf(records, title = "Relatório consolidado de atividades de campo") {
    if (!window.jspdf?.jsPDF) {
      showToast("O módulo de PDF não foi carregado.", "error");
      return;
    }
    showToast("Gerando PDF com fotos...");
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
    const width = doc.internal.pageSize.getWidth();
    const height = doc.internal.pageSize.getHeight();
    const recordIds = new Set(records.map((record) => record.id));
    const relatedActions = state.actions.filter((action) => recordIds.has(action.activityId));
    const photoCount = records.reduce(
      (sum, record) => sum + Math.max(record.photoCount || 0, record.photos?.length || 0),
      0,
    );

    doc.setFillColor(15, 91, 255);
    doc.rect(0, 0, width, 44, "F");
    doc.setFillColor(244, 20, 125);
    doc.rect(width - 42, 0, 42, 44, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(21);
    doc.text(title, 16, 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text(`Gerado em ${new Date().toLocaleString("pt-BR")}`, 16, 30);
    doc.setTextColor(28, 43, 69);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("Resumo executivo", 16, 57);
    const metrics = [
      ["Atividades", records.length],
      ["Concluídas", records.filter((record) => normalize(record.status) === "concluida").length],
      ["Pendências", relatedActions.filter((action) => !["concluida", "cancelada"].includes(normalize(action.status))).length],
      ["Não conformidades", records.filter((record) => record.hasNonConformity).length],
      ["Fotos", photoCount],
    ];
    metrics.forEach(([label, value], index) => {
      const x = 16 + index * 53;
      doc.setFillColor(246, 248, 252);
      doc.roundedRect(x, 64, 46, 25, 3, 3, "F");
      doc.setTextColor(15, 91, 255);
      doc.setFontSize(15);
      doc.text(String(value), x + 5, 75);
      doc.setTextColor(103, 116, 139);
      doc.setFontSize(7);
      doc.text(label, x + 5, 83);
    });

    const rows = records.map((record) => [
      record.id,
      formatDate(record.activityDate),
      record.technician || "Não informado",
      record.activityType,
      record.partner || "—",
      record.projectNumber || "—",
      record.region || "—",
      record.status,
      String(Math.max(record.photoCount || 0, record.photos?.length || 0)),
    ]);
    if (typeof doc.autoTable === "function") {
      doc.autoTable({
        startY: 98,
        head: [["ID", "Data", "Técnico", "Atividade", "Parceira", "Projeto", "Região", "Status", "Fotos"]],
        body: rows,
        theme: "grid",
        styles: { fontSize: 6.8, cellPadding: 2, valign: "middle", overflow: "linebreak" },
        headStyles: { fillColor: [21, 79, 130], textColor: 255, fontStyle: "bold" },
        alternateRowStyles: { fillColor: [246, 248, 252] },
        columnStyles: { 0: { cellWidth: 28 }, 3: { cellWidth: 49 }, 4: { cellWidth: 32 } },
        margin: { left: 12, right: 12 },
      });
    }

    if (relatedActions.length) {
      doc.addPage();
      doc.setFontSize(16);
      doc.setTextColor(28, 43, 69);
      doc.text("Ações e pendências", 14, 18);
      doc.autoTable?.({
        startY: 25,
        head: [["ID", "Atividade", "Descrição", "Responsável", "Empresa", "Prazo", "Prioridade", "Status"]],
        body: relatedActions.map((action) => [
          action.id,
          action.activityId,
          action.description,
          action.owner,
          action.company,
          formatDate(action.deadline),
          action.priority,
          action.status,
        ]),
        styles: { fontSize: 7, cellPadding: 2, overflow: "linebreak" },
        headStyles: { fillColor: [21, 79, 130], textColor: 255 },
        columnStyles: { 2: { cellWidth: 72 } },
      });
    }

    for (const record of records) {
      const realPhotos = (record.photos || []).filter((photo) => photo.dataUrl);
      if (!realPhotos.length) continue;
      doc.addPage();
      doc.setTextColor(28, 43, 69);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(15);
      doc.text(`Evidências — ${record.id}`, 14, 17);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.text(
        `${record.activityType} • ${record.technician || "Técnico não informado"} • ${formatDate(record.activityDate)}`,
        14,
        24,
      );
      realPhotos.slice(0, 6).forEach((photo, index) => {
        const col = index % 3;
        const row = Math.floor(index / 3);
        const x = 14 + col * 91;
        const y = 32 + row * 74;
        try {
          doc.addImage(photo.dataUrl, "JPEG", x, y, 84, 58, undefined, "FAST");
          doc.setFontSize(7);
          doc.text((photo.legend || photo.name || `Foto ${index + 1}`).slice(0, 54), x, y + 63);
        } catch {
          doc.setFillColor(240, 243, 248);
          doc.rect(x, y, 84, 58, "F");
          doc.text("Imagem indisponível", x + 5, y + 29);
        }
      });
    }

    for (let page = 1; page <= doc.getNumberOfPages(); page += 1) {
      doc.setPage(page);
      doc.setTextColor(120, 130, 145);
      doc.setFontSize(7);
      doc.text(`Página ${page} de ${doc.getNumberOfPages()}`, width - 34, height - 6);
    }
    doc.save(`Relatorio_Atividades_Campo_${todayISO()}.pdf`);
    showToast("PDF gerado com sucesso.", "success");
  }

  async function importWorkbook(file) {
    if (!file || typeof ExcelJS === "undefined") return;
    showToast("Analisando as abas do Excel...");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const imported = {};
    workbook.eachSheet((worksheet) => {
      const rows = [];
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        const values = row.values.slice(1).map((value) => {
          if (value && typeof value === "object" && "text" in value) return value.text;
          if (value && typeof value === "object" && "result" in value) return value.result;
          return value ?? null;
        });
        if (values.some((value) => value !== null && value !== "")) rows.push(values);
      });
      imported[worksheet.name.trim()] = rows;
    });

    let recordsImported = 0;
    if (imported.REGISTROS_ATIVIDADES) {
      for (const row of toObjects(imported.REGISTROS_ATIVIDADES)) {
        await dbPut(STORE_RECORDS, normalizeRecord(row));
        recordsImported += 1;
      }
    }
    if (imported.ACOES_PENDENCIAS) {
      for (const row of toObjects(imported.ACOES_PENDENCIAS)) {
        await dbPut(STORE_ACTIONS, normalizeAction(row));
      }
    }
    if (imported.TIPOS_ATIVIDADE?.length) {
      state.model.TIPOS_ATIVIDADE = imported.TIPOS_ATIVIDADE;
    }
    if (imported.PARCEIRAS?.length) {
      state.model.PARCEIRAS = imported.PARCEIRAS;
    }
    await loadListsFromCurrentModel();
    await refreshState();
    populateLists();
    renderAll();
    showToast(
      recordsImported
        ? `${recordsImported} registro(s) importado(s) e listas atualizadas.`
        : "Listas e estrutura do arquivo importadas.",
      "success",
    );
  }

  async function loadListsFromCurrentModel() {
    const types = toObjects(state.model.TIPOS_ATIVIDADE);
    if (types.length) state.activityTypes = types;
    const partners = toObjects(state.model.PARCEIRAS);
    if (partners.length) state.partners = partners;
  }

  function updateConnectionBadge() {
    const badge = $("#connectionBadge");
    const text = $("#connectionText");
    const status = $("#sheetStatus");
    if (!navigator.onLine) {
      badge.className = "connection-badge offline";
      text.textContent = "Sem internet";
    } else if (state.settings.endpoint) {
      badge.className = "connection-badge online";
      text.textContent = "Google Sheets configurado";
    } else {
      badge.className = "connection-badge";
      text.textContent = "Dados locais";
    }
    if (status) {
      status.textContent = state.settings.endpoint ? "Configurado" : "Não conectado";
      status.className = `status-pill ${state.settings.endpoint ? "success" : "neutral"}`;
    }
  }

  async function testConnection() {
    const endpoint = $("#sheetEndpoint").value.trim();
    if (!endpoint) {
      showToast("Informe a URL publicada do Google Apps Script.", "error");
      return;
    }
    showToast("Testando a conexão...");
    const url = new URL(endpoint);
    url.searchParams.set("action", "ping");
    if ($("#sheetToken").value) url.searchParams.set("token", $("#sheetToken").value);
    const response = await fetch(url);
    const result = await response.json();
    if (!response.ok || result.ok === false) throw new Error(result.error || "Conexão recusada.");
    showToast("Conexão com o Google Sheets confirmada.", "success");
  }

  function saveSettings() {
    state.settings = {
      endpoint: $("#sheetEndpoint").value.trim(),
      token: $("#sheetToken").value.trim(),
    };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(state.settings));
    updateConnectionBadge();
    showToast("Configuração salva neste dispositivo.", "success");
  }

  async function syncOneRecord(record) {
    if (!state.settings.endpoint) throw new Error("Integração não configurada.");
    const action = state.actions.find((item) => item.activityId === record.id) || null;
    const response = await fetch(state.settings.endpoint, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "saveRecord",
        token: state.settings.token || "",
        record,
        pendingAction: action,
      }),
    });
    const result = await response.json();
    if (!response.ok || result.ok === false) throw new Error(result.error || "Falha ao sincronizar.");
    record.syncStatus = "Sincronizado";
    record.updatedAt = new Date().toISOString();
    await dbPut(STORE_RECORDS, record);
  }

  async function syncPending() {
    if (!state.settings.endpoint) {
      switchView("settings");
      showToast("Configure primeiro a integração com o Google Sheets.", "error");
      return;
    }
    const pending = state.records.filter((record) => normalize(record.syncStatus) !== "sincronizado");
    if (!pending.length) {
      showToast("Todos os registros locais já estão sincronizados.", "success");
      return;
    }
    showToast(`Sincronizando ${pending.length} registro(s)...`);
    let success = 0;
    for (const record of pending) {
      try {
        await syncOneRecord(record);
        success += 1;
      } catch {
        // Mantém o item pendente para nova tentativa.
      }
    }
    await refreshState();
    renderAll();
    showToast(
      success === pending.length
        ? "Sincronização concluída."
        : `${success} de ${pending.length} registros sincronizados.`,
      success ? "success" : "error",
    );
  }

  async function pullRecords() {
    if (!state.settings.endpoint) {
      showToast("Configure a URL do Google Apps Script.", "error");
      return;
    }
    const url = new URL(state.settings.endpoint);
    url.searchParams.set("action", "list");
    if (state.settings.token) url.searchParams.set("token", state.settings.token);
    showToast("Importando registros do Google Sheets...");
    const response = await fetch(url);
    const result = await response.json();
    if (!response.ok || result.ok === false) throw new Error(result.error || "Falha ao importar.");
    for (const row of result.records || []) await dbPut(STORE_RECORDS, normalizeRecord(row));
    for (const row of result.actions || []) await dbPut(STORE_ACTIONS, normalizeAction(row));
    await refreshState();
    renderAll();
    showToast("Dados do Google Sheets importados.", "success");
  }

  function setReportPreset(preset) {
    const now = new Date();
    let start = "";
    let end = "";
    if (preset === "first-half") {
      start = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
      end = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-15`;
    } else if (preset === "second-half") {
      start = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-16`;
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0).toISOString().slice(0, 10);
    } else if (preset === "month") {
      start = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0).toISOString().slice(0, 10);
    } else if (preset === "previous-month") {
      const previous = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      start = `${previous.getFullYear()}-${String(previous.getMonth() + 1).padStart(2, "0")}-01`;
      end = new Date(now.getFullYear(), now.getMonth(), 0).toISOString().slice(0, 10);
    }
    $("#reportStart").value = start;
    $("#reportEnd").value = end;
    renderReportPreview();
  }

  function showToast(message, type = "") {
    const toast = $("#toast");
    toast.textContent = message;
    toast.className = `toast show ${type}`.trim();
    clearTimeout(showToast.timer);
    showToast.timer = setTimeout(() => {
      toast.className = "toast";
    }, 3600);
  }

  function bindEvents() {
    document.addEventListener("click", (event) => {
      const viewButton = event.target.closest("[data-view-target]");
      if (viewButton) switchView(viewButton.dataset.viewTarget);

      if (event.target.closest("[data-open-record]")) openRecordDialog();
      if (event.target.closest("[data-close-dialog]")) closeDialog($("#recordDialog"));
      if (event.target.closest("[data-close-detail]")) closeDialog($("#detailDialog"));

      const detailButton = event.target.closest("[data-detail-id]");
      if (detailButton) openDetail(detailButton.dataset.detailId);

      const removePhoto = event.target.closest("[data-remove-photo]");
      if (removePhoto) {
        state.selectedPhotos.splice(Number(removePhoto.dataset.removePhoto), 1);
        renderPhotoPreview();
      }
    });

    $("#recordForm").addEventListener("submit", (event) => {
      event.preventDefault();
      saveRecord(false).catch((error) => showToast(error.message || "Não foi possível salvar.", "error"));
    });
    $("#saveDraft").addEventListener("click", () =>
      saveRecord(true).catch((error) => showToast(error.message || "Não foi possível salvar.", "error")),
    );
    $("#deleteRecord").addEventListener("click", () =>
      deleteSelectedRecord().catch((error) => showToast(error.message || "Não foi possível excluir.", "error")),
    );
    $("#captureLocation").addEventListener("click", captureLocation);
    $("#photoInput").addEventListener("change", (event) => {
      handlePhotos(event.target.files).finally(() => {
        event.target.value = "";
      });
    });

    $("#activityType").addEventListener("change", () => {
      $("#otherActivityField").classList.toggle("hidden", $("#activityType").value !== "Outras");
    });
    $("#partner").addEventListener("change", () => {
      const isOther = $("#partner").value === "Outras";
      $("#otherPartnerField").classList.toggle("hidden", !isOther);
      const partner = state.partners.find((item) => item.NOME_PARCEIRA === $("#partner").value);
      if (partner) {
        $("#contractNumber").value = partner.NUMERO_CONTRATO || "";
        if (partner.REGIAO_ATENDIDA && !$("#region").value) $("#region").value = partner.REGIAO_ATENDIDA;
        $("#pendingCompany").value = partner.NOME_PARCEIRA || "";
      }
    });
    $("#hasNonConformity").addEventListener("change", () => {
      $("#nonConformityField").classList.toggle("hidden", !$("#hasNonConformity").checked);
    });
    $("#hasPending").addEventListener("change", () => {
      $("#pendingFields").classList.toggle("hidden", !$("#hasPending").checked);
    });

    const recordFilterIds = [
      "filterSearch",
      "filterDateStart",
      "filterDateEnd",
      "filterPartner",
      "filterActivity",
      "filterRegion",
      "filterStatus",
    ];
    recordFilterIds.forEach((id) => {
      $(`#${id}`).addEventListener(id === "filterSearch" ? "input" : "change", renderRecords);
    });
    $("#resetFilters").addEventListener("click", () => {
      recordFilterIds.forEach((id) => {
        $(`#${id}`).value = "";
      });
      renderRecords();
    });

    const pendingFilterIds = [
      "pendingSearch",
      "pendingDeadlineFilter",
      "pendingCompanyFilter",
      "pendingPriorityFilter",
      "pendingStatusFilter",
    ];
    pendingFilterIds.forEach((id) => {
      $(`#${id}`).addEventListener(id === "pendingSearch" ? "input" : "change", renderActions);
    });
    $("#resetPendingFilters").addEventListener("click", () => {
      pendingFilterIds.forEach((id) => {
        $(`#${id}`).value = "";
      });
      renderActions();
    });

    $("#exportExcel").addEventListener("click", () => exportExcel(filteredRecords()));
    $("#exportPdf").addEventListener("click", () => exportPdf(filteredRecords()));
    $("#exportPeriod").addEventListener("click", () => {
      const start = $("#filterDateStart").value;
      const end = $("#filterDateEnd").value;

      if (!start || !end) {
        showToast("Selecione a data inicial e a data final para exportar o período.", "warning");
        return;
      }

      if (start > end) {
        showToast("A data inicial não pode ser posterior à data final.", "error");
        return;
      }

      const records = filteredRecords();
      if (!records.length) {
        showToast("Não há registros no período selecionado.", "warning");
        return;
      }

      exportExcel(
        records,
        `Atividades de campo — ${formatDate(start)} a ${formatDate(end)}`,
      );
    });
    $("#openOnlineReport").addEventListener("click", () => {
      const start = $("#filterDateStart").value;
      const end = $("#filterDateEnd").value;

      if (start) $("#reportStart").value = start;
      if (end) $("#reportEnd").value = end;
      if (start || end) {
        $("#reportPreset").value = "custom";
        $("#reportTitle").value = `Relatório online — ${
          start ? formatDate(start) : "Início"
        } a ${end ? formatDate(end) : "Hoje"}`;
      }

      renderReportPreview();
      switchView("reports");
    });
    $("#reportExcel").addEventListener("click", () =>
      exportExcel(
        recordsInPeriod($("#reportStart").value, $("#reportEnd").value),
        $("#reportTitle").value,
      ),
    );
    $("#reportPdf").addEventListener("click", () =>
      exportPdf(
        recordsInPeriod($("#reportStart").value, $("#reportEnd").value),
        $("#reportTitle").value,
      ),
    );
    $("#reportPreset").addEventListener("change", (event) => setReportPreset(event.target.value));
    ["reportStart", "reportEnd", "reportTitle"].forEach((id) =>
      $(`#${id}`).addEventListener(id === "reportTitle" ? "input" : "change", renderReportPreview),
    );

    $$(".import-workbook-mirror").forEach((input) =>
      input.addEventListener("change", (event) => {
        importWorkbook(event.target.files[0]).catch((error) =>
          showToast(error.message || "Não foi possível importar o arquivo.", "error"),
        );
        event.target.value = "";
      }),
    );
    $("#saveSettings").addEventListener("click", saveSettings);
    $("#testConnection").addEventListener("click", () =>
      testConnection().catch((error) => showToast(error.message || "Falha ao conectar.", "error")),
    );
    $("#pullRecords").addEventListener("click", () =>
      pullRecords().catch((error) => showToast(error.message || "Falha ao importar.", "error")),
    );
    $("#syncButton").addEventListener("click", () =>
      syncPending().catch((error) => showToast(error.message || "Falha ao sincronizar.", "error")),
    );
    $("#clearLocalRecords").addEventListener("click", async () => {
      if (!window.confirm("Limpar todos os registros e pendências salvos neste dispositivo?")) return;
      await dbClear(STORE_RECORDS);
      await dbClear(STORE_ACTIONS);
      await refreshState();
      renderAll();
      showToast("Dados locais removidos.", "success");
    });

    window.addEventListener("online", updateConnectionBadge);
    window.addEventListener("offline", updateConnectionBadge);
  }

  async function bootstrap() {
    try {
      loadSettings();
      await loadModel();
      state.db = await openDatabase();
      await seedExamples();
      await refreshState();
      populateLists();
      $("#sheetEndpoint").value = state.settings.endpoint || "";
      $("#sheetToken").value = state.settings.token || "";
      $("#reportStart").value = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}-01`;
      $("#reportEnd").value = todayISO();
      bindEvents();
      renderAll();
      if ("serviceWorker" in navigator) {
        navigator.serviceWorker.register("./service-worker.js").catch(() => undefined);
      }
    } catch (error) {
      console.error(error);
      showToast(error.message || "Falha ao iniciar o aplicativo.", "error");
    }
  }

  document.addEventListener("DOMContentLoaded", bootstrap);
})();
