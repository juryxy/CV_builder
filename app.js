const STORAGE_KEY = "offline-cv-builder-state-v4";
const HANDLE_DB_NAME = "cv-builder-fs";
const HANDLE_STORE_NAME = "handles";
const HANDLE_KEY = "root-folder";
const STORAGE_PROMPT_KEY = "cv-builder-storage-prompt-seen";

const demoPerson = {
  id: crypto.randomUUID(),
  folderName: null,
  profile: {
    name: "Prof. Dr. Alex Example",
    headline: "Research Scientist",
    summary: "This starter layout mirrors the attached CV style more closely: a formal title, section headings, label-value rows for personal details, date-description rows for career history, and publication-style list rows."
  },
  auditLog: [
    {
      id: crypto.randomUUID(),
      action: "Initialized demo CV registry",
      detail: "Loaded starter sections inspired by a formal academic CV template.",
      timestamp: new Date().toISOString()
    }
  ],
  sections: [
    {
      id: crypto.randomUUID(),
      title: "Personal information",
      subtitle: "",
      renderMode: "pair-list",
      includeInCv: true,
      fields: [
        { id: crypto.randomUUID(), label: "Label", key: "label", type: "text", width: "third" },
        { id: crypto.randomUUID(), label: "Value", key: "value", type: "text", width: "full" }
      ],
      entries: [
        { id: crypto.randomUUID(), values: { label: "Name", value: "Prof. Dr. Alex Example" } },
        { id: crypto.randomUUID(), values: { label: "Address", value: "Institute for Example Research, Sample Street 1, 12345 Sample City, Germany" } },
        { id: crypto.randomUUID(), values: { label: "E-mail", value: "alex@example.org" } },
        { id: crypto.randomUUID(), values: { label: "Nationality", value: "German" } }
      ]
    },
    {
      id: crypto.randomUUID(),
      title: "Professional experience",
      subtitle: "",
      renderMode: "date-list",
      includeInCv: true,
      fields: [
        { id: crypto.randomUUID(), label: "Period", key: "period", type: "text", width: "third" },
        { id: crypto.randomUUID(), label: "Role", key: "role", type: "text", width: "third" },
        { id: crypto.randomUUID(), label: "Institution", key: "institution", type: "text", width: "third" },
        { id: crypto.randomUUID(), label: "Details", key: "details", type: "textarea", width: "full" }
      ],
      entries: [
        {
          id: crypto.randomUUID(),
          values: {
            period: "2022-2025",
            role: "Professor of Example Science",
            institution: "University of Examples",
            details: "Led an interdisciplinary lab focused on translational methods and digital biomarkers."
          }
        }
      ]
    },
    {
      id: crypto.randomUUID(),
      title: "Selected publications",
      subtitle: "",
      renderMode: "publication-list",
      includeInCv: true,
      fields: [
        { id: crypto.randomUUID(), label: "No.", key: "number", type: "text", width: "quarter" },
        { id: crypto.randomUUID(), label: "Citation", key: "citation", type: "textarea", width: "full" },
        { id: crypto.randomUUID(), label: "Year", key: "year", type: "year", width: "quarter" }
      ],
      entries: [
        {
          id: crypto.randomUUID(),
          values: {
            number: "1",
            citation: "Example A, Researcher B, Sample C. A publication entry can be formatted here to match the compact table-like style of the Word CV.",
            year: "2025"
          }
        }
      ]
    }
  ]
};

const demoState = {
  activePersonId: demoPerson.id,
  people: [demoPerson]
};

let state = loadState();
let editingSectionId = null;
let hasPromptedPersonSelection = false;
let draggedEntryId = null;
let draggedSectionId = null;

const profileNameInput = document.getElementById("profileName");
const profileHeadlineInput = document.getElementById("profileHeadline");
const profileSummaryInput = document.getElementById("profileSummary");
const addEntryFlowBtn = document.getElementById("addEntryFlowBtn");
const addSectionBtn = document.getElementById("addSectionBtn");
const importBtn = document.getElementById("importBtn");
const addPersonalPresetBtn = document.getElementById("addPersonalPresetBtn");
const addTimelinePresetBtn = document.getElementById("addTimelinePresetBtn");
const addPublicationPresetBtn = document.getElementById("addPublicationPresetBtn");
const activePersonCard = document.getElementById("activePersonCard");
const switchPersonBtn = document.getElementById("switchPersonBtn");
const addPersonBtn = document.getElementById("addPersonBtn");
const connectFolderBtn = document.getElementById("connectFolderBtn");
const storageStatus = document.getElementById("storageStatus");
const exportJsonBtn = document.getElementById("exportJsonBtn");
const restoreJsonInput = document.getElementById("restoreJsonInput");
const printBtn = document.getElementById("printBtn");
const resetDemoBtn = document.getElementById("resetDemoBtn");
const sectionsContainer = document.getElementById("sectionsContainer");
const cvPreview = document.getElementById("cvPreview");
const showGuidesToggle = document.getElementById("showGuidesToggle");
const statsGrid = document.getElementById("statsGrid");
const sectionRegistry = document.getElementById("sectionRegistry");
const activityList = document.getElementById("activityList");
const activityLogPanel = document.getElementById("activityLogPanel");

const sectionDialog = document.getElementById("sectionDialog");
const sectionDialogTitle = document.getElementById("sectionDialogTitle");
const sectionForm = document.getElementById("sectionForm");
const sectionTitleInput = document.getElementById("sectionTitleInput");
const sectionSubtitleInput = document.getElementById("sectionSubtitleInput");
const sectionRenderModeInput = document.getElementById("sectionRenderModeInput");
const sectionIncludeInput = document.getElementById("sectionIncludeInput");
const fieldRows = document.getElementById("fieldRows");
const fieldRowTemplate = document.getElementById("fieldRowTemplate");
const addFieldBtn = document.getElementById("addFieldBtn");

const importDialog = document.getElementById("importDialog");
const importSectionSelect = document.getElementById("importSectionSelect");
const importDelimiterSelect = document.getElementById("importDelimiterSelect");
const importTextarea = document.getElementById("importTextarea");
const runImportBtn = document.getElementById("runImportBtn");
const entryDialog = document.getElementById("entryDialog");
const entrySectionSelect = document.getElementById("entrySectionSelect");
const entryFormFields = document.getElementById("entryFormFields");
const submitEntryBtn = document.getElementById("submitEntryBtn");
const entryFollowupDialog = document.getElementById("entryFollowupDialog");
const followupSameSectionBtn = document.getElementById("followupSameSectionBtn");
const followupDifferentSectionBtn = document.getElementById("followupDifferentSectionBtn");
const followupBackOverviewBtn = document.getElementById("followupBackOverviewBtn");
const personPickerDialog = document.getElementById("personPickerDialog");
const personPickerList = document.getElementById("personPickerList");
const personPickerAddBtn = document.getElementById("personPickerAddBtn");
const personDialog = document.getElementById("personDialog");
const personNameInput = document.getElementById("personNameInput");
const personHeadlineInput = document.getElementById("personHeadlineInput");
const personSummaryInput = document.getElementById("personSummaryInput");
const savePersonBtn = document.getElementById("savePersonBtn");
const storagePromptDialog = document.getElementById("storagePromptDialog");
const storagePromptYesBtn = document.getElementById("storagePromptYesBtn");
const storagePromptNoBtn = document.getElementById("storagePromptNoBtn");

let lastEntrySectionId = null;
let folderHandle = null;
let saveDebounceId = null;
let storageInfo = {
  mode: "browser-only",
  detail: "Using browser backup only.",
  lastSaved: ""
};

bindEvents();
initApp();

function on(element, eventName, handler) {
  if (element) element.addEventListener(eventName, handler);
}

function bindEvents() {
  [profileNameInput, profileHeadlineInput, profileSummaryInput].forEach((element) => {
    element.addEventListener("input", () => {
      const person = currentPerson();
      if (!person) return;
      person.profile.name = profileNameInput.value;
      person.profile.headline = profileHeadlineInput.value;
      person.profile.summary = profileSummaryInput.value;
      persist();
      renderDashboard();
      renderPreview();
    });
  });

  on(addEntryFlowBtn, "click", () => openEntryDialog());
  on(addSectionBtn, "click", () => openSectionDialog());
  on(switchPersonBtn, "click", openPersonPickerDialog);
  on(addPersonBtn, "click", openPersonDialog);
  on(connectFolderBtn, "click", connectFolder);
  on(importBtn, "click", openImportDialog);
  on(addPersonalPresetBtn, "click", () => addPresetSection("personal"));
  on(addTimelinePresetBtn, "click", () => addPresetSection("timeline"));
  on(addPublicationPresetBtn, "click", () => addPresetSection("publication"));
  on(exportJsonBtn, "click", exportState);
  on(restoreJsonInput, "change", restoreState);
  on(printBtn, "click", () => window.print());
  on(resetDemoBtn, "click", () => {
    state = structuredClone(demoState);
    persist();
    render();
  });
  on(showGuidesToggle, "change", renderPreview);
  on(addFieldBtn, "click", () => addFieldRow());

  on(fieldRows, "click", (event) => {
    const button = event.target.closest("[data-action='remove-field']");
    if (button) {
      button.closest(".field-row")?.remove();
    }
  });

  on(sectionForm, "submit", (event) => {
    event.preventDefault();
    saveSectionFromDialog();
  });

  on(runImportBtn, "click", importRowsIntoSection);
  on(entrySectionSelect, "change", renderEntryDialogFields);
  on(submitEntryBtn, "click", submitNewEntryFromDialog);
  on(followupSameSectionBtn, "click", () => reopenEntryDialogAfterAdd("same"));
  on(followupDifferentSectionBtn, "click", () => reopenEntryDialogAfterAdd("different"));
  on(followupBackOverviewBtn, "click", () => {
    entryFollowupDialog.close();
    activateTab("overviewTab");
  });
  on(personPickerAddBtn, "click", () => {
    personPickerDialog.close();
    openPersonDialog();
  });
  on(savePersonBtn, "click", saveNewPerson);
  on(storagePromptYesBtn, "click", async () => {
    localStorage.setItem(STORAGE_PROMPT_KEY, "seen");
    storagePromptDialog.close();
    await connectFolder();
  });
  on(storagePromptNoBtn, "click", () => {
    localStorage.setItem(STORAGE_PROMPT_KEY, "seen");
    storagePromptDialog.close();
  });
  on(sectionsContainer, "click", handleSectionClicks);
  on(sectionsContainer, "input", handleSectionInputs);
  on(sectionsContainer, "dragstart", handleEntryDragStart);
  on(sectionsContainer, "dragover", handleEntryDragOver);
  on(sectionsContainer, "drop", handleEntryDrop);
  on(sectionsContainer, "dragend", handleEntryDragEnd);

  document.querySelectorAll("[data-close-dialog]").forEach((button) => {
    button.addEventListener("click", () => {
      document.getElementById(button.dataset.closeDialog)?.close();
    });
  });

  document.querySelectorAll("[data-tab-target]").forEach((button) => {
    button.addEventListener("click", () => activateTab(button.dataset.tabTarget));
  });
}

async function initApp() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {}
  await hydrateFromConnectedFolder();
  render();
  maybePromptFolderConnection();
}

function clearRuntimeCache() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (error) {
    console.warn("Could not clear runtime cache", error);
  }
}

function loadState() {
  return structuredClone(demoState);
}

function normalizePerson(person) {
  if (!person.id) person.id = crypto.randomUUID();
  if (!person.profile) person.profile = { name: "", headline: "", summary: "" };
  if (!Array.isArray(person.sections)) person.sections = [];
  if (!Array.isArray(person.auditLog)) person.auditLog = [];
  if (!Object.prototype.hasOwnProperty.call(person, "folderName")) person.folderName = null;
}

function persist() {
  queueFilesystemSave();
}

function currentPerson() {
  return state.people.find((person) => person.id === state.activePersonId) || state.people[0] || null;
}

function render() {
  const person = currentPerson();
  if (profileNameInput) profileNameInput.value = person?.profile.name || "";
  if (profileHeadlineInput) profileHeadlineInput.value = person?.profile.headline || "";
  if (profileSummaryInput) profileSummaryInput.value = person?.profile.summary || "";
  renderSections();
  renderDashboard();
  renderPreview();
  renderActivePersonCard();
  renderStorageStatus();
  if (!hasPromptedPersonSelection && !personPickerDialog.open) {
    ensurePersonSelected();
  }
}

function renderSections() {
  const person = currentPerson();
  if (!person?.sections.length) {
    sectionsContainer.innerHTML = '<div class="empty-state">No sections yet. Create your first section to define a CV block.</div>';
    return;
  }

  sectionsContainer.innerHTML = person.sections.map(renderSectionCard).join("");
}

function renderSectionCard(section) {
  const entriesHtml = section.entries.length
    ? section.entries.map((entry) => renderEntryCard(section, entry)).join("")
    : '<div class="empty-state">No entries yet. Add one below or import rows from Excel.</div>';

  return `
    <article class="section-card" data-section-id="${section.id}">
      <div class="section-card-header">
        <div>
          <h3>${escapeHtml(section.title)}</h3>
          <div class="section-meta">
            <span>${section.fields.length} fields</span>
            <span>${section.entries.length} entries</span>
            <span>${renderModeLabel(section.renderMode || "grid")}</span>
            <span>${section.includeInCv ? "Included in output" : "Hidden from output"}</span>
          </div>
        </div>
        <div class="button-stack">
          <button data-action="toggle-include">${section.includeInCv ? "Hide in PDF" : "Show in PDF"}</button>
          <button data-action="edit-section">Edit schema</button>
          <button data-action="delete-section" class="danger">Delete section</button>
        </div>
      </div>
      <div class="section-card-body"><div class="entries-list">${entriesHtml}</div></div>
    </article>
  `;
}

function renderDashboard() {
  const person = currentPerson();
  if (!person) return;
  const includedSections = person.sections.filter((section) => section.includeInCv).length;
  const totalEntries = person.sections.reduce((sum, section) => sum + section.entries.length, 0);
  const emptySections = person.sections.filter((section) => !section.entries.some((entry) => Object.values(entry.values || {}).some((value) => String(value).trim()))).length;
  const lastUpdated = person.auditLog[0]?.timestamp ? formatTimestamp(person.auditLog[0].timestamp) : "No edits yet";

  statsGrid.innerHTML = [
    statCard("Sections", person.sections.length, "Configured CV sections"),
    statCard("Entries", totalEntries, "Stored records across all sections"),
    statCard("Active in PDF", includedSections, "Sections enabled for export"),
    statCard("Last update", lastUpdated, emptySections ? `${emptySections} empty sections need attention` : "All sections contain at least one entry")
  ].join("");

  sectionRegistry.innerHTML = person.sections.length
    ? person.sections.map((section) => `
      <div class="registry-row">
        <div>
          <strong>${escapeHtml(section.title)}</strong>
          <div class="registry-meta">${section.entries.length} entries, ${renderModeLabel(section.renderMode || "grid")}</div>
        </div>
        <span>${section.includeInCv ? "Included" : "Hidden"}</span>
      </div>
    `).join("")
    : '<div class="empty-state">No sections available yet.</div>';

  const recentActivity = person.auditLog.slice(0, 6);
  const activityMarkup = recentActivity.length
    ? recentActivity.map((item) => `
      <div class="activity-row">
        <strong>${escapeHtml(item.action)}</strong>
        <div>${escapeHtml(item.detail || "")}</div>
        <div class="activity-meta">${formatTimestamp(item.timestamp)}</div>
      </div>
    `).join("")
    : '<div class="empty-state">No recent activity yet.</div>';

  activityList.innerHTML = activityMarkup;
  activityLogPanel.innerHTML = activityMarkup.replace(/activity-row/g, "audit-row").replace(/activity-meta/g, "audit-time");
}

function renderEntryCard(section, entry) {
  const fieldsHtml = section.fields.map((field) => {
    const value = entry.values?.[field.key] ?? "";
    return `<label class="entry-field" data-width="${field.width}">
      <span>${escapeHtml(field.label)}</span>
      ${renderFieldInput(section.id, entry.id, field, value)}
    </label>`;
  }).join("");

  return `
    <section class="entry-card" data-entry-id="${entry.id}" draggable="true" title="Drag to reorder entries">
      <div class="entry-header">
        <strong>${escapeHtml(summarizeEntry(section, entry))}</strong>
        <div class="entry-actions">
          <span class="entry-drag-hint" aria-hidden="true">Drag to reorder</span>
          <button data-action="delete-entry" class="danger small">Delete entry</button>
        </div>
      </div>
      <div class="entry-body">${fieldsHtml}</div>
    </section>
  `;
}

function renderFieldInput(sectionId, entryId, field, value) {
  const encoded = escapeHtml(value);
  const common = `data-input-field="${field.key}" data-section-id="${sectionId}" data-entry-id="${entryId}"`;
  if (field.type === "textarea" || field.type === "bullets") {
    return `<textarea rows="4" ${common}>${encoded}</textarea>`;
  }
  return `<input type="${inputTypeFor(field.type)}" value="${encoded}" ${common}>`;
}

function handleSectionClicks(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  const action = button.dataset.action;
  const sectionEl = button.closest("[data-section-id]");
  const person = currentPerson();
  const section = person?.sections.find((item) => item.id === sectionEl?.dataset.sectionId);
  if (!section) return;

  if (action === "toggle-include") {
    section.includeInCv = !section.includeInCv;
    pushAudit("Toggled section visibility", `${section.title} is now ${section.includeInCv ? "included" : "hidden"} in the PDF.`);
    persist();
    render();
  } else if (action === "edit-section") {
    openSectionDialog(section.id);
  } else if (action === "delete-section") {
    if (!window.confirm(`Delete section "${section.title}"?`)) return;
    person.sections = person.sections.filter((item) => item.id !== section.id);
    pushAudit("Deleted section", section.title);
    persist();
    render();
  } else if (action === "delete-entry") {
    const entryEl = button.closest("[data-entry-id]");
    section.entries = section.entries.filter((item) => item.id !== entryEl?.dataset.entryId);
    pushAudit("Deleted entry", `Removed an entry from ${section.title}.`);
    persist();
    render();
  }
}

function handleSectionInputs(event) {
  const target = event.target;
  if (!target.dataset.inputField) return;
  const person = currentPerson();
  const section = person?.sections.find((item) => item.id === target.dataset.sectionId);
  const entry = section?.entries.find((item) => item.id === target.dataset.entryId);
  if (!entry) return;
  entry.values[target.dataset.inputField] = target.value;
  persist();
  renderPreview();
  renderDashboard();
}
function handleEntryDragStart(event) {
  const entryCard = event.target.closest(".entry-card");
  const sectionEl = event.target.closest("[data-section-id]");
  if (!entryCard || !sectionEl) return;

  draggedEntryId = entryCard.dataset.entryId;
  draggedSectionId = sectionEl.dataset.sectionId;
  entryCard.classList.add("is-dragging");

  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", draggedEntryId);
  }
}

function handleEntryDragOver(event) {
  const sectionEl = event.target.closest("[data-section-id]");
  const entriesList = event.target.closest(".entries-list");
  if (!sectionEl || !entriesList || !draggedEntryId || sectionEl.dataset.sectionId !== draggedSectionId) return;

  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
}

function handleEntryDrop(event) {
  const sectionEl = event.target.closest("[data-section-id]");
  const entriesList = event.target.closest(".entries-list");
  if (!sectionEl || !entriesList || !draggedEntryId || sectionEl.dataset.sectionId !== draggedSectionId) {
    clearEntryDragState();
    return;
  }

  event.preventDefault();

  const person = currentPerson();
  const section = person?.sections.find((item) => item.id === sectionEl.dataset.sectionId);
  if (!section) {
    clearEntryDragState();
    return;
  }

  const fromIndex = section.entries.findIndex((item) => item.id === draggedEntryId);
  if (fromIndex < 0) {
    clearEntryDragState();
    return;
  }

  const targetEntryEl = event.target.closest(".entry-card");
  let targetIndex = section.entries.length;

  if (targetEntryEl) {
    const hoveredEntryId = targetEntryEl.dataset.entryId;
    targetIndex = section.entries.findIndex((item) => item.id === hoveredEntryId);
    if (targetIndex < 0) {
      clearEntryDragState();
      return;
    }

    const rect = targetEntryEl.getBoundingClientRect();
    const insertAfter = event.clientY > rect.top + (rect.height / 2);
    if (insertAfter) targetIndex += 1;
  }

  if (targetIndex > fromIndex) targetIndex -= 1;
  if (targetIndex === fromIndex || targetIndex < 0) {
    clearEntryDragState();
    return;
  }

  const [movedEntry] = section.entries.splice(fromIndex, 1);
  section.entries.splice(Math.min(targetIndex, section.entries.length), 0, movedEntry);

  pushAudit("Reordered entries", `Updated entry order in ${section.title}.`);
  persist();
  render();
  clearEntryDragState();
}

function handleEntryDragEnd() {
  clearEntryDragState();
}

function clearEntryDragState() {
  document.querySelectorAll(".entry-card.is-dragging").forEach((card) => card.classList.remove("is-dragging"));
  draggedEntryId = null;
  draggedSectionId = null;
}


function openSectionDialog(sectionId = null) {
  editingSectionId = sectionId;
  const person = currentPerson();
  const section = person?.sections.find((item) => item.id === sectionId);
  sectionDialogTitle.textContent = section ? "Edit section" : "New section";
  sectionTitleInput.value = section?.title || "";
  sectionSubtitleInput.value = section?.subtitle || "";
  sectionRenderModeInput.value = section?.renderMode || "grid";
  sectionIncludeInput.checked = section?.includeInCv ?? true;
  fieldRows.innerHTML = "";

  const fields = section?.fields?.length ? section.fields : [
    { label: "Title", key: "title", type: "text", width: "half" },
    { label: "Details", key: "details", type: "textarea", width: "full" }
  ];
  fields.forEach((field) => addFieldRow(field));
  sectionDialog.showModal();
}

function addFieldRow(field = {}) {
  const row = fieldRowTemplate.content.firstElementChild.cloneNode(true);
  row.querySelector('[data-prop="label"]').value = field.label || "";
  row.querySelector('[data-prop="key"]').value = field.key || "";
  row.querySelector('[data-prop="type"]').value = field.type || "text";
  row.querySelector('[data-prop="width"]').value = field.width || "full";
  fieldRows.append(row);
}

function saveSectionFromDialog() {
  const fields = Array.from(fieldRows.querySelectorAll(".field-row")).map((row) => ({
    id: crypto.randomUUID(),
    label: row.querySelector('[data-prop="label"]').value.trim(),
    key: slugify(row.querySelector('[data-prop="key"]').value.trim()),
    type: row.querySelector('[data-prop="type"]').value,
    width: row.querySelector('[data-prop="width"]').value
  })).filter((field) => field.label && field.key);

  if (!fields.length) return window.alert("Please add at least one field.");
  if (fields.some((field, index) => fields.findIndex((item) => item.key === field.key) !== index)) {
    return window.alert("Field keys must be unique inside a section.");
  }

  const nextSection = {
    id: editingSectionId || crypto.randomUUID(),
    title: sectionTitleInput.value.trim(),
    subtitle: sectionSubtitleInput.value.trim(),
    renderMode: sectionRenderModeInput.value,
    includeInCv: sectionIncludeInput.checked,
    fields,
    entries: []
  };

  if (!nextSection.title) return window.alert("Section title is required.");

  if (editingSectionId) {
    const currentSection = currentPerson()?.sections.find((item) => item.id === editingSectionId);
    currentSection.title = nextSection.title;
    currentSection.subtitle = nextSection.subtitle;
    currentSection.renderMode = nextSection.renderMode;
    currentSection.includeInCv = nextSection.includeInCv;
    currentSection.fields = fields;
    currentSection.entries = currentSection.entries.map((entry) => ({
      ...entry,
      values: Object.fromEntries(fields.map((field) => [field.key, entry.values?.[field.key] ?? ""]))
    }));
  } else {
    nextSection.entries.push({
      id: crypto.randomUUID(),
      values: Object.fromEntries(fields.map((field) => [field.key, ""]))
    });
    currentPerson().sections.push(nextSection);
  }

  pushAudit(editingSectionId ? "Updated section schema" : "Created section", nextSection.title);
  persist();
  render();
  sectionDialog.close();
}

function openEntryDialog(preselectedSectionId = null) {
  const person = currentPerson();
  if (!person?.sections.length) {
    window.alert("Create a section first so the new entry has somewhere to go.");
    return;
  }
  entrySectionSelect.innerHTML = person.sections.map((section) => `<option value="${section.id}">${escapeHtml(section.title)}</option>`).join("");
  entrySectionSelect.value = preselectedSectionId && person.sections.some((section) => section.id === preselectedSectionId)
    ? preselectedSectionId
    : person.sections[0].id;
  renderEntryDialogFields();
  entryDialog.showModal();
}

function renderEntryDialogFields() {
  const section = currentPerson()?.sections.find((item) => item.id === entrySectionSelect.value);
  if (!section) {
    entryFormFields.innerHTML = '<div class="empty-state">No section selected.</div>';
    return;
  }

  entryFormFields.innerHTML = `
    <div class="entry-create-grid">
      ${section.fields.map((field) => `
        <label class="entry-field" data-width="${field.width}">
          <span>${escapeHtml(field.label)}</span>
          ${renderNewEntryInput(field)}
        </label>
      `).join("")}
    </div>
  `;
}

function renderNewEntryInput(field) {
  const common = `data-entry-create-field="${field.key}"`;
  if (field.type === "textarea" || field.type === "bullets") {
    return `<textarea rows="4" ${common}></textarea>`;
  }
  return `<input type="${inputTypeFor(field.type)}" ${common}>`;
}

function submitNewEntryFromDialog() {
  const section = currentPerson()?.sections.find((item) => item.id === entrySectionSelect.value);
  if (!section) return;

  const values = {};
  entryFormFields.querySelectorAll("[data-entry-create-field]").forEach((input) => {
    values[input.dataset.entryCreateField] = input.value;
  });

  const hasValue = Object.values(values).some((value) => String(value || "").trim());
  if (!hasValue) {
    window.alert("Please enter at least one value before adding the entry.");
    return;
  }

  section.entries.push({ id: crypto.randomUUID(), values });
  lastEntrySectionId = section.id;
  pushAudit("Added entry", `New entry created in ${section.title}.`);
  persist();
  render();
  entryDialog.close();
  entryFollowupDialog.showModal();
}

function reopenEntryDialogAfterAdd(mode) {
  entryFollowupDialog.close();
  if (mode === "same") {
    openEntryDialog(lastEntrySectionId);
    return;
  }
  openEntryDialog();
}

function openImportDialog() {
  const person = currentPerson();
  if (!person?.sections.length) {
    window.alert("Create a section first so the importer has a target.");
    return;
  }
  importSectionSelect.innerHTML = person.sections.map((section) => `<option value="${section.id}">${escapeHtml(section.title)}</option>`).join("");
  importTextarea.value = "";
  importDelimiterSelect.value = "auto";
  importDialog.showModal();
}

function importRowsIntoSection() {
  const section = currentPerson()?.sections.find((item) => item.id === importSectionSelect.value);
  if (!section) return;
  const raw = importTextarea.value.trim();
  if (!raw) return window.alert("Paste a table first.");

  const delimiter = detectDelimiter(raw, importDelimiterSelect.value);
  const rows = raw.split(/\r?\n/).map((line) => splitDelimitedLine(line, delimiter));
  if (rows.length < 2) return window.alert("Please include a header row and at least one data row.");

  const headers = rows[0].map(normalizeKey);
  const importedEntries = rows.slice(1).filter((row) => row.some(Boolean)).map((row) => {
    const values = {};
    section.fields.forEach((field) => {
      const headerIndex = headers.findIndex((header) => header === normalizeKey(field.label) || header === normalizeKey(field.key));
      values[field.key] = headerIndex >= 0 ? (row[headerIndex] || "").trim() : "";
    });
    return { id: crypto.randomUUID(), values };
  });

  section.entries.push(...importedEntries);
  pushAudit("Imported rows", `${importedEntries.length} row(s) imported into ${section.title}.`);
  persist();
  render();
  importDialog.close();
}

function renderPreview() {
  const person = currentPerson();
  if (!person) {
    cvPreview.innerHTML = "";
    return;
  }
  if (!cvPreview) return;
  cvPreview.classList.toggle("show-guides", !!showGuidesToggle?.checked);
  cvPreview.innerHTML = `
    <h1 class="cv-document-title">Curriculum Vitae</h1>
    <header class="cv-header">
      <h1>${escapeHtml(person.profile.name || "Your Name")}</h1>
      <p class="cv-headline">${escapeHtml(person.profile.headline || "Professional Headline")}</p>
      ${person.profile.summary ? `<div class="cv-summary">${escapeHtml(person.profile.summary)}</div>` : ""}
    </header>
    ${person.sections.filter((section) => section.includeInCv).map(renderPreviewSection).join("")}
  `;
}

function renderPreviewSection(section) {
  const visibleEntries = section.entries.filter((entry) => Object.values(entry.values || {}).some((value) => String(value).trim()));
  if (!visibleEntries.length) return "";
  return `
    <section class="cv-section">
      <div class="cv-section-title">
        <h2>${escapeHtml(section.title)}</h2>
        <span>${escapeHtml(section.subtitle || "")}</span>
      </div>
      ${visibleEntries.map((entry) => renderPreviewEntry(section, entry)).join("")}
    </section>
  `;
}

function renderPreviewEntry(section, entry) {
  const mode = section.renderMode || "grid";
  if (mode === "pair-list") return renderPairListEntry(section, entry);
  if (mode === "date-list") return renderDateListEntry(section, entry);
  if (mode === "publication-list") return renderPublicationEntry(section, entry);

  return `
    <article class="cv-entry">
      <div class="cv-entry-grid">
        ${section.fields.map((field) => renderPreviewCell(field, entry.values?.[field.key] || "")).join("")}
      </div>
    </article>
  `;
}

function renderPairListEntry(section, entry) {
  const [labelField, valueField, ...rest] = section.fields;
  const label = labelField ? formatFieldOutput(labelField, entry.values?.[labelField.key] || "") : "";
  const value = valueField ? formatFieldOutput(valueField, entry.values?.[valueField.key] || "") : "";
  const extras = rest.map((field) => renderNamedInline(field, entry.values?.[field.key] || "")).filter(Boolean).join("");
  if (!label && !value && !extras) return "";
  return `
    <article class="cv-entry cv-entry-pair">
      <div class="cv-entry-label">${label || "&nbsp;"}</div>
      <div class="cv-entry-body">${value}${extras ? `<div class="cv-entry-note">${extras}</div>` : ""}</div>
    </article>
  `;
}

function renderDateListEntry(section, entry) {
  const [dateField, mainField, ...rest] = section.fields;
  const dateValue = dateField ? formatFieldOutput(dateField, entry.values?.[dateField.key] || "") : "";
  const mainValue = mainField ? formatFieldOutput(mainField, entry.values?.[mainField.key] || "") : "";
  const extras = rest.map((field) => renderNamedInline(field, entry.values?.[field.key] || "")).filter(Boolean).join("");
  if (!dateValue && !mainValue && !extras) return "";
  return `
    <article class="cv-entry cv-entry-date">
      <div class="cv-entry-label">${dateValue || "&nbsp;"}</div>
      <div class="cv-entry-body">${mainValue}${extras ? `<div class="cv-entry-note">${extras}</div>` : ""}</div>
    </article>
  `;
}

function renderPublicationEntry(section, entry) {
  if (section.fields.length < 3) {
    return renderPreviewEntry({ ...section, renderMode: "grid" }, entry);
  }
  const first = section.fields[0];
  const last = section.fields[section.fields.length - 1];
  const middle = section.fields.slice(1, -1);
  const leftValue = formatFieldOutput(first, entry.values?.[first.key] || "");
  const rightValue = formatFieldOutput(last, entry.values?.[last.key] || "");
  const body = middle.map((field, index) => {
    const rendered = formatFieldOutput(field, entry.values?.[field.key] || "");
    if (!rendered) return "";
    return index === 0 ? rendered : `<div class="cv-entry-note">${renderNamedInline(field, entry.values?.[field.key] || "")}</div>`;
  }).filter(Boolean).join("");
  if (!leftValue && !body && !rightValue) return "";
  return `
    <article class="cv-entry cv-entry-publication">
      <div class="cv-entry-index">${leftValue || "&nbsp;"}</div>
      <div class="cv-entry-body">${body}</div>
      <div class="cv-entry-year">${rightValue || "&nbsp;"}</div>
    </article>
  `;
}

function renderPreviewCell(field, rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) return "";
  const renderedValue = formatFieldOutput(field, value);
  return `<div class="cv-cell ${field.width}"><span class="cv-label">${escapeHtml(field.label)}</span><div class="cv-value">${renderedValue}</div></div>`;
}

function exportState() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `cv-data-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

function restoreState(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  file.text().then((text) => {
    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed.people) || !parsed.people.length) throw new Error("Invalid backup file.");
    parsed.people.forEach(normalizePerson);
    if (!parsed.activePersonId || !parsed.people.some((person) => person.id === parsed.activePersonId)) {
      parsed.activePersonId = parsed.people[0].id;
    }
    state = parsed;
    pushAudit("Restored backup", "Loaded CV registry from a JSON backup.");
    persist();
    render();
  }).catch(() => window.alert("Could not restore JSON backup."))
    .finally(() => { restoreJsonInput.value = ""; });
}

function summarizeEntry(section, entry) {
  const primaryField = section.fields[0];
  return primaryField ? (entry.values?.[primaryField.key]?.trim() || "Untitled entry") : "Entry";
}

function detectDelimiter(raw, selectedValue) {
  if (selectedValue === "tab") return "\t";
  if (selectedValue !== "auto") return selectedValue;
  const firstLine = raw.split(/\r?\n/, 1)[0];
  return ["\t", ";", ","].sort((a, b) => firstLine.split(b).length - firstLine.split(a).length)[0];
}

function splitDelimitedLine(line, delimiter) {
  if (delimiter === "\t") return line.split("\t");
  const result = [];
  let current = "";
  let insideQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const nextChar = line[i + 1];
    if (char === '"') {
      if (insideQuotes && nextChar === '"') {
        current += '"';
        i += 1;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === delimiter && !insideQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}

function inputTypeFor(type) {
  if (type === "date") return "date";
  if (type === "month") return "month";
  if (type === "year") return "number";
  if (type === "email") return "email";
  if (type === "url") return "url";
  return "text";
}

function formatValue(type, value) {
  if (type === "month") {
    const [year, month] = value.split("-");
    if (year && month) {
      return new Date(Number(year), Number(month) - 1, 1).toLocaleDateString(undefined, { year: "numeric", month: "long" });
    }
  }
  if (type === "date") {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return date.toLocaleDateString();
  }
  return value;
}

function formatFieldOutput(field, value) {
  const cleanValue = String(value || "").trim();
  if (!cleanValue) return "";
  if (field.type === "bullets") {
    return `<ul>${cleanValue.split(/\r?\n/).filter(Boolean).map((line) => `<li>${escapeHtml(line.trim())}</li>`).join("")}</ul>`;
  }
  if (field.type === "url") {
    return `<a href="${escapeAttribute(cleanValue)}" target="_blank" rel="noreferrer">${escapeHtml(cleanValue)}</a>`;
  }
  if (field.type === "email") {
    return `<a href="mailto:${escapeAttribute(cleanValue)}">${escapeHtml(cleanValue)}</a>`;
  }
  return escapeHtml(formatValue(field.type, cleanValue));
}

function renderNamedInline(field, value) {
  const rendered = formatFieldOutput(field, value);
  if (!rendered) return "";
  return `<span><strong>${escapeHtml(field.label)}:</strong> ${rendered}</span>`;
}

function renderModeLabel(mode) {
  if (mode === "pair-list") return "Label/value rows";
  if (mode === "date-list") return "Date/detail rows";
  if (mode === "publication-list") return "Publication rows";
  return "Flexible grid";
}

function activateTab(tabId) {
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.tabTarget === tabId);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === tabId);
  });
}

function renderActivePersonCard() {
  const person = currentPerson();
  if (!person) {
    activePersonCard.innerHTML = '<div class="empty-state">No person selected.</div>';
    return;
  }
  activePersonCard.innerHTML = `
    <div class="registry-row">
      <div>
        <strong>${escapeHtml(person.profile.name || "Unnamed person")}</strong>
        <div class="registry-meta">${escapeHtml(person.profile.headline || "No headline")}</div>
      </div>
      <span>${person.sections.length} sections</span>
    </div>
  `;
}

function openPersonPickerDialog() {
  hasPromptedPersonSelection = true;
  personPickerList.innerHTML = state.people.map((person) => `
    <div class="registry-row">
      <div>
        <strong>${escapeHtml(person.profile.name || "Unnamed person")}</strong>
        <div class="registry-meta">${escapeHtml(person.profile.headline || "No headline")}</div>
      </div>
      <button type="button" data-select-person="${person.id}">Edit this CV</button>
    </div>
  `).join("");

  personPickerList.querySelectorAll("[data-select-person]").forEach((button) => {
    button.addEventListener("click", () => {
      state.activePersonId = button.dataset.selectPerson;
      persist();
      render();
      personPickerDialog.close();
    });
  });

  personPickerDialog.showModal();
}

function openPersonDialog() {
  personNameInput.value = "";
  personHeadlineInput.value = "";
  personSummaryInput.value = "";
  personDialog.showModal();
}

function saveNewPerson() {
  const name = personNameInput.value.trim();
  if (!name) {
    window.alert("Please enter a name for the person.");
    return;
  }
  const person = {
    id: crypto.randomUUID(),
    folderName: null,
    profile: {
      name,
      headline: personHeadlineInput.value.trim(),
      summary: personSummaryInput.value.trim()
    },
    auditLog: [{
      id: crypto.randomUUID(),
      action: "Created person",
      detail: `Initialized CV registry for ${name}.`,
      timestamp: new Date().toISOString()
    }],
    sections: []
  };
  state.people.push(person);
  state.activePersonId = person.id;
  hasPromptedPersonSelection = true;
  persist();
  render();
  personDialog.close();
}

function ensurePersonSelected() {
  if (state.people.length > 1) {
    openPersonPickerDialog();
  } else {
    hasPromptedPersonSelection = true;
  }
}

function pushAudit(action, detail) {
  const person = currentPerson();
  if (!person) return;
  if (!Array.isArray(person.auditLog)) person.auditLog = [];
  person.auditLog.unshift({
    id: crypto.randomUUID(),
    action,
    detail,
    timestamp: new Date().toISOString()
  });
  person.auditLog = person.auditLog.slice(0, 100);
}

function formatTimestamp(timestamp) {
  const date = new Date(timestamp);
  if (Number.isNaN(date.getTime())) return timestamp;
  return date.toLocaleString();
}

function statCard(label, value, helper) {
  return `<article class="stat-card"><span class="eyebrow">${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong><div class="registry-meta">${escapeHtml(helper)}</div></article>`;
}

function addPresetSection(kind) {
  const preset = createPresetSection(kind);
  currentPerson().sections.push(preset);
  pushAudit("Added preset section", preset.title);
  persist();
  render();
  activateTab("sectionsTab");
}

function createPresetSection(kind) {
  if (kind === "personal") {
    return {
      id: crypto.randomUUID(),
      title: "Personal information",
      subtitle: "",
      renderMode: "pair-list",
      includeInCv: true,
      fields: [
        { id: crypto.randomUUID(), label: "Label", key: "label", type: "text", width: "third" },
        { id: crypto.randomUUID(), label: "Value", key: "value", type: "text", width: "full" }
      ],
      entries: [{ id: crypto.randomUUID(), values: { label: "", value: "" } }]
    };
  }
  if (kind === "publication") {
    return {
      id: crypto.randomUUID(),
      title: "Selected publications",
      subtitle: "",
      renderMode: "publication-list",
      includeInCv: true,
      fields: [
        { id: crypto.randomUUID(), label: "No.", key: "number", type: "text", width: "quarter" },
        { id: crypto.randomUUID(), label: "Citation", key: "citation", type: "textarea", width: "full" },
        { id: crypto.randomUUID(), label: "Year", key: "year", type: "year", width: "quarter" }
      ],
      entries: [{ id: crypto.randomUUID(), values: { number: "", citation: "", year: "" } }]
    };
  }
  return {
    id: crypto.randomUUID(),
    title: "Timeline section",
    subtitle: "",
    renderMode: "date-list",
    includeInCv: true,
    fields: [
      { id: crypto.randomUUID(), label: "Period", key: "period", type: "text", width: "third" },
      { id: crypto.randomUUID(), label: "Main entry", key: "main", type: "text", width: "full" },
      { id: crypto.randomUUID(), label: "Details", key: "details", type: "textarea", width: "full" }
    ],
    entries: [{ id: crypto.randomUUID(), values: { period: "", main: "", details: "" } }]
  };
}

function renderStorageStatus() {
  const folderName = folderHandle?.name || "No folder connected";
  const detail = storageInfo.lastSaved ? `Last saved ${storageInfo.lastSaved}` : storageInfo.detail;
  storageStatus.innerHTML = `
    <div class="registry-row">
      <div>
        <strong>${escapeHtml(folderName)}</strong>
        <div class="registry-meta">${escapeHtml(detail)}</div>
      </div>
      <span>${escapeHtml(storageInfo.mode)}</span>
    </div>
  `;
  connectFolderBtn.textContent = folderHandle ? "Reconnect CV_builder folder" : "Connect CV_builder folder";
}

function queueFilesystemSave() {
  clearTimeout(saveDebounceId);
  if (!folderHandle) {
    renderStorageStatus();
    return;
  }
  saveDebounceId = window.setTimeout(() => {
    saveStateToFolder().catch((error) => {
      console.error("Filesystem save failed", error);
      storageInfo.mode = "needs-reconnect";
      storageInfo.detail = "Reconnect the CV_builder folder to resume autosave.";
      renderStorageStatus();
    });
  }, 400);
}

async function connectFolder() {
  if (!("showDirectoryPicker" in window)) {
    window.alert("This browser does not support direct folder autosave. Please use a Chromium-based browser.");
    return;
  }
  try {
    const handle = await window.showDirectoryPicker({ id: "cv-builder-folder", mode: "readwrite" });
    const granted = await ensureHandlePermission(handle, "readwrite", true);
    if (!granted) {
      storageInfo.mode = "permission-needed";
      storageInfo.detail = "Folder access was not granted.";
      renderStorageStatus();
      return;
    }
    folderHandle = handle;
    await saveRootHandle(handle);
    const loadedState = await loadStateFromFolder(handle);
    if (loadedState) {
      state = loadedState;
    } else {
      await saveStateToFolder();
    }
    storageInfo.mode = "folder-connected";
    storageInfo.detail = "Autosave is active for this folder.";
    render();
  } catch (error) {
    if (error?.name !== "AbortError") {
      console.error(error);
      window.alert("Could not connect the folder.");
    }
  }
}

async function hydrateFromConnectedFolder() {
  if (!("showDirectoryPicker" in window)) {
    storageInfo.mode = "browser-only";
    storageInfo.detail = "This browser stores data locally in-browser only.";
    return;
  }
  folderHandle = await loadRootHandle();
  if (!folderHandle) {
    storageInfo.mode = "browser-only";
    storageInfo.detail = "Connect the CV_builder folder to enable autosave JSON files.";
    return;
  }
  const granted = await ensureHandlePermission(folderHandle, "readwrite", true);
  if (!granted) {
    storageInfo.mode = "needs-reconnect";
    storageInfo.detail = "Browser permission for the saved CV_builder folder was not granted on startup.";
    return;
  }
  const loadedState = await loadStateFromFolder(folderHandle);
  if (loadedState) {
    state = loadedState;
  }
  storageInfo.mode = "folder-connected";
  storageInfo.detail = "Autosave is active for this folder.";
}

async function saveStateToFolder() {
  if (!folderHandle) return false;
  const granted = await ensureHandlePermission(folderHandle, "readwrite", true);
  if (!granted) {
    storageInfo.mode = "needs-reconnect";
    storageInfo.detail = "Reconnect the CV_builder folder to resume autosave.";
    renderStorageStatus();
    return false;
  }

  const peopleDir = await folderHandle.getDirectoryHandle("people", { create: true });
  const registry = { activePersonId: state.activePersonId, people: [] };

  for (const person of state.people) {
    normalizePerson(person);
    if (!person.folderName) person.folderName = buildPersonFolderName(person);
    const personDir = await peopleDir.getDirectoryHandle(person.folderName, { create: true });
    await writeJsonFile(personDir, "cv.json", person);
    registry.people.push({ id: person.id, folderName: person.folderName, name: person.profile.name || "" });
  }

  await writeJsonFile(folderHandle, "cv_registry.json", registry);
  storageInfo.mode = "folder-connected";
  storageInfo.detail = "Autosave is active for this folder.";
  storageInfo.lastSaved = formatTimestamp(new Date().toISOString());
  renderStorageStatus();
  return true;
}

async function loadStateFromFolder(handle) {
  try {
    const registry = await readJsonFile(handle, "cv_registry.json");
    if (!registry || !Array.isArray(registry.people) || !registry.people.length) return null;
    const peopleDir = await handle.getDirectoryHandle("people");
    const people = [];

    for (const meta of registry.people) {
      try {
        const personDir = await peopleDir.getDirectoryHandle(meta.folderName);
        const person = await readJsonFile(personDir, "cv.json");
        if (!person) continue;
        normalizePerson(person);
        person.id = meta.id || person.id || crypto.randomUUID();
        person.folderName = meta.folderName || person.folderName || buildPersonFolderName(person);
        people.push(person);
      } catch (error) {
        console.warn("Skipping unreadable person record", error);
      }
    }

    if (!people.length) return null;
    return {
      activePersonId: people.some((person) => person.id === registry.activePersonId) ? registry.activePersonId : people[0].id,
      people
    };
  } catch (error) {
    console.warn("Could not load state from folder", error);
    return null;
  }
}

async function writeJsonFile(directoryHandle, fileName, value) {
  const fileHandle = await directoryHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(JSON.stringify(value, null, 2));
  await writable.close();
}

async function readJsonFile(directoryHandle, fileName) {
  const fileHandle = await directoryHandle.getFileHandle(fileName);
  const file = await fileHandle.getFile();
  return JSON.parse(await file.text());
}

function buildPersonFolderName(person) {
  const base = slugify(person.profile?.name || "person") || "person";
  const suffix = (person.id || crypto.randomUUID()).slice(0, 8);
  return `${base}-${suffix}`;
}

async function ensureHandlePermission(handle, mode, shouldRequest) {
  try {
    const options = { mode };
    const current = await handle.queryPermission(options);
    if (current === "granted") return true;
    if (!shouldRequest) return false;
    return (await handle.requestPermission(options)) === "granted";
  } catch {
    return false;
  }
}

async function openHandleDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(HANDLE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(HANDLE_STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveRootHandle(handle) {
  const db = await openHandleDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(HANDLE_STORE_NAME, "readwrite");
    tx.objectStore(HANDLE_STORE_NAME).put(handle, HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadRootHandle() {
  try {
    const db = await openHandleDatabase();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(HANDLE_STORE_NAME, "readonly");
      const request = tx.objectStore(HANDLE_STORE_NAME).get(HANDLE_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  } catch {
    return null;
  }
}

function maybePromptFolderConnection() {
  if (folderHandle) return;
  if (!("showDirectoryPicker" in window)) return;
  if (localStorage.getItem(STORAGE_PROMPT_KEY)) return;
  storagePromptDialog.showModal();
}

function normalizeKey(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function slugify(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function escapeHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;").replace(/'/g, "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/\"/g, "&quot;");
}
