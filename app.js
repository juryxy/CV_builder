const STORAGE_KEY = "offline-cv-builder-state-v5";
const HANDLE_DB_NAME = "cv-builder-fs";
const HANDLE_STORE_NAME = "handles";
const HANDLE_KEY = "root-folder";
const STORAGE_PROMPT_KEY = "cv-builder-storage-prompt-seen";

const SECTION_TEMPLATES = [
  {
    key: "personal",
    label: "Personal information",
    create: () => ({
      title: "Personal information",
      subtitle: "",
      renderMode: "pair-list",
      includeInCv: true,
      fields: [
        makeField("Label", "label", "text", "third"),
        makeField("Value", "value", "text", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { label: "", value: "" } }]
    })
  },
  {
    key: "experience",
    label: "Professional experience",
    create: () => ({
      title: "Professional experience",
      subtitle: "",
      renderMode: "date-list",
      includeInCv: true,
      fields: [
        makeField("Period", "period", "text", "third"),
        makeField("Role", "role", "text", "third"),
        makeField("Institution", "institution", "text", "third"),
        makeField("Details", "details", "textarea", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { period: "", role: "", institution: "", details: "" } }]
    })
  },
  {
    key: "education",
    label: "Education",
    create: () => ({
      title: "Education",
      subtitle: "",
      renderMode: "date-list",
      includeInCv: true,
      fields: [
        makeField("Period", "period", "text", "third"),
        makeField("Degree", "degree", "text", "third"),
        makeField("Institution", "institution", "text", "third"),
        makeField("Details", "details", "textarea", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { period: "", degree: "", institution: "", details: "" } }]
    })
  },
  {
    key: "skills",
    label: "Skills",
    create: () => ({
      title: "Skills",
      subtitle: "",
      renderMode: "pair-list",
      includeInCv: true,
      fields: [
        makeField("Category", "category", "text", "third"),
        makeField("Details", "details", "text", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { category: "", details: "" } }]
    })
  },
  {
    key: "awards",
    label: "Awards and honors",
    create: () => ({
      title: "Awards and honors",
      subtitle: "",
      renderMode: "date-list",
      includeInCv: true,
      fields: [
        makeField("Year", "year", "year", "quarter"),
        makeField("Award", "award", "text", "third"),
        makeField("Organization", "organization", "text", "third"),
        makeField("Details", "details", "textarea", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { year: "", award: "", organization: "", details: "" } }]
    })
  },
  {
    key: "publications",
    label: "Selected publications",
    create: () => ({
      title: "Selected publications",
      subtitle: "",
      renderMode: "publication-list",
      includeInCv: true,
      fields: [
        makeField("No.", "number", "text", "quarter"),
        makeField("Citation", "citation", "textarea", "full"),
        makeField("Year", "year", "year", "quarter")
      ],
      entries: [{ id: crypto.randomUUID(), values: { number: "", citation: "", year: "" } }]
    })
  },
  {
    key: "projects",
    label: "Projects",
    create: () => ({
      title: "Projects",
      subtitle: "",
      renderMode: "date-list",
      includeInCv: true,
      fields: [
        makeField("Period", "period", "text", "third"),
        makeField("Project", "project", "text", "third"),
        makeField("Role", "role", "text", "third"),
        makeField("Details", "details", "textarea", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { period: "", project: "", role: "", details: "" } }]
    })
  },
  {
    key: "languages",
    label: "Languages",
    create: () => ({
      title: "Languages",
      subtitle: "",
      renderMode: "pair-list",
      includeInCv: true,
      fields: [
        makeField("Language", "language", "text", "third"),
        makeField("Level", "level", "text", "third"),
        makeField("Notes", "notes", "text", "full")
      ],
      entries: [{ id: crypto.randomUUID(), values: { language: "", level: "", notes: "" } }]
    })
  }
];

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
    populateTemplate(SECTION_TEMPLATES.find((item) => item.key === "personal").create(), {
      entries: [
        { id: crypto.randomUUID(), values: { label: "Name", value: "Prof. Dr. Alex Example" } },
        { id: crypto.randomUUID(), values: { label: "Address", value: "Institute for Example Research, Sample Street 1, 12345 Sample City, Germany" } },
        { id: crypto.randomUUID(), values: { label: "E-mail", value: "alex@example.org" } },
        { id: crypto.randomUUID(), values: { label: "Nationality", value: "German" } }
      ]
    }),
    populateTemplate(SECTION_TEMPLATES.find((item) => item.key === "experience").create(), {
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
    }),
    populateTemplate(SECTION_TEMPLATES.find((item) => item.key === "publications").create(), {
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
    })
  ]
};

const demoState = {
  activePersonId: demoPerson.id,
  people: [demoPerson]
};

const emptyState = {
  activePersonId: null,
  people: []
};

let state = loadState();
let editingSectionId = null;
let hasPromptedPersonSelection = false;

const profileNameInput = document.getElementById("profileName");
const profileHeadlineInput = document.getElementById("profileHeadline");
const profileSummaryInput = document.getElementById("profileSummary");
const addEntryFlowBtn = document.getElementById("addEntryFlowBtn");
const addSectionBtn = document.getElementById("addSectionBtn");
const importBtn = document.getElementById("importBtn");
const addTemplateSectionBtn = document.getElementById("addTemplateSectionBtn");
const sectionTemplateSelect = document.getElementById("sectionTemplateSelect");
const activePersonCard = document.getElementById("activePersonCard");
const switchPersonBtn = document.getElementById("switchPersonBtn");
const addPersonBtn = document.getElementById("addPersonBtn");
const connectFolderBtn = document.getElementById("connectFolderBtn");
const storageStatus = document.getElementById("storageStatus");
const storagePanel = document.getElementById("storagePanel");
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
let importSectionHint = null;
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

let draggedSectionId = null;
let draggedEntryId = null;
let draggedEntrySectionId = null;

bindEvents();
initApp();

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
      renderActivePersonCard();
    });
  });

  addEntryFlowBtn.addEventListener("click", () => openEntryDialog());
  addSectionBtn.addEventListener("click", () => openSectionDialog());
  addTemplateSectionBtn.addEventListener("click", addSelectedTemplateSection);
  switchPersonBtn.addEventListener("click", openPersonPickerDialog);
  addPersonBtn.addEventListener("click", openPersonDialog);
  connectFolderBtn.addEventListener("click", connectFolder);
  importBtn.addEventListener("click", openImportDialog);
  exportJsonBtn.addEventListener("click", exportState);
  restoreJsonInput.addEventListener("change", restoreState);
  printBtn.addEventListener("click", () => {
    window.alert('For a clean PDF, disable "Headers and footers" in the browser print dialog.');
    window.print();
  });
  resetDemoBtn.addEventListener("click", loadDemoPerson);
  showGuidesToggle.addEventListener("change", renderPreview);
  addFieldBtn.addEventListener("click", () => addFieldRow());

  fieldRows.addEventListener("click", (event) => {
    const button = event.target.closest("[data-action='remove-field']");
    if (button) {
      button.closest(".field-row")?.remove();
    }
  });

  sectionForm.addEventListener("submit", (event) => {
    event.preventDefault();
    saveSectionFromDialog();
  });

  runImportBtn.addEventListener("click", importRowsIntoSection);
  importSectionSelect.addEventListener("change", renderImportDialogHint);
  entrySectionSelect.addEventListener("change", renderEntryDialogFields);
  submitEntryBtn.addEventListener("click", submitNewEntryFromDialog);
  followupSameSectionBtn.addEventListener("click", () => reopenEntryDialogAfterAdd("same"));
  followupDifferentSectionBtn.addEventListener("click", () => reopenEntryDialogAfterAdd("different"));
  followupBackOverviewBtn.addEventListener("click", () => {
    entryFollowupDialog.close();
    activateTab("overviewTab");
  });
  personPickerAddBtn.addEventListener("click", () => {
    personPickerDialog.close();
    openPersonDialog();
  });
  savePersonBtn.addEventListener("click", saveNewPerson);
  storagePromptYesBtn.addEventListener("click", async () => {
    localStorage.setItem(STORAGE_PROMPT_KEY, "seen");
    storagePromptDialog.close();
    await connectFolder();
  });
  storagePromptNoBtn.addEventListener("click", () => {
    localStorage.setItem(STORAGE_PROMPT_KEY, "seen");
    storagePromptDialog.close();
  });
  sectionsContainer.addEventListener("click", handleSectionClicks);
  sectionsContainer.addEventListener("input", handleSectionInputs);
  sectionsContainer.addEventListener("dragstart", handleDragStart);
  sectionsContainer.addEventListener("dragover", handleDragOver);
  sectionsContainer.addEventListener("drop", handleDrop);
  sectionsContainer.addEventListener("dragend", handleDragEnd);

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
  populateTemplateSelector();
  await hydrateFromConnectedFolder();
  render();
  maybePromptFolderConnection();
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return structuredClone(emptyState);
    const parsed = JSON.parse(raw);
    return sanitizeState(parsed);
  } catch (error) {
    console.warn("Could not load local state", error);
    return structuredClone(emptyState);
  }
}

function sanitizeState(nextState) {
  const candidate = nextState && Array.isArray(nextState.people)
    ? { activePersonId: nextState.activePersonId || null, people: nextState.people }
    : structuredClone(emptyState);

  candidate.people.forEach(normalizePerson);
  if (!candidate.people.length) {
    candidate.activePersonId = null;
    return candidate;
  }
  if (!candidate.activePersonId || !candidate.people.some((person) => person.id === candidate.activePersonId)) {
    candidate.activePersonId = candidate.people[0].id;
  }
  return candidate;
}

function normalizePerson(person) {
  if (!person.id) person.id = crypto.randomUUID();
  if (!person.profile) person.profile = { name: "", headline: "", summary: "" };
  if (!Array.isArray(person.sections)) person.sections = [];
  if (!Array.isArray(person.auditLog)) person.auditLog = [];
  if (!Object.prototype.hasOwnProperty.call(person, "folderName")) person.folderName = null;
  person.sections = person.sections.map(normalizeSection);
}

function normalizeSection(section) {
  const nextSection = {
    id: section.id || crypto.randomUUID(),
    title: section.title || "Untitled section",
    subtitle: section.subtitle || "",
    renderMode: section.renderMode || "grid",
    includeInCv: section.includeInCv !== false,
    fields: Array.isArray(section.fields) ? section.fields.map((field) => ({
      id: field.id || crypto.randomUUID(),
      label: field.label || "Field",
      key: slugify(field.key || field.label || crypto.randomUUID()),
      type: field.type || "text",
      width: field.width || "full"
    })) : [],
    entries: Array.isArray(section.entries) ? section.entries.map((entry) => ({
      id: entry.id || crypto.randomUUID(),
      values: { ...(entry.values || {}) }
    })) : []
  };
  return nextSection;
}

function persist() {
  persistToLocalStorage();
  queueFilesystemSave();
}

function persistToLocalStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (error) {
    console.warn("Could not persist local state", error);
  }
}

function currentPerson() {
  return state.people.find((person) => person.id === state.activePersonId) || state.people[0] || null;
}

function render() {
  const person = currentPerson();
  profileNameInput.value = person?.profile.name || "";
  profileHeadlineInput.value = person?.profile.headline || "";
  profileSummaryInput.value = person?.profile.summary || "";
  profileNameInput.disabled = !person;
  profileHeadlineInput.disabled = !person;
  profileSummaryInput.disabled = !person;
  renderSections();
  renderDashboard();
  renderPreview();
  renderActivePersonCard();
  renderStorageStatus();
  populateTemplateSelector();
  if (!hasPromptedPersonSelection && !personPickerDialog.open) {
    ensurePersonSelected();
  }
}

function renderSections() {
  const person = currentPerson();
  if (!person) {
    sectionsContainer.innerHTML = '<div class="empty-state">Add a person to start building a CV.</div>';
    return;
  }
  if (!person.sections.length) {
    sectionsContainer.innerHTML = '<div class="empty-state">No sections yet. Create your first section or add a template section.</div>';
    return;
  }

  sectionsContainer.innerHTML = person.sections.map(renderSectionCard).join("");
}

function renderSectionCard(section) {
  const entriesHtml = section.entries.length
    ? section.entries.map((entry) => renderEntryCard(section, entry)).join("")
    : '<div class="empty-state">No entries yet. Add one below or import rows from Excel.</div>';

  return `
    <article class="section-card" data-section-id="${section.id}" draggable="true">
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
        <div class="button-stack action-stack-inline">
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
  if (!person) {
    statsGrid.innerHTML = [
      statCard("People", state.people.length, "Create a person to start"),
      statCard("Sections", 0, "No active CV selected"),
      statCard("Entries", 0, "No records yet"),
      statCard("Storage", storageInfo.mode, storageInfo.detail)
    ].join("");
    sectionRegistry.innerHTML = '<div class="empty-state">No sections available yet.</div>';
    activityList.innerHTML = '<div class="empty-state">No recent activity yet.</div>';
    activityLogPanel.innerHTML = '<div class="empty-state">No recent activity yet.</div>';
    return;
  }

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
    <section class="entry-card" data-entry-id="${entry.id}" draggable="true">
      <div class="entry-header">
        <strong>${escapeHtml(summarizeEntry(section, entry))}</strong>
        <button data-action="delete-entry" class="danger small">Delete entry</button>
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

function handleDragStart(event) {
  const entryCard = event.target.closest('.entry-card');
  if (entryCard) {
    const sectionCard = entryCard.closest('[data-section-id]');
    draggedEntryId = entryCard.dataset.entryId;
    draggedEntrySectionId = sectionCard?.dataset.sectionId || null;
    draggedSectionId = null;
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', draggedEntryId);
    return;
  }

  const sectionCard = event.target.closest('.section-card');
  if (sectionCard) {
    draggedSectionId = sectionCard.dataset.sectionId;
    draggedEntryId = null;
    draggedEntrySectionId = null;
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', draggedSectionId);
  }
}

function handleDragOver(event) {
  if (event.target.closest('.entry-card') || event.target.closest('.section-card')) {
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
  }
}

function handleDrop(event) {
  event.preventDefault();
  const person = currentPerson();
  if (!person) return;

  const targetEntry = event.target.closest('.entry-card');
  if (draggedEntryId && draggedEntrySectionId && targetEntry) {
    const targetSectionId = targetEntry.closest('[data-section-id]')?.dataset.sectionId;
    if (!targetSectionId || targetSectionId !== draggedEntrySectionId) {
      resetDragState();
      return;
    }
    const section = person.sections.find((item) => item.id === targetSectionId);
    if (!section) {
      resetDragState();
      return;
    }
    const fromIndex = section.entries.findIndex((item) => item.id === draggedEntryId);
    const toIndex = section.entries.findIndex((item) => item.id === targetEntry.dataset.entryId);
    if (fromIndex >= 0 && toIndex >= 0 && fromIndex !== toIndex) {
      const [moved] = section.entries.splice(fromIndex, 1);
      section.entries.splice(toIndex, 0, moved);
      pushAudit('Reordered entries', `Updated entry order in ${section.title}.`);
      persist();
      render();
    }
    resetDragState();
    return;
  }

  const targetSection = event.target.closest('.section-card');
  if (draggedSectionId && targetSection) {
    const fromIndex = person.sections.findIndex((item) => item.id === draggedSectionId);
    const toIndex = person.sections.findIndex((item) => item.id === targetSection.dataset.sectionId);
    if (fromIndex >= 0 && toIndex >= 0 && fromIndex !== toIndex) {
      const [moved] = person.sections.splice(fromIndex, 1);
      person.sections.splice(toIndex, 0, moved);
      pushAudit('Reordered sections', 'Updated section order.');
      persist();
      render();
    }
  }
  resetDragState();
}

function handleDragEnd() {
  resetDragState();
}

function resetDragState() {
  draggedSectionId = null;
  draggedEntryId = null;
  draggedEntrySectionId = null;
}

function openSectionDialog(sectionId = null) {
  if (!currentPerson()) {
    window.alert("Add a person first before creating sections.");
    return;
  }
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

function populateTemplateSelector() {
  if (!sectionTemplateSelect) return;
  sectionTemplateSelect.innerHTML = SECTION_TEMPLATES.map((template) => `<option value="${template.key}">${escapeHtml(template.label)}</option>`).join("");
}

function addSelectedTemplateSection() {
  const person = currentPerson();
  if (!person) {
    window.alert("Add a person first before adding template sections.");
    return;
  }
  const template = SECTION_TEMPLATES.find((item) => item.key === sectionTemplateSelect.value) || SECTION_TEMPLATES[0];
  const preset = populateTemplate(template.create());
  person.sections.push(preset);
  pushAudit("Added template section", preset.title);
  persist();
  render();
  activateTab("sectionsTab");
}

function openEntryDialog(preselectedSectionId = null) {
  const person = currentPerson();
  if (!person) {
    window.alert("Add a person first so the new entry has somewhere to go.");
    return;
  }
  if (!person.sections.length) {
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
  if (!person) {
    window.alert("Add a person first so the importer has a target.");
    return;
  }
  if (!person.sections.length) {
    window.alert("Create a section first so the importer has a target.");
    return;
  }
  importSectionSelect.innerHTML = person.sections.map((section) => `<option value="${section.id}">${escapeHtml(section.title)}</option>`).join("");
  importTextarea.value = "";
  importDelimiterSelect.value = "auto";
  ensureImportDialogHint();
  renderImportDialogHint();
  importDialog.showModal();
}

function ensureImportDialogHint() {
  if (importSectionHint?.isConnected) return importSectionHint;
  const helper = importDialog.querySelector('.helper-text');
  importSectionHint = document.createElement('div');
  importSectionHint.id = 'importSectionHint';
  importSectionHint.className = 'import-section-hint';
  if (helper?.parentNode) {
    helper.parentNode.insertBefore(importSectionHint, helper.nextSibling);
  } else {
    const actions = importDialog.querySelector('.dialog-actions');
    if (actions?.parentNode) {
      actions.parentNode.insertBefore(importSectionHint, actions);
    } else {
      importDialog.append(importSectionHint);
    }
  }
  return importSectionHint;
}

function renderImportDialogHint() {
  const hint = ensureImportDialogHint();
  const section = currentPerson()?.sections.find((item) => item.id === importSectionSelect.value);
  if (!section) {
    hint.innerHTML = '<div class="helper-text">Select a section to see the expected import format.</div>';
    return;
  }

  const exampleHeaders = section.fields.map((field) => field.label).join(';');
  const exampleValues = section.fields.map((field, index) => exampleValueForField(field, index)).join(';');

  hint.innerHTML = `
    <div class="import-hint-card">
      <strong>Import format for ${escapeHtml(section.title)}</strong>
      <div class="helper-text">Use the first row as column headers. Best match: the same order and names as this section.</div>
      <div class="import-hint-fields">${section.fields.map((field) => `<span>${escapeHtml(field.label)}</span>`).join('')}</div>
      <pre class="import-hint-example">${escapeHtml(exampleHeaders)}
${escapeHtml(exampleValues)}</pre>
    </div>
  `;
}

function exampleValueForField(field, index) {
  const key = normalizeKey(field.key || field.label);
  if (key.includes('period') || key.includes('year') || key.includes('date')) return index === 0 ? '2023-present' : '2024';
  if (key.includes('role') || key.includes('title') || key.includes('position')) return 'Professor for Digital Biomarkers';
  if (key.includes('institution') || key.includes('organization') || key.includes('company')) return 'Heinrich Heine University Duesseldorf';
  if (key.includes('degree')) return 'PhD in Neuroscience';
  if (key.includes('project')) return 'Digital Biomarker Platform';
  if (key.includes('award')) return 'Research Award';
  if (key.includes('citation')) return 'Author A, Author B. Example publication title. Journal, 2024.';
  if (key.includes('number')) return '1';
  if (key.includes('language')) return 'English';
  if (key.includes('level')) return 'Fluent';
  if (key.includes('category')) return 'Programming';
  if (key.includes('label')) return 'Email';
  if (key.includes('value')) return 'name@example.org';
  if (key.includes('details') || key.includes('notes')) return 'Short description or additional details';
  return `Example ${index + 1}`;
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
    cvPreview.innerHTML = `
      <h1 class="cv-document-title">Curriculum Vitae</h1>
      <div class="empty-state">Add a person and some sections to render the CV preview.</div>
    `;
    return;
  }
  cvPreview.classList.toggle("show-guides", showGuidesToggle.checked);
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
  const extras = rest.map((field) => renderNamedInline(field, entry.values?.[field.key] || "")).filter(Boolean).join(" · ");
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
  const extras = rest.map((field) => renderNamedInline(field, entry.values?.[field.key] || "")).filter(Boolean).join(" · ");
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
  }).filter(Boolean).join(" · ");
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
    state = buildStateFromBackup(parsed);
    if (!state.people.length) throw new Error("Invalid backup file.");
    pushAudit("Restored backup", "Loaded CV registry from a JSON backup.");
    persist();
    render();
  }).catch(() => window.alert("Could not restore JSON backup."))
    .finally(() => { restoreJsonInput.value = ""; });
}

function buildStateFromBackup(parsed) {
  if (parsed && Array.isArray(parsed.people)) {
    return sanitizeState(parsed);
  }

  if (parsed && parsed.profile && Array.isArray(parsed.sections)) {
    const singlePerson = { ...parsed };
    normalizePerson(singlePerson);
    return sanitizeState({
      activePersonId: singlePerson.id,
      people: [singlePerson]
    });
  }

  throw new Error("Invalid backup file.");
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
  if (!state.people.length) {
    personPickerList.innerHTML = '<div class="empty-state">No people available yet.</div>';
    personPickerDialog.showModal();
    return;
  }

  personPickerList.innerHTML = state.people.map((person) => `
    <div class="registry-row registry-row-actions">
      <div>
        <strong>${escapeHtml(person.profile.name || "Unnamed person")}</strong>
        <div class="registry-meta">${escapeHtml(person.profile.headline || "No headline")}</div>
      </div>
      <div class="button-stack action-stack-inline">
        <button type="button" data-select-person="${person.id}">Edit this CV</button>
        <button type="button" class="danger" data-delete-person="${person.id}">Delete</button>
      </div>
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

  personPickerList.querySelectorAll("[data-delete-person]").forEach((button) => {
    button.addEventListener("click", () => {
      deletePerson(button.dataset.deletePerson);
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

async function deletePerson(personId) {
  const person = state.people.find((item) => item.id === personId);
  if (!person) return;

  const label = person.profile?.name || "this person";
  if (!window.confirm(`Delete ${label} and all associated CV data?`)) return;

  state.people = state.people.filter((item) => item.id !== personId);

  if (state.activePersonId === personId) {
    state.activePersonId = state.people[0]?.id || null;
  }

  if (!state.people.length) {
    hasPromptedPersonSelection = true;
  }

  if (folderHandle && person.folderName) {
    try {
      const peopleDir = await folderHandle.getDirectoryHandle("people");
      await peopleDir.removeEntry(person.folderName, { recursive: true });
    } catch (error) {
      console.warn("Could not remove person folder", error);
    }
  }

  persist();
  render();

  if (state.people.length) {
    openPersonPickerDialog();
  } else {
    personPickerDialog.close();
  }
}

function ensurePersonSelected() {
  if (state.people.length > 1) {
    openPersonPickerDialog();
  } else {
    hasPromptedPersonSelection = true;
  }
}

function loadDemoPerson() {
  const existingDemo = state.people.find(isDemoPerson);
  if (existingDemo) {
    state.activePersonId = existingDemo.id;
    persist();
    render();
    return;
  }

  const demo = structuredClone(demoPerson);
  demo.id = crypto.randomUUID();
  demo.folderName = null;
  demo.sections = demo.sections.map(normalizeSection);
  demo.auditLog = [{
    id: crypto.randomUUID(),
    action: "Loaded demo person",
    detail: "Added the demo CV without removing existing people.",
    timestamp: new Date().toISOString()
  }];

  state.people.push(demo);
  state.activePersonId = demo.id;
  hasPromptedPersonSelection = true;
  persist();
  render();
}

function isDemoPerson(person) {
  return person?.profile?.name === demoPerson.profile.name
    && person?.profile?.headline === demoPerson.profile.headline;
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

function renderStorageStatus() {
  const folderName = folderHandle?.name || "No folder connected";
  const detail = storageInfo.lastSaved ? `Last saved ${storageInfo.lastSaved}` : storageInfo.detail;
  const needsAttention = storageInfo.mode !== "folder-connected";
  storagePanel.classList.toggle("storage-alert", needsAttention);
  storageStatus.innerHTML = `
    <div class="registry-row ${needsAttention ? "storage-status-alert" : ""}">
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
      storageInfo.lastSaved = "";
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
      storageInfo.lastSaved = "";
      renderStorageStatus();
      return;
    }
    folderHandle = handle;
    await saveRootHandle(handle);
    const loadedState = await loadStateFromFolder(handle);
    if (loadedState) {
      state = loadedState;
      persistToLocalStorage();
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
  const granted = await ensureHandlePermission(folderHandle, "readwrite", false);
  if (!granted) {
    storageInfo.mode = "needs-reconnect";
    storageInfo.detail = "Browser permission for the saved CV_builder folder was not granted on startup.";
    return;
  }
  const loadedState = await loadStateFromFolder(folderHandle);
  if (loadedState) {
    state = loadedState;
    persistToLocalStorage();
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
    storageInfo.lastSaved = "";
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

function populateTemplate(section, overrides = {}) {
  return {
    id: overrides.id || crypto.randomUUID(),
    title: overrides.title || section.title,
    subtitle: overrides.subtitle ?? section.subtitle,
    renderMode: overrides.renderMode || section.renderMode,
    includeInCv: overrides.includeInCv ?? section.includeInCv,
    fields: (overrides.fields || section.fields).map((field) => ({ ...field, id: field.id || crypto.randomUUID() })),
    entries: overrides.entries || section.entries || []
  };
}

function makeField(label, key, type, width) {
  return { id: crypto.randomUUID(), label, key, type, width };
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
