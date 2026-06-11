const STORAGE_KEY = "budget-app-v2";
const STANDARD_CATEGORY_NAMES = [
  "Wohnen",
  "Lebensmittel",
  "Transport",
  "Auto",
  "Versicherung",
  "Gesundheit",
  "Freizeit",
  "Reisen",
  "Abos",
  "Schulden",
  "Einkommen",
  "Sonstiges",
];
const ENTRY_STATUSES = ["open", "paid", "ended"];

const defaults = {
  categories: STANDARD_CATEGORY_NAMES.map((name) => ({ id: createId(), name, budget: defaultBudgetForCategory(name) })),
  entries: [
  ],
  debts: [],
  assets: [],
  people: [
    { id: createId(), name: "Gemeinsam" },
  ],
  selectedPersonId: "all",
};

const state = loadState();
const formatMoney = new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR" });
ensurePeople();
ensureStandardCategories();
repairCopiedAssets();

const elements = {
  monthInput: document.querySelector("#monthInput"),
  entryForm: document.querySelector("#entryForm"),
  assetForm: document.querySelector("#assetForm"),
  assetFormTitle: document.querySelector("#assetFormTitle"),
  assetId: document.querySelector("#assetId"),
  assetPersonInput: document.querySelector("#assetPersonInput"),
  assetNameInput: document.querySelector("#assetNameInput"),
  assetAmountInput: document.querySelector("#assetAmountInput"),
  assetNoteInput: document.querySelector("#assetNoteInput"),
  assetModalActions: document.querySelector("#assetModalActions"),
  assetCancelButton: document.querySelector("#assetCancelButton"),
  assetDeleteButton: document.querySelector("#assetDeleteButton"),
  formTitle: document.querySelector("#formTitle"),
  entryId: document.querySelector("#entryId"),
  typeInput: document.querySelector("#typeInput"),
  typeField: document.querySelector("[data-type-field]"),
  personInput: document.querySelector("#personInput"),
  personNameInput: document.querySelector("#personNameInput"),
  addPersonButton: document.querySelector("#addPersonButton"),
  personSummaryList: document.querySelector("#personSummaryList"),
  toggleFormButton: document.querySelector("#toggleFormButton"),
  dateInput: document.querySelector("#dateInput"),
  categoryInput: document.querySelector("#categoryInput"),
  customCategoryInput: document.querySelector("#customCategoryInput"),
  amountInput: document.querySelector("#amountInput"),
  descriptionInput: document.querySelector("#descriptionInput"),
  paymentInput: document.querySelector("#paymentInput"),
  recurringInput: document.querySelector("#recurringInput"),
  entryEndInput: document.querySelector("#entryEndInput"),
  entryStatusInput: document.querySelector("#entryStatusInput"),
  resetFormButton: document.querySelector("#resetFormButton"),
  debtId: document.querySelector("#debtId"),
  creditorInput: document.querySelector("#creditorInput"),
  debtTotalInput: document.querySelector("#debtTotalInput"),
  paidSoFarInput: document.querySelector("#paidSoFarInput"),
  debtPaymentInput: document.querySelector("#debtPaymentInput"),
  debtStartInput: document.querySelector("#debtStartInput"),
  debtEndInput: document.querySelector("#debtEndInput"),
  debtStatusInput: document.querySelector("#debtStatusInput"),
  debtPaymentMethodInput: document.querySelector("#debtPaymentMethodInput"),
  debtAccountInput: document.querySelector("#debtAccountInput"),
  principalInput: document.querySelector("#principalInput"),
  interestInput: document.querySelector("#interestInput"),
  nominalRateInput: document.querySelector("#nominalRateInput"),
  termInput: document.querySelector("#termInput"),
  finalPaymentInput: document.querySelector("#finalPaymentInput"),
  normalFields: document.querySelectorAll("[data-normal-field]"),
  debtFields: document.querySelectorAll("[data-debt-field]"),
  formMessage: document.querySelector("#formMessage"),
  debtList: document.querySelector("#debtList"),
  assetList: document.querySelector("#assetList"),
  addDebtButton: document.querySelector("#addDebtButton"),
  addAssetButton: document.querySelector("#addAssetButton"),
  addIncomeButton: document.querySelector("#addIncomeButton"),
  addExpenseButton: document.querySelector("#addExpenseButton"),
  debtTitle: document.querySelector("#debtTitle"),
  assetTitle: document.querySelector("#assetTitle"),
  entryTitle: document.querySelector("#entryTitle"),
  categoryList: document.querySelector("#categoryList"),
  categoryTemplate: document.querySelector("#categoryTemplate"),
  addCategoryButton: document.querySelector("#addCategoryButton"),
  entryMenuButton: document.querySelector("#entryMenuButton"),
  entryMenuPanel: document.querySelector("#entryMenuPanel"),
  filterInput: document.querySelector("#filterInput"),
  entryList: document.querySelector("#entryList"),
  exportButton: document.querySelector("#exportButton"),
  importInput: document.querySelector("#importInput"),
  clearButton: document.querySelector("#clearButton"),
  remainingValue: document.querySelector("#remainingValue"),
  totalDebtValue: document.querySelector("#totalDebtValue"),
  totalAssetValue: document.querySelector("#totalAssetValue"),
  netWorthValue: document.querySelector("#netWorthValue"),
  currentMonthValue: document.querySelector("#currentMonthValue"),
  previousMonthValue: document.querySelector("#previousMonthValue"),
  monthDifferenceValue: document.querySelector("#monthDifferenceValue"),
  topCategoryList: document.querySelector("#topCategoryList"),
  recurringInsightList: document.querySelector("#recurringInsightList"),
  restBudgetValue: document.querySelector("#restBudgetValue"),
  restBudgetHint: document.querySelector("#restBudgetHint"),
  carryoverCount: document.querySelector("#carryoverCount"),
  carryoverHint: document.querySelector("#carryoverHint"),
  summaryCards: document.querySelectorAll("[data-summary]"),
  summaryBreakdown: document.querySelector("#summaryBreakdown"),
  summaryBreakdownTitle: document.querySelector("#summaryBreakdownTitle"),
  summaryBreakdownList: document.querySelector("#summaryBreakdownList"),
  summaryBreakdownClose: document.querySelector("#summaryBreakdownClose"),
  debtTotalMini: document.querySelector("#debtTotalMini"),
  assetTotalMini: document.querySelector("#assetTotalMini"),
  entryIncomeMini: document.querySelector("#entryIncomeMini"),
  entryExpenseMini: document.querySelector("#entryExpenseMini"),
  clearDebtsButton: document.querySelector("#clearDebtsButton"),
  clearAssetsButton: document.querySelector("#clearAssetsButton"),
  clearEntriesButton: document.querySelector("#clearEntriesButton"),
  confirmDialog: document.querySelector("#confirmDialog"),
  confirmDialogTitle: document.querySelector("#confirmDialogTitle"),
  confirmDialogMessage: document.querySelector("#confirmDialogMessage"),
  confirmDialogOk: document.querySelector("#confirmDialogOk"),
  confirmDialogCancel: document.querySelector("#confirmDialogCancel"),
  confirmDialogBackdrop: document.querySelector("#confirmDialog .confirm-dialog-backdrop"),
  choiceDialog: document.querySelector("#choiceDialog"),
  choiceDialogTitle: document.querySelector("#choiceDialogTitle"),
  choiceDialogMessage: document.querySelector("#choiceDialogMessage"),
  choiceDialogOptions: document.querySelector("#choiceDialogOptions"),
  choiceDialogCancel: document.querySelector("#choiceDialogCancel"),
  choiceDialogBackdrop: document.querySelector("#choiceDialog .confirm-dialog-backdrop"),
  promptDialog: document.querySelector("#promptDialog"),
  promptDialogTitle: document.querySelector("#promptDialogTitle"),
  promptDialogMessage: document.querySelector("#promptDialogMessage"),
  promptDialogInput: document.querySelector("#promptDialogInput"),
  promptDialogOk: document.querySelector("#promptDialogOk"),
  promptDialogCancel: document.querySelector("#promptDialogCancel"),
  promptDialogBackdrop: document.querySelector("#promptDialog .confirm-dialog-backdrop"),
  personContextMenu: document.querySelector("#personContextMenu"),
  editModal: document.querySelector("#editModal"),
  editModalBackdrop: document.querySelector("#editModalBackdrop"),
  editModalSlot: document.querySelector("#editModalSlot"),
  editModalActions: document.querySelector("#editModalActions"),
  editCancelButton: document.querySelector("#editCancelButton"),
  editDeleteButton: document.querySelector("#editDeleteButton"),
};

let formTypeChoiceVisible = true;
let activeSummaryType = "";
let activeFormKey = "";
let activeEditContext = null;

elements.monthInput.value = currentLocalMonth();
elements.typeInput.value = "expense";

elements.entryForm.addEventListener("submit", saveEntry);
elements.entryForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.entryForm.addEventListener("change", (event) => clearFieldError(event.target));
elements.assetForm.addEventListener("submit", saveAsset);
elements.assetForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.assetForm.addEventListener("change", (event) => clearFieldError(event.target));
elements.resetFormButton.addEventListener("click", () => resetForm());
elements.toggleFormButton.addEventListener("click", toggleForm);
elements.addDebtButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  openTypedForm("debt");
});
elements.addIncomeButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  openTypedForm("income");
});
elements.addExpenseButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  openTypedForm("expense");
});
elements.addAssetButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  openAssetForm();
});
elements.addPersonButton.addEventListener("click", addPerson);
elements.personNameInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    addPerson();
  }
});
elements.typeInput.addEventListener("change", syncFormMode);
elements.recurringInput.addEventListener("change", syncRecurringFields);
elements.categoryInput.addEventListener("change", () => {
  if (elements.categoryInput.value) elements.customCategoryInput.value = "";
});
elements.addCategoryButton.addEventListener("click", addCategory);
elements.filterInput.addEventListener("change", () => {
  render();
  closeEntryMenu();
});
elements.monthInput.addEventListener("change", render);
elements.summaryCards.forEach((card) => {
  card.addEventListener("click", () => toggleSummaryBreakdown(card.dataset.summary));
});
elements.summaryBreakdownClose.addEventListener("click", () => closeSummaryBreakdown());
elements.entryMenuButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  toggleEntryMenu();
});
elements.entryMenuPanel.addEventListener("click", (event) => event.stopPropagation());
elements.exportButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeEntryMenu();
  exportBackup();
});
elements.importInput.addEventListener("change", importBackup);
elements.clearButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  clearAllData();
});
elements.clearDebtsButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  clearDebtsForContext();
});
elements.clearAssetsButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  clearAssetsForContext();
});
elements.clearEntriesButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeEntryMenu();
  clearEntriesForContext();
});
document.querySelectorAll(".collapsible-panel .list-actions").forEach((node) => {
  node.addEventListener("click", (event) => event.stopPropagation());
});
elements.filterInput.addEventListener("click", (event) => event.stopPropagation());
document.addEventListener("click", (event) => {
  if (elements.entryMenuPanel.hidden) return;
  if (elements.entryMenuPanel.contains(event.target) || elements.entryMenuButton.contains(event.target)) return;
  closeEntryMenu();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeEntryMenu();
    if (!elements.editModal.hidden) closeEditModal();
  }
});
elements.editModalBackdrop.addEventListener("click", closeEditModal);
elements.editModal.addEventListener("click", (event) => {
  if (event.target === elements.editModal) closeEditModal();
});
elements.editCancelButton.addEventListener("click", closeEditModal);
elements.editDeleteButton.addEventListener("click", deleteActiveEditItem);
elements.assetCancelButton.addEventListener("click", closeEditModal);
elements.assetDeleteButton.addEventListener("click", deleteActiveEditItem);
[elements.amountInput, elements.debtTotalInput, elements.paidSoFarInput, elements.debtPaymentInput, elements.principalInput, elements.interestInput, elements.finalPaymentInput].forEach((input) => {
  input.addEventListener("blur", () => formatMoneyField(input));
  input.addEventListener("focus", () => input.select());
});
elements.assetAmountInput.addEventListener("blur", () => formatMoneyField(elements.assetAmountInput));
elements.assetAmountInput.addEventListener("focus", () => elements.assetAmountInput.select());
elements.nominalRateInput.addEventListener("blur", () => {
  elements.nominalRateInput.value = formatPercentInput(elements.nominalRateInput.value);
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").then((registration) => {
      registration.update();
      if (registration.waiting) {
        registration.waiting.postMessage({ type: "SKIP_WAITING" });
      }
      registration.addEventListener("updatefound", () => {
        const worker = registration.installing;
        if (!worker) return;
        worker.addEventListener("statechange", () => {
          if (worker.state === "installed" && navigator.serviceWorker.controller && !sessionStorage.getItem("budget-app-reloaded")) {
            sessionStorage.setItem("budget-app-reloaded", "1");
            window.location.reload();
          }
        });
      });
    });
  });
}

render();
syncFormMode();

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return clone(defaults);

  try {
    const parsed = JSON.parse(saved);
    return {
      categories: Array.isArray(parsed.categories) ? parsed.categories : clone(defaults.categories),
      entries: Array.isArray(parsed.entries) ? parsed.entries.map(normalizeEntry) : clone(defaults.entries),
      debts: Array.isArray(parsed.debts) ? parsed.debts.map(normalizeDebt) : [],
      assets: Array.isArray(parsed.assets) ? parsed.assets.map(normalizeAsset) : [],
      people: Array.isArray(parsed.people) ? parsed.people.map(normalizePerson) : clone(defaults.people),
      selectedPersonId: parsed.selectedPersonId || "all",
    };
  } catch {
    return clone(defaults);
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function defaultPersonId() {
  return state.people[0]?.id || "gemeinsam";
}

function fallbackPersonId() {
  return defaults.people[0]?.id || "gemeinsam";
}

function cleanCopyLabel(value) {
  const text = String(value || "").trim();
  const cleaned = text.replace(/\s*(?:\(|\[)?kopie(?:\)|\])?\s*$/i, "").trim();
  return cleaned || text;
}

function ensurePeople() {
  let changed = false;
  if (!Array.isArray(state.people) || !state.people.length) {
    state.people = clone(defaults.people);
    changed = true;
  }
  state.people.forEach((person) => {
    const cleanName = cleanCopyLabel(person.name);
    if (cleanName && cleanName !== person.name) {
      person.name = cleanName;
      changed = true;
    }
  });
  if (!state.selectedPersonId) {
    state.selectedPersonId = "all";
    changed = true;
  }
  const validIds = new Set(state.people.map((person) => person.id));
  if (state.selectedPersonId !== "all" && !validIds.has(state.selectedPersonId)) {
    state.selectedPersonId = "all";
    changed = true;
  }
  if (state.selectedPersonId === defaultPersonId()) {
    state.selectedPersonId = "all";
    changed = true;
  }
  [...state.entries, ...state.debts, ...state.assets].forEach((item) => {
    if (!item.personId || !validIds.has(item.personId)) {
      item.personId = defaultPersonId();
      changed = true;
    }
  });
  if (changed) persist();
}

function ensureStandardCategories() {
  let changed = false;
  STANDARD_CATEGORY_NAMES.forEach((name) => {
    if (!state.categories.some((category) => sameCategory(category.name, name))) {
      state.categories.push({ id: createId(), name, budget: defaultBudgetForCategory(name) });
      changed = true;
    }
  });
  state.entries.forEach((entry) => {
    const name = String(entry.category || "").trim();
    if (name && !state.categories.some((category) => sameCategory(category.name, name))) {
      state.categories.push({ id: createId(), name, budget: 0 });
      changed = true;
    }
  });
  state.categories.forEach((category) => {
    if (!category.id) {
      category.id = createId();
      changed = true;
    }
    category.name = String(category.name || "").trim() || "Sonstiges";
    category.budget = Number(category.budget || 0);
  });
  if (changed) persist();
}

function toggleEntryMenu() {
  elements.entryMenuPanel.hidden = !elements.entryMenuPanel.hidden;
}

function closeEntryMenu() {
  elements.entryMenuPanel.hidden = true;
}

function repairCopiedAssets() {
  const originalsBySignature = new Map();
  let changed = false;

  state.assets.forEach((asset) => {
    const signature = assetSignature(asset);
    if (!signature) return;

    if (asset.duplicateOf) {
      if (!originalsBySignature.has(signature)) originalsBySignature.set(signature, asset);
      return;
    }

    const original = originalsBySignature.get(signature);
    if (original && (original.personId || defaultPersonId()) !== (asset.personId || defaultPersonId())) {
      asset.duplicateOf = original.duplicateOf || original.id;
      changed = true;
      return;
    }

    originalsBySignature.set(signature, asset);
  });

  if (changed) persist();
}

function assetSignature(asset) {
  const name = String(asset.name || "").trim().toLowerCase();
  const note = String(asset.note || "").trim().toLowerCase();
  const amount = Number(asset.amount || 0).toFixed(2);
  if (!name && amount === "0.00") return "";
  return `${name}|${amount}|${note}`;
}

function closeForm() {
  elements.entryForm.hidden = true;
  elements.entryForm.classList.remove("modal-form");
  elements.editModalActions.hidden = true;
  elements.editDeleteButton.hidden = false;
  elements.assetForm.hidden = true;
  elements.assetForm.classList.remove("modal-form");
  elements.assetModalActions.hidden = true;
  elements.assetDeleteButton.hidden = false;
  elements.toggleFormButton.textContent = "+ Neuer Eintrag";
  activeFormKey = "";
  activeEditContext = null;
  clearFieldErrors();
}

function toggleForm() {
  if (isModalOpenFor("main")) {
    closeEditModal();
    return;
  }
  resetForm({ showTypeChoice: true, type: "expense" });
  openEntryModal({ key: "main", mode: "create", type: "expense", showTypeChoice: true });
}

function openTypedForm(type) {
  const key = `typed:${type}`;
  if (isModalOpenFor(key)) {
    closeEditModal();
    return;
  }
  resetForm({ showTypeChoice: false, type });
  openEntryModal({ key, mode: "create", type, showTypeChoice: false });
  fillPersonSelect();
  elements.personInput.value = targetPersonIdForNewItem();
  if (type === "income") elements.categoryInput.value = "Gehalt";
  syncFormMode();
}

function targetPersonIdForNewItem() {
  return state.selectedPersonId === "all" ? defaultPersonId() : state.selectedPersonId;
}

function isModalOpenFor(key) {
  return !elements.editModal.hidden && activeFormKey === key;
}

function openEntryModal({ key, mode, type, id = "", showTypeChoice = false }) {
  if (isModalOpenFor(key)) {
    closeEditModal();
    return false;
  }
  activeEditContext = { kind: type === "debt" ? "debt" : "entry", mode, id };
  activeFormKey = key;
  elements.editModalSlot.append(elements.entryForm);
  elements.assetForm.hidden = true;
  elements.entryForm.hidden = false;
  elements.entryForm.classList.add("modal-form");
  elements.editModalActions.hidden = false;
  elements.editDeleteButton.hidden = mode !== "edit";
  elements.toggleFormButton.textContent = key === "main" ? "× Formular schließen" : "+ Neuer Eintrag";
  elements.editModal.hidden = false;
  document.body.classList.add("modal-open");
  setTypeChoiceVisible(showTypeChoice);
  elements.typeInput.value = type;
  fillPersonSelect();
  syncFormMode();
  updateFormTitle();
  focusModalForm(elements.entryForm);
  return true;
}

function openEditForm({ type, id }) {
  const key = `edit-${type}:${id}`;
  return openEntryModal({ key, mode: "edit", type, id, showTypeChoice: false });
}

function closeEditModal() {
  if (elements.editModal.hidden) return;
  elements.editModal.hidden = true;
  document.body.classList.remove("modal-open");
  closeForm();
}

function focusModalForm(form) {
  window.setTimeout(() => form.querySelector("input:not([type='hidden']), select")?.focus({ preventScroll: true }), 30);
}

function setTypeChoiceVisible(visible) {
  formTypeChoiceVisible = visible;
  elements.typeField.hidden = !visible;
  updateFormTitle();
}

function updateFormTitle() {
  if (activeEditContext?.mode === "edit") {
    elements.formTitle.textContent = activeEditContext.kind === "debt" ? "Schuld bearbeiten" : "Eintrag bearbeiten";
    return;
  }
  if (formTypeChoiceVisible) {
    elements.formTitle.textContent = "Neuer Eintrag";
    return;
  }

  const titles = {
    debt: "Neue Schuld",
    income: "Neue Einnahme",
    expense: "Neue Ausgabe",
  };
  elements.formTitle.textContent = titles[elements.typeInput.value] || "Eintragen";
}

function addPerson() {
  const name = elements.personNameInput.value.trim();
  if (!name) {
    showMessage("Bitte Namen eingeben.");
    return;
  }
  const person = { id: createId(), name };
  state.people.push(person);
  state.selectedPersonId = person.id;
  elements.personNameInput.value = "";
  persist();
  render();
  elements.personInput.value = person.id;
  showMessage("Person gespeichert.");
}

function saveEntry(event) {
  event.preventDefault();
  clearFieldErrors();
  if (elements.typeInput.value === "debt") {
    saveDebtFromMainForm();
    return;
  }

  const amount = parseMoneyInput(elements.amountInput.value);
  if (!Number.isFinite(amount) || amount <= 0) {
    setFieldError(elements.amountInput, "Bitte Betrag eingeben.");
    return;
  }
  if (elements.recurringInput.checked && elements.entryEndInput.value && dateToMonthNumber(elements.entryEndInput.value) < dateToMonthNumber(elements.dateInput.value || monthToDate(elements.monthInput.value))) {
    setFieldError(elements.entryEndInput, "Enddatum muss nach dem Startdatum liegen.");
    return;
  }

  const category = resolveCategoryInput(elements.typeInput.value);

  const entry = {
    id: elements.entryId.value || createId(),
    personId: elements.personInput.value || defaultPersonId(),
    type: elements.typeInput.value,
    date: elements.dateInput.value || monthToDate(elements.monthInput.value),
    category,
    description: elements.descriptionInput.value.trim() || (elements.typeInput.value === "income" ? "Einnahme" : "Ausgabe"),
    payment: elements.paymentInput.value,
    amount,
    recurring: elements.recurringInput.checked,
    endDate: elements.recurringInput.checked ? elements.entryEndInput.value : "",
    status: normalizeStatus(elements.entryStatusInput.value),
    updatedAt: Date.now(),
  };

  const existingIndex = state.entries.findIndex((item) => item.id === entry.id);
  if (existingIndex >= 0) {
    state.entries[existingIndex] = entry;
  } else {
    state.entries.push(entry);
  }

  persist();
  resetForm({ showTypeChoice: true, type: "expense" });
  render();
  showMessage("Gespeichert.", "success");
  if (activeEditContext) closeEditModal();
  else closeForm();
}

function resetForm({ showTypeChoice = formTypeChoiceVisible, type } = {}) {
  const nextType = type || (showTypeChoice ? "expense" : elements.typeInput.value || "expense");
  elements.entryForm.reset();
  elements.entryId.value = "";
  elements.debtId.value = "";
  elements.typeInput.value = nextType;
  elements.paymentInput.value = "Karte";
  elements.recurringInput.checked = false;
  elements.entryEndInput.value = "";
  elements.entryStatusInput.value = "open";
  elements.customCategoryInput.value = "";
  elements.debtPaymentMethodInput.value = "";
  elements.debtStatusInput.value = "open";
  fillCategorySelect();
  setTypeChoiceVisible(showTypeChoice);
  syncFormMode();
}

function saveDebtFromMainForm() {
  const totalAmount = parseMoneyInput(elements.debtTotalInput.value);
  const paidSoFar = parseMoneyInput(elements.paidSoFarInput.value);
  const monthlyPayment = parseMoneyInput(elements.debtPaymentInput.value);
  const principalAmount = parseMoneyInput(elements.principalInput.value);
  const interestAmount = parseMoneyInput(elements.interestInput.value);
  const nominalRate = parsePercentInput(elements.nominalRateInput.value);
  const finalPayment = parseMoneyInput(elements.finalPaymentInput.value);
  const hasAnyDebtInput = [
    elements.creditorInput.value,
    elements.debtTotalInput.value,
    elements.paidSoFarInput.value,
    elements.debtPaymentInput.value,
    elements.debtStartInput.value,
    elements.debtEndInput.value,
    elements.debtPaymentMethodInput.value,
    elements.debtAccountInput.value,
    elements.principalInput.value,
    elements.interestInput.value,
    elements.nominalRateInput.value,
    elements.termInput.value,
    elements.finalPaymentInput.value,
  ].some((value) => String(value).trim());

  if (!hasAnyDebtInput) {
    setFieldError(elements.creditorInput, "Bitte erst etwas eintragen.");
    return;
  }
  if (elements.debtTotalInput.value.trim() && (!Number.isFinite(totalAmount) || totalAmount < 0)) {
    setFieldError(elements.debtTotalInput, "Gesamtschuld ist keine gültige Zahl.");
    return;
  }
  if (elements.debtPaymentInput.value.trim() && (!Number.isFinite(monthlyPayment) || monthlyPayment < 0)) {
    setFieldError(elements.debtPaymentInput, "Monatlich ist keine gültige Zahl.");
    return;
  }

  const debt = {
    id: elements.debtId.value || createId(),
    personId: elements.personInput.value || defaultPersonId(),
    creditor: elements.creditorInput.value.trim(),
    totalAmount: Number.isFinite(totalAmount) ? totalAmount : 0,
    paidSoFar: Number.isFinite(paidSoFar) ? paidSoFar : 0,
    monthlyPayment: Number.isFinite(monthlyPayment) ? monthlyPayment : 0,
    startDate: elements.debtStartInput.value,
    endDate: elements.debtEndInput.value,
    status: normalizeStatus(elements.debtStatusInput.value),
    paymentMethod: elements.debtPaymentMethodInput.value,
    account: elements.debtAccountInput.value.trim(),
    principalAmount: Number.isFinite(principalAmount) ? principalAmount : 0,
    interestAmount: Number.isFinite(interestAmount) ? interestAmount : 0,
    nominalRate: Number.isFinite(nominalRate) ? nominalRate : 0,
    termMonths: Number(elements.termInput.value) || 0,
    finalPayment: Number.isFinite(finalPayment) ? finalPayment : 0,
    note: "",
  };

  if (debt.endDate && debt.startDate && dateToMonthNumber(debt.endDate) < dateToMonthNumber(debt.startDate)) {
    setFieldError(elements.debtEndInput, "Enddatum muss nach dem Startdatum liegen.");
    return;
  }

  const existingIndex = state.debts.findIndex((item) => item.id === debt.id);
  if (existingIndex >= 0) {
    state.debts[existingIndex] = debt;
  } else {
    state.debts.push(debt);
  }

  persist();
  resetForm();
  render();
  showMessage("Gespeichert.", "success");
  if (activeEditContext) closeEditModal();
  else closeForm();
}

function addCategory() {
  state.categories.push({ id: createId(), name: "Neue Kategorie", budget: 0 });
  persist();
  render();
}

function updateCategory(id, patch) {
  const category = state.categories.find((item) => item.id === id);
  if (!category) return;
  Object.assign(category, patch);
  persist();
  renderSummary();
  fillCategorySelect();
}

function removeCategory(id) {
  state.categories = state.categories.filter((item) => item.id !== id);
  persist();
  render();
}

function render() {
  fillPersonSelect();
  fillCategorySelect();
  renderCategories();
  renderPeople();
  renderDetailTitles();
  renderDebts();
  renderAssets();
  renderEntries();
  renderSummary();
}

function fillPersonSelect() {
  const existing = elements.personInput.value;
  const options = [{ id: defaultPersonId(), name: "Gesamt" }, ...visiblePeople()];
  elements.personInput.innerHTML = options
    .map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)}</option>`)
    .join("");
  if (existing && options.some((person) => person.id === existing)) {
    elements.personInput.value = existing;
  } else if (state.selectedPersonId !== "all" && options.some((person) => person.id === state.selectedPersonId)) {
    elements.personInput.value = state.selectedPersonId;
  } else {
    elements.personInput.value = options[0].id;
  }
}

function renderPeople() {
  const month = elements.monthInput.value;
  const familyTotals = monthlyTotalsForFamily(month);
  const familyClass = familyTotals.balance >= 0 ? "positive-text" : "negative-text";
  const familyCard = `
    <article class="person-card family-card ${state.selectedPersonId === "all" ? "selected" : ""}" data-person-select="all">
      <div class="person-copy">
        <strong>Gesamt</strong>
        <b class="${familyClass}">${formatMoney.format(familyTotals.balance)}</b>
      </div>
      <div class="person-chips-row">
        <span class="person-chip income-chip"><span>Einnahmen</span><strong>${formatMoney.format(familyTotals.income)}</strong></span>
        <span class="person-chip expense-chip"><span>Ausgaben</span><strong>${formatMoney.format(familyTotals.expense)}</strong></span>
        <span class="person-chip asset-chip"><span>Vermögen</span><strong>${formatMoney.format(familyTotals.assets)}</strong></span>
        <span class="person-chip debt-chip"><span>Schulden</span><strong>${formatMoney.format(familyTotals.debt)}</strong></span>
      </div>
    </article>
  `;
  const personCards = visiblePeople().map((person) => {
    const totals = monthlyTotalsForPerson(person.id, month);
    const balanceClass = totals.balance >= 0 ? "positive-text" : "negative-text";
    return `
      <article class="person-card ${state.selectedPersonId === person.id ? "selected" : ""}" data-person-select="${escapeHtml(person.id)}">
        <div class="person-copy">
          <strong>${escapeHtml(person.name)}</strong>
          <b class="${balanceClass}">${formatMoney.format(totals.balance)}</b>
        </div>
        <div class="person-chips-row">
          <span class="person-chip income-chip"><span>Einnahmen</span><strong>${formatMoney.format(totals.income)}</strong></span>
          <span class="person-chip expense-chip"><span>Ausgaben</span><strong>${formatMoney.format(totals.expense)}</strong></span>
          <span class="person-chip asset-chip"><span>Vermögen</span><strong>${formatMoney.format(totals.assets)}</strong></span>
          <span class="person-chip debt-chip"><span>Schulden</span><strong>${formatMoney.format(totals.debt)}</strong></span>
        </div>
        <div class="person-controls">
          <button class="icon-button" type="button" title="Person bearbeiten" aria-label="Person bearbeiten" data-person-edit="${person.id}">✎</button>
          <button class="icon-button" type="button" title="Person löschen" aria-label="Person löschen" data-person-delete="${person.id}">×</button>
        </div>
      </article>
    `;
  }).join("");

  elements.personSummaryList.innerHTML = familyCard + personCards;

  elements.personSummaryList.querySelectorAll("[data-person-select]").forEach((card) => {
    card.addEventListener("click", (event) => {
      if (event.target.closest("button")) return;
      selectPerson(card.dataset.personSelect);
    });
  });

  elements.personSummaryList.querySelectorAll("[data-person-edit]").forEach((button) => {
    button.addEventListener("click", () => editPerson(button.dataset.personEdit));
  });
  elements.personSummaryList.querySelectorAll("[data-person-delete]").forEach((button) => {
    button.addEventListener("click", () => deletePerson(button.dataset.personDelete));
  });
}

function selectPerson(id) {
  state.selectedPersonId = id || "all";
  persist();
  render();
}

function renderDetailTitles() {
  const suffix = selectedPersonName();
  const month = elements.monthInput.value;
  const totals = selectedMonthlyTotals(month);
  const debt = selectedDebtTotal(month);
  const assetTotal = state.assets
  .filter(includeInSelectedTotal)
  .reduce((total, asset) => total + Number(asset.amount || 0), 0);
  elements.debtTitle.textContent = `Schulden - ${suffix}`;
  elements.debtTotalMini.textContent = formatMoney.format(debt);
  elements.assetTitle.textContent = `Vermögen - ${suffix}`;
  elements.assetTotalMini.textContent = formatMoney.format(assetTotal);
  elements.entryTitle.textContent = `Einnahmen und Ausgaben - ${suffix}`;
  elements.entryIncomeMini.textContent = `Einn. ${formatMoney.format(totals.income)}`;
  elements.entryExpenseMini.textContent = `Ausg. ${formatMoney.format(totals.expense)}`;
}

function editPerson(id) {
  if (id === defaultPersonId()) {
    showMessage("Gesamt kann nicht bearbeitet werden.");
    return;
  }
  const person = state.people.find((item) => item.id === id);
  if (!person) return;
  const nextName = window.prompt("Name ändern", person.name);
  if (nextName === null) return;
  const name = nextName.trim();
  if (!name) {
    showMessage("Name darf nicht leer sein.");
    return;
  }
  person.name = name;
  persist();
  render();
  showMessage("Person geändert.");
}

async function deletePerson(id) {
  if (id === defaultPersonId()) {
    showMessage("Gesamt kann nicht gelöscht werden.");
    return;
  }
  const person = state.people.find((item) => item.id === id);
  if (!person) return;

  const remainingPeople = state.people.filter((item) => item.id !== id);
  const targetId = defaultPersonId();
  const targetName = "Gesamt";
  const ok = await confirmCentered(
    `${person.name} löschen? Die Einträge bleiben und gehen zu ${targetName}.`,
    { title: `${person.name} löschen` }
  );
  if (!ok) return;

  state.entries = state.entries.filter((entry) => !(entry.personId === id && entry.duplicateOf));
  state.debts = state.debts.filter((debt) => !(debt.personId === id && debt.duplicateOf));
  state.assets = state.assets.filter((asset) => !(asset.personId === id && asset.duplicateOf));

  [...state.entries, ...state.debts, ...state.assets].forEach((item) => {
    if (item.personId === id) item.personId = targetId;
  });
  state.people = remainingPeople;
  if (state.selectedPersonId === id) state.selectedPersonId = "all";
  persist();
  render();
  showMessage("Person gelöscht.");
}

function monthlyTotalsForPerson(personId, month) {
  const entries = state.entries.filter((entry) => (entry.personId || defaultPersonId()) === personId && isEntryActiveInMonth(entry, month));
  const income = sum(entries.filter((entry) => entry.type === "income"));
  const entryExpense = sum(entries.filter((entry) => entry.type === "expense"));
  const debtExpense = state.debts
    .filter((debt) => (debt.personId || defaultPersonId()) === personId && isDebtActiveInMonth(debt, month))
    .reduce((total, debt) => total + Number(debt.monthlyPayment || 0), 0);
  const expense = entryExpense + debtExpense;
  const assets = state.assets
    .filter((asset) => (asset.personId || defaultPersonId()) === personId)
    .reduce((total, asset) => total + Number(asset.amount || 0), 0);
  const debt = state.debts
    .filter((d) => (d.personId || defaultPersonId()) === personId)
    .reduce((total, d) => total + remainingDebt(d, month), 0);
  return { income, expense, assets, debt, balance: income - expense };
}

function monthlyTotalsForFamily(month) {
  const income = state.entries
    .filter((entry) => !entry.duplicateOf && isEntryActiveInMonth(entry, month) && entry.type === "income")
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
  const entryExpense = state.entries
    .filter((entry) => !entry.duplicateOf && isEntryActiveInMonth(entry, month) && entry.type === "expense")
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
  const debtExpense = state.debts
    .filter((debt) => !debt.duplicateOf && isDebtActiveInMonth(debt, month))
    .reduce((total, debt) => total + Number(debt.monthlyPayment || 0), 0);
  const expense = entryExpense + debtExpense;
  const assets = state.assets
    .filter((asset) => !asset.duplicateOf)
    .reduce((total, asset) => total + Number(asset.amount || 0), 0);
  const debt = state.debts
    .filter((d) => !d.duplicateOf)
    .reduce((total, d) => total + remainingDebt(d, month), 0);
  return { income, expense, assets, debt, balance: income - expense };
}

function syncFormMode() {
  const isDebt = elements.typeInput.value === "debt";
  const isIncome = elements.typeInput.value === "income";
  elements.normalFields.forEach((field) => {
    field.hidden = isDebt;
  });
  elements.debtFields.forEach((field) => {
    field.hidden = !isDebt;
  });

  elements.categoryInput.required = false;
  elements.amountInput.required = false;
  elements.dateInput.required = false;
  elements.creditorInput.required = false;
  elements.debtTotalInput.required = false;
  elements.paidSoFarInput.required = false;
  elements.debtPaymentInput.required = false;
  elements.debtStartInput.required = false;
  elements.debtEndInput.required = false;
  elements.debtStatusInput.required = false;
  elements.debtPaymentMethodInput.required = false;
  elements.principalInput.required = false;
  elements.interestInput.required = false;
  elements.nominalRateInput.required = false;
  elements.termInput.required = false;
  elements.finalPaymentInput.required = false;

  if (isIncome && elements.categoryInput.value !== "Gehalt") {
    elements.categoryInput.value = "Gehalt";
  }
  if (isDebt) {
    elements.creditorInput.required = true;
    elements.debtTotalInput.required = true;
  } else {
    elements.categoryInput.required = true;
    elements.amountInput.required = true;
  }
  syncRecurringFields();
  updateFormTitle();
}

function syncRecurringFields() {
  if (!elements.entryEndInput) return;
  const recurring = elements.recurringInput.checked;
  elements.entryEndInput.disabled = !recurring;
  elements.entryEndInput.closest("label")?.classList.toggle("muted-field", !recurring);
  if (!recurring) elements.entryEndInput.value = "";
}

function fillCategorySelect() {
  const existing = elements.categoryInput.value;
  const names = uniqueCategoryNames([...STANDARD_CATEGORY_NAMES, ...state.categories.map((category) => category.name), "Gehalt"]);
  elements.categoryInput.innerHTML = names
    .filter(Boolean)
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("");
  if (names.includes(existing)) elements.categoryInput.value = existing;
}

function renderCategories() {
  elements.categoryList.innerHTML = "";
  state.categories.forEach((category) => {
    const item = elements.categoryTemplate.content.firstElementChild.cloneNode(true);
    const nameInput = item.querySelector(".category-name");
    const budgetInput = item.querySelector(".category-budget");
    const deleteButton = item.querySelector(".category-delete");

    nameInput.value = category.name;
    budgetInput.value = category.budget;
    nameInput.addEventListener("change", () => updateCategory(category.id, { name: nameInput.value.trim() || "Kategorie" }));
    budgetInput.addEventListener("change", () => updateCategory(category.id, { budget: Number(budgetInput.value) || 0 }));
    deleteButton.addEventListener("click", () => removeCategory(category.id));
    elements.categoryList.append(item);
  });
}

function renderEntries() {
  const month = elements.monthInput.value;
  const filter = elements.filterInput.value;
  const entries = state.entries
    .filter((entry) => isEntryActiveInMonth(entry, month))
    .filter(includeInSelectedTotal)
    .filter((entry) => filter === "all" || entry.type === filter || (filter === "recurring" && entry.recurring))
    .sort((a, b) => {
      if (a.type !== b.type) return a.type === "income" ? -1 : 1;
      const diff = Number(b.updatedAt || 0) - Number(a.updatedAt || 0);
      if (diff !== 0) return diff;
      return b.date.localeCompare(a.date);
    });

  if (!entries.length) {
    elements.entryList.innerHTML = `<p class="empty-state">Keine Einnahmen oder Ausgaben für ${escapeHtml(selectedPersonName())} in diesem Monat.</p>`;
    return;
  }

  elements.entryList.innerHTML = entries.map((entry) => {
    const sign = entry.type === "income" ? "+" : "-";
    const title = entry.description || entry.category;
    const meta = [personName(entry.personId), entry.category, displayText(entry.payment)].filter(Boolean).join(" · ");
    const badges = entryBadges(entry, month);
    return `
      <article class="entry-row ${entry.type}">
        <time class="entry-date" datetime="${entryOccurrenceDate(entry, month)}">${formatDate(entryOccurrenceDate(entry, month))}</time>
        <div class="entry-main">
          <div class="row-heading">
            <strong>${escapeHtml(title)}</strong>
            <span class="entry-controls inline-controls">
              <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-edit="${entry.id}">✎</button>
              <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-delete="${entry.id}">×</button>
            </span>
          </div>
          <span>${escapeHtml(meta)}</span>
          ${badges ? `<div class="status-badges">${badges}</div>` : ""}
        </div>
        <div class="entry-amount ${entry.type}">${sign}${formatMoney.format(entry.amount)}</div>
      </article>
    `;
  }).join("");

  elements.entryList.querySelectorAll("[data-edit]").forEach((button) => {
    button.addEventListener("click", () => editEntry(button.dataset.edit));
  });
  elements.entryList.querySelectorAll("[data-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteEntry(button.dataset.delete));
  });
}

function renderDebts() {
  const month = elements.monthInput.value;
  const debts = state.debts
    .filter((debt) => isDebtVisibleInMonth(debt, month))
    .filter(includeInSelectedTotal)
    .sort((a, b) => a.creditor.localeCompare(b.creditor));

  if (!debts.length) {
    elements.debtList.innerHTML = `<p class="empty-state">Keine Schulden für ${escapeHtml(selectedPersonName())} eingetragen.</p>`;
    return;
  }

  elements.debtList.innerHTML = debts.map((debt) => {
    const remaining = remainingDebt(debt, month);
    const paid = paidDebt(debt, month);
    const active = isDebtActiveInMonth(debt, month);
    const term = debt.termMonths || monthsFromDates(debt.startDate, debt.endDate);
    const progress = debtProgressPercent(debt, month);
    const badges = debtBadges(debt, month, active);
    const meta = [
      personName(debt.personId),
      debt.startDate || debt.endDate ? `${debt.startDate ? formatDateFull(debt.startDate) : "kein Start"} bis ${debt.endDate ? formatDateFull(debt.endDate) : "offen"}` : "",
      displayText(debt.paymentMethod),
      debt.account ? `über ${debt.account}` : "",
      debt.note,
    ].filter(Boolean).join(" · ");
    return `
      <article class="debt-row ${active ? "active" : ""}">
        <div class="debt-main">
          <div class="row-heading">
            <strong>${escapeHtml(debt.creditor)}</strong>
            <span class="entry-controls inline-controls">
              <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-debt-edit="${debt.id}">✎</button>
              <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-debt-delete="${debt.id}">×</button>
            </span>
          </div>
          <span>${escapeHtml(meta)}</span>
          ${badges ? `<div class="status-badges">${badges}</div>` : ""}
        </div>
        <div class="debt-numbers">
          ${debtNumberLine("Restschuld", formatMoney.format(remaining), "remaining", "strong")}
          ${debtNumberLine("Monatliche Rate", formatMoney.format(debt.monthlyPayment))}
          ${debtNumberLine("Bereits bezahlt", formatMoney.format(paid), "paid")}
          <div class="debt-progress-line">
            <span>Fortschritt</span>
            <strong>${Math.round(progress)}%</strong>
          </div>
          <div class="debt-progress" aria-hidden="true">
            <span style="width: ${progress}%"></span>
          </div>
        </div>
      </article>
    `;
  }).join("");

  elements.debtList.querySelectorAll("[data-debt-edit]").forEach((button) => {
    button.addEventListener("click", () => editDebt(button.dataset.debtEdit));
  });
  elements.debtList.querySelectorAll("[data-debt-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteDebt(button.dataset.debtDelete));
  });
}

function debtNumberLine(label, value, className = "", tag = "span") {
  const classes = ["debt-number-row", className].filter(Boolean).join(" ");
  return `<${tag} class="${classes}"><span>${escapeHtml(label)}:</span><b>${escapeHtml(value)}</b></${tag}>`;
}

function debtProgressPercent(debt, month) {
  const total = Number(debt.totalAmount || 0);
  if (total <= 0) return 0;
  return Math.min(100, Math.max(0, (paidDebt(debt, month) / total) * 100));
}

function renderAssets() {
  const assets = state.assets
    .filter(includeInSelectedTotal)
    .sort((a, b) => a.name.localeCompare(b.name));

  if (!assets.length) {
    elements.assetList.innerHTML = `<p class="empty-state">Kein Vermögen für ${escapeHtml(selectedPersonName())} eingetragen.</p>`;
    return;
  }

  elements.assetList.innerHTML = assets.map((asset) => `
    <article class="asset-row">
      <div class="debt-main">
        <div class="row-heading">
          <strong>${escapeHtml(asset.name)}</strong>
          <span class="entry-controls inline-controls">
            <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-asset-edit="${asset.id}">✎</button>
            <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-asset-delete="${asset.id}">×</button>
          </span>
        </div>
        <span>${escapeHtml([personName(asset.personId), asset.note || "Plus-Wert"].filter(Boolean).join(" · "))}</span>
      </div>
      <strong class="asset-amount">${formatMoney.format(asset.amount || 0)}</strong>
    </article>
  `).join("");

  elements.assetList.querySelectorAll("[data-asset-edit]").forEach((button) => {
    button.addEventListener("click", () => editAsset(button.dataset.assetEdit));
  });
  elements.assetList.querySelectorAll("[data-asset-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteAsset(button.dataset.assetDelete));
  });
}

function openAssetForm(id = "") {
  const mode = id ? "edit" : "create";
  const key = id ? `edit-asset:${id}` : "typed:asset";
  if (isModalOpenFor(key)) {
    closeEditModal();
    return;
  }

  resetAssetForm();
  activeEditContext = { kind: "asset", mode, id };
  activeFormKey = key;
  elements.editModalSlot.append(elements.assetForm);
  elements.entryForm.hidden = true;
  elements.assetForm.hidden = false;
  elements.assetForm.classList.add("modal-form");
  elements.assetModalActions.hidden = false;
  elements.assetDeleteButton.hidden = mode !== "edit";
  elements.assetFormTitle.textContent = mode === "edit" ? "Vermögen bearbeiten" : "Neues Vermögen";
  fillAssetPersonSelect();

  if (id) {
    const asset = state.assets.find((item) => item.id === id);
    if (!asset) return;
    elements.assetId.value = asset.id;
    elements.assetPersonInput.value = asset.personId || defaultPersonId();
    elements.assetNameInput.value = asset.name;
    elements.assetAmountInput.value = formatMoneyInput(asset.amount);
    elements.assetNoteInput.value = asset.note || "";
  } else {
    elements.assetPersonInput.value = targetPersonIdForNewItem();
  }

  elements.toggleFormButton.textContent = "+ Neuer Eintrag";
  elements.editModal.hidden = false;
  document.body.classList.add("modal-open");
  focusModalForm(elements.assetForm);
}

function resetAssetForm() {
  elements.assetForm.reset();
  elements.assetId.value = "";
  elements.assetModalActions.hidden = true;
  elements.assetDeleteButton.hidden = false;
  clearFieldErrors();
}

function fillAssetPersonSelect() {
  const existing = elements.assetPersonInput.value;
  const options = [{ id: defaultPersonId(), name: "Gesamt" }, ...visiblePeople()];
  elements.assetPersonInput.innerHTML = options
    .map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)}</option>`)
    .join("");
  if (existing && options.some((person) => person.id === existing)) {
    elements.assetPersonInput.value = existing;
  } else if (state.selectedPersonId !== "all" && options.some((person) => person.id === state.selectedPersonId)) {
    elements.assetPersonInput.value = state.selectedPersonId;
  } else {
    elements.assetPersonInput.value = options[0].id;
  }
}

function saveAsset(event) {
  event.preventDefault();
  clearFieldErrors();
  const name = elements.assetNameInput.value.trim();
  const amount = parseMoneyInput(elements.assetAmountInput.value);
  if (!name) {
    setFieldError(elements.assetNameInput, "Bitte Namen eingeben.");
    return;
  }
  if (!Number.isFinite(amount) || amount < 0) {
    setFieldError(elements.assetAmountInput, "Betrag ist keine gültige Zahl.");
    return;
  }

  const asset = {
    id: elements.assetId.value || createId(),
    personId: elements.assetPersonInput.value || defaultPersonId(),
    name,
    amount,
    note: elements.assetNoteInput.value.trim() || "manuell",
  };
  const existingIndex = state.assets.findIndex((item) => item.id === asset.id);
  if (existingIndex >= 0) {
    state.assets[existingIndex] = { ...state.assets[existingIndex], ...asset };
  } else {
    state.assets.push(asset);
  }
  persist();
  render();
  showMessage(existingIndex >= 0 ? "Vermögen geändert." : "Vermögen gespeichert.", "success");
  closeEditModal();
}

function editAsset(id) {
  openAssetForm(id);
}

function deleteAsset(id) {
  state.assets = state.assets.filter((asset) => asset.id !== id);
  persist();
  render();
  if (activeEditContext?.kind === "asset" && activeEditContext.id === id) closeEditModal();
  showMessage("Vermögen gelöscht.");
}

function renderSummary() {
  const month = elements.monthInput.value;
  const summary = summaryTotals(month);
  const remaining = summary.balance;
  const totalDebt = summary.totalDebt;
  const totalAssets = summary.totalAssets;
  const netWorth = summary.netWorth;

  elements.remainingValue.textContent = formatMoney.format(remaining);
  elements.remainingValue.parentElement.classList.toggle("positive", remaining >= 0);
  elements.remainingValue.parentElement.classList.toggle("negative", remaining < 0);
  elements.totalDebtValue.textContent = formatMoney.format(totalDebt);
  elements.totalAssetValue.textContent = formatMoney.format(totalAssets);
  elements.netWorthValue.textContent = formatMoney.format(netWorth);
  elements.netWorthValue.parentElement.classList.toggle("positive", netWorth >= 0);
  elements.netWorthValue.parentElement.classList.toggle("negative", netWorth < 0);
  renderMonthComparison(month);
  renderInsights(month, summary);
  if (activeSummaryType) renderSummaryBreakdown();
}

function renderInsights(month, summary) {
  const monthEntries = state.entries
    .filter(includeInSelectedTotal)
    .filter((entry) => isEntryActiveInMonth(entry, month));
  const expenseByCategory = new Map();
  monthEntries
    .filter((entry) => entry.type === "expense")
    .forEach((entry) => {
      const key = entry.category || "Sonstiges";
      expenseByCategory.set(key, (expenseByCategory.get(key) || 0) + Number(entry.amount || 0));
    });

  const topCategories = [...expenseByCategory.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3);
  elements.topCategoryList.innerHTML = topCategories.length
    ? topCategories.map(([name, amount]) => insightLine(name, formatMoney.format(amount))).join("")
    : `<p class="empty-state compact-empty">Keine Ausgaben in diesem Monat</p>`;

  const recurringRows = monthEntries
    .filter((entry) => entry.recurring)
    .sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0))
    .slice(0, 3);
  elements.recurringInsightList.innerHTML = recurringRows.length
    ? recurringRows.map((entry) => insightLine(entry.description || entry.category, formatMoney.format(entry.amount), entry.endDate ? `endet ${formatMonthName(entry.endDate.slice(0, 7))}` : "läuft weiter")).join("")
    : `<p class="empty-state compact-empty">Keine Abos im aktuellen Monat</p>`;

  const daysLeft = daysLeftInMonth(month);
  const dailyBudget = daysLeft > 0 ? summary.balance / daysLeft : summary.balance;
  elements.restBudgetValue.textContent = formatMoney.format(summary.balance);
  elements.restBudgetValue.classList.toggle("positive-text", summary.balance >= 0);
  elements.restBudgetValue.classList.toggle("negative-text", summary.balance < 0);
  elements.restBudgetHint.textContent = `${formatMoney.format(dailyBudget)} pro verbleibendem Tag`;

  const carryoverItems = carryoverItemsForNextMonth(month);
  elements.carryoverCount.textContent = String(carryoverItems.length);
  elements.carryoverHint.textContent = carryoverItems.length
    ? `${formatMonthName(shiftMonth(month, 1))}: ${carryoverItems.slice(0, 2).map((item) => item.label).join(", ")}${carryoverItems.length > 2 ? " ..." : ""}`
    : `Keine offenen Übernahmen nach ${formatMonthName(shiftMonth(month, 1))}`;
}

function insightLine(label, value, hint = "") {
  return `
    <div class="insight-line">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      ${hint ? `<small>${escapeHtml(hint)}</small>` : ""}
    </div>
  `;
}

function summaryTotals(month) {
  const monthEntries = state.entries
    .filter(includeInSelectedTotal)
    .filter((entry) => isEntryActiveInMonth(entry, month));
  const income = sum(monthEntries.filter((entry) => entry.type === "income"));
  const entryExpense = sum(monthEntries.filter((entry) => entry.type === "expense"));
  const debtMonthlyCost = activeDebts(month)
    .filter(includeInSelectedTotal)
    .reduce((total, debt) => total + Number(debt.monthlyPayment || 0), 0);
  const totalDebt = state.debts
    .filter(includeInSelectedTotal)
    .reduce((total, debt) => total + remainingDebt(debt, month), 0);
  const totalAssets = state.assets
    .filter(includeInSelectedTotal)
    .reduce((total, asset) => total + Number(asset.amount || 0), 0);
  const expense = entryExpense + debtMonthlyCost;
  const balance = income - expense;
  const netWorth = totalAssets - totalDebt;
  return { income, entryExpense, debtMonthlyCost, expense, balance, totalDebt, totalAssets, netWorth };
}

function renderMonthComparison(month) {
  const current = summaryTotals(month).balance;
  const previous = summaryTotals(shiftMonth(month, -1)).balance;
  const difference = current - previous;
  setSignedAmount(elements.currentMonthValue, current);
  setSignedAmount(elements.previousMonthValue, previous);
  setSignedAmount(elements.monthDifferenceValue, difference);
}

function setSignedAmount(element, amount) {
  if (!element) return;
  element.textContent = `${amount >= 0 ? "+" : ""}${formatMoney.format(amount)}`;
  element.classList.toggle("positive-text", amount >= 0);
  element.classList.toggle("negative-text", amount < 0);
}

function shiftMonth(month, offset) {
  const [year, monthIndex] = month.split("-").map(Number);
  const date = new Date(year, monthIndex - 1 + offset, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function toggleSummaryBreakdown(type) {
  if (activeSummaryType === type && !elements.summaryBreakdown.hidden) {
    closeSummaryBreakdown();
    return;
  }
  activeSummaryType = type;
  renderSummaryBreakdown();
}

function closeSummaryBreakdown() {
  activeSummaryType = "";
  elements.summaryBreakdown.hidden = true;
  elements.summaryCards.forEach((card) => card.classList.remove("selected"));
}

function renderSummaryBreakdown() {
  const data = summaryBreakdownData(activeSummaryType);
  if (!data) return;

  elements.summaryBreakdown.hidden = false;
  elements.summaryBreakdownTitle.textContent = data.title;
  elements.summaryCards.forEach((card) => {
    card.classList.toggle("selected", card.dataset.summary === activeSummaryType);
  });

  const rows = data.rows.length
    ? data.rows.map(summaryBreakdownRow).join("")
    : `<p class="empty-state">Keine Einträge in diesem Monat.</p>`;

  const totalClass = data.totalClass || (data.total >= 0 ? "positive-text" : "negative-text");

  elements.summaryBreakdownList.innerHTML = `
    ${rows}
    <article class="summary-breakdown-row summary-total">
      <span>Summe</span>
      <strong class="${totalClass}">${formatMoney.format(data.total)}</strong>
    </article>
  `;
}

function summaryBreakdownRow(row) {
  const amountClass = row.amountClass || (row.amount >= 0 ? "positive-text" : "negative-text");
  return `
    <article class="summary-breakdown-row">
      <span>
        <strong>${escapeHtml(row.title)}</strong>
        <small>${escapeHtml(row.meta)}</small>
      </span>
      <strong class="${amountClass}">${formatMoney.format(row.amount)}</strong>
    </article>
  `;
}

function summaryBreakdownData(type) {
  const month = elements.monthInput.value;
  const entries = state.entries
    .filter(includeInSelectedTotal)
    .filter((entry) => isEntryActiveInMonth(entry, month));
  const incomeRows = entries
    .filter((entry) => entry.type === "income")
    .map((entry) => entrySummaryRow(entry, Number(entry.amount || 0)));
  const expenseRows = entries
    .filter((entry) => entry.type === "expense")
    .map((entry) => entrySummaryRow(entry, Number(entry.amount || 0)));
  const debtRateRows = activeDebts(month)
    .filter(includeInSelectedTotal)
    .filter((debt) => Number(debt.monthlyPayment || 0) > 0)
    .map((debt) => debtSummaryRow(debt, Number(debt.monthlyPayment || 0), "monatliche Rate"));
  const debtRows = state.debts
    .filter(includeInSelectedTotal)
    .filter((debt) => isDebtVisibleInMonth(debt, month))
    .map((debt) => debtSummaryRow(debt, remainingDebt(debt, month), "Restschuld"));
  const assetRows = state.assets
    .filter(includeInSelectedTotal)
    .map((asset) => ({
      title: asset.name || "Vermögen",
      meta: [personName(asset.personId), asset.note || ""].filter(Boolean).join(" · "),
      amount: Number(asset.amount || 0),
    }));
  const netRows = [
    ...assetRows,
    ...debtRows.map((row) => ({ ...row, amount: -Math.abs(row.amount), amountClass: "negative-text" })),
  ];

  if (type === "income") {
    return buildSummaryData("Einnahmen gesamt", incomeRows);
  }
  if (type === "expense") {
    const rows = [...expenseRows, ...debtRateRows].map((row) => ({ ...row, amountClass: "negative-text" }));
    return buildSummaryData("Ausgaben gesamt", rows, "negative-text");
  }
  if (type === "balance") {
    const balanceRows = [
      ...incomeRows,
      ...expenseRows.map((row) => ({ ...row, amount: -Math.abs(row.amount) })),
      ...debtRateRows.map((row) => ({ ...row, amount: -Math.abs(row.amount) })),
    ];
    return buildSummaryData("Monatliches Ergebnis", balanceRows);
  }
  if (type === "debt") {
    return buildSummaryData("Schulden gesamt", debtRows);
  }
  if (type === "asset") {
    return buildSummaryData("Vermögen gesamt", assetRows);
  }
  if (type === "net") {
    return buildSummaryData("Netto-Status", netRows);
  }
  return null;
}

function buildSummaryData(title, rows, totalClass = "") {
  const total = rows.reduce((sumValue, row) => sumValue + Number(row.amount || 0), 0);
  return { title, rows, total, totalClass };
}

function entrySummaryRow(entry, amount) {
  return {
    title: entry.description || entry.category || "Eintrag",
    meta: [personName(entry.personId), entry.category, displayText(entry.payment), entry.recurring ? "monatlich" : "", statusLabel(entry.status)].filter(Boolean).join(" · "),
    amount,
  };
}

function debtSummaryRow(debt, amount, label) {
  return {
    title: debt.creditor || "Schuld",
    meta: [personName(debt.personId), label, debt.account ? `über ${debt.account}` : ""].filter(Boolean).join(" · "),
    amount,
  };
}

function activeDebts(month) {
  return state.debts.filter((debt) => isDebtActiveInMonth(debt, month));
}

function editEntry(id) {
  const entry = state.entries.find((item) => item.id === id);
  if (!entry) return;

  resetForm({ showTypeChoice: false, type: entry.type });
  if (!openEditForm({ type: "entry", id })) return;
  elements.entryId.value = entry.id;
  elements.debtId.value = "";
  elements.typeInput.value = entry.type;
  elements.personInput.value = entry.personId || defaultPersonId();
  elements.dateInput.value = entry.date;
  elements.categoryInput.value = entry.category;
  elements.customCategoryInput.value = state.categories.some((category) => sameCategory(category.name, entry.category)) || entry.category === "Gehalt" ? "" : entry.category;
  elements.amountInput.value = formatMoneyInput(entry.amount);
  elements.descriptionInput.value = entry.description;
  elements.paymentInput.value = entry.payment;
  elements.recurringInput.checked = entry.recurring;
  elements.entryEndInput.value = entry.endDate || "";
  elements.entryStatusInput.value = normalizeStatus(entry.status);
  syncFormMode();
}

function deleteEntry(id) {
  state.entries = state.entries.filter((entry) => entry.id !== id);
  persist();
  render();
  if (activeEditContext?.kind === "entry" && activeEditContext.id === id) closeEditModal();
}

function editDebt(id) {
  const debt = state.debts.find((item) => item.id === id);
  if (!debt) return;

  resetForm({ showTypeChoice: false, type: "debt" });
  if (!openEditForm({ type: "debt", id })) return;
  elements.debtId.value = debt.id;
  elements.entryId.value = "";
  elements.typeInput.value = "debt";
  elements.personInput.value = debt.personId || defaultPersonId();
  elements.creditorInput.value = debt.creditor;
  elements.debtTotalInput.value = formatMoneyInput(debt.totalAmount);
  elements.paidSoFarInput.value = formatMoneyInput(debt.paidSoFar || "");
  elements.debtPaymentInput.value = formatMoneyInput(debt.monthlyPayment);
  elements.debtStartInput.value = debt.startDate;
  elements.debtEndInput.value = debt.endDate;
  elements.debtStatusInput.value = normalizeStatus(debt.status);
  elements.debtPaymentMethodInput.value = debt.paymentMethod || "";
  elements.debtAccountInput.value = debt.account || "";
  elements.principalInput.value = formatMoneyInput(debt.principalAmount || "");
  elements.interestInput.value = formatMoneyInput(debt.interestAmount || "");
  elements.nominalRateInput.value = formatPercentInput(debt.nominalRate || "");
  elements.termInput.value = debt.termMonths || "";
  elements.finalPaymentInput.value = formatMoneyInput(debt.finalPayment || "");
  syncFormMode();
}

function deleteDebt(id) {
  state.debts = state.debts.filter((debt) => debt.id !== id);
  persist();
  render();
  if (activeEditContext?.kind === "debt" && activeEditContext.id === id) closeEditModal();
}

function deleteActiveEditItem() {
  if (!activeEditContext) return;
  const { kind, id } = activeEditContext;
  if (kind === "entry") {
    deleteEntry(id);
    showMessage("Eintrag gelöscht.");
  } else if (kind === "debt") {
    deleteDebt(id);
    showMessage("Schuld gelöscht.");
  } else if (kind === "asset") {
    deleteAsset(id);
  }
}

async function exportBackup() {
  const options = [
    { label: "Gesamt – alle Personen", value: "all", variant: "all" },
    ...state.people.map((person) => ({ label: person.name, value: person.id })),
  ];
  const choice = await chooseCentered(
    "Backup exportieren",
    "Wessen Daten möchtest du sichern?",
    options
  );
  if (choice === null) return;

  let payload;
  let suffix;
  if (choice === "all") {
    payload = state;
    suffix = "gesamt";
  } else {
    const person = state.people.find((p) => p.id === choice);
    const belongs = (item) => (item.personId || defaultPersonId()) === choice;
    payload = {
      categories: state.categories,
      entries: state.entries.filter(belongs),
      debts: state.debts.filter(belongs),
      assets: state.assets.filter(belongs),
      people: person ? [person] : [],
      selectedPersonId: choice,
    };
    suffix = (person?.name || "person").toLowerCase().replace(/[^a-z0-9-]+/g, "-").replace(/(^-|-$)/g, "") || "person";
  }

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `budget-backup-${suffix}-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  URL.revokeObjectURL(url);
  showMessage(choice === "all" ? "Backup (Gesamt) exportiert." : `Backup für ${state.people.find((p) => p.id === choice)?.name || "Person"} exportiert.`);
}

function importBackup(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.addEventListener("load", async () => {
    try {
      const lowerName = file.name.toLowerCase();
      if (lowerName.endsWith(".xlsx")) {
        importXlsx(reader.result);
      } else if (lowerName.endsWith(".csv")) {
        const text = String(reader.result);
        importCsv(text);
      } else {
        const text = String(reader.result);
        const imported = JSON.parse(text);
        if (!Array.isArray(imported.categories) || !Array.isArray(imported.entries)) return;
        const ok = await confirmCentered(
          "Daten wiederherstellen ersetzt deine aktuellen Daten. Vorher am besten Daten sichern.",
          { title: "Daten wiederherstellen" }
        );
        if (!ok) return;
        state.categories = imported.categories;
        state.entries = imported.entries.map(normalizeEntry);
        state.debts = Array.isArray(imported.debts) ? imported.debts.map(normalizeDebt) : [];
        state.assets = Array.isArray(imported.assets) ? imported.assets.map(normalizeAsset) : [];
        state.people = Array.isArray(imported.people) ? imported.people.map(normalizePerson) : state.people;
        showMessage("Backup importiert.");
      }
      persist();
      ensurePeople();
      repairCopiedAssets();
      render();
    } catch (error) {
      showMessage(`Import nicht erkannt: ${error.message || "unbekannter Fehler"}`);
    } finally {
      elements.importInput.value = "";
    }
  });
  if (file.name.toLowerCase().endsWith(".xlsx")) {
    reader.readAsArrayBuffer(file);
  } else {
    reader.readAsText(file);
  }
}

function importXlsx(arrayBuffer) {
  if (!window.XLSX) {
    showMessage("Excel-Import ist ohne XLSX-Bibliothek nicht verfügbar.", "error");
    return;
  }
  const workbook = window.XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = window.XLSX.utils.decode_range(firstSheet["!ref"]);
  const rows = [];
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = firstSheet[address];
      row.push({ value: cell ? cell.v : "", type: cell ? cell.t : "" });
    }
    rows.push(row);
  }
  importBudgetRows(rows);
}

function importCsv(text) {
  const rows = parseCsv(text);
  importBudgetRows(rows);
}

function importBudgetRows(rows) {
  if (rows.length < 2) {
    showMessage("Datei ist leer.");
    return;
  }

  if (looksLikePersonalBudget(rows)) {
    importPersonalBudgetLayout(rows);
    return;
  }

  const headers = rows[0].map((header) => normalizeHeader(cellValue(header)));
  let importedEntries = 0;
  let importedDebts = 0;
  let importedAssets = 0;
  rows.slice(1).forEach((row) => {
    const item = Object.fromEntries(headers.map((header, index) => [header, cellValue(row[index]) || ""]));
    const creditor = item["bei wem"] || item["gläubiger"] || item["glaeubiger"] || item["schuld"] || "";
    const totalDebt = parseMoneyInput(item["gesamtschuld"] || item["schuld gesamt"] || "");
    const monthlyDebt = parseMoneyInput(item["monatlich"] || item["monatsrate"] || "");

    const looksLikeDebt = Boolean(
      creditor ||
      item["gesamtschuld"] ||
      item["schuld gesamt"] ||
      item["monatsrate"] ||
      item["bei wem"] ||
      item["gläubiger"] ||
      item["glaeubiger"]
    );

    if (looksLikeDebt) {
      state.debts.push({
        id: createId(),
        personId: defaultPersonId(),
        creditor,
        totalAmount: Number.isFinite(totalDebt) ? totalDebt : 0,
        monthlyPayment: Number.isFinite(monthlyDebt) ? monthlyDebt : 0,
        startDate: normalizeDate(item["startdatum"] || item["start"] || ""),
        endDate: normalizeDate(item["enddatum"] || item["bis datum"] || item["bis"] || ""),
        status: normalizeStatus(item["status"] || ""),
        paymentMethod: item["zahlungsart"] || "",
        account: item["von konto"] || item["konto"] || "",
        note: "",
      });
      importedDebts += 1;
      return;
    }

    const amount = parseMoneyInput(item["betrag"] || item["summe"] || "");
    if (!Number.isFinite(amount) || amount <= 0) return;
    const typeText = (item["typ"] || item["art"] || "").toLowerCase();
      state.entries.push({
        id: createId(),
        personId: defaultPersonId(),
        type: typeText.includes("einnah") || typeText.includes("income") ? "income" : "expense",
        date: normalizeDate(item["datum"] || "") || monthToDate(elements.monthInput.value),
        category: item["kategorie"] || "",
        description: item["beschreibung"] || item["text"] || item["name"] || "",
        payment: item["zahlungsart"] || "",
        amount,
        recurring: /ja|true|monat/i.test(item["monatlich"] || ""),
        endDate: normalizeDate(item["enddatum"] || item["bis datum"] || item["bis"] || ""),
        status: normalizeStatus(item["status"] || ""),
      });
    importedEntries += 1;
  });
  showMessage(`${importedEntries} Buchungen, ${importedDebts} Schulden importiert.`);
}

function looksLikePersonalBudget(rows) {
  return rows.some((row) => row.some((cell) => String(cellValue(cell)).includes("Persönliches Monatsbudget"))) ||
    rows.some((row) => row.some((cell) => normalizeHeader(cellValue(cell)).includes("aktuelle schulden")));
}

function importPersonalBudgetLayout(rows) {
  let importedEntries = 0;
  let importedDebts = 0;
  let importedAssets = 0;
  const skipLabels = [
    "einnahmen",
    "wohnen",
    "betrag",
    "tatsächliche kosten",
    "spalte1",
    "monatlichen einkünfte",
    "ergebnis",
    "summe der tatsächlichen ausgaben",
    "tatsächlicher saldo",
    "aktuelle schulden",
    "gesamtschulden",
  ];

  rows.forEach((row) => {
    const leftName = String(cellValue(row[1]) || "").trim();
    const leftAmount = parseMoneyInput(cellValue(row[2]) || "");
    const rightName = String(cellValue(row[6]) || "").trim();
    const rightAmount = parseMoneyInput(cellValue(row[7]) || "");

    if (leftName && isExcelLikeNumber(row[2]) && Number.isFinite(leftAmount) && !skipLabels.includes(normalizeHeader(leftName))) {
      const isIncome = rowIndexBeforeLabel(rows, row, "WOHNEN");
      state.entries.push({
        id: createId(),
        personId: defaultPersonId(),
        type: isIncome ? "income" : "expense",
        date: monthToDate(elements.monthInput.value),
        category: isIncome ? "Einnahme" : "Ausgabe",
        description: leftName,
        payment: "",
        amount: Math.abs(leftAmount),
        recurring: true,
        endDate: "",
        status: "open",
      });
      importedEntries += 1;
    }

    if (rightName && isExcelLikeNumber(row[7]) && Number.isFinite(rightAmount) && !skipLabels.includes(normalizeHeader(rightName))) {
      if (rightAmount < 0) {
        state.debts.push({
          id: createId(),
          personId: defaultPersonId(),
          creditor: rightName,
          totalAmount: Math.abs(rightAmount),
          monthlyPayment: 0,
          startDate: "",
          endDate: "",
          status: "open",
          paymentMethod: "",
          account: "",
          note: "",
        });
        importedDebts += 1;
      } else {
        state.assets.push({
          id: createId(),
          personId: defaultPersonId(),
          name: rightName,
          amount: rightAmount,
          note: "aus Excel",
        });
        importedAssets += 1;
      }
    }
  });

  showMessage(`${importedEntries} Buchungen, ${importedDebts} Schulden, ${importedAssets} Vermögen importiert.`);
}

function rowIndexBeforeLabel(rows, targetRow, label) {
  const targetIndex = rows.indexOf(targetRow);
  const labelIndex = rows.findIndex((row) => row.some((cell) => String(cellValue(cell)).trim() === label));
  return labelIndex === -1 ? false : targetIndex < labelIndex;
}

function cellValue(cell) {
  return cell && typeof cell === "object" && Object.prototype.hasOwnProperty.call(cell, "value") ? cell.value : cell;
}

function cellType(cell) {
  return cell && typeof cell === "object" && Object.prototype.hasOwnProperty.call(cell, "type") ? cell.type : "";
}

function isExcelLikeNumber(cell) {
  const type = cellType(cell);
  const value = cellValue(cell);
  if (type === "n") return true;
  const text = String(value || "").trim();
  if (!text) return false;
  if (text.includes(",")) return true;
  return /^-?\d+$/.test(text.replace(/\s/g, "").replace(/€/g, ""));
}

function parseCsv(text) {
  const delimiter = text.includes(";") ? ";" : ",";
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function normalizeHeader(value) {
  return String(value).trim().toLowerCase();
}

function normalizeDate(value) {
  const text = String(value).trim();
  if (!text) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (!match) return "";
  return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
}

async function clearAllData() {
  const isAll = state.selectedPersonId === "all";
  const name = selectedPersonName();
  const message = isAll
    ? "Wirklich ALLES löschen?\nAlle Personen, alle Buchungen, Schulden und Vermögen werden entfernt."
    : `Möchtest du bei ${name} alles löschen?\nBuchungen, Schulden und Vermögen von ${name} werden entfernt.`;
  const title = isAll ? "Alles löschen" : `${name}: Alles löschen`;
  const ok = await confirmCentered(message, { title });
  if (!ok) return;

  if (isAll) {
    state.entries = [];
    state.debts = [];
    state.assets = [];
  } else {
    const pid = state.selectedPersonId;
    const belongs = (item) => (item.personId || defaultPersonId()) === pid;
    state.entries = state.entries.filter((entry) => !belongs(entry));
    state.debts = state.debts.filter((debt) => !belongs(debt));
    state.assets = state.assets.filter((asset) => !belongs(asset));
  }

  persist();
  resetForm();
  render();
  showMessage(isAll ? "Alles gelöscht." : `${name}: Alles gelöscht.`);
}

async function clearDebtsForContext() {
  const isAll = state.selectedPersonId === "all";
  const name = selectedPersonName();
  const message = isAll ? "Alle Schulden löschen?" : `Schulden von ${name} löschen?`;
  const ok = await confirmCentered(message, { title: "Schulden leeren" });
  if (!ok) return;
  if (isAll) {
    state.debts = [];
  } else {
    const pid = state.selectedPersonId;
    state.debts = state.debts.filter((debt) => (debt.personId || defaultPersonId()) !== pid);
  }
  persist();
  render();
  showMessage(isAll ? "Schulden geleert." : `${name}: Schulden geleert.`);
}

async function clearAssetsForContext() {
  const isAll = state.selectedPersonId === "all";
  const name = selectedPersonName();
  const message = isAll ? "Alle Vermögen löschen?" : `Vermögen von ${name} löschen?`;
  const ok = await confirmCentered(message, { title: "Vermögen leeren" });
  if (!ok) return;
  if (isAll) {
    state.assets = [];
  } else {
    const pid = state.selectedPersonId;
    state.assets = state.assets.filter((asset) => (asset.personId || defaultPersonId()) !== pid);
  }
  persist();
  render();
  showMessage(isAll ? "Vermögen geleert." : `${name}: Vermögen geleert.`);
}

async function clearEntriesForContext() {
  const isAll = state.selectedPersonId === "all";
  const name = selectedPersonName();
  const message = isAll ? "Alle Buchungen (Einnahmen + Ausgaben) löschen?" : `Buchungen von ${name} löschen?`;
  const ok = await confirmCentered(message, { title: "Buchungen leeren" });
  if (!ok) return;
  if (isAll) {
    state.entries = [];
  } else {
    const pid = state.selectedPersonId;
    state.entries = state.entries.filter((entry) => (entry.personId || defaultPersonId()) !== pid);
  }
  persist();
  render();
  showMessage(isAll ? "Buchungen geleert." : `${name}: Buchungen geleert.`);
}

let confirmDialogResolver = null;

function confirmCentered(message, { title = "Bestätigen", okLabel = "Ja, löschen", cancelLabel = "Abbrechen" } = {}) {
  return new Promise((resolve) => {
    if (confirmDialogResolver) {
      confirmDialogResolver(false);
    }
    confirmDialogResolver = resolve;
    elements.confirmDialogTitle.textContent = title;
    elements.confirmDialogMessage.textContent = message;
    elements.confirmDialogOk.textContent = okLabel;
    elements.confirmDialogCancel.textContent = cancelLabel;
    elements.confirmDialog.hidden = false;
    document.body.style.overflow = "hidden";
    window.setTimeout(() => elements.confirmDialogOk.focus(), 0);
  });
}

function resolveConfirmDialog(result) {
  if (!confirmDialogResolver) {
    elements.confirmDialog.hidden = true;
    document.body.style.overflow = "";
    return;
  }
  const resolver = confirmDialogResolver;
  confirmDialogResolver = null;
  elements.confirmDialog.hidden = true;
  document.body.style.overflow = "";
  resolver(result);
}

elements.confirmDialogOk.addEventListener("click", () => resolveConfirmDialog(true));
elements.confirmDialogCancel.addEventListener("click", () => resolveConfirmDialog(false));
elements.confirmDialogBackdrop.addEventListener("click", () => resolveConfirmDialog(false));
document.addEventListener("keydown", (event) => {
  if (elements.confirmDialog.hidden) return;
  if (event.key === "Escape") resolveConfirmDialog(false);
  if (event.key === "Enter") resolveConfirmDialog(true);
});

let choiceDialogResolver = null;

function chooseCentered(title, message, options) {
  return new Promise((resolve) => {
    if (choiceDialogResolver) choiceDialogResolver(null);
    choiceDialogResolver = resolve;
    elements.choiceDialogTitle.textContent = title;
    elements.choiceDialogMessage.textContent = message || "";
    elements.choiceDialogMessage.hidden = !message;
    elements.choiceDialogOptions.innerHTML = "";
    options.forEach((option) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `choice-dialog-option ${option.variant ? `choice-${option.variant}` : ""}`.trim();
      button.textContent = option.label;
      button.addEventListener("click", () => resolveChoiceDialog(option.value));
      elements.choiceDialogOptions.appendChild(button);
    });
    elements.choiceDialog.hidden = false;
    document.body.style.overflow = "hidden";
  });
}

function resolveChoiceDialog(value) {
  if (!choiceDialogResolver) {
    elements.choiceDialog.hidden = true;
    document.body.style.overflow = "";
    return;
  }
  const resolver = choiceDialogResolver;
  choiceDialogResolver = null;
  elements.choiceDialog.hidden = true;
  document.body.style.overflow = "";
  resolver(value);
}

elements.choiceDialogCancel.addEventListener("click", () => resolveChoiceDialog(null));
elements.choiceDialogBackdrop.addEventListener("click", () => resolveChoiceDialog(null));
document.addEventListener("keydown", (event) => {
  if (elements.choiceDialog.hidden) return;
  if (event.key === "Escape") resolveChoiceDialog(null);
});

let promptDialogResolver = null;

function promptCentered(title, message, defaultValue = "", { okLabel = "OK", cancelLabel = "Abbrechen", placeholder = "" } = {}) {
  return new Promise((resolve) => {
    if (promptDialogResolver) promptDialogResolver(null);
    promptDialogResolver = resolve;
    elements.promptDialogTitle.textContent = title;
    elements.promptDialogMessage.textContent = message || "";
    elements.promptDialogMessage.hidden = !message;
    elements.promptDialogInput.value = defaultValue;
    elements.promptDialogInput.placeholder = placeholder;
    elements.promptDialogOk.textContent = okLabel;
    elements.promptDialogCancel.textContent = cancelLabel;
    elements.promptDialog.hidden = false;
    document.body.style.overflow = "hidden";
    window.setTimeout(() => {
      elements.promptDialogInput.focus();
      elements.promptDialogInput.select();
    }, 30);
  });
}

function resolvePromptDialog(value) {
  if (!promptDialogResolver) {
    elements.promptDialog.hidden = true;
    document.body.style.overflow = "";
    return;
  }
  const resolver = promptDialogResolver;
  promptDialogResolver = null;
  elements.promptDialog.hidden = true;
  document.body.style.overflow = "";
  resolver(value);
}

elements.promptDialogOk.addEventListener("click", () => resolvePromptDialog(elements.promptDialogInput.value));
elements.promptDialogCancel.addEventListener("click", () => resolvePromptDialog(null));
elements.promptDialogBackdrop.addEventListener("click", () => resolvePromptDialog(null));
elements.promptDialogInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    resolvePromptDialog(elements.promptDialogInput.value);
  } else if (event.key === "Escape") {
    event.preventDefault();
    resolvePromptDialog(null);
  }
});

let contextMenuPersonId = null;
let longPressTimer = null;

function showPersonContextMenu(clientX, clientY, personId) {
  contextMenuPersonId = personId;
  const menu = elements.personContextMenu;
  menu.hidden = false;
  menu.style.left = "0px";
  menu.style.top = "0px";
  const rect = menu.getBoundingClientRect();
  const x = Math.min(clientX, window.innerWidth - rect.width - 8);
  const y = Math.min(clientY, window.innerHeight - rect.height - 8);
  menu.style.left = `${Math.max(8, x)}px`;
  menu.style.top = `${Math.max(8, y)}px`;
}

function hidePersonContextMenu() {
  contextMenuPersonId = null;
  elements.personContextMenu.hidden = true;
}

document.addEventListener("click", (event) => {
  if (!elements.personContextMenu.hidden && !elements.personContextMenu.contains(event.target)) {
    hidePersonContextMenu();
  }
});
document.addEventListener("keydown", (event) => {
  if (!elements.personContextMenu.hidden && event.key === "Escape") hidePersonContextMenu();
});
window.addEventListener("scroll", () => {
  if (!elements.personContextMenu.hidden) hidePersonContextMenu();
}, true);
window.addEventListener("resize", () => {
  if (!elements.personContextMenu.hidden) hidePersonContextMenu();
});

elements.personContextMenu.addEventListener("click", (event) => {
  const button = event.target.closest("[data-action]");
  if (!button) return;
  const action = button.dataset.action;
  const personId = contextMenuPersonId;
  hidePersonContextMenu();
  if (!personId) return;
  if (action === "duplicate") {
    if (personId === "all") duplicateAllAsPerson();
    else duplicatePerson(personId);
  } else if (action === "rename") renamePersonCentered(personId);
  else if (action === "delete") deletePerson(personId);
});

function updateContextMenuItems(personId) {
  const isAll = personId === "all";
  elements.personContextMenu.querySelectorAll("[data-action]").forEach((button) => {
    if (button.dataset.action === "duplicate") {
      button.hidden = false;
      button.textContent = isAll ? "Alles als neue Person duplizieren" : "Person duplizieren";
    } else {
      button.hidden = isAll;
    }
  });
}

elements.personSummaryList.addEventListener("contextmenu", (event) => {
  const card = event.target.closest("[data-person-select]");
  if (!card) return;
  const id = card.dataset.personSelect;
  event.preventDefault();
  updateContextMenuItems(id);
  showPersonContextMenu(event.clientX, event.clientY, id);
});

elements.personSummaryList.addEventListener("touchstart", (event) => {
  const card = event.target.closest("[data-person-select]");
  if (!card) return;
  const id = card.dataset.personSelect;
  const touch = event.touches[0];
  if (longPressTimer) clearTimeout(longPressTimer);
  longPressTimer = window.setTimeout(() => {
    updateContextMenuItems(id);
    showPersonContextMenu(touch.clientX, touch.clientY, id);
  }, 550);
}, { passive: true });

["touchend", "touchmove", "touchcancel"].forEach((evt) => {
  elements.personSummaryList.addEventListener(evt, () => {
    if (longPressTimer) {
      clearTimeout(longPressTimer);
      longPressTimer = null;
    }
  });
});

async function duplicateAllAsPerson() {
  const name = await promptCentered(
    "Alles als neue Person duplizieren",
    "Alle Einnahmen, Ausgaben, Schulden und Vermögen werden auf eine neue Person übernommen. Diese Daten zählen nicht erneut zur Gesamtsumme.",
    "Neue Person",
    { okLabel: "Duplizieren", placeholder: "Neuer Name" }
  );
  if (name === null) return;
  const trimmed = cleanCopyLabel(name);
  if (!trimmed) {
    showMessage("Name darf nicht leer sein.");
    return;
  }

  const newPerson = { id: createId(), name: trimmed };
  state.people.push(newPerson);

  const now = Date.now();
  let counter = 0;
  state.entries.slice().forEach((entry) => {
    if (entry.duplicateOf) return;
    state.entries.push({ ...entry, id: createId(), personId: newPerson.id, updatedAt: now + counter, duplicateOf: entry.id });
    counter += 1;
  });
  state.debts.slice().forEach((debt) => {
    if (debt.duplicateOf) return;
    state.debts.push({ ...debt, id: createId(), personId: newPerson.id, duplicateOf: debt.id });
  });
  state.assets.slice().forEach((asset) => {
    if (asset.duplicateOf) return;
    state.assets.push({ ...asset, id: createId(), personId: newPerson.id, duplicateOf: asset.id });
  });

  state.selectedPersonId = newPerson.id;
  persist();
  render();
  showMessage(`Gesamt → ${newPerson.name}: dupliziert.`);
}

async function duplicatePerson(sourceId) {
  const source = state.people.find((p) => p.id === sourceId);
  if (!source) return;
  const suggested = cleanCopyLabel(source.name);
  const name = await promptCentered(
    "Person duplizieren",
    `Alle Einnahmen, Ausgaben, Schulden und Vermögen von ${source.name} werden auf eine neue Person übernommen. Diese Daten zählen nur in der neuen Person, nicht erneut zur Gesamtsumme.`,
    suggested,
    { okLabel: "Duplizieren", placeholder: "Neuer Name" }
  );
  if (name === null) return;
  const trimmed = cleanCopyLabel(name);
  if (!trimmed) {
    showMessage("Name darf nicht leer sein.");
    return;
  }

  const newPerson = { id: createId(), name: trimmed };
  state.people.push(newPerson);

  const now = Date.now();
  let counter = 0;
  state.entries
  .filter((entry) => (entry.personId || defaultPersonId()) === sourceId)
  .forEach((entry) => {
    state.entries.push({
      ...entry,
      id: createId(),
      personId: newPerson.id,
      updatedAt: now + counter,
      duplicateOf: entry.duplicateOf || entry.id,
    });
    counter += 1;
  });

state.debts
  .filter((debt) => (debt.personId || defaultPersonId()) === sourceId)
  .forEach((debt) => {
    state.debts.push({
      ...debt,
      id: createId(),
      personId: newPerson.id,
      duplicateOf: debt.duplicateOf || debt.id,
    });
  });

state.assets
  .filter((asset) => (asset.personId || defaultPersonId()) === sourceId)
  .forEach((asset) => {
    state.assets.push({
      ...asset,
      id: createId(),
      personId: newPerson.id,
      duplicateOf: asset.duplicateOf || asset.id,
    });
  });

  state.selectedPersonId = newPerson.id;
  persist();
  render();
  showMessage(`${source.name} → ${newPerson.name}: dupliziert.`);
}

async function renamePersonCentered(id) {
  const person = state.people.find((p) => p.id === id);
  if (!person) return;
  const next = await promptCentered("Person umbenennen", "Neuer Name:", person.name, { placeholder: "Name" });
  if (next === null) return;
  const trimmed = next.trim();
  if (!trimmed) {
    showMessage("Name darf nicht leer sein.");
    return;
  }
  person.name = trimmed;
  persist();
  render();
  showMessage("Person geändert.");
}

function showMessage(message, type = "info") {
  elements.formMessage.textContent = message;
  elements.formMessage.classList.toggle("success", type === "success");
  elements.formMessage.classList.toggle("error", type === "error");
  window.setTimeout(() => {
    if (elements.formMessage.textContent === message) {
      elements.formMessage.textContent = "";
      elements.formMessage.classList.remove("success", "error");
    }
  }, 2500);
}

function setFieldError(input, message) {
  const label = input?.closest("label");
  if (label) label.classList.add("field-error");
  input?.classList.add("input-error");
  input?.setAttribute("aria-invalid", "true");
  showMessage(message, "error");
  input?.focus({ preventScroll: true });
}

function clearFieldError(target) {
  if (!target || !target.closest) return;
  target.closest("label")?.classList.remove("field-error");
  target.classList?.remove("input-error");
  target.removeAttribute?.("aria-invalid");
}

function clearFieldErrors() {
  [elements.entryForm, elements.assetForm].forEach((form) => {
    form.querySelectorAll(".field-error").forEach((label) => label.classList.remove("field-error"));
    form.querySelectorAll(".input-error, [aria-invalid='true']").forEach((input) => {
      input.classList.remove("input-error");
      input.removeAttribute("aria-invalid");
    });
  });
}

function parseMoneyInput(value) {
  const cleaned = String(value)
    .trim()
    .replace(/\s/g, "")
    .replace(/€/g, "");
  if (!cleaned) return NaN;

  if (cleaned.includes(",")) {
    return Number(cleaned.replace(/\./g, "").replace(",", "."));
  }

  if (/^-?\d+\.\d{1,2}$/.test(cleaned)) {
    return Number(cleaned);
  }

  return Number(cleaned.replace(/\./g, ""));
}

function formatMoneyInput(value) {
  const number = typeof value === "number" ? value : parseMoneyInput(value);
  if (!Number.isFinite(number)) return "";
  return new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(number);
}

function formatMoneyField(input) {
  input.value = formatMoneyInput(input.value);
}

function parsePercentInput(value) {
  const cleaned = String(value).trim().replace("%", "").replace(",", ".");
  if (!cleaned) return NaN;
  return Number(cleaned);
}

function formatPercentInput(value) {
  const number = typeof value === "number" ? value : parsePercentInput(value);
  if (!Number.isFinite(number)) return "";
  return new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(number);
}

function formatPercentOutput(value) {
  const number = typeof value === "number" ? value : parsePercentInput(value);
  if (!Number.isFinite(number)) return "";
  return `${formatPercentInput(number)} %`;
}

function sum(entries) {
  return entries.reduce((total, entry) => total + Number(entry.amount || 0), 0);
}

function formatDate(value) {
  return new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit" }).format(new Date(`${value}T12:00:00`));
}

function remainingDebt(debt, month) {
  if (!isDebtVisibleInMonth(debt, month) || isDebtClosed(debt)) return 0;
  return Math.max(0, Number(debt.totalAmount || 0) - paidDebt(debt, month));
}

function paidDebt(debt, month) {
  if (monthToNumber(month) < debtStartMonthNumber(debt)) return 0;
  const paidMonths = paidMonthsUntil(debt, month);
  const paidAmount = paidMonths * Number(debt.monthlyPayment || 0);
  return Math.min(Number(debt.totalAmount || 0), Number(debt.paidSoFar || 0) + paidAmount);
}

function paidMonthsUntil(debt, month) {
  if (!debt.startDate || monthToNumber(month) < dateToMonthNumber(debt.startDate)) return 0;
  const cappedMonth = debt.endDate && monthToNumber(month) > dateToMonthNumber(debt.endDate) ? debt.endDate.slice(0, 7) : month;
  return Math.max(0, monthsBetween(debt.startDate.slice(0, 7), cappedMonth) + 1);
}

function isDebtActiveInMonth(debt, month) {
  return isDebtVisibleInMonth(debt, month) && !isDebtClosed(debt) && remainingDebt(debt, month) > 0;
}

function isDebtVisibleInMonth(debt, month) {
  const selected = monthToNumber(month);
  const start = debtStartMonthNumber(debt);
  const end = debt.endDate ? dateToMonthNumber(debt.endDate) : Infinity;
  if (selected < start || selected > end) return false;
  if (!debt.endDate && isDebtClosed(debt)) return selected === start;
  return true;
}

function isDebtClosed(debt) {
  const total = Number(debt.totalAmount || 0);
  return normalizeStatus(debt.status) !== "open" || (total > 0 && Number(debt.paidSoFar || 0) >= total);
}

function debtStartMonthNumber(debt) {
  return debt.startDate ? dateToMonthNumber(debt.startDate) : -Infinity;
}

function isEntryActiveInMonth(entry, month) {
  const entryMonth = entry.date?.slice(0, 7);
  if (!entryMonth) return false;
  if (!entry.recurring) return entryMonth === month;
  if (normalizeStatus(entry.status) === "ended" && !entry.endDate) return entryMonth === month;
  const selected = monthToNumber(month);
  const start = monthToNumber(entryMonth);
  const end = entry.endDate ? dateToMonthNumber(entry.endDate) : Infinity;
  return selected >= start && selected <= end;
}

function entryOccurrenceDate(entry, month) {
  if (!entry.recurring || entry.date?.startsWith(month)) return entry.date;
  const day = Math.min(Number(entry.date?.slice(8, 10)) || 1, daysInMonth(month));
  return `${month}-${String(day).padStart(2, "0")}`;
}

function entryBadges(entry, month) {
  return [
    statusBadge(statusLabel(entry.status), `status-${normalizeStatus(entry.status)}`),
    entry.recurring ? statusBadge("wiederkehrend", "status-recurring") : "",
    entry.recurring && entry.endDate ? statusBadge(`endet ${formatMonthName(entry.endDate.slice(0, 7))}`, "status-ended") : "",
    entry.recurring && !entry.date.startsWith(month) ? statusBadge("übernommen", "status-carry") : "",
  ].filter(Boolean).join("");
}

function debtBadges(debt, month, active) {
  const status = isDebtClosed(debt) || remainingDebt(debt, month) <= 0 ? (normalizeStatus(debt.status) === "ended" ? "ended" : "paid") : "open";
  return [
    statusBadge(statusLabel(status), `status-${status}`),
    active ? statusBadge("offen im Monat", "status-carry") : "",
    debt.endDate ? statusBadge(`endet ${formatMonthName(debt.endDate.slice(0, 7))}`, "status-ended") : statusBadge("ohne Enddatum", "status-recurring"),
  ].filter(Boolean).join("");
}

function statusBadge(label, className) {
  return `<span class="status-badge ${escapeHtml(className)}">${escapeHtml(label)}</span>`;
}

function monthsBetween(startMonth, endMonth) {
  const [startYear, start] = startMonth.split("-").map(Number);
  const [endYear, end] = endMonth.split("-").map(Number);
  return (endYear - startYear) * 12 + (end - start);
}

function monthsFromDates(startDate, endDate) {
  if (!startDate || !endDate) return 0;
  return Math.max(0, monthsBetween(startDate.slice(0, 7), endDate.slice(0, 7)) + 1);
}

function monthToNumber(month) {
  const [year, value] = month.split("-").map(Number);
  return year * 12 + value;
}

function dateToMonthNumber(date) {
  return monthToNumber(date.slice(0, 7));
}

function monthToDate(month) {
  return `${month}-01`;
}

function daysInMonth(month) {
  const [year, value] = month.split("-").map(Number);
  return new Date(year, value, 0).getDate();
}

function daysLeftInMonth(month) {
  const [year, value] = month.split("-").map(Number);
  const now = new Date();
  const isCurrentMonth = now.getFullYear() === year && now.getMonth() + 1 === value;
  return isCurrentMonth ? Math.max(1, daysInMonth(month) - now.getDate() + 1) : daysInMonth(month);
}

function formatMonthName(month) {
  return new Intl.DateTimeFormat("de-DE", { month: "long", year: "numeric" }).format(new Date(`${month}-01T12:00:00`));
}

function carryoverItemsForNextMonth(month) {
  const nextMonth = shiftMonth(month, 1);
  const recurring = state.entries
    .filter(includeInSelectedTotal)
    .filter((entry) => entry.recurring && isEntryActiveInMonth(entry, nextMonth))
    .map((entry) => ({ label: entry.description || entry.category || "Buchung" }));
  const debts = state.debts
    .filter(includeInSelectedTotal)
    .filter((debt) => isDebtActiveInMonth(debt, nextMonth))
    .map((debt) => ({ label: debt.creditor || "Schuld" }));
  return [...recurring, ...debts];
}

function addMonths(month, amount) {
  const [year, value] = month.split("-").map(Number);
  const date = new Date(year, value - 1 + amount, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function currentLocalMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

function defaultBudgetForCategory(name) {
  const budgets = {
    Wohnen: 950,
    Lebensmittel: 350,
    Transport: 80,
    Auto: 160,
    Versicherung: 120,
    Gesundheit: 80,
    Freizeit: 150,
    Reisen: 100,
    Abos: 60,
    Schulden: 0,
    Einkommen: 0,
    Sonstiges: 100,
  };
  return budgets[name] || 0;
}

function uniqueCategoryNames(names) {
  const seen = new Set();
  return names
    .map((name) => String(name || "").trim())
    .filter((name) => {
      const key = name.toLowerCase();
      if (!name || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function sameCategory(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function resolveCategoryInput(type) {
  const custom = elements.customCategoryInput.value.trim();
  const selected = elements.categoryInput.value.trim();
  const fallback = type === "income" ? "Einkommen" : "Sonstiges";
  const categoryName = custom || selected || fallback;
  if (!state.categories.some((category) => sameCategory(category.name, categoryName))) {
    state.categories.push({ id: createId(), name: categoryName, budget: 0 });
  }
  return categoryName;
}

function normalizeStatus(value) {
  const status = String(value || "").trim().toLowerCase();
  if (status === "bezahlt" || status === "paid") return "paid";
  if (status === "beendet" || status === "erledigt" || status === "ended" || status === "done") return "ended";
  return ENTRY_STATUSES.includes(status) ? status : "open";
}

function statusLabel(status) {
  const labels = {
    open: "offen",
    paid: "bezahlt",
    ended: "beendet",
  };
  return labels[normalizeStatus(status)] || "offen";
}

function normalizeDebt(debt) {
  return {
    id: debt.id || createId(),
    personId: debt.personId || fallbackPersonId(),
    duplicateOf: debt.duplicateOf || "",
    creditor: debt.creditor || "",
    totalAmount: Number(debt.totalAmount || 0),
    paidSoFar: Number(debt.paidSoFar || 0),
    monthlyPayment: Number(debt.monthlyPayment || 0),
    startDate: debt.startDate || (debt.startMonth ? monthToDate(debt.startMonth) : ""),
    endDate: debt.endDate || (debt.endMonth ? monthToDate(debt.endMonth) : ""),
    status: normalizeStatus(debt.status || (Number(debt.paidSoFar || 0) >= Number(debt.totalAmount || 0) && Number(debt.totalAmount || 0) > 0 ? "paid" : "open")),
    paymentMethod: debt.paymentMethod || "",
    account: debt.account || "",
    principalAmount: Number(debt.principalAmount || 0),
    interestAmount: Number(debt.interestAmount || 0),
    nominalRate: Number(debt.nominalRate || 0),
    termMonths: Number(debt.termMonths || 0),
    finalPayment: Number(debt.finalPayment || 0),
    note: debt.note || "",
  };
}

function normalizeAsset(asset) {
  return {
    id: asset.id || createId(),
    personId: asset.personId || fallbackPersonId(),
    duplicateOf: asset.duplicateOf || "",
    name: asset.name || "",
    amount: Number(asset.amount || 0),
    note: asset.note || "",
  };
}

function normalizeEntry(entry) {
  const date = entry.date || monthToDate(currentLocalMonth());
  const fallbackUpdatedAt = Date.parse(`${date}T12:00:00`) || 0;
  return {
    id: entry.id || createId(),
    personId: entry.personId || fallbackPersonId(),
    duplicateOf: entry.duplicateOf || "",
    type: entry.type === "income" ? "income" : "expense",
    date,
    category: entry.category || "",
    description: entry.description || "",
    payment: entry.payment || "",
    amount: Number(entry.amount || 0),
    recurring: Boolean(entry.recurring),
    endDate: entry.endDate || (entry.endMonth ? monthToDate(entry.endMonth) : ""),
    status: normalizeStatus(entry.status),
    updatedAt: Number(entry.updatedAt) > 0 ? Number(entry.updatedAt) : fallbackUpdatedAt,
  };
}

function normalizePerson(person) {
  return {
    id: person.id || createId(),
    name: cleanCopyLabel(person.name || "Gemeinsam") || "Gemeinsam",
  };
}

function personName(id) {
  if (!id || id === defaultPersonId()) return "Gesamt";
  return state.people.find((person) => person.id === id)?.name || state.people[0]?.name || "Gemeinsam";
}

function visiblePeople() {
  const hiddenId = defaultPersonId();
  return state.people.filter((person) => person.id !== hiddenId);
}

function selectedPersonName() {
  if (state.selectedPersonId === "all") return "Gesamt";
  return personName(state.selectedPersonId);
}

function matchesSelectedPerson(personId) {
  if (state.selectedPersonId === "all") return true;
  return (personId || defaultPersonId()) === state.selectedPersonId;
}

function includeInSelectedTotal(item) {
  if (!matchesSelectedPerson(item.personId)) return false;
  if (state.selectedPersonId === "all" && item.duplicateOf) return false;
  return true;
}

function selectedMonthlyTotals(month) {
  if (state.selectedPersonId === "all") return monthlyTotalsForFamily(month);
  return monthlyTotalsForPerson(state.selectedPersonId, month);
}

function selectedDebtTotal(month) {
  return state.debts
    .filter(includeInSelectedTotal)
    .reduce((total, debt) => total + remainingDebt(debt, month), 0);
}

function displayText(value) {
  return String(value || "")
    .replaceAll("Ueberweisung", "Überweisung")
    .replaceAll("Vermoegen", "Vermögen")
    .replaceAll("Loeschen", "Löschen")
    .replaceAll("fuer", "für")
    .replaceAll("ueber", "über");
}

function formatDateFull(value) {
  return new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" }).format(new Date(`${value}T12:00:00`));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function createId() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}
