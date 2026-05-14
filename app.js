const STORAGE_KEY = "budget-app-v2";

const defaults = {
  categories: [
    { id: createId(), name: "Miete", budget: 950 },
    { id: createId(), name: "Lebensmittel", budget: 350 },
    { id: createId(), name: "Transport", budget: 80 },
    { id: createId(), name: "Versicherung", budget: 120 },
    { id: createId(), name: "Freizeit", budget: 150 },
    { id: createId(), name: "Schulden", budget: 0 },
    { id: createId(), name: "Sonstiges", budget: 100 },
  ],
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
ensureDebtCategory();

const elements = {
  monthInput: document.querySelector("#monthInput"),
  entryForm: document.querySelector("#entryForm"),
  mainFormSlot: document.querySelector("#mainFormSlot"),
  debtFormSlot: document.querySelector("#debtFormSlot"),
  entryFormSlot: document.querySelector("#entryFormSlot"),
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
  amountInput: document.querySelector("#amountInput"),
  descriptionInput: document.querySelector("#descriptionInput"),
  paymentInput: document.querySelector("#paymentInput"),
  recurringInput: document.querySelector("#recurringInput"),
  resetFormButton: document.querySelector("#resetFormButton"),
  debtId: document.querySelector("#debtId"),
  creditorInput: document.querySelector("#creditorInput"),
  debtTotalInput: document.querySelector("#debtTotalInput"),
  paidSoFarInput: document.querySelector("#paidSoFarInput"),
  debtPaymentInput: document.querySelector("#debtPaymentInput"),
  debtStartInput: document.querySelector("#debtStartInput"),
  debtEndInput: document.querySelector("#debtEndInput"),
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
  filterInput: document.querySelector("#filterInput"),
  entryList: document.querySelector("#entryList"),
  exportButton: document.querySelector("#exportButton"),
  importInput: document.querySelector("#importInput"),
  clearButton: document.querySelector("#clearButton"),
  familyIncomeValue: document.querySelector("#familyIncomeValue"),
  familyExpenseValue: document.querySelector("#familyExpenseValue"),
  remainingValue: document.querySelector("#remainingValue"),
  totalDebtValue: document.querySelector("#totalDebtValue"),
  totalAssetValue: document.querySelector("#totalAssetValue"),
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
};

let formTypeChoiceVisible = true;
let activeSummaryType = "";
let activeFormKey = "";

elements.monthInput.value = new Date().toISOString().slice(0, 7);
elements.typeInput.value = "expense";

elements.entryForm.addEventListener("submit", saveEntry);
elements.entryForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.entryForm.addEventListener("change", (event) => clearFieldError(event.target));
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
  addAsset();
});
elements.addPersonButton.addEventListener("click", addPerson);
elements.personNameInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    addPerson();
  }
});
elements.typeInput.addEventListener("change", syncFormMode);
elements.addCategoryButton.addEventListener("click", addCategory);
elements.filterInput.addEventListener("change", render);
elements.monthInput.addEventListener("change", render);
elements.summaryCards.forEach((card) => {
  card.addEventListener("click", () => toggleSummaryBreakdown(card.dataset.summary));
});
elements.summaryBreakdownClose.addEventListener("click", () => closeSummaryBreakdown());
elements.exportButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
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
  clearEntriesForContext();
});
document.querySelectorAll(".collapsible-panel .list-actions").forEach((node) => {
  node.addEventListener("click", (event) => event.stopPropagation());
});
elements.filterInput.addEventListener("click", (event) => event.stopPropagation());
[elements.amountInput, elements.debtTotalInput, elements.paidSoFarInput, elements.debtPaymentInput, elements.principalInput, elements.interestInput, elements.finalPaymentInput].forEach((input) => {
  input.addEventListener("blur", () => formatMoneyField(input));
  input.addEventListener("focus", () => input.select());
});
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

function ensurePeople() {
  let changed = false;
  if (!Array.isArray(state.people) || !state.people.length) {
    state.people = clone(defaults.people);
    changed = true;
  }
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

function ensureDebtCategory() {
  if (!state.categories.some((category) => category.name.toLowerCase() === "schulden")) {
    state.categories.push({ id: createId(), name: "Schulden", budget: 0 });
    persist();
  }
}

function openForm({ showTypeChoice = true, slot = elements.mainFormSlot, key = "main" } = {}) {
  moveFormTo(slot);
  activeFormKey = key;
  elements.entryForm.hidden = false;
  elements.toggleFormButton.textContent = "× Formular schließen";
  setTypeChoiceVisible(showTypeChoice);
  fillPersonSelect();
  syncFormMode();
}

function closeForm() {
  elements.entryForm.hidden = true;
  elements.toggleFormButton.textContent = "+ Neuer Eintrag";
  activeFormKey = "";
  clearFieldErrors();
}

function toggleForm() {
  if (elements.entryForm.hidden || activeFormKey !== "main") {
    resetForm({ showTypeChoice: true, type: "expense" });
    openForm({ showTypeChoice: true, slot: elements.mainFormSlot, key: "main" });
  } else {
    closeForm();
  }
}

function openTypedForm(type) {
  const key = `typed:${type}`;
  if (!elements.entryForm.hidden && activeFormKey === key) {
    closeForm();
    return;
  }
  resetForm({ showTypeChoice: false, type });
  openForm({ showTypeChoice: false, slot: formSlotForType(type), key });
  fillPersonSelect();
  elements.personInput.value = targetPersonIdForNewItem();
  if (type === "income") elements.categoryInput.value = "Gehalt";
  syncFormMode();
}

function targetPersonIdForNewItem() {
  return state.selectedPersonId === "all" ? defaultPersonId() : state.selectedPersonId;
}

function formSlotForType(type) {
  return type === "debt" ? elements.debtFormSlot : elements.entryFormSlot;
}

function moveFormTo(slot) {
  if (!slot) return;
  slot.after(elements.entryForm);
}

function setTypeChoiceVisible(visible) {
  formTypeChoiceVisible = visible;
  elements.typeField.hidden = !visible;
  updateFormTitle();
}

function updateFormTitle() {
  if (formTypeChoiceVisible) {
    elements.formTitle.textContent = "Einfach eintragen";
    return;
  }

  const titles = {
    debt: "Schuld eintragen",
    income: "Einnahme eintragen",
    expense: "Ausgabe eintragen",
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

  const entry = {
    id: elements.entryId.value || createId(),
    personId: elements.personInput.value || defaultPersonId(),
    type: elements.typeInput.value,
    date: elements.dateInput.value || monthToDate(elements.monthInput.value),
    category: elements.categoryInput.value || (elements.typeInput.value === "income" ? "Einnahme" : "Ausgabe"),
    description: elements.descriptionInput.value.trim() || (elements.typeInput.value === "income" ? "Einnahme" : "Ausgabe"),
    payment: elements.paymentInput.value,
    amount,
    recurring: elements.recurringInput.checked,
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
  closeForm();
}

function resetForm({ showTypeChoice = formTypeChoiceVisible, type } = {}) {
  const nextType = type || (showTypeChoice ? "expense" : elements.typeInput.value || "expense");
  elements.entryForm.reset();
  elements.entryId.value = "";
  elements.debtId.value = "";
  elements.typeInput.value = nextType;
  elements.paymentInput.value = "Karte";
  elements.recurringInput.checked = false;
  elements.debtPaymentMethodInput.value = "";
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
  closeForm();
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
  const entries = state.entries.filter((entry) => (entry.personId || defaultPersonId()) === personId && entry.date.startsWith(month));
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
    .filter((entry) => !entry.duplicateOf && entry.date.startsWith(month) && entry.type === "income")
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
  const entryExpense = state.entries
    .filter((entry) => !entry.duplicateOf && entry.date.startsWith(month) && entry.type === "expense")
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
  elements.debtPaymentMethodInput.required = false;
  elements.principalInput.required = false;
  elements.interestInput.required = false;
  elements.nominalRateInput.required = false;
  elements.termInput.required = false;
  elements.finalPaymentInput.required = false;

  if (isIncome && elements.categoryInput.value !== "Gehalt") {
    elements.categoryInput.value = "Gehalt";
  }
  updateFormTitle();
}

function fillCategorySelect() {
  const existing = elements.categoryInput.value;
  const names = [...state.categories.map((category) => category.name), "Gehalt"];
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
    .filter((entry) => entry.date.startsWith(month))
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
    const meta = [personName(entry.personId), entry.category, displayText(entry.payment), entry.recurring ? "monatlich" : ""].filter(Boolean).join(" · ");
    return `
      <article class="entry-row ${entry.type}">
        <time class="entry-date" datetime="${entry.date}">${formatDate(entry.date)}</time>
        <div class="entry-main">
          <strong>${escapeHtml(title)}</strong>
          <span>${escapeHtml(meta)}</span>
        </div>
        <div class="entry-amount ${entry.type}">${sign}${formatMoney.format(entry.amount)}</div>
        <div class="entry-controls">
          <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-edit="${entry.id}">✎</button>
          <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-delete="${entry.id}">×</button>
        </div>
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
          <strong>${escapeHtml(debt.creditor)}</strong>
          <span>${escapeHtml(meta)}</span>
        </div>
        <div class="debt-numbers">
          <span>Gesamt ${formatMoney.format(debt.totalAmount || 0)}</span>
          ${debt.paidSoFar ? `<span class="paid">Bisher bezahlt ${formatMoney.format(debt.paidSoFar)}</span>` : ""}
          <span>Rate ${formatMoney.format(debt.monthlyPayment)}</span>
          ${debt.interestAmount ? `<span>Zinsen ${formatMoney.format(debt.interestAmount)}</span>` : ""}
          ${debt.nominalRate ? `<span>Sollzins ${formatPercentOutput(debt.nominalRate)}</span>` : ""}
          ${term ? `<span>Laufzeit ${term} Monate</span>` : ""}
          <span class="paid">Gezahlt ${formatMoney.format(paid)}</span>
          <strong class="remaining">Rest ${formatMoney.format(remaining)}</strong>
        </div>
        <div class="entry-controls">
          <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-debt-edit="${debt.id}">✎</button>
          <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-debt-delete="${debt.id}">×</button>
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
        <strong>${escapeHtml(asset.name)}</strong>
        <span>${escapeHtml([personName(asset.personId), asset.note || "Plus-Wert"].filter(Boolean).join(" · "))}</span>
      </div>
      <strong class="asset-amount">${formatMoney.format(asset.amount || 0)}</strong>
      <div class="entry-controls">
        <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-asset-edit="${asset.id}">✎</button>
        <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-asset-delete="${asset.id}">×</button>
      </div>
    </article>
  `).join("");

  elements.assetList.querySelectorAll("[data-asset-edit]").forEach((button) => {
    button.addEventListener("click", () => editAsset(button.dataset.assetEdit));
  });
  elements.assetList.querySelectorAll("[data-asset-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteAsset(button.dataset.assetDelete));
  });
}

function addAsset() {
  const personId = targetPersonIdForNewItem();
  const owner = personName(personId);
  const nextName = window.prompt(`Vermögen für ${owner}: Name`, "");
  if (nextName === null) return;
  const name = nextName.trim();
  if (!name) {
    showMessage("Bitte Namen eingeben.");
    return;
  }

  const nextAmount = window.prompt("Betrag", "");
  if (nextAmount === null) return;
  const amount = parseMoneyInput(nextAmount);
  if (!Number.isFinite(amount) || amount < 0) {
    showMessage("Betrag ist keine gültige Zahl.");
    return;
  }

  state.assets.push({
    id: createId(),
    personId,
    name,
    amount,
    note: "manuell",
  });
  persist();
  render();
  showMessage("Vermögen gespeichert.");
}

function editAsset(id) {
  const asset = state.assets.find((item) => item.id === id);
  if (!asset) return;
  const nextName = window.prompt("Name ändern", asset.name);
  if (nextName === null) return;
  const name = nextName.trim();
  if (!name) {
    showMessage("Name darf nicht leer sein.");
    return;
  }

  const nextAmount = window.prompt("Betrag ändern", formatMoneyInput(asset.amount));
  if (nextAmount === null) return;
  const amount = parseMoneyInput(nextAmount);
  if (!Number.isFinite(amount) || amount < 0) {
    showMessage("Betrag ist keine gültige Zahl.");
    return;
  }

  asset.name = name;
  asset.amount = amount;
  persist();
  render();
  showMessage("Vermögen geändert.");
}

function deleteAsset(id) {
  state.assets = state.assets.filter((asset) => asset.id !== id);
  persist();
  render();
  showMessage("Vermögen gelöscht.");
}

function renderSummary() {
  const month = elements.monthInput.value;
  const summary = summaryTotals(month);
  const income = summary.income;
  const entryExpense = summary.entryExpense;
  const debtMonthlyCost = summary.debtMonthlyCost;
  const expense = entryExpense + debtMonthlyCost;
  const remaining = income - expense;
  const totalDebt = summary.totalDebt;
  const totalAssets = summary.totalAssets;

  elements.familyIncomeValue.textContent = formatMoney.format(income);
  elements.familyExpenseValue.textContent = formatMoney.format(expense);
  elements.remainingValue.textContent = formatMoney.format(remaining);
  elements.remainingValue.parentElement.classList.toggle("positive", remaining >= 0);
  elements.remainingValue.parentElement.classList.toggle("negative", remaining < 0);
  elements.totalDebtValue.textContent = formatMoney.format(totalDebt);
  elements.totalAssetValue.textContent = formatMoney.format(totalAssets);
  if (activeSummaryType) renderSummaryBreakdown();
}

function summaryTotals(month) {
  const monthEntries = state.entries
    .filter(includeInSelectedTotal)
    .filter((entry) => entry.date.startsWith(month));
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
  return { income, entryExpense, debtMonthlyCost, totalDebt, totalAssets };
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
    .filter((entry) => entry.date.startsWith(month));
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
    .map((debt) => debtSummaryRow(debt, remainingDebt(debt, month), "Restschuld"));
  const assetRows = state.assets
    .filter(includeInSelectedTotal)
    .map((asset) => ({
      title: asset.name || "Vermögen",
      meta: [personName(asset.personId), asset.note || ""].filter(Boolean).join(" · "),
      amount: Number(asset.amount || 0),
    }));

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
    return buildSummaryData("Plus / Minus", balanceRows);
  }
  if (type === "debt") {
    return buildSummaryData("Restschulden", debtRows);
  }
  if (type === "asset") {
    return buildSummaryData("Vermögen", assetRows);
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
    meta: [personName(entry.personId), entry.category, displayText(entry.payment), entry.recurring ? "monatlich" : ""].filter(Boolean).join(" · "),
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

  const key = `edit-entry:${id}`;
  if (!elements.entryForm.hidden && activeFormKey === key) {
    closeForm();
    return;
  }
  resetForm({ showTypeChoice: false, type: entry.type });
  openForm({ showTypeChoice: false, slot: elements.entryFormSlot, key });
  elements.entryId.value = entry.id;
  elements.debtId.value = "";
  elements.typeInput.value = entry.type;
  elements.personInput.value = entry.personId || defaultPersonId();
  elements.dateInput.value = entry.date;
  elements.categoryInput.value = entry.category;
  elements.amountInput.value = formatMoneyInput(entry.amount);
  elements.descriptionInput.value = entry.description;
  elements.paymentInput.value = entry.payment;
  elements.recurringInput.checked = entry.recurring;
  syncFormMode();
}

function deleteEntry(id) {
  state.entries = state.entries.filter((entry) => entry.id !== id);
  persist();
  render();
}

function editDebt(id) {
  const debt = state.debts.find((item) => item.id === id);
  if (!debt) return;

  const key = `edit-debt:${id}`;
  if (!elements.entryForm.hidden && activeFormKey === key) {
    closeForm();
    return;
  }
  resetForm({ showTypeChoice: false, type: "debt" });
  openForm({ showTypeChoice: false, slot: elements.debtFormSlot, key });
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
        state.categories = imported.categories;
        state.entries = imported.entries.map(normalizeEntry);
        state.debts = Array.isArray(imported.debts) ? imported.debts.map(normalizeDebt) : [];
        state.assets = Array.isArray(imported.assets) ? imported.assets.map(normalizeAsset) : [];
        state.people = Array.isArray(imported.people) ? imported.people.map(normalizePerson) : state.people;
        showMessage("Backup importiert.");
      }
      persist();
      ensurePeople();
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
    showMessage("Excel-Import braucht Internet beim ersten Laden.");
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
    "Alle Einnahmen, Ausgaben, Schulden und Vermögen von ALLEN Personen werden auf eine neue Person kopiert. Diese Kopien zählen nicht erneut zur Gesamtsumme.",
    "Gesamt (Kopie)",
    { okLabel: "Duplizieren", placeholder: "Neuer Name" }
  );
  if (name === null) return;
  const trimmed = name.trim();
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
  const suggested = `${source.name} (Kopie)`;
  const name = await promptCentered(
    "Person duplizieren",
    `Alle Einnahmen, Ausgaben, Schulden und Vermögen von ${source.name} werden auf eine neue Person kopiert. Die Kopien zählen nur in der neuen Person, nicht erneut zur Gesamtsumme.`,
    suggested,
    { okLabel: "Duplizieren", placeholder: "Neuer Name" }
  );
  if (name === null) return;
  const trimmed = name.trim();
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
  elements.entryForm.querySelectorAll(".field-error").forEach((label) => label.classList.remove("field-error"));
  elements.entryForm.querySelectorAll(".input-error, [aria-invalid='true']").forEach((input) => {
    input.classList.remove("input-error");
    input.removeAttribute("aria-invalid");
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
  return Math.max(0, Number(debt.totalAmount || 0) - paidDebt(debt, month));
}

function paidDebt(debt, month) {
  const paidMonths = paidMonthsUntil(debt, month);
  const paidAmount = paidMonths * Number(debt.monthlyPayment || 0);
  return Math.min(Number(debt.totalAmount || 0), Number(debt.paidSoFar || 0) + paidAmount);
}

function paidMonthsUntil(debt, month) {
  if (!debt.startDate || !debt.endDate || monthToNumber(month) < dateToMonthNumber(debt.startDate)) return 0;
  const endMonth = debt.endDate.slice(0, 7);
  const cappedMonth = monthToNumber(month) > monthToNumber(endMonth) ? endMonth : month;
  return Math.max(0, monthsBetween(debt.startDate.slice(0, 7), cappedMonth) + 1);
}

function isDebtActiveInMonth(debt, month) {
  const afterStart = debt.startDate ? monthToNumber(month) >= dateToMonthNumber(debt.startDate) : true;
  const beforeEnd = debt.endDate ? monthToNumber(month) <= dateToMonthNumber(debt.endDate) : true;
  return afterStart && beforeEnd && remainingDebt(debt, month) > 0;
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

function addMonths(month, amount) {
  const [year, value] = month.split("-").map(Number);
  const date = new Date(year, value - 1 + amount, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
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
  const date = entry.date || monthToDate(new Date().toISOString().slice(0, 7));
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
    updatedAt: Number(entry.updatedAt) > 0 ? Number(entry.updatedAt) : fallbackUpdatedAt,
  };
}

function normalizePerson(person) {
  return {
    id: person.id || createId(),
    name: String(person.name || "Gemeinsam").trim() || "Gemeinsam",
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
