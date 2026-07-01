const STORAGE_KEY = "budget-app-v2";
const STANDARD_CATEGORY_NAMES = [
  "Wohnen",
  "Lebensmittel",
  "Mobilität",
  "Abos",
  "Versicherungen",
  "Gesundheit",
  "Bildung",
  "Freizeit",
  "Reisen",
  "Familie",
  "Schulden",
  "Einkommen",
  "Investments",
  "Sonstiges",
];
const ENTRY_STATUSES = ["open", "paid", "ended"];
const RECURRENCE_TYPES = ["none", "daily", "weekly", "monthly", "yearly"];
const RECURRENCE_LABELS = {
  none: "keine",
  daily: "täglich",
  weekly: "wöchentlich",
  monthly: "monatlich",
  yearly: "jährlich",
};
const BANK_PROVIDERS = [
  { id: "gocardless", name: "GoCardless", status: "geplant" },
  { id: "tink", name: "Tink", status: "geplant" },
  { id: "finapi", name: "finAPI", status: "geplant" },
];
const ASSET_TYPES = {
  cash: "Bargeld",
  bank: "Bankkonto",
  savings: "Sparkonto",
  investment: "Investment",
  property: "Sachwert",
  other: "Sonstiges",
};

const defaults = {
  categories: STANDARD_CATEGORY_NAMES.map((name) => ({ id: createId(), name, budget: defaultBudgetForCategory(name) })),
  entries: [
  ],
  debts: [],
  assets: [],
  bankConnections: [],
  people: [
    { id: createId(), name: "Gemeinsam" },
  ],
  selectedPersonId: "all",
};

const state = loadState();
const formatMoney = new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR" });
ensurePeople();
ensureStandardCategories();
repairCopiedItems();

const elements = {
  monthInput: document.querySelector("#monthInput"),
  contextSelect: document.querySelector("#contextSelect"),
  settingsButton: document.querySelector("#settingsButton"),
  settingsModal: document.querySelector("#settingsModal"),
  settingsModalBackdrop: document.querySelector("#settingsModalBackdrop"),
  settingsCloseButton: document.querySelector("#settingsCloseButton"),
  formTemplates: document.querySelector(".form-templates"),
  entryForm: document.querySelector("#entryForm"),
  assetForm: document.querySelector("#assetForm"),
  assetFormTitle: document.querySelector("#assetFormTitle"),
  assetId: document.querySelector("#assetId"),
  assetPersonInput: document.querySelector("#assetPersonInput"),
  assetTypeInput: document.querySelector("#assetTypeInput"),
  assetNameInput: document.querySelector("#assetNameInput"),
  assetAmountInput: document.querySelector("#assetAmountInput"),
  assetNoteInput: document.querySelector("#assetNoteInput"),
  assetModalActions: document.querySelector("#assetModalActions"),
  assetCancelButton: document.querySelector("#assetCancelButton"),
  assetDeleteButton: document.querySelector("#assetDeleteButton"),
  categoryForm: document.querySelector("#categoryForm"),
  categoryFormTitle: document.querySelector("#categoryFormTitle"),
  categoryIdInput: document.querySelector("#categoryIdInput"),
  categoryNameInput: document.querySelector("#categoryNameInput"),
  categoryBudgetInput: document.querySelector("#categoryBudgetInput"),
  categoryModalActions: document.querySelector("#categoryModalActions"),
  categoryCancelButton: document.querySelector("#categoryCancelButton"),
  categoryDeleteButton: document.querySelector("#categoryDeleteButton"),
  appViews: document.querySelectorAll("[data-view]"),
  navButtons: document.querySelectorAll("[data-view-target]"),
  accountStructureList: document.querySelector("#accountStructureList"),
  bankStatusText: document.querySelector("#bankStatusText"),
  bankProviderList: document.querySelector("#bankProviderList"),
  bankConnectButton: document.querySelector("#bankConnectButton"),
  moreBankConnectButton: document.querySelector("#moreBankConnectButton"),
  moreSettingsButton: document.querySelector("#moreSettingsButton"),
  bankInfoModal: document.querySelector("#bankInfoModal"),
  bankInfoBackdrop: document.querySelector("#bankInfoBackdrop"),
  bankInfoCloseButton: document.querySelector("#bankInfoCloseButton"),
  viewTitle: document.querySelector("#viewTitle"),
  fabButton: document.querySelector("#fabButton"),
  fabMenu: document.querySelector("#fabMenu"),
  calendarPanel: document.querySelector(".calendar-head-panel"),
  formTitle: document.querySelector("#formTitle"),
  entryId: document.querySelector("#entryId"),
  typeInput: document.querySelector("#typeInput"),
  typeField: document.querySelector("[data-type-field]"),
  personInput: document.querySelector("#personInput"),
  personNameInput: document.querySelector("#personNameInput"),
  addPersonButton: document.querySelector("#addPersonButton"),
  personSummaryList: document.querySelector("#personSummaryList"),
  dateInput: document.querySelector("#dateInput"),
  categoryInput: document.querySelector("#categoryInput"),
  customCategoryInput: document.querySelector("#customCategoryInput"),
  amountInput: document.querySelector("#amountInput"),
  descriptionInput: document.querySelector("#descriptionInput"),
  paymentInput: document.querySelector("#paymentInput"),
  recurrenceInput: document.querySelector("#recurrenceInput"),
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
  budgetList: document.querySelector("#budgetList"),
  budgetOverviewMini: document.querySelector("#budgetOverviewMini"),
  categoryList: document.querySelector("#categoryList"),
  categoryTemplate: document.querySelector("#categoryTemplate"),
  addCategoryButton: document.querySelector("#addCategoryButton"),
  quickAddInput: document.querySelector("#quickAddInput"),
  quickAddButton: document.querySelector("#quickAddButton"),
  quickAddTypeInputs: document.querySelectorAll("input[name='quickAddType']"),
  entryMenuButton: document.querySelector("#entryMenuButton"),
  entryMenuPanel: document.querySelector("#entryMenuPanel"),
  filterInput: document.querySelector("#filterInput"),
  transactionSearchInput: document.querySelector("#transactionSearchInput"),
  entryList: document.querySelector("#entryList"),
  exportButton: document.querySelector("#exportButton"),
  importInput: document.querySelector("#importInput"),
  clearButton: document.querySelector("#clearButton"),
  remainingValue: document.querySelector("#remainingValue"),
  totalDebtValue: document.querySelector("#totalDebtValue"),
  totalAssetValue: document.querySelector("#totalAssetValue"),
  overviewIncomeValue: document.querySelector("#overviewIncomeValue"),
  overviewExpenseValue: document.querySelector("#overviewExpenseValue"),
  overviewMonthLabel: document.querySelector("#overviewMonthLabel"),
  heroIncomeValue: document.querySelector("#heroIncomeValue"),
  heroExpenseValue: document.querySelector("#heroExpenseValue"),
  heroBudgetValue: document.querySelector("#heroBudgetValue"),
  overviewMoneyButton: document.querySelector("#overviewMoneyButton"),
  overviewAssetShortcut: document.querySelector("#overviewAssetShortcut"),
  overviewDebtShortcut: document.querySelector("#overviewDebtShortcut"),
  overviewAssetShortcutValue: document.querySelector("#overviewAssetShortcutValue"),
  overviewDebtShortcutValue: document.querySelector("#overviewDebtShortcutValue"),
  overviewAssetShortcutHint: document.querySelector("#overviewAssetShortcutHint"),
  overviewDebtShortcutHint: document.querySelector("#overviewDebtShortcutHint"),
  overviewRecurringValue: document.querySelector("#overviewRecurringValue"),
  overviewRecurringHint: document.querySelector("#overviewRecurringHint"),
  accountsNetWorthValue: document.querySelector("#accountsNetWorthValue"),
  accountsDebtValue: document.querySelector("#accountsDebtValue"),
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
  biggestExpenseValue: document.querySelector("#biggestExpenseValue"),
  biggestExpenseHint: document.querySelector("#biggestExpenseHint"),
  topSpendValue: document.querySelector("#topSpendValue"),
  topSpendHint: document.querySelector("#topSpendHint"),
  nextPaymentValue: document.querySelector("#nextPaymentValue"),
  nextPaymentHint: document.querySelector("#nextPaymentHint"),
  weekSpendValue: document.querySelector("#weekSpendValue"),
  weekSpendHint: document.querySelector("#weekSpendHint"),
  reportsRatioValue: document.querySelector("#reportsRatioValue"),
  reportsRatioHint: document.querySelector("#reportsRatioHint"),
  reportsNetWorthValue: document.querySelector("#reportsNetWorthValue"),
  reportsNetWorthHint: document.querySelector("#reportsNetWorthHint"),
  reportsOptimizationList: document.querySelector("#reportsOptimizationList"),
  reportsFrequentList: document.querySelector("#reportsFrequentList"),
  reportsLargestList: document.querySelector("#reportsLargestList"),
  reportsDebtValue: document.querySelector("#reportsDebtValue"),
  reportsDebtHint: document.querySelector("#reportsDebtHint"),
  reportsMonthMini: document.querySelector("#reportsMonthMini"),
  reportsTrendList: document.querySelector("#reportsTrendList"),
  calendarTitle: document.querySelector("#calendarTitle"),
  calendarSubtitle: document.querySelector("#calendarSubtitle"),
  calendarPrevButton: document.querySelector("#calendarPrevButton"),
  calendarNextButton: document.querySelector("#calendarNextButton"),
  calendarTodayButton: document.querySelector("#calendarTodayButton"),
  calendarGrid: document.querySelector("#calendarGrid"),
  calendarDayTitle: document.querySelector("#calendarDayTitle"),
  calendarDayTotal: document.querySelector("#calendarDayTotal"),
  calendarDayList: document.querySelector("#calendarDayList"),
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
let activeView = "overview";
let selectedCalendarDay = localDateString(new Date());

elements.monthInput.value = currentLocalMonth();
elements.typeInput.value = "expense";

elements.entryForm.addEventListener("submit", saveEntry);
elements.entryForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.entryForm.addEventListener("change", (event) => clearFieldError(event.target));
elements.assetForm.addEventListener("submit", saveAsset);
elements.assetForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.assetForm.addEventListener("change", (event) => clearFieldError(event.target));
elements.categoryForm.addEventListener("submit", saveCategoryFromForm);
elements.categoryForm.addEventListener("input", (event) => clearFieldError(event.target));
elements.categoryForm.addEventListener("change", (event) => clearFieldError(event.target));
elements.resetFormButton.addEventListener("click", () => resetForm());
elements.navButtons.forEach((button) => {
  button.addEventListener("click", () => setActiveView(button.dataset.viewTarget));
});
elements.fabButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  toggleFabMenu();
});
elements.fabMenu.addEventListener("click", (event) => event.stopPropagation());
elements.addDebtButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeFabMenu();
  openTypedForm("debt");
});
elements.addIncomeButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeFabMenu();
  openTypedForm("income");
});
elements.addExpenseButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeFabMenu();
  openTypedForm("expense");
});
elements.addAssetButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeFabMenu();
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
elements.recurrenceInput.addEventListener("change", syncRecurringFields);
elements.categoryInput.addEventListener("change", () => {
  if (elements.categoryInput.value) elements.customCategoryInput.value = "";
});
elements.addCategoryButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  closeFabMenu();
  openCategoryForm();
});
elements.contextSelect.addEventListener("change", () => {
  selectPerson(elements.contextSelect.value || "all");
});
elements.settingsButton.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  openSettingsModal();
});
elements.settingsModalBackdrop.addEventListener("click", closeSettingsModal);
elements.settingsCloseButton.addEventListener("click", closeSettingsModal);
elements.moreSettingsButton?.addEventListener("click", (event) => {
  event.preventDefault();
  openSettingsModal();
});
elements.bankConnectButton?.addEventListener("click", openBankInfoModal);
elements.moreBankConnectButton?.addEventListener("click", openBankInfoModal);
elements.bankInfoBackdrop?.addEventListener("click", closeBankInfoModal);
elements.bankInfoCloseButton?.addEventListener("click", closeBankInfoModal);
[elements.overviewMoneyButton, elements.overviewAssetShortcut, elements.overviewDebtShortcut].forEach((button) => {
  button?.addEventListener("click", () => setActiveView("accounts"));
});
document.querySelectorAll("[data-more-action]").forEach((button) => {
  button.addEventListener("click", () => handleMoreAction(button.dataset.moreAction));
});
elements.quickAddButton.addEventListener("click", saveQuickAdd);
elements.quickAddInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    saveQuickAdd();
  }
});
elements.filterInput.addEventListener("change", () => {
  render();
  closeEntryMenu();
});
elements.transactionSearchInput.addEventListener("input", renderEntries);
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
elements.clearDebtsButton?.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  clearDebtsForContext();
});
elements.clearAssetsButton?.addEventListener("click", (event) => {
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
  if (!elements.fabMenu.hidden && !elements.fabMenu.contains(event.target) && !elements.fabButton.contains(event.target)) {
    closeFabMenu();
  }
  if (elements.entryMenuPanel.hidden) return;
  if (elements.entryMenuPanel.contains(event.target) || elements.entryMenuButton.contains(event.target)) return;
  closeEntryMenu();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeFabMenu();
    closeEntryMenu();
    if (!elements.editModal.hidden) closeEditModal();
    if (!elements.settingsModal.hidden) closeSettingsModal();
    if (elements.bankInfoModal && !elements.bankInfoModal.hidden) closeBankInfoModal();
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
elements.categoryCancelButton.addEventListener("click", closeEditModal);
elements.categoryDeleteButton.addEventListener("click", deleteActiveEditItem);
[elements.amountInput, elements.debtTotalInput, elements.paidSoFarInput, elements.debtPaymentInput, elements.principalInput, elements.interestInput, elements.finalPaymentInput].forEach((input) => {
  input.addEventListener("blur", () => formatMoneyField(input));
  input.addEventListener("focus", () => input.select());
});
elements.assetAmountInput.addEventListener("blur", () => formatMoneyField(elements.assetAmountInput));
elements.assetAmountInput.addEventListener("focus", () => elements.assetAmountInput.select());
elements.categoryBudgetInput.addEventListener("blur", () => formatMoneyField(elements.categoryBudgetInput));
elements.categoryBudgetInput.addEventListener("focus", () => elements.categoryBudgetInput.select());
elements.nominalRateInput.addEventListener("blur", () => {
  elements.nominalRateInput.value = formatPercentInput(elements.nominalRateInput.value);
});
elements.calendarPrevButton.addEventListener("click", () => {
  elements.monthInput.value = shiftMonth(elements.monthInput.value, -1);
  selectedCalendarDay = monthToDate(elements.monthInput.value);
  render();
});
elements.calendarNextButton.addEventListener("click", () => {
  elements.monthInput.value = shiftMonth(elements.monthInput.value, 1);
  selectedCalendarDay = monthToDate(elements.monthInput.value);
  render();
});
elements.calendarTodayButton.addEventListener("click", () => {
  elements.monthInput.value = currentLocalMonth();
  selectedCalendarDay = localDateString(new Date());
  render();
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").then((registration) => {
      if (!registration) return;
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
    }).catch(() => {});
  });
}

render();
setActiveView(activeView);
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
      bankConnections: Array.isArray(parsed.bankConnections) ? parsed.bankConnections.map(normalizeBankConnection) : [],
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

function repairCopiedItems() {
  let changed = false;
  changed = markDuplicateLineage(state.entries, entrySignature) || changed;
  changed = markDuplicateLineage(state.debts, debtSignature) || changed;
  changed = markDuplicateLineage(state.assets, assetSignature) || changed;
  changed = removeOrphanDuplicates(state.entries) || changed;
  changed = removeOrphanDuplicates(state.debts) || changed;
  changed = removeOrphanDuplicates(state.assets) || changed;
  if (changed) persist();
}

function markDuplicateLineage(items, signatureFn) {
  const groups = new Map();

  items.forEach((item) => {
    const signature = signatureFn(item);
    if (!signature) return;
    if (!groups.has(signature)) groups.set(signature, []);
    groups.get(signature).push(item);
  });

  let changed = false;
  groups.forEach((group) => {
    const referencedIds = new Set(group.map((item) => item.duplicateOf).filter(Boolean));
    const root = group.find((item) => referencedIds.has(item.id)) || group.find((item) => !item.duplicateOf) || group[0];
    const rootOwner = root.personId || defaultPersonId();
    const rootId = root.duplicateOf && root.duplicateOf !== root.id ? root.duplicateOf : root.id;

    group.forEach((item) => {
      const owner = item.personId || defaultPersonId();
      if (item.id === rootId || item.id === root.id) {
        if (item.duplicateOf === item.id) {
          delete item.duplicateOf;
          changed = true;
        }
        return;
      }
      if (owner !== rootOwner && item.duplicateOf !== rootId) {
        item.duplicateOf = rootId;
        changed = true;
      }
    });
  });

  return changed;
}

function removeOrphanDuplicates(items) {
  const ids = new Set(items.map((item) => item.id));
  const before = items.length;
  for (let index = items.length - 1; index >= 0; index -= 1) {
    const duplicateOf = items[index].duplicateOf;
    if (duplicateOf && !ids.has(duplicateOf)) items.splice(index, 1);
  }
  return items.length !== before;
}

function lineageId(item) {
  return item?.duplicateOf || item?.id || "";
}

function findItemById(id) {
  return [...state.entries, ...state.debts, ...state.assets].find((item) => item.id === id);
}

function sameLineage(item, id) {
  if (!id) return false;
  const target = findItemById(id);
  const targetLineage = target ? lineageId(target) : id;
  return item.id === targetLineage || item.duplicateOf === targetLineage || lineageId(item) === targetLineage;
}

function entrySignature(entry) {
  const type = entry.type || "expense";
  const date = entry.date || "";
  const category = String(entry.category || "").trim().toLowerCase();
  const description = String(entry.description || "").trim().toLowerCase();
  const amount = Number(entry.amount || 0).toFixed(2);
  const recurring = entryRecurrence(entry);
  const endDate = entry.endDate || "";
  if (!description && amount === "0.00") return "";
  return `${type}|${date}|${category}|${description}|${amount}|${recurring}|${endDate}`;
}

function debtSignature(debt) {
  const creditor = String(debt.creditor || "").trim().toLowerCase();
  const total = Number(debt.totalAmount || 0).toFixed(2);
  const monthly = Number(debt.monthlyPayment || 0).toFixed(2);
  const start = debt.startDate || "";
  const end = debt.endDate || "";
  if (!creditor && total === "0.00") return "";
  return `${creditor}|${total}|${monthly}|${start}|${end}`;
}

function assetSignature(asset) {
  const name = String(asset.name || "").trim().toLowerCase();
  const note = String(asset.note || "").trim().toLowerCase();
  const type = normalizeAssetType(asset.type || inferAssetType(asset));
  const amount = Number(asset.amount || 0).toFixed(2);
  if (!name && amount === "0.00") return "";
  return `${name}|${amount}|${note}|${type}`;
}

function getActiveContext(id = state.selectedPersonId) {
  const contextId = id || "all";
  if (contextId === "all") return { id: "all", label: "Gesamt", isAll: true };
  const person = state.people.find((item) => item.id === contextId);
  if (!person) return { id: "all", label: "Gesamt", isAll: true };
  return { id: person.id, label: person.name, isAll: false };
}

function contextMatchesItem(item, context = getActiveContext()) {
  // Kopien bleiben in ihrer Einzelperson sichtbar, werden in Gesamt aber nicht erneut addiert.
  if (context.isAll) return !item.duplicateOf;
  return (item.personId || defaultPersonId()) === context.id;
}

function shouldIncludeEntryInView(entry, context = getActiveContext(), monthEntries = []) {
  if (!context.isAll) return (entry.personId || defaultPersonId()) === context.id;
  if (entry.duplicateOf) return false;
  return !isLikelyDuplicateAggregateEntry(entry, monthEntries);
}

function shouldIncludeItemInView(item, context = getActiveContext()) {
  return contextMatchesItem(item, context);
}

function isLikelyDuplicateAggregateEntry(entry, monthEntries) {
  const ownerId = entry.personId || defaultPersonId();
  if (!isAggregatePersonId(ownerId)) return false;
  const key = comparableEntryKey(entry);
  return monthEntries.some((candidate) => {
    if (candidate === entry || candidate.duplicateOf) return false;
    const candidateOwner = candidate.personId || defaultPersonId();
    return !isAggregatePersonId(candidateOwner) && comparableEntryKey(candidate) === key;
  });
}

function isAggregatePersonId(id) {
  const person = state.people.find((item) => item.id === id);
  const name = String(person?.name || "").trim().toLowerCase();
  return !id || id === defaultPersonId() || ["gesamt", "gemeinsam", "alle"].includes(name);
}

function comparableEntryKey(entry) {
  return [
    entry.type || "expense",
    entryOccurrenceDate(entry, entry.occurrenceDate?.slice(0, 7) || entry.date?.slice(0, 7) || elements.monthInput.value),
    normalizeComparableText(entry.category),
    normalizeComparableText(entry.description || entry.category),
  ].join("|");
}

function normalizeComparableText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function preserveCopyTracking(nextItem, existingItem) {
  return existingItem?.duplicateOf ? { ...nextItem, duplicateOf: existingItem.duplicateOf } : nextItem;
}

function getEntriesForMonth(month, context = getActiveContext()) {
  const monthEntries = state.entries.flatMap((entry) => expandEntryOccurrences(entry, month));
  return monthEntries.filter((entry) => shouldIncludeEntryInView(entry, context, monthEntries));
}

function getDebtsForMonth(month, context = getActiveContext(), { activeOnly = false, visibleOnly = true } = {}) {
  return state.debts
    .filter((debt) => shouldIncludeItemInView(debt, context))
    .filter((debt) => !visibleOnly || isDebtVisibleInMonth(debt, month))
    .filter((debt) => !activeOnly || isDebtActiveInMonth(debt, month));
}

function getAssetsForContext(context = getActiveContext()) {
  return state.assets.filter((asset) => shouldIncludeItemInView(asset, context));
}

function getBankConnections() {
  return Array.isArray(state.bankConnections) ? state.bankConnections : [];
}

function calculateMoneyStructure(month, context = getActiveContext()) {
  const assets = getAssetsForContext(context);
  const debts = getDebtsForMonth(month, context);
  const buckets = [
    { key: "bank", label: "Bankkonten", hint: "Giro- und Zahlungskonten", amount: 0, count: 0 },
    { key: "cash", label: "Bargeld", hint: "Cash und Haushaltskasse", amount: 0, count: 0 },
    { key: "savings", label: "Sparkonten", hint: "Rücklagen und Sparziele", amount: 0, count: 0 },
    { key: "investment", label: "Investments", hint: "Depot, ETF, Krypto oder Beteiligungen", amount: 0, count: 0 },
    { key: "property", label: "Sachwerte", hint: "Wertgegenstände und sonstige Assets", amount: 0, count: 0 },
    { key: "liability", label: "Schulden", hint: "Kredite, Ratenzahlungen, Kreditkarten", amount: 0, count: 0, negative: true },
  ];
  const byKey = new Map(buckets.map((bucket) => [bucket.key, bucket]));
  assets.forEach((asset) => {
    const key = normalizeAssetType(asset.type || inferAssetType(asset));
    const bucket = byKey.get(key) || byKey.get("property");
    bucket.amount += Number(asset.amount || 0);
    bucket.count += 1;
  });
  debts.forEach((debt) => {
    const bucket = byKey.get("liability");
    bucket.amount += remainingDebt(debt, month);
    bucket.count += 1;
  });
  return buckets;
}

function calculateMonthSummary(month, context = getActiveContext()) {
  const monthEntries = getEntriesForMonth(month, context);
  const income = sum(monthEntries.filter((entry) => entry.type === "income"));
  const entryExpense = sum(monthEntries.filter((entry) => entry.type === "expense"));
  const debtMonthlyCost = getDebtsForMonth(month, context, { activeOnly: true })
    .reduce((total, debt) => total + debtPaymentForMonth(debt, month), 0);
  const totalDebt = getDebtsForMonth(month, context)
    .reduce((total, debt) => total + remainingDebt(debt, month), 0);
  const totalAssets = getAssetsForContext(context)
    .reduce((total, asset) => total + Number(asset.amount || 0), 0);
  const expense = entryExpense + debtMonthlyCost;
  const balance = income - expense;
  const netWorth = totalAssets - totalDebt;
  return { income, entryExpense, debtMonthlyCost, expense, balance, totalDebt, totalAssets, netWorth };
}

function calculateInsightSummary(month, context = getActiveContext()) {
  const entries = getEntriesForMonth(month, context);
  const expenses = entries.filter((entry) => entry.type === "expense");
  const recurringExpenses = expenses.filter(isEntryRecurring);
  const byCategory = totalsBy(expenses, (entry) => entry.category || "Sonstiges");
  const byDescription = totalsBy(expenses, (entry) => entry.description || entry.category || "Ausgabe");
  const topCategory = topMapEntry(byCategory);
  const topDescription = topMapEntry(byDescription);
  const largestPayments = expenses
    .slice()
    .sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0))
    .slice(0, 3);
  const recurringTotal = recurringExpenses.reduce((total, entry) => total + Number(entry.amount || 0), 0);
  const summary = calculateMonthSummary(month, context);
  return {
    entries,
    expenses,
    byCategory,
    byDescription,
    topCategory,
    topDescription,
    largestPayments,
    recurringCount: recurringExpenses.length,
    recurringTotal,
    recurringShare: summary.expense > 0 ? Math.round((recurringTotal / summary.expense) * 100) : 0,
    summary,
  };
}

function calculateBudgetUsage(month, context = getActiveContext()) {
  const expenses = getEntriesForMonth(month, context).filter((entry) => entry.type === "expense");
  const spentByCategory = new Map();
  expenses.forEach((entry) => {
    const name = entry.category || "Sonstiges";
    spentByCategory.set(name, (spentByCategory.get(name) || 0) + Number(entry.amount || 0));
  });
  const categories = uniqueCategoryNames(state.categories.map((category) => category.name))
    .map((name) => state.categories.find((category) => sameCategory(category.name, name)))
    .filter(Boolean);
  const totalBudget = categories.reduce((total, category) => total + Number(category.budget || 0), 0);
  const totalSpent = [...spentByCategory.values()].reduce((total, value) => total + value, 0);
  return { categories, spentByCategory, totalBudget, totalSpent, remaining: totalBudget - totalSpent };
}

function getCalendarDaySummary(date, context = getActiveContext()) {
  const items = calendarItemsForDate(date, context);
  const entries = items.filter((item) => item.kind !== "debt");
  const debtExpense = items
    .filter((item) => item.kind === "debt")
    .reduce((total, item) => total + Number(item.amount || 0), 0);
  const income = sum(entries.filter((entry) => entry.type === "income"));
  const entryExpense = sum(entries.filter((entry) => entry.type === "expense"));
  const expense = entryExpense + debtExpense;
  return {
    items,
    income,
    entryExpense,
    debtExpense,
    expense,
    net: income - expense,
    recurring: entries.some(isEntryRecurring),
  };
}

function normalizeView(view) {
  const valid = ["overview", "accounts", "budgets", "transactions", "reports", "more"];
  return valid.includes(view) ? view : "overview";
}

function setActiveView(view, { resetScroll = true } = {}) {
  const previousView = activeView;
  activeView = normalizeView(view);
  const titles = {
    overview: "Start",
    accounts: "Geld",
    budgets: "Budgets",
    transactions: "Buchungen",
    reports: "Analyse",
    more: "Mehr",
  };
  elements.appViews.forEach((section) => {
    section.hidden = section.dataset.view !== activeView;
  });
  if (elements.calendarPanel) {
    elements.calendarPanel.hidden = activeView !== "transactions";
  }
  elements.navButtons.forEach((button) => {
    const selected = button.dataset.viewTarget === activeView;
    button.classList.toggle("active", selected);
    button.setAttribute("aria-current", selected ? "page" : "false");
  });
  elements.viewTitle.textContent = titles[activeView] || "BudgetUp";
  closeFabMenu();
  closeSummaryBreakdown();
  if (resetScroll && previousView !== activeView) resetViewScroll();
}

function resetViewScroll() {
  requestAnimationFrame(() => {
    const shell = document.querySelector(".app-shell");
    if (shell && shell.scrollTop) shell.scrollTop = 0;
    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
    window.scrollTo(0, 0);
  });
}

function toggleFabMenu() {
  elements.fabMenu.hidden = !elements.fabMenu.hidden;
  elements.fabButton.classList.toggle("open", !elements.fabMenu.hidden);
  elements.fabButton.setAttribute("aria-expanded", String(!elements.fabMenu.hidden));
}

function closeFabMenu() {
  elements.fabMenu.hidden = true;
  elements.fabButton.classList.remove("open");
  elements.fabButton.setAttribute("aria-expanded", "false");
}

function openSettingsModal() {
  closeFabMenu();
  closeEntryMenu();
  elements.settingsModal.hidden = false;
  document.body.classList.add("modal-open");
  window.setTimeout(() => elements.settingsCloseButton.focus({ preventScroll: true }), 30);
}

function closeSettingsModal() {
  if (elements.settingsModal.hidden) return;
  elements.settingsModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function openBankInfoModal(event) {
  event?.preventDefault();
  closeFabMenu();
  closeEntryMenu();
  if (!elements.bankInfoModal) return;
  elements.bankInfoModal.hidden = false;
  document.body.classList.add("modal-open");
  window.setTimeout(() => elements.bankInfoCloseButton?.focus({ preventScroll: true }), 30);
}

function closeBankInfoModal() {
  if (!elements.bankInfoModal || elements.bankInfoModal.hidden) return;
  elements.bankInfoModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function handleMoreAction(action) {
  if (action === "category") {
    openCategoryForm();
    return;
  }
  if (action === "export") {
    exportBackup();
    return;
  }
  openSettingsModal();
}

function prepareModalForm(form) {
  elements.editModalSlot.replaceChildren(form);
  const title = form.querySelector("h2");
  if (title?.id) elements.editModal.setAttribute("aria-labelledby", title.id);
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
  elements.categoryForm.hidden = true;
  elements.categoryForm.classList.remove("modal-form");
  elements.categoryModalActions.hidden = true;
  elements.categoryDeleteButton.hidden = false;
  elements.formTemplates.append(elements.entryForm, elements.assetForm, elements.categoryForm);
  elements.editModal.setAttribute("aria-labelledby", "formTitle");
  activeFormKey = "";
  activeEditContext = null;
  clearFieldErrors();
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
  return state.selectedPersonId === "all" ? (visiblePeople()[0]?.id || defaultPersonId()) : state.selectedPersonId;
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
  prepareModalForm(elements.entryForm);
  elements.assetForm.hidden = true;
  elements.categoryForm.hidden = true;
  elements.entryForm.hidden = false;
  elements.entryForm.classList.add("modal-form");
  elements.editModalActions.hidden = false;
  elements.editDeleteButton.hidden = mode !== "edit";
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
  const recurrence = selectedRecurrence();
  if (recurrence !== "none" && elements.entryEndInput.value && compareDates(elements.entryEndInput.value, elements.dateInput.value || monthToDate(elements.monthInput.value)) < 0) {
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
    recurrence,
    recurring: recurrence !== "none",
    endDate: recurrence !== "none" ? elements.entryEndInput.value : "",
    status: normalizeStatus(elements.entryStatusInput.value),
    updatedAt: Date.now(),
  };

  const existingIndex = state.entries.findIndex((item) => item.id === entry.id);
  if (existingIndex >= 0) {
    state.entries[existingIndex] = preserveCopyTracking(entry, state.entries[existingIndex]);
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

function saveQuickAdd() {
  const raw = elements.quickAddInput.value.trim();
  const type = selectedQuickAddType();
  if (!raw) {
    showMessage("Beispiel: 12,50 Kaffee Lebensmittel", "error");
    elements.quickAddInput.focus();
    return;
  }
  const parsed = parseQuickAdd(raw, type, getActiveContext(), elements.monthInput.value);
  if (!parsed) {
    showMessage("Bitte mit Betrag starten, z. B. 12,50 Kaffee Lebensmittel.", "error");
    elements.quickAddInput.focus();
    return;
  }
  if ((type === "expense" || type === "income") && !state.categories.some((category) => sameCategory(category.name, parsed.category))) {
    state.categories.push({ id: createId(), name: parsed.category, budget: 0 });
  }
  const personId = targetPersonIdForNewItem();
  if (type === "debt") {
    state.debts.push({
      id: createId(),
      personId,
      creditor: parsed.description,
      totalAmount: parsed.amount,
      paidSoFar: 0,
      monthlyPayment: parsed.monthlyPayment || 0,
      startDate: quickAddDate(),
      endDate: "",
      status: "open",
      paymentMethod: "",
      account: "",
      liabilityType: inferLiabilityType({ creditor: parsed.description, note: "Quick Add", account: "" }),
      principalAmount: 0,
      interestAmount: 0,
      nominalRate: 0,
      termMonths: 0,
      finalPayment: 0,
      note: "Quick Add",
    });
  } else if (type === "asset") {
    state.assets.push({
      id: createId(),
      personId,
      type: inferAssetType({ name: parsed.description, note: parsed.category }),
      name: parsed.description,
      amount: parsed.amount,
      note: "Quick Add",
    });
  } else {
    state.entries.push({
      id: createId(),
      personId,
      type,
      date: quickAddDate(),
      category: parsed.category,
      description: parsed.description,
      payment: "Karte",
      amount: parsed.amount,
      recurrence: "none",
      recurring: false,
      endDate: "",
      status: "open",
      updatedAt: Date.now(),
    });
  }
  elements.quickAddInput.value = "";
  persist();
  render();
  showMessage(`${quickAddTypeLabel(type)} gespeichert.`, "success");
}

function selectedQuickAddType() {
  return [...elements.quickAddTypeInputs].find((input) => input.checked)?.value || "expense";
}

function quickAddTypeLabel(type) {
  return {
    expense: "Ausgabe",
    income: "Einnahme",
    debt: "Schuld",
    asset: "Vermögen",
  }[type] || "Eintrag";
}

function parseQuickAdd(raw, type = "expense", context = getActiveContext(), month = elements.monthInput.value) {
  const match = String(raw || "").trim().match(/^([+-]?\d+(?:[.,]\d{1,2})?)(?:\s*€)?\s+(.+)$/);
  if (!match) return null;
  const amount = parseMoneyInput(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return null;
  const words = match[2].trim().split(/\s+/).filter(Boolean);
  const monthlyIndex = words.findIndex((word) => /^monatlich$/i.test(word));
  let monthlyPayment = 0;
  if (monthlyIndex > 0) {
    const possibleRate = parseMoneyInput(words[monthlyIndex - 1]);
    if (Number.isFinite(possibleRate) && possibleRate > 0) {
      monthlyPayment = possibleRate;
      words.splice(monthlyIndex - 1, 2);
    }
  }
  const categoryNames = uniqueCategoryNames([...state.categories.map((category) => category.name), ...STANDARD_CATEGORY_NAMES, "Einkommen", "Gehalt"]);
  const categoryWord = words.find((word) => categoryNames.some((name) => sameCategory(name, word)));
  const category = categoryNames.find((name) => sameCategory(name, categoryWord)) || (type === "income" ? "Einkommen" : "Sonstiges");
  const descriptionWords = categoryWord ? words.filter((word) => !sameCategory(word, categoryWord)) : words;
  const description = descriptionWords.join(" ").trim() || defaultQuickAddDescription(type, category);
  return {
    amount: Math.abs(amount),
    type,
    category,
    description,
    monthlyPayment,
    contextId: context.id,
    date: monthToDate(month),
  };
}

function defaultQuickAddDescription(type, category) {
  if (type === "debt") return "Schuld";
  if (type === "asset") return "Vermögen";
  return type === "income" ? "Einnahme" : category || "Ausgabe";
}

function quickAddDate() {
  return selectedCalendarDate();
}

function selectedCalendarDate() {
  const month = elements.monthInput.value || currentLocalMonth();
  return selectedCalendarDay?.slice(0, 7) === month ? selectedCalendarDay : monthToDate(month);
}

function resetForm({ showTypeChoice = formTypeChoiceVisible, type } = {}) {
  const nextType = type || (showTypeChoice ? "expense" : elements.typeInput.value || "expense");
  elements.entryForm.reset();
  elements.entryId.value = "";
  elements.debtId.value = "";
  elements.typeInput.value = nextType;
  elements.paymentInput.value = "Karte";
  elements.dateInput.value = selectedCalendarDate();
  elements.debtStartInput.value = selectedCalendarDate();
  elements.recurrenceInput.value = "none";
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
    liabilityType: inferLiabilityType({ creditor: elements.creditorInput.value.trim(), account: elements.debtAccountInput.value.trim(), note: "" }),
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
    state.debts[existingIndex] = preserveCopyTracking(debt, state.debts[existingIndex]);
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
  openCategoryForm();
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
  if (activeEditContext?.kind === "category" && activeEditContext.id === id) closeEditModal();
}

function render() {
  fillContextSelect();
  fillPersonSelect();
  fillCategorySelect();
  renderCategories();
  renderPeople();
  renderDetailTitles();
  renderMoneyStructure();
  renderBankConnect();
  renderDebts();
  renderAssets();
  renderEntries();
  renderSummary();
  renderBudgets();
  renderCalendar();
  renderReports();
}

function fillContextSelect() {
  const options = [
    { id: "all", name: "Gesamt" },
    ...visiblePeople().map((person) => ({ id: person.id, name: person.name })),
  ];
  elements.contextSelect.innerHTML = options
    .map((person) => `<option value="${escapeHtml(person.id)}">${escapeHtml(person.name)}</option>`)
    .join("");
  if (options.some((person) => person.id === state.selectedPersonId)) {
    elements.contextSelect.value = state.selectedPersonId;
  } else {
    elements.contextSelect.value = "all";
  }
}

function fillPersonSelect() {
  const existing = elements.personInput.value;
  const people = visiblePeople();
  const options = people.length ? people : [{ id: defaultPersonId(), name: "Gemeinsam" }];
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
  const totals = calculateMonthSummary(month);
  elements.debtTitle.textContent = `Schulden - ${suffix}`;
  elements.debtTotalMini.textContent = formatMoney.format(totals.totalDebt);
  elements.assetTitle.textContent = `Vermögen - ${suffix}`;
  elements.assetTotalMini.textContent = formatMoney.format(totals.totalAssets);
  elements.entryTitle.textContent = `Einnahmen und Ausgaben - ${suffix}`;
  elements.entryIncomeMini.textContent = `Einn. ${formatMoney.format(totals.income)}`;
  elements.entryExpenseMini.textContent = `Ausg. ${formatMoney.format(totals.expense)}`;
}

function renderMoneyStructure() {
  if (!elements.accountStructureList) return;
  const month = elements.monthInput.value;
  const buckets = calculateMoneyStructure(month);
  elements.accountStructureList.innerHTML = buckets.map((bucket) => {
    const amount = bucket.negative ? -Math.abs(bucket.amount) : bucket.amount;
    const valueClass = amount < 0 ? "negative-text" : "positive-text";
    return `
      <article class="account-type-row">
        <span class="account-type-icon">${escapeHtml(accountTypeInitial(bucket.label))}</span>
        <div>
          <strong>${escapeHtml(bucket.label)}</strong>
          <small>${escapeHtml(bucket.count ? `${bucket.count} Eintrag${bucket.count === 1 ? "" : "e"} · ${bucket.hint}` : bucket.hint)}</small>
        </div>
        <b class="${valueClass}">${formatMoney.format(amount)}</b>
      </article>
    `;
  }).join("");
}

function renderBankConnect() {
  const connections = getBankConnections();
  const active = connections.filter((connection) => connection.status === "connected");
  if (elements.bankStatusText) {
    elements.bankStatusText.textContent = active.length
      ? `${active.length} Verbindung${active.length === 1 ? "" : "en"} vorbereitet. Letzter Sync: ${formatNullableDate(active[0].lastSyncAt)}`
      : "Noch nicht verbunden. CSV-Import bleibt die sichere Zwischenlösung.";
  }
  if (elements.bankProviderList) {
    elements.bankProviderList.innerHTML = BANK_PROVIDERS.map((provider) => `
      <article class="provider-row">
        <span>${escapeHtml(provider.name)}</span>
        <strong>${escapeHtml(provider.status)}</strong>
      </article>
    `).join("");
  }
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
  const summary = calculateMonthSummary(month, getActiveContext(personId));
  return { income: summary.income, expense: summary.expense, assets: summary.totalAssets, debt: summary.totalDebt, balance: summary.balance };
}

function monthlyTotalsForFamily(month) {
  const summary = calculateMonthSummary(month, getActiveContext("all"));
  return { income: summary.income, expense: summary.expense, assets: summary.totalAssets, debt: summary.totalDebt, balance: summary.balance };
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
  const recurring = selectedRecurrence() !== "none";
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
  if (!elements.categoryList) return;
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

function renderBudgets() {
  const month = elements.monthInput.value;
  const usage = calculateBudgetUsage(month);
  const rows = usage.categories;
  elements.budgetOverviewMini.textContent = `${formatMoney.format(usage.remaining)} übrig`;

  if (!rows.length) {
    elements.budgetList.innerHTML = `<p class="empty-state">Noch keine Budgets. Nutze den Plus-Button, um eine Kategorie anzulegen.</p>`;
    return;
  }

  elements.budgetList.innerHTML = rows.map((category) => {
    const spent = usage.spentByCategory.get(category.name) || 0;
    const budget = Number(category.budget || 0);
    const remaining = budget - spent;
    const percent = budget > 0 ? Math.min(100, Math.max(0, (spent / budget) * 100)) : 0;
    const status = budgetStatus(spent, budget);
    const group = categoryGroup(category.name);
    return `
      <article class="budget-row ${status.className}">
        <div class="budget-copy">
          <small>${escapeHtml(group)}</small>
          <strong>${escapeHtml(category.name)}</strong>
          <span>${formatMoney.format(spent)} ausgegeben · ${formatMoney.format(remaining)} Rest</span>
          <div class="budget-progress" aria-hidden="true"><span style="width:${percent}%"></span></div>
        </div>
        <div class="budget-side">
          <b>${formatMoney.format(budget)}</b>
          <span class="status-badge ${status.badgeClass}">${escapeHtml(status.label)}</span>
          <button class="icon-button" type="button" title="Budget bearbeiten" aria-label="Budget bearbeiten" data-category-edit="${escapeHtml(category.id)}">✎</button>
        </div>
      </article>
    `;
  }).join("");

  elements.budgetList.querySelectorAll("[data-category-edit]").forEach((button) => {
    button.addEventListener("click", () => openCategoryForm(button.dataset.categoryEdit));
  });
}

function openCategoryForm(id = "") {
  const mode = id ? "edit" : "create";
  const key = id ? `edit-category:${id}` : "typed:category";
  if (isModalOpenFor(key)) {
    closeEditModal();
    return;
  }
  resetCategoryForm();
  activeEditContext = { kind: "category", mode, id };
  activeFormKey = key;
  prepareModalForm(elements.categoryForm);
  elements.entryForm.hidden = true;
  elements.assetForm.hidden = true;
  elements.categoryForm.hidden = false;
  elements.categoryForm.classList.add("modal-form");
  elements.categoryModalActions.hidden = false;
  elements.categoryDeleteButton.hidden = mode !== "edit";
  elements.categoryFormTitle.textContent = mode === "edit" ? "Budget bearbeiten" : "Budget/Kategorie";
  if (id) {
    const category = state.categories.find((item) => item.id === id);
    if (!category) return;
    elements.categoryIdInput.value = category.id;
    elements.categoryNameInput.value = category.name;
    elements.categoryBudgetInput.value = formatMoneyInput(category.budget || 0);
  }
  elements.editModal.hidden = false;
  document.body.classList.add("modal-open");
  focusModalForm(elements.categoryForm);
}

function resetCategoryForm() {
  elements.categoryForm.reset();
  elements.categoryIdInput.value = "";
  elements.categoryModalActions.hidden = true;
  elements.categoryDeleteButton.hidden = false;
  clearFieldErrors();
}

function saveCategoryFromForm(event) {
  event.preventDefault();
  clearFieldErrors();
  const name = elements.categoryNameInput.value.trim();
  const budget = parseMoneyInput(elements.categoryBudgetInput.value || "0");
  if (!name) {
    setFieldError(elements.categoryNameInput, "Bitte Kategorie eingeben.");
    return;
  }
  if (elements.categoryBudgetInput.value.trim() && (!Number.isFinite(budget) || budget < 0)) {
    setFieldError(elements.categoryBudgetInput, "Budget ist keine gültige Zahl.");
    return;
  }

  const id = elements.categoryIdInput.value;
  const duplicate = state.categories.find((category) => sameCategory(category.name, name) && category.id !== id);
  if (duplicate) {
    setFieldError(elements.categoryNameInput, "Diese Kategorie gibt es schon.");
    return;
  }
  if (id) updateCategory(id, { name, budget: Number.isFinite(budget) ? budget : 0 });
  else state.categories.push({ id: createId(), name, budget: Number.isFinite(budget) ? budget : 0 });
  persist();
  render();
  showMessage("Budget gespeichert.", "success");
  closeEditModal();
}

function renderEntries() {
  const month = elements.monthInput.value;
  const filter = elements.filterInput.value;
  const search = elements.transactionSearchInput.value.trim().toLowerCase();
  const entries = getEntriesForMonth(month)
    .filter((entry) => {
      if (filter === "all") return true;
      if (filter === "recurring") return isEntryRecurring(entry);
      if (filter === "open" || filter === "paid") return normalizeStatus(entry.status) === filter;
      return entry.type === filter;
    })
    .filter((entry) => {
      if (!search) return true;
      return [entry.description, entry.category, entry.payment, personName(entry.personId)]
        .some((value) => String(value || "").toLowerCase().includes(search));
    })
    .sort((a, b) => {
      const dateDiff = entryOccurrenceDate(b, month).localeCompare(entryOccurrenceDate(a, month));
      if (dateDiff !== 0) return dateDiff;
      const diff = Number(b.updatedAt || 0) - Number(a.updatedAt || 0);
      if (diff !== 0) return diff;
      if (a.type !== b.type) return a.type === "income" ? -1 : 1;
      return (a.description || "").localeCompare(b.description || "");
    });

  if (!entries.length) {
    elements.entryList.innerHTML = `<p class="empty-state">Keine passenden Buchungen für ${escapeHtml(selectedPersonName())} in diesem Monat.</p>`;
    return;
  }

  elements.entryList.innerHTML = groupedTransactionHtml(entries, month);
  wireEntryRowActions(elements.entryList);
}

function groupedTransactionHtml(entries, month) {
  const groups = new Map();
  entries.forEach((entry) => {
    const date = entryOccurrenceDate(entry, month);
    if (!groups.has(date)) groups.set(date, []);
    groups.get(date).push(entry);
  });
  return [...groups.entries()].map(([date, rows]) => {
    const net = rows.reduce((total, entry) => total + (entry.type === "income" ? Number(entry.amount || 0) : -Number(entry.amount || 0)), 0);
    const dailyTotal = rows.length > 1
      ? `<strong class="entry-day-total ${net >= 0 ? "positive-text" : "negative-text"}"><span>Tagesbilanz</span>${net >= 0 ? "+" : ""}${formatMoney.format(net)}</strong>`
      : "";
    return `
      <section class="entry-date-group">
        <header class="entry-date-group-title">
          <span>${formatDateFull(date)}</span>
          ${dailyTotal}
        </header>
        ${rows.map((entry) => transactionRowHtml(entry, month)).join("")}
      </section>
    `;
  }).join("");
}

function transactionRowHtml(entry, monthOrDate) {
  const month = monthOrDate.length === 10 ? monthOrDate.slice(0, 7) : monthOrDate;
  const sign = entry.type === "income" ? "+" : "-";
  const title = entry.description || entry.category;
  const meta = [personName(entry.personId), entry.category, displayText(entry.payment)].filter(Boolean).join(" · ");
  const badges = entryBadges(entry, month);
  const occurrence = monthOrDate.length === 10 ? monthOrDate : entryOccurrenceDate(entry, month);
  return `
      <article class="entry-row ${entry.type} actionable-row" role="button" tabindex="0" data-row-edit-kind="entry" data-row-edit-id="${escapeHtml(entry.id)}">
      <time class="entry-date" datetime="${occurrence}">${formatDate(occurrence)}</time>
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
}

function wireEntryRowActions(root) {
  root.querySelectorAll("[data-edit]").forEach((button) => {
    button.addEventListener("click", () => editEntry(button.dataset.edit));
  });
  root.querySelectorAll("[data-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteEntry(button.dataset.delete));
  });
  root.querySelectorAll("[data-debt-edit]").forEach((button) => {
    button.addEventListener("click", () => editDebt(button.dataset.debtEdit));
  });
  root.querySelectorAll("[data-debt-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteDebt(button.dataset.debtDelete));
  });
  root.querySelectorAll("[data-asset-edit]").forEach((button) => {
    button.addEventListener("click", () => editAsset(button.dataset.assetEdit));
  });
  root.querySelectorAll("[data-asset-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteAsset(button.dataset.assetDelete));
  });
  root.querySelectorAll("[data-row-edit-kind]").forEach((row) => {
    const open = () => {
      const { rowEditKind, rowEditId } = row.dataset;
      if (rowEditKind === "entry") editEntry(rowEditId);
      if (rowEditKind === "debt") editDebt(rowEditId);
      if (rowEditKind === "asset") editAsset(rowEditId);
    };
    row.addEventListener("click", (event) => {
      if (event.target.closest("button, a, input, select, textarea")) return;
      open();
    });
    row.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      if (event.target.closest("button, a, input, select, textarea")) return;
      event.preventDefault();
      open();
    });
  });
}

function renderDebts() {
  const month = elements.monthInput.value;
  const debts = getDebtsForMonth(month)
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
      <article class="debt-row ${active ? "active" : ""} actionable-row" role="button" tabindex="0" data-row-edit-kind="debt" data-row-edit-id="${escapeHtml(debt.id)}">
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

  wireEntryRowActions(elements.debtList);
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
  const assets = getAssetsForContext()
    .sort((a, b) => a.name.localeCompare(b.name));

  if (!assets.length) {
    elements.assetList.innerHTML = `<p class="empty-state">Kein Vermögen für ${escapeHtml(selectedPersonName())} eingetragen.</p>`;
    return;
  }

  elements.assetList.innerHTML = assets.map((asset) => `
    <article class="asset-row actionable-row" role="button" tabindex="0" data-row-edit-kind="asset" data-row-edit-id="${escapeHtml(asset.id)}">
      <div class="debt-main">
        <div class="row-heading">
          <strong>${escapeHtml(asset.name)}</strong>
          <span class="entry-controls inline-controls">
            <button class="icon-button" type="button" title="Bearbeiten" aria-label="Bearbeiten" data-asset-edit="${asset.id}">✎</button>
            <button class="icon-button" type="button" title="Löschen" aria-label="Löschen" data-asset-delete="${asset.id}">×</button>
          </span>
        </div>
        <span>${escapeHtml([personName(asset.personId), assetTypeLabel(asset), asset.note || "Plus-Wert"].filter(Boolean).join(" · "))}</span>
      </div>
      <strong class="asset-amount">${formatMoney.format(asset.amount || 0)}</strong>
    </article>
  `).join("");

  wireEntryRowActions(elements.assetList);
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
  prepareModalForm(elements.assetForm);
  elements.entryForm.hidden = true;
  elements.categoryForm.hidden = true;
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
    elements.assetTypeInput.value = normalizeAssetType(asset.type || inferAssetType(asset));
    elements.assetNameInput.value = asset.name;
    elements.assetAmountInput.value = formatMoneyInput(asset.amount);
    elements.assetNoteInput.value = asset.note || "";
  } else {
    elements.assetPersonInput.value = targetPersonIdForNewItem();
    elements.assetTypeInput.value = "cash";
  }

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
    type: normalizeAssetType(elements.assetTypeInput.value),
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
  state.assets = state.assets.filter((asset) => !sameLineage(asset, id));
  persist();
  render();
  if (activeEditContext?.kind === "asset" && sameLineage(activeEditContext, id)) closeEditModal();
  showMessage("Vermögen gelöscht.");
}

function renderSummary() {
  const month = elements.monthInput.value;
  const summary = summaryTotals(month);
  const remaining = summary.balance;
  const totalDebt = summary.totalDebt;
  const totalAssets = summary.totalAssets;
  const netWorth = summary.netWorth;

  elements.overviewMonthLabel.textContent = formatMonthName(month);
  elements.overviewIncomeValue.textContent = formatMoney.format(summary.income);
  elements.overviewExpenseValue.textContent = formatMoney.format(summary.expense);
  elements.heroIncomeValue.textContent = formatMoney.format(summary.income);
  elements.heroExpenseValue.textContent = formatMoney.format(summary.expense);
  elements.heroBudgetValue.textContent = formatMoney.format(calculateBudgetUsage(month, getActiveContext()).totalBudget || summary.income);
  elements.remainingValue.textContent = formatMoney.format(remaining);
  elements.remainingValue.nextElementSibling.textContent = remaining >= 0 ? "Bleibt am Monatsende" : "Über Budget";
  elements.remainingValue.parentElement.classList.toggle("positive", remaining >= 0);
  elements.remainingValue.parentElement.classList.toggle("negative", remaining < 0);
  elements.totalDebtValue.textContent = formatMoney.format(totalDebt);
  elements.totalAssetValue.textContent = formatMoney.format(totalAssets);
  elements.netWorthValue.textContent = formatMoney.format(netWorth);
  elements.accountsNetWorthValue.textContent = formatMoney.format(netWorth);
  elements.accountsDebtValue.textContent = formatMoney.format(totalDebt);
  elements.overviewAssetShortcutValue.textContent = formatMoney.format(totalAssets);
  elements.overviewDebtShortcutValue.textContent = formatMoney.format(totalDebt);
  elements.overviewAssetShortcutHint.textContent = `${getAssetsForContext().length} Wert${getAssetsForContext().length === 1 ? "" : "e"}`;
  const openDebts = getDebtsForMonth(month, getActiveContext(), { activeOnly: true }).length;
  elements.overviewDebtShortcutHint.textContent = `${openDebts} offen`;
  elements.netWorthValue.parentElement.classList.toggle("positive", netWorth >= 0);
  elements.netWorthValue.parentElement.classList.toggle("negative", netWorth < 0);
  renderMonthComparison(month);
  renderInsights(month, summary);
  if (activeSummaryType) renderSummaryBreakdown();
}

function renderInsights(month, summary) {
  const monthEntries = getEntriesForMonth(month);
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
    .filter(isEntryRecurring)
    .sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0))
    .slice(0, 3);
  const recurringTotal = monthEntries
    .filter((entry) => isEntryRecurring(entry) && entry.type === "expense")
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
  elements.overviewRecurringValue.textContent = formatMoney.format(recurringTotal);
  elements.overviewRecurringHint.textContent = recurringRows.length
    ? recurringRows.slice(0, 2).map((entry) => entry.description || entry.category).join(", ")
    : "keine aktiven Abos";
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

  const expenses = monthEntries.filter((entry) => entry.type === "expense");
  const biggestExpense = expenses
    .slice()
    .sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0))[0];
  elements.biggestExpenseValue.textContent = biggestExpense ? formatMoney.format(biggestExpense.amount) : formatMoney.format(0);
  elements.biggestExpenseHint.textContent = biggestExpense
    ? `${biggestExpense.description || biggestExpense.category} · ${formatDateFull(entryOccurrenceDate(biggestExpense, month))}`
    : "Noch keine Ausgaben";

  const topCategory = topCategories[0];
  elements.topSpendValue.textContent = topCategory ? formatMoney.format(topCategory[1]) : formatMoney.format(0);
  elements.topSpendHint.textContent = topCategory ? topCategory[0] : "Noch keine Kategorie";

  const nextPayment = nextPaymentForMonth(month);
  elements.nextPaymentValue.textContent = nextPayment ? formatMoney.format(nextPayment.amount) : formatMoney.format(0);
  elements.nextPaymentHint.textContent = nextPayment ? `${nextPayment.label} · ${formatDateFull(nextPayment.date)}` : "Keine geplanten Zahlungen";

  const weekSpend = weeklyExpenseTotal(localDateString(new Date()));
  elements.weekSpendValue.textContent = formatMoney.format(weekSpend);
  elements.weekSpendValue.classList.toggle("negative-text", weekSpend > 0);
  elements.weekSpendHint.textContent = weekSpend > 0 ? "Ausgaben der laufenden Woche" : "Diese Woche noch ruhig";
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

function renderCalendar() {
  const month = elements.monthInput.value;
  const selectedMonth = selectedCalendarDay.slice(0, 7) === month ? selectedCalendarDay : monthToDate(month);
  if (selectedCalendarDay.slice(0, 7) !== month) selectedCalendarDay = selectedMonth;
  elements.calendarTitle.textContent = formatMonthName(month);
  elements.calendarSubtitle.textContent = "Tage mit Buchungen werden hervorgehoben";

  const weekdays = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"];
  const [year, monthIndex] = month.split("-").map(Number);
  const days = daysInMonth(month);
  const firstDate = new Date(year, monthIndex - 1, 1);
  const leading = (firstDate.getDay() + 6) % 7;
  const cells = [];
  weekdays.forEach((day) => cells.push(`<div class="calendar-weekday">${day}</div>`));
  for (let index = 0; index < leading; index += 1) {
    cells.push(`<div class="calendar-day empty"></div>`);
  }
  for (let day = 1; day <= days; day += 1) {
    const date = `${month}-${String(day).padStart(2, "0")}`;
    const totals = dailyTotals(date);
    const net = totals.income - totals.expense;
    const isSelected = date === selectedCalendarDay;
    const isToday = date === localDateString(new Date());
    const weekday = new Date(`${date}T12:00:00`).getDay();
    const markers = [
      totals.income ? `<i class="calendar-marker income" title="Einnahmen"></i>` : "",
      totals.entryExpense ? `<i class="calendar-marker expense" title="Ausgaben"></i>` : "",
      totals.debtExpense ? `<i class="calendar-marker debt" title="Schulden"></i>` : "",
      totals.recurring ? `<i class="calendar-marker recurring" title="Wiederkehrend"></i>` : "",
    ].filter(Boolean).join("");
    cells.push(`
      <button class="calendar-day ${isSelected ? "selected" : ""} ${isToday ? "today" : ""} ${weekday === 0 || weekday === 6 ? "weekend" : ""} ${totals.income ? "has-income" : ""} ${totals.expense ? "has-expense" : ""} ${totals.debtExpense ? "has-debt" : ""}" type="button" data-calendar-day="${date}">
        <span>${day}</span>
        ${totals.income || totals.expense ? `<small class="${net >= 0 ? "positive-text" : "negative-text"}">${formatCompactMoney(net)}</small>` : ""}
        ${markers ? `<em class="calendar-markers" aria-hidden="true">${markers}</em>` : ""}
      </button>
    `);
  }
  elements.calendarGrid.innerHTML = cells.join("");
  elements.calendarGrid.querySelectorAll("[data-calendar-day]").forEach((button) => {
    button.addEventListener("click", () => {
      selectedCalendarDay = button.dataset.calendarDay;
      renderCalendar();
    });
  });
  renderCalendarDayList();
}

function renderCalendarDayList() {
  const daySummary = getCalendarDaySummary(selectedCalendarDay);
  const items = daySummary.items
    .sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0));
  const total = daySummary.net;
  elements.calendarDayTitle.textContent = formatDateFull(selectedCalendarDay);
  elements.calendarDayTotal.textContent = `${total >= 0 ? "+" : ""}${formatMoney.format(total)}`;
  elements.calendarDayTotal.classList.toggle("income-pill", total >= 0);
  elements.calendarDayTotal.classList.toggle("expense-pill", total < 0);
  if (!items.length) {
    elements.calendarDayList.innerHTML = `<p class="empty-state calendar-empty">Keine Buchungen an diesem Tag. Nutze +, wenn du etwas erfassen möchtest.</p>`;
    return;
  }
  elements.calendarDayList.innerHTML = items.map((item) => item.kind === "debt" ? calendarDebtRowHtml(item) : transactionRowHtml(item, selectedCalendarDay)).join("");
  wireEntryRowActions(elements.calendarDayList);
}

function dailyTotals(date) {
  return getCalendarDaySummary(date);
}

function calendarItemsForDate(date, context = getActiveContext()) {
  const month = date.slice(0, 7);
  const entries = entriesForDate(date, context)
    .map((entry) => ({ ...entry, kind: "entry" }));
  const debtItems = activeDebts(month)
    .filter((debt) => contextMatchesItem(debt, context))
    .filter((debt) => debtPaymentForMonth(debt, month) > 0 && debtOccurrenceDate(debt, month) === date)
    .map((debt) => ({
      kind: "debt",
      id: `debt-${debt.id}`,
      sourceId: debt.id,
      type: "expense",
      personId: debt.personId,
      date,
      category: "Schulden",
      description: debt.creditor || "Schuld",
      payment: debt.paymentMethod || "Rate",
      amount: debtPaymentForMonth(debt, month),
      updatedAt: Date.parse(`${date}T12:00:00`) || 0,
    }));
  return [...entries, ...debtItems];
}

function calendarDebtRowHtml(item) {
  const meta = [personName(item.personId), item.category, displayText(item.payment)].filter(Boolean).join(" · ");
  return `
    <article class="entry-row debt-calendar expense actionable-row" role="button" tabindex="0" data-row-edit-kind="debt" data-row-edit-id="${escapeHtml(item.sourceId)}">
      <time class="entry-date" datetime="${item.date}">${formatDate(item.date)}</time>
      <div class="entry-main">
        <div class="row-heading">
          <strong>${escapeHtml(item.description)}</strong>
          <span class="entry-controls inline-controls">
            <button class="icon-button" type="button" title="Schuld bearbeiten" aria-label="Schuld bearbeiten" data-debt-edit="${escapeHtml(item.sourceId)}">✎</button>
            <button class="icon-button" type="button" title="Schuld löschen" aria-label="Schuld löschen" data-debt-delete="${escapeHtml(item.sourceId)}">×</button>
          </span>
        </div>
        <span>${escapeHtml(meta)}</span>
        <div class="status-badges"><span class="status-badge status-carry">Schuldenrate</span></div>
      </div>
      <div class="entry-amount expense">-${formatMoney.format(item.amount)}</div>
    </article>
  `;
}

function entriesForDate(date, context = getActiveContext()) {
  const month = date.slice(0, 7);
  return getEntriesForMonth(month, context)
    .filter((entry) => entryOccurrenceDate(entry, month) === date);
}

function renderReports() {
  const month = elements.monthInput.value;
  const insights = calculateInsightSummary(month);
  const summary = insights.summary;
  const ratio = summary.income > 0 ? Math.round((summary.expense / summary.income) * 100) : 0;
  elements.reportsRatioValue.textContent = `${ratio} %`;
  elements.reportsRatioHint.textContent = `${formatMoney.format(summary.expense)} von ${formatMoney.format(summary.income)} ausgegeben`;
  elements.reportsNetWorthValue.textContent = formatMoney.format(summary.netWorth);
  elements.reportsNetWorthHint.textContent = `${formatMoney.format(summary.totalAssets)} Vermögen · ${formatMoney.format(summary.totalDebt)} Schulden`;
  elements.reportsDebtValue.textContent = formatMoney.format(summary.totalDebt);
  elements.reportsDebtHint.textContent = summary.debtMonthlyCost > 0
    ? `${formatMoney.format(summary.debtMonthlyCost)} monatliche Raten`
    : "keine aktiven Raten";
  elements.reportsMonthMini.textContent = formatMoney.format(summary.balance);

  elements.reportsOptimizationList.innerHTML = optimizationInsights(insights)
    .map((item) => insightLine(item.label, item.value, item.hint))
    .join("") || `<p class="empty-state compact-empty">Noch zu wenig Daten für Hinweise.</p>`;

  const frequent = [...insights.byDescription.entries()]
    .sort((a, b) => b[1].count - a[1].count || b[1].amount - a[1].amount)
    .slice(0, 3);
  elements.reportsFrequentList.innerHTML = frequent.length
    ? frequent.map(([name, item]) => insightLine(name, `${item.count}×`, formatMoney.format(item.amount))).join("")
    : `<p class="empty-state compact-empty">Keine Ausgaben gefunden.</p>`;

  elements.reportsLargestList.innerHTML = insights.largestPayments.length
    ? insights.largestPayments.map((entry) => insightLine(entry.description || entry.category, formatMoney.format(entry.amount), formatDateFull(entryOccurrenceDate(entry, month)))).join("")
    : `<p class="empty-state compact-empty">Keine Einzelzahlungen gefunden.</p>`;

  const months = Array.from({ length: 6 }, (_, index) => shiftMonth(month, index - 5));
  elements.reportsTrendList.innerHTML = months.map((itemMonth) => {
    const item = summaryTotals(itemMonth);
    return `
      <article class="summary-breakdown-row">
        <span>
          <strong>${formatMonthName(itemMonth)}</strong>
          <small>${formatMoney.format(item.income)} Einnahmen · ${formatMoney.format(item.expense)} Ausgaben</small>
        </span>
        <strong class="${item.balance >= 0 ? "positive-text" : "negative-text"}">${formatMoney.format(item.balance)}</strong>
      </article>
    `;
  }).join("");
}

function summaryTotals(month) {
  return calculateMonthSummary(month);
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
  const entries = getEntriesForMonth(month);
  const incomeRows = entries
    .filter((entry) => entry.type === "income")
    .map((entry) => entrySummaryRow(entry, Number(entry.amount || 0)));
  const expenseRows = entries
    .filter((entry) => entry.type === "expense")
    .map((entry) => entrySummaryRow(entry, Number(entry.amount || 0)));
  const debtRateRows = getDebtsForMonth(month, getActiveContext(), { activeOnly: true })
    .filter((debt) => Number(debt.monthlyPayment || 0) > 0)
    .map((debt) => debtSummaryRow(debt, Number(debt.monthlyPayment || 0), "monatliche Rate"));
  const debtRows = getDebtsForMonth(month)
    .map((debt) => debtSummaryRow(debt, remainingDebt(debt, month), "Restschuld"));
  const assetRows = getAssetsForContext()
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
    meta: [personName(entry.personId), entry.category, displayText(entry.payment), recurrenceLabel(entry), statusLabel(entry.status)].filter(Boolean).join(" · "),
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
  elements.recurrenceInput.value = entryRecurrence(entry);
  elements.entryEndInput.value = entry.endDate || "";
  elements.entryStatusInput.value = normalizeStatus(entry.status);
  syncFormMode();
}

function deleteEntry(id) {
  state.entries = state.entries.filter((entry) => !sameLineage(entry, id));
  persist();
  render();
  if (activeEditContext?.kind === "entry" && sameLineage(activeEditContext, id)) closeEditModal();
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
  state.debts = state.debts.filter((debt) => !sameLineage(debt, id));
  persist();
  render();
  if (activeEditContext?.kind === "debt" && sameLineage(activeEditContext, id)) closeEditModal();
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
  } else if (kind === "category") {
    removeCategory(id);
    showMessage("Budget gelöscht.");
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
      bankConnections: getBankConnections(),
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
        state.bankConnections = Array.isArray(imported.bankConnections) ? imported.bankConnections.map(normalizeBankConnection) : [];
        state.people = Array.isArray(imported.people) ? imported.people.map(normalizePerson) : state.people;
        showMessage("Backup importiert.");
      }
      persist();
      ensurePeople();
      repairCopiedItems();
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
        recurrence: normalizeRecurrence(item["wiederholung"] || item["rhythmus"] || item["intervall"] || "", /ja|true|monat/i.test(item["monatlich"] || "")),
        recurring: normalizeRecurrence(item["wiederholung"] || item["rhythmus"] || item["intervall"] || "", /ja|true|monat/i.test(item["monatlich"] || "")) !== "none",
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
        recurrence: "monthly",
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
          liabilityType: "other",
          note: "",
        });
        importedDebts += 1;
      } else {
        state.assets.push({
          id: createId(),
          personId: defaultPersonId(),
          type: inferAssetType({ name: rightName, note: "aus Excel" }),
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
    state.bankConnections = [];
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
  [elements.entryForm, elements.assetForm, elements.categoryForm].forEach((form) => {
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

function totalsBy(items, keyFn) {
  const map = new Map();
  items.forEach((item) => {
    const key = keyFn(item);
    const current = map.get(key) || { amount: 0, count: 0 };
    current.amount += Number(item.amount || 0);
    current.count += 1;
    map.set(key, current);
  });
  return map;
}

function topMapEntry(map) {
  const rows = [...map.entries()].sort((a, b) => b[1].amount - a[1].amount);
  return rows[0] || null;
}

function normalizeAssetType(type) {
  return ASSET_TYPES[type] ? type : "other";
}

function inferAssetType(asset) {
  const text = `${asset.name || ""} ${asset.note || ""}`.toLowerCase();
  if (/bar|cash|kasse/.test(text)) return "cash";
  if (/spar|rücklage|tagesgeld|reserve/.test(text)) return "savings";
  if (/depot|etf|aktie|fonds|krypto|investment|invest/.test(text)) return "investment";
  if (/auto|immobil|wohnung|haus|sachwert|gold/.test(text)) return "property";
  if (/konto|bank|giro|n26|sparkasse|ing/.test(text)) return "bank";
  return "other";
}

function assetTypeLabel(asset) {
  return ASSET_TYPES[normalizeAssetType(asset.type || inferAssetType(asset))] || ASSET_TYPES.other;
}

function inferLiabilityType(debt) {
  const text = `${debt.creditor || ""} ${debt.note || ""} ${debt.account || ""}`.toLowerCase();
  if (/kreditkarte|credit card|visa|mastercard/.test(text)) return "credit_card";
  if (/rate|raten|finanzierung/.test(text)) return "installment";
  if (/privat|freund|familie/.test(text)) return "private_debt";
  if (/kredit|loan|bank/.test(text)) return "loan";
  return "other";
}

function accountTypeInitial(label) {
  return String(label || "?").slice(0, 1).toUpperCase();
}

function categoryGroup(name) {
  const key = String(name || "").toLowerCase();
  if (/miete|wohnen|strom|gas|nebenkosten/.test(key)) return "Wohnen";
  if (/lebensmittel|essen|food|supermarkt/.test(key)) return "Lebensmittel";
  if (/mobil|transport|auto|bahn|taxi|fuel|benzin/.test(key)) return "Mobilität";
  if (/abo|stream|software|subscription/.test(key)) return "Abos";
  if (/versicherung/.test(key)) return "Versicherungen";
  if (/gesund|arzt|apotheke/.test(key)) return "Gesundheit";
  if (/bildung|kurs|schule|uni|buch/.test(key)) return "Bildung";
  if (/freizeit|kino|restaurant|kaffee|hobby/.test(key)) return "Freizeit";
  if (/reise|urlaub|hotel|flug/.test(key)) return "Reisen";
  if (/familie|kind/.test(key)) return "Familie";
  if (/schuld|kredit|rate/.test(key)) return "Schulden";
  if (/einkommen|gehalt|lohn/.test(key)) return "Einkommen";
  if (/invest|depot|etf|aktie/.test(key)) return "Investments";
  return "Sonstiges";
}

function budgetStatus(spent, budget) {
  if (budget <= 0 && spent <= 0) return { label: "bereit", className: "budget-neutral", badgeClass: "status-recurring" };
  if (budget <= 0 && spent > 0) return { label: "ohne Limit", className: "budget-watch", badgeClass: "status-open" };
  const ratio = spent / budget;
  if (ratio > 1) return { label: "überschritten", className: "budget-over", badgeClass: "status-open" };
  if (ratio >= 0.8) return { label: "knapp", className: "budget-watch", badgeClass: "status-carry" };
  return { label: "gut", className: "budget-good", badgeClass: "status-paid" };
}

function optimizationInsights(insights) {
  const rows = [];
  if (insights.topCategory) {
    rows.push({
      label: `${insights.topCategory[0]} ist am größten`,
      value: formatMoney.format(insights.topCategory[1].amount),
      hint: "prüfe hier zuerst dein Budget",
    });
  }
  if (insights.recurringCount) {
    rows.push({
      label: "Wiederkehrende Kosten",
      value: `${insights.recurringShare} %`,
      hint: `${insights.recurringCount} aktive Zahlung${insights.recurringCount === 1 ? "" : "en"}`,
    });
  }
  const daily = daysLeftInMonth(elements.monthInput.value) > 0
    ? insights.summary.balance / daysLeftInMonth(elements.monthInput.value)
    : insights.summary.balance;
  rows.push({
    label: "Rest pro Tag",
    value: formatMoney.format(daily),
    hint: daily >= 0 ? "aktueller Puffer" : "Budget ist überzogen",
  });
  return rows.slice(0, 3);
}

function formatNullableDate(value) {
  if (!value) return "noch nie";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "unbekannt";
  return new Intl.DateTimeFormat("de-DE", { dateStyle: "medium" }).format(date);
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
  return Math.min(Number(debt.totalAmount || 0), paidDebtBeforeMonth(debt, month) + debtPaymentForMonth(debt, month));
}

function paidDebtBeforeMonth(debt, month) {
  if (!debt.startDate || monthToNumber(month) < dateToMonthNumber(debt.startDate)) return Math.min(Number(debt.totalAmount || 0), Number(debt.paidSoFar || 0));
  const cappedMonth = debt.endDate && monthToNumber(month) > dateToMonthNumber(debt.endDate) ? debt.endDate.slice(0, 7) : month;
  const elapsedMonthsBeforeSelected = Math.max(0, monthsBetween(debt.startDate.slice(0, 7), cappedMonth));
  const scheduledPaid = elapsedMonthsBeforeSelected * Number(debt.monthlyPayment || 0);
  return Math.min(Number(debt.totalAmount || 0), Number(debt.paidSoFar || 0) + scheduledPaid);
}

function debtPaymentForMonth(debt, month) {
  if (!isDebtVisibleInMonth(debt, month) || isDebtClosed(debt)) return 0;
  const monthlyPayment = Number(debt.monthlyPayment || 0);
  if (monthlyPayment <= 0) return 0;
  const remainingBeforeMonth = Math.max(0, Number(debt.totalAmount || 0) - paidDebtBeforeMonth(debt, month));
  return Math.min(monthlyPayment, remainingBeforeMonth);
}

function paidMonthsUntil(debt, month) {
  if (!debt.startDate || monthToNumber(month) < dateToMonthNumber(debt.startDate)) return 0;
  const cappedMonth = debt.endDate && monthToNumber(month) > dateToMonthNumber(debt.endDate) ? debt.endDate.slice(0, 7) : month;
  return Math.max(0, monthsBetween(debt.startDate.slice(0, 7), cappedMonth) + 1);
}

function isDebtActiveInMonth(debt, month) {
  return isDebtVisibleInMonth(debt, month) && !isDebtClosed(debt) && debtPaymentForMonth(debt, month) > 0;
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

function selectedRecurrence() {
  return normalizeRecurrence(elements.recurrenceInput?.value || "none");
}

function normalizeRecurrence(value, legacyRecurring = false) {
  const normalized = String(value || "").trim().toLowerCase();
  if (["daily", "taeglich", "täglich"].includes(normalized)) return "daily";
  if (["weekly", "woechentlich", "wöchentlich"].includes(normalized)) return "weekly";
  if (["monthly", "monatlich", "true", "ja"].includes(normalized)) return "monthly";
  if (["yearly", "annual", "jaehrlich", "jährlich"].includes(normalized)) return "yearly";
  return legacyRecurring ? "monthly" : "none";
}

function entryRecurrence(entry) {
  return normalizeRecurrence(entry?.recurrence, Boolean(entry?.recurring));
}

function isEntryRecurring(entry) {
  return entryRecurrence(entry) !== "none";
}

function isEntryActiveInMonth(entry, month) {
  return expandEntryOccurrences(entry, month).length > 0;
}

function expandEntryOccurrences(entry, month) {
  const dates = occurrenceDatesForEntryInMonth(entry, month);
  return dates.map((date) => ({
    ...entry,
    sourceId: entry.sourceId || entry.id,
    occurrenceDate: date,
    recurring: isEntryRecurring(entry),
    recurrence: entryRecurrence(entry),
  }));
}

function occurrenceDatesForEntryInMonth(entry, month) {
  const startDate = normalizeDate(entry.date || "");
  if (!startDate) return [];
  const recurrence = entryRecurrence(entry);
  if (normalizeStatus(entry.status) === "ended" && !entry.endDate && recurrence !== "none") {
    return startDate.startsWith(month) ? [startDate] : [];
  }
  const endDate = normalizeDate(entry.endDate || "") || "";
  const monthStart = monthToDate(month);
  const monthEnd = `${month}-${String(daysInMonth(month)).padStart(2, "0")}`;
  if (compareDates(monthEnd, startDate) < 0) return [];
  if (endDate && compareDates(monthStart, endDate) > 0) return [];

  if (recurrence === "none") return startDate.startsWith(month) ? [startDate] : [];

  const rangeStart = compareDates(startDate, monthStart) > 0 ? startDate : monthStart;
  const rangeEnd = endDate && compareDates(endDate, monthEnd) < 0 ? endDate : monthEnd;
  if (compareDates(rangeEnd, rangeStart) < 0) return [];

  if (recurrence === "daily") return datesBetween(rangeStart, rangeEnd);

  if (recurrence === "weekly") {
    const offset = positiveModulo(dayDiff(startDate, rangeStart), 7);
    let current = offset === 0 ? rangeStart : addDays(rangeStart, 7 - offset);
    const dates = [];
    while (compareDates(current, rangeEnd) <= 0) {
      dates.push(current);
      current = addDays(current, 7);
    }
    return dates;
  }

  if (recurrence === "monthly") {
    const occurrence = monthlyOccurrenceDate(startDate, month);
    return compareDates(occurrence, rangeStart) >= 0 && compareDates(occurrence, rangeEnd) <= 0 ? [occurrence] : [];
  }

  if (recurrence === "yearly") {
    const [, startMonth] = startDate.split("-").map(Number);
    const [, targetMonth] = month.split("-").map(Number);
    if (startMonth !== targetMonth) return [];
    const occurrence = monthlyOccurrenceDate(startDate, month);
    return compareDates(occurrence, rangeStart) >= 0 && compareDates(occurrence, rangeEnd) <= 0 ? [occurrence] : [];
  }

  return [];
}

function entryOccurrenceDate(entry, month) {
  if (entry.occurrenceDate) return entry.occurrenceDate;
  const dates = occurrenceDatesForEntryInMonth(entry, month);
  return dates[0] || entry.date;
}

function debtOccurrenceDate(debt, month) {
  const sourceDate = debt.startDate || monthToDate(month);
  const day = Math.min(Number(sourceDate.slice(8, 10)) || 1, daysInMonth(month));
  return `${month}-${String(day).padStart(2, "0")}`;
}

function entryBadges(entry, month) {
  return [
    statusBadge(statusLabel(entry.status), `status-${normalizeStatus(entry.status)}`),
    isEntryRecurring(entry) ? statusBadge(recurrenceLabel(entry), "status-recurring") : "",
    isEntryRecurring(entry) && entry.endDate ? statusBadge(`endet ${formatMonthName(entry.endDate.slice(0, 7))}`, "status-ended") : "",
    isEntryRecurring(entry) && !entryOccurrenceDate(entry, month).startsWith(entry.date?.slice(0, 7) || "") ? statusBadge("übernommen", "status-carry") : "",
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

function monthlyOccurrenceDate(startDate, month) {
  const day = Math.min(Number(startDate?.slice(8, 10)) || 1, daysInMonth(month));
  return `${month}-${String(day).padStart(2, "0")}`;
}

function compareDates(a, b) {
  return String(a || "").localeCompare(String(b || ""));
}

function parseDateParts(value) {
  const [year, month, day] = String(value || "").split("-").map(Number);
  return { year, month, day };
}

function dateToEpochDay(value) {
  const { year, month, day } = parseDateParts(value);
  return Math.floor(Date.UTC(year, month - 1, day) / 86400000);
}

function epochDayToDate(epochDay) {
  const date = new Date(epochDay * 86400000);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function dayDiff(startDate, endDate) {
  return dateToEpochDay(endDate) - dateToEpochDay(startDate);
}

function addDays(date, amount) {
  return epochDayToDate(dateToEpochDay(date) + amount);
}

function datesBetween(startDate, endDate) {
  const dates = [];
  let current = startDate;
  while (compareDates(current, endDate) <= 0) {
    dates.push(current);
    current = addDays(current, 1);
  }
  return dates;
}

function positiveModulo(value, divisor) {
  return ((value % divisor) + divisor) % divisor;
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
  const recurring = getEntriesForMonth(nextMonth)
    .filter(isEntryRecurring)
    .map((entry) => ({ label: entry.description || entry.category || "Buchung" }));
  const debts = getDebtsForMonth(nextMonth, getActiveContext(), { activeOnly: true })
    .map((debt) => ({ label: debt.creditor || "Schuld" }));
  return [...recurring, ...debts];
}

function nextPaymentForMonth(month) {
  const today = localDateString(new Date());
  const candidates = [
    ...getEntriesForMonth(month)
      .filter((entry) => isEntryRecurring(entry) && entry.type === "expense")
      .map((entry) => ({
        date: entryOccurrenceDate(entry, month),
        amount: Number(entry.amount || 0),
        label: entry.description || entry.category || "Wiederkehrend",
      })),
    ...getDebtsForMonth(month, getActiveContext(), { activeOnly: true })
      .filter((debt) => Number(debt.monthlyPayment || 0) > 0)
      .map((debt) => ({
        date: debtOccurrenceDate(debt, month),
        amount: Number(debt.monthlyPayment || 0),
        label: debt.creditor || "Schuld",
      })),
  ].filter((item) => item.date >= (today.slice(0, 7) === month ? today : monthToDate(month)));
  return candidates.sort((a, b) => a.date.localeCompare(b.date) || b.amount - a.amount)[0] || null;
}

function weeklyExpenseTotal(date) {
  const anchor = new Date(`${date}T12:00:00`);
  const day = (anchor.getDay() + 6) % 7;
  const start = new Date(anchor);
  start.setDate(anchor.getDate() - day);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const startDate = localDateString(start);
  const endDate = localDateString(end);
  return getEntriesForMonth(date.slice(0, 7))
    .filter((entry) => entry.type === "expense")
    .filter((entry) => {
      const occurrence = entryOccurrenceDate(entry, date.slice(0, 7));
      return occurrence >= startDate && occurrence <= endDate;
    })
    .reduce((total, entry) => total + Number(entry.amount || 0), 0);
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

function localDateString(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatCompactMoney(value) {
  const number = Number(value || 0);
  const sign = number > 0 ? "+" : "";
  return `${sign}${new Intl.NumberFormat("de-DE", {
    maximumFractionDigits: 0,
  }).format(number)} €`;
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

function recurrenceLabel(entryOrValue) {
  const value = typeof entryOrValue === "string" ? normalizeRecurrence(entryOrValue) : entryRecurrence(entryOrValue);
  return value === "none" ? "" : RECURRENCE_LABELS[value] || "wiederkehrend";
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
    liabilityType: debt.liabilityType || inferLiabilityType(debt),
    principalAmount: Number(debt.principalAmount || 0),
    interestAmount: Number(debt.interestAmount || 0),
    nominalRate: Number(debt.nominalRate || 0),
    termMonths: Number(debt.termMonths || 0),
    finalPayment: Number(debt.finalPayment || 0),
    note: debt.note || "",
  };
}

function normalizeAsset(asset) {
  const type = normalizeAssetType(asset.type || inferAssetType(asset));
  return {
    id: asset.id || createId(),
    personId: asset.personId || fallbackPersonId(),
    duplicateOf: asset.duplicateOf || "",
    type,
    name: asset.name || "",
    amount: Number(asset.amount || 0),
    note: asset.note || "",
  };
}

function normalizeBankConnection(connection) {
  return {
    id: connection.id || createId(),
    provider: connection.provider || "",
    status: connection.status || "not_connected",
    lastSyncAt: connection.lastSyncAt || "",
    consentExpiresAt: connection.consentExpiresAt || "",
    accountIds: Array.isArray(connection.accountIds) ? connection.accountIds : [],
  };
}

function normalizeEntry(entry) {
  const date = entry.date || monthToDate(currentLocalMonth());
  const fallbackUpdatedAt = Date.parse(`${date}T12:00:00`) || 0;
  const recurrence = normalizeRecurrence(entry.recurrence || entry.repeatInterval || entry.repeat || "", Boolean(entry.recurring));
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
    recurrence,
    recurring: recurrence !== "none",
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
  if (!id || id === defaultPersonId()) return "Gemeinsam";
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
  return contextMatchesItem(item);
}

function selectedMonthlyTotals(month) {
  return calculateMonthSummary(month);
}

function selectedDebtTotal(month) {
  return getDebtsForMonth(month)
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

function deleteEntryLineageForTest(id) {
  state.entries = state.entries.filter((entry) => !sameLineage(entry, id));
}

function deleteDebtLineageForTest(id) {
  state.debts = state.debts.filter((debt) => !sameLineage(debt, id));
}

function deleteAssetLineageForTest(id) {
  state.assets = state.assets.filter((asset) => !sameLineage(asset, id));
}

function runBudgetUpSelfTest() {
  const snapshot = clone({
    categories: state.categories,
    entries: state.entries,
    debts: state.debts,
    assets: state.assets,
    bankConnections: state.bankConnections,
    people: state.people,
    selectedPersonId: state.selectedPersonId,
  });
  const previousMonth = elements.monthInput.value;
  const previousSelectedCalendarDay = selectedCalendarDay;
  const results = [];
  const assert = (name, condition, detail = "") => {
    results.push({ name, pass: Boolean(condition), detail });
  };
  const closeEnough = (actual, expected, tolerance = 0.01) => Math.abs(actual - expected) <= tolerance;

  try {
    state.people = [
      { id: "shared-test", name: "Gemeinsam" },
      { id: "p1-test", name: "Test Person" },
      { id: "p2-test", name: "Andere Person" },
      { id: "debt-test", name: "Schulden Verlauf" },
      { id: "simple-test", name: "Einfacher Fall" },
      { id: "case-test", name: "Faruk" },
      { id: "case-copy", name: "Faruk Kopie" },
    ];
    state.selectedPersonId = "p1-test";
    state.categories = [
      { id: "cat-food", name: "Lebensmittel", budget: 300 },
      { id: "cat-income", name: "Einkommen", budget: 0 },
      { id: "cat-other", name: "Sonstiges", budget: 100 },
    ];
    state.entries = [
      { id: "e-income", personId: "p1-test", type: "income", date: "2026-06-01", category: "Einkommen", description: "Gehalt", payment: "Überweisung", amount: 1000, recurring: false, endDate: "", status: "open", updatedAt: 1 },
      { id: "e-expense", personId: "p1-test", type: "expense", date: "2026-06-05", category: "Lebensmittel", description: "Einkauf", payment: "Karte", amount: 100, recurring: false, endDate: "", status: "open", updatedAt: 2 },
      { id: "e-recurring", personId: "p1-test", type: "expense", date: "2026-05-10", category: "Sonstiges", description: "Abo", payment: "Karte", amount: 50, recurring: true, endDate: "", status: "open", updatedAt: 3 },
      { id: "e-ended", personId: "p1-test", type: "expense", date: "2026-05-15", category: "Sonstiges", description: "Endendes Abo", payment: "Karte", amount: 25, recurring: true, endDate: "2026-06-15", status: "open", updatedAt: 4 },
      { id: "e-other", personId: "p2-test", type: "expense", date: "2026-06-05", category: "Sonstiges", description: "Andere Person", payment: "Karte", amount: 999, recurring: false, endDate: "", status: "open", updatedAt: 5 },
    ];
    state.debts = [
      { id: "d-open", personId: "p1-test", creditor: "Bank", totalAmount: 500, paidSoFar: 100, monthlyPayment: 50, startDate: "2026-06-03", endDate: "", status: "open", paymentMethod: "Überweisung", account: "", principalAmount: 0, interestAmount: 0, nominalRate: 0, termMonths: 0, finalPayment: 0, note: "" },
      { id: "d-paid", personId: "p1-test", creditor: "Paid Bank", totalAmount: 300, paidSoFar: 300, monthlyPayment: 30, startDate: "2026-06-03", endDate: "", status: "paid", paymentMethod: "Überweisung", account: "", principalAmount: 0, interestAmount: 0, nominalRate: 0, termMonths: 0, finalPayment: 0, note: "" },
      { id: "d-final", personId: "debt-test", creditor: "Final Rate", totalAmount: 120, paidSoFar: 0, monthlyPayment: 50, startDate: "2026-06-15", endDate: "", status: "open", paymentMethod: "Überweisung", account: "", principalAmount: 0, interestAmount: 0, nominalRate: 0, termMonths: 0, finalPayment: 0, note: "" },
    ];
    state.assets = [
      { id: "a-cash", personId: "p1-test", name: "Bargeld", amount: 200, note: "" },
      { id: "a-other", personId: "p2-test", name: "Fremd", amount: 999, note: "" },
    ];
    elements.monthInput.value = "2026-06";

    const context = getActiveContext("p1-test");
    const june = calculateMonthSummary("2026-06", context);
    const july = calculateMonthSummary("2026-07", context);
    const budget = calculateBudgetUsage("2026-06", context);
    const dayExpense = getCalendarDaySummary("2026-06-05", context);
    const dayRecurring = getCalendarDaySummary("2026-06-10", context);
    const dayDebt = getCalendarDaySummary("2026-06-03", context);
    const family = calculateMonthSummary("2026-06", getActiveContext("all"));

    assert("einmalige Einnahme", june.income === 1000, `income=${june.income}`);
    assert("einmalige und wiederkehrende Ausgaben", june.entryExpense === 175, `entryExpense=${june.entryExpense}`);
    assert("offene Schuld als Monatskosten", june.debtMonthlyCost === 50, `debtMonthlyCost=${june.debtMonthlyCost}`);
    assert("Saldo korrekt", june.balance === 775, `balance=${june.balance}`);
    assert("wiederkehrende Ausgabe ohne Enddatum läuft weiter", july.entryExpense === 50, `julyEntryExpense=${july.entryExpense}`);
    assert("wiederkehrende Ausgabe mit Enddatum endet", !getEntriesForMonth("2026-07", context).some((entry) => entry.id === "e-ended"), "ended recurring should be absent in July");
    assert("bezahlte Schuld wird nicht fortgeführt", !getDebtsForMonth("2026-06", context, { activeOnly: true }).some((debt) => debt.id === "d-paid"), "paid debt should be inactive");
    assert("offene Restschuld korrekt", june.totalDebt === 350, `totalDebt=${june.totalDebt}`);
    assert("Vermögen Kontext korrekt", june.totalAssets === 200, `totalAssets=${june.totalAssets}`);
    assert("Nettovermögen korrekt", june.netWorth === -150, `netWorth=${june.netWorth}`);
    assert("Kontext Person filtert andere Person", june.entryExpense !== family.entryExpense && family.entryExpense === 1174, `familyEntryExpense=${family.entryExpense}`);
    assert("Budget geplant/ausgegeben/übrig", budget.totalBudget === 400 && budget.totalSpent === 175 && budget.remaining === 225 && budget.spentByCategory.get("Lebensmittel") === 100 && budget.spentByCategory.get("Sonstiges") === 75, `totalSpent=${budget.totalSpent}, remaining=${budget.remaining}`);
    assert("Kalender Tageswert Ausgabe", dayExpense.expense === 100 && dayExpense.net === -100, JSON.stringify(dayExpense));
    assert("Kalender Tageswert wiederkehrend", dayRecurring.expense === 50 && dayRecurring.recurring, JSON.stringify(dayRecurring));
    assert("Kalender Tageswert Schuld", dayDebt.debtExpense === 50 && dayDebt.net === -50, JSON.stringify(dayDebt));
    const debtContext = getActiveContext("debt-test");
    const finalDebtJune = calculateMonthSummary("2026-06", debtContext);
    const finalDebtJuly = calculateMonthSummary("2026-07", debtContext);
    const finalDebtAugust = calculateMonthSummary("2026-08", debtContext);
    const finalDebtSeptember = calculateMonthSummary("2026-09", debtContext);
    assert("Schuldenrate reduziert Restschuld monatlich", finalDebtJune.debtMonthlyCost === 50 && finalDebtJune.totalDebt === 70 && finalDebtJuly.debtMonthlyCost === 50 && finalDebtJuly.totalDebt === 20, JSON.stringify({ finalDebtJune, finalDebtJuly }));
    assert("Letzte Schuldenrate wird auf Restbetrag begrenzt", finalDebtAugust.debtMonthlyCost === 20 && finalDebtAugust.totalDebt === 0 && finalDebtSeptember.debtMonthlyCost === 0 && finalDebtSeptember.totalDebt === 0, JSON.stringify({ finalDebtAugust, finalDebtSeptember }));
    assert("Kalender zeigt letzte Schuldenrate nur mit Restbetrag", getCalendarDaySummary("2026-08-15", debtContext).debtExpense === 20, JSON.stringify(getCalendarDaySummary("2026-08-15", debtContext)));
    render();
    assert("Schuldenzeile ist direkt bearbeitbar", Boolean(elements.debtList.querySelector('[data-row-edit-kind="debt"][data-row-edit-id="d-open"]')), elements.debtList.innerHTML);
    assert("Vermoegenszeile ist direkt bearbeitbar", Boolean(elements.assetList.querySelector('[data-row-edit-kind="asset"][data-row-edit-id="a-cash"]')), elements.assetList.innerHTML);
    selectedCalendarDay = "2026-06-03";
    renderCalendarDayList();
    assert("Kalender-Schuldenrate ist direkt bearbeitbar", Boolean(elements.calendarDayList.querySelector('[data-row-edit-kind="debt"][data-row-edit-id="d-open"]')), elements.calendarDayList.innerHTML);

    const quickExpense = parseQuickAdd("12,50 Kaffee Lebensmittel", "expense", context, "2026-06");
    const quickIncome = parseQuickAdd("200 Upwork Einkommen", "income", context, "2026-06");
    const quickDebt = parseQuickAdd("500 Targobank Kredit 50 monatlich", "debt", context, "2026-06");
    const quickAsset = parseQuickAdd("300 Bargeld", "asset", context, "2026-06");
    assert("QuickAdd Ausgabe Parser", quickExpense?.amount === 12.5 && quickExpense.category === "Lebensmittel" && quickExpense.description === "Kaffee", JSON.stringify(quickExpense));
    assert("QuickAdd Einnahme Parser", quickIncome?.amount === 200 && quickIncome.category === "Einkommen" && quickIncome.description === "Upwork", JSON.stringify(quickIncome));
    assert("QuickAdd Schuld Parser", quickDebt?.amount === 500 && quickDebt.monthlyPayment === 50 && quickDebt.description === "Targobank Kredit", JSON.stringify(quickDebt));
    assert("QuickAdd Vermögen Parser", quickAsset?.amount === 300 && quickAsset.description === "Bargeld", JSON.stringify(quickAsset));

    state.entries.push(
      { id: "simple-income", personId: "simple-test", type: "income", date: "2026-05-01", category: "Einkommen", description: "Gehalt", payment: "Überweisung", amount: 2000, recurring: false, endDate: "", status: "open", updatedAt: 9 },
      { id: "simple-expense", personId: "simple-test", type: "expense", date: "2026-05-02", category: "Wohnen", description: "Miete", payment: "Karte", amount: 500, recurring: false, endDate: "", status: "open", updatedAt: 10 },
      { id: "case-income", personId: "case-test", type: "income", date: "2026-05-01", category: "Einkommen", description: "Gehalt Faruk", payment: "Überweisung", amount: 2000, recurring: false, endDate: "", status: "open", updatedAt: 11 },
      { id: "case-expense", personId: "case-test", type: "expense", date: "2026-05-01", category: "Wohnen", description: "Miete Faruk", payment: "Karte", amount: 500, recurring: false, endDate: "", status: "open", updatedAt: 11 },
      { id: "case-o2", personId: "case-test", type: "expense", date: "2026-05-21", category: "Abos", description: "O2 Tarif", payment: "Lastschrift", amount: 27.99, recurring: true, endDate: "2028-01-21", status: "open", updatedAt: 12 },
      { id: "case-income-copy", personId: "case-copy", type: "income", date: "2026-05-01", category: "Einkommen", description: "Gehalt Faruk", payment: "Überweisung", amount: 2000, recurring: false, endDate: "", status: "open", updatedAt: 13, duplicateOf: "case-income" },
      { id: "case-expense-copy", personId: "case-copy", type: "expense", date: "2026-05-01", category: "Wohnen", description: "Miete Faruk", payment: "Karte", amount: 500, recurring: false, endDate: "", status: "open", updatedAt: 14, duplicateOf: "case-expense" }
    );
    state.debts.push({ id: "case-debt", personId: "case-test", creditor: "Test Bank", totalAmount: 1000, paidSoFar: 0, monthlyPayment: 100, startDate: "2026-05-01", endDate: "", status: "open", paymentMethod: "Überweisung", account: "", principalAmount: 0, interestAmount: 0, nominalRate: 0, termMonths: 0, finalPayment: 0, note: "" });
    state.assets.push(
      { id: "case-asset", personId: "case-test", name: "Depot", amount: 10000, note: "" },
      { id: "case-asset-copy", personId: "case-copy", name: "Depot", amount: 10000, note: "", duplicateOf: "case-asset" }
    );

    const simpleMay = calculateMonthSummary("2026-05", getActiveContext("simple-test"));
    const caseContext = getActiveContext("case-test");
    const caseJune = calculateMonthSummary("2026-06", caseContext);
    const caseApril = calculateMonthSummary("2026-04", caseContext);
    const caseFebruary2028 = calculateMonthSummary("2028-02", caseContext);
    const caseMayDay = getCalendarDaySummary("2026-05-01", caseContext);
    assert("Fall A Einnahmen/Ausgaben/Saldo", simpleMay.income === 2000 && simpleMay.entryExpense === 500 && simpleMay.balance === 1500, JSON.stringify(simpleMay));
    assert("Fall B wiederkehrend startet und endet korrekt", closeEnough(caseJune.entryExpense, 27.99) && caseApril.entryExpense === 0 && caseFebruary2028.entryExpense === 0, JSON.stringify({ caseJune, caseApril, caseFebruary2028 }));
    assert("Fall C Tagesbilanz ohne Doppelzählung", caseMayDay.income === 2000 && caseMayDay.entryExpense === 500 && caseMayDay.debtExpense === 100 && caseMayDay.net === 1400 && caseMayDay.items.length === 3, JSON.stringify(caseMayDay));

    state.entries.push({ ...state.entries[1], id: "e-expense-copy", personId: "p2-test", duplicateOf: "e-expense" });
    state.debts.push({ ...state.debts[0], id: "d-open-copy", personId: "p2-test", duplicateOf: "d-open" });
    state.assets.push({ ...state.assets[0], id: "a-cash-copy", personId: "p2-test", duplicateOf: "a-cash" });
    const familyBeforeDelete = calculateMonthSummary("2026-06", getActiveContext("all"));
    assert("Gesamt zählt Kopien nicht doppelt", closeEnough(familyBeforeDelete.entryExpense, 1201.99) && familyBeforeDelete.totalDebt === 1220 && familyBeforeDelete.totalAssets === 11199, JSON.stringify(familyBeforeDelete));
    const editedCopy = preserveCopyTracking({ ...state.entries.find((entry) => entry.id === "case-income-copy"), amount: 2100 }, state.entries.find((entry) => entry.id === "case-income-copy"));
    assert("Bearbeiten bewahrt duplicateOf", editedCopy.duplicateOf === "case-income", JSON.stringify(editedCopy));
    assert("Fall D/E Kopien bleiben aus Gesamt heraus", calculateMonthSummary("2026-05", getActiveContext("case-copy")).income === 2000 && calculateMonthSummary("2026-05", getActiveContext("all")).totalAssets === 11199, JSON.stringify(calculateMonthSummary("2026-05", getActiveContext("all"))));
    deleteEntryLineageForTest("e-expense-copy");
    deleteDebtLineageForTest("d-open-copy");
    deleteAssetLineageForTest("a-cash-copy");
    assert("Löschen entfernt Original und Kopien", !state.entries.some((entry) => sameLineage(entry, "e-expense")) && !state.debts.some((debt) => sameLineage(debt, "d-open")) && !state.assets.some((asset) => sameLineage(asset, "a-cash")), "lineage should be gone");

    state.entries = [];
    state.debts = [];
    state.assets = [];
    const farukExpenses = [
      ["A-200", 200],
      ["A-100", 100],
      ["A-456", 456.86],
      ["A-5", 5],
      ["SIGMA KREDITBANK", 232.7],
      ["kreditBest check24", 97.32],
      ["IPHONE 15 PRO MAX 256GB", 42],
      ["Vattenfall Strom", 31.58],
      ["Telefónica O2", 27.99],
      ["Miete Einbecker", 477.4],
      ["Vodafone Kabel", 40],
      ["A-80", 80],
      ["A-200b", 200],
      ["A-1000", 1000],
    ];
    state.entries.push(
      { id: "faruk-jobcenter", personId: "case-test", type: "income", date: "2026-06-01", category: "Einkommen", description: "JobCenter", payment: "Überweisung", amount: 2066, recurrence: "none", recurring: false, endDate: "", status: "open", updatedAt: 20 },
      { id: "faruk-roya", personId: "case-test", type: "income", date: "2026-06-28", category: "Einkommen", description: "Roya", payment: "Überweisung", amount: 200, recurrence: "none", recurring: false, endDate: "", status: "open", updatedAt: 21 },
      { id: "bahar-rest", personId: "p2-test", type: "expense", date: "2026-06-02", category: "Sonstiges", description: "Andere Juni Ausgaben", payment: "Karte", amount: 1758.06, recurrence: "none", recurring: false, endDate: "", status: "open", updatedAt: 22 },
      ...farukExpenses.map(([description, amount], index) => ({ id: `faruk-expense-${index}`, personId: "case-test", type: "expense", date: "2026-06-01", category: "Fixkosten", description, payment: "Lastschrift", amount, recurrence: "none", recurring: false, endDate: "", status: "open", updatedAt: 30 + index })),
      ...[
        ["IPHONE 15 PRO MAX 256GB", 42],
        ["kreditBest check24", 97.32],
        ["Miete Einbecker", 477.4],
        ["SIGMA KREDITBANK", 232.7],
        ["Telefónica O2", 27.99],
        ["Vattenfall Strom", 31.58],
        ["Vodafone Kabel", 30],
      ].map(([description, amount], index) => ({ id: `aggregate-dup-${index}`, personId: defaultPersonId(), type: "expense", date: "2026-06-01", category: "Fixkosten", description, payment: "Lastschrift", amount, recurrence: "none", recurring: false, endDate: "", status: "open", updatedAt: 60 + index }))
    );
    const farukJune = calculateMonthSummary("2026-06", getActiveContext("case-test"));
    const farukJuneDay = getCalendarDaySummary("2026-06-01", getActiveContext("case-test"));
    const totalJune = calculateMonthSummary("2026-06", getActiveContext("all"));
    assert("Fall A Faruk Juni Einnahmen", closeEnough(farukJune.income, 2266), JSON.stringify(farukJune));
    assert("Fall A Faruk Juni Ausgaben", closeEnough(farukJune.entryExpense, 2990.85), JSON.stringify(farukJune));
    assert("Fall A Faruk Tagesbilanz 01.06.", closeEnough(farukJuneDay.net, -924.85), JSON.stringify(farukJuneDay));
    assert("Fall A Faruk Monatssaldo", closeEnough(farukJune.balance, -724.85), JSON.stringify(farukJune));
    assert("Fall B Gesamt ohne alte Gesamt-Dubletten", closeEnough(totalJune.entryExpense, 4748.91), JSON.stringify(totalJune));

    state.entries.push(
      { id: "repeat-weekly", personId: "case-test", type: "expense", date: "2026-06-03", category: "Lebensmittel", description: "Lebensmittel wöchentlich", payment: "Karte", amount: 10, recurrence: "weekly", recurring: true, endDate: "", status: "open", updatedAt: 80 },
      { id: "repeat-yearly", personId: "case-test", type: "expense", date: "2026-06-15", category: "Versicherungen", description: "Versicherung jährlich", payment: "Überweisung", amount: 120, recurrence: "yearly", recurring: true, endDate: "", status: "open", updatedAt: 81 },
      { id: "repeat-daily", personId: "case-test", type: "expense", date: "2026-06-01", category: "Mobilität", description: "Parken täglich", payment: "Karte", amount: 2, recurrence: "daily", recurring: true, endDate: "2026-06-05", status: "open", updatedAt: 82 },
      { id: "repeat-month-end", personId: "case-test", type: "expense", date: "2026-01-31", category: "Abos", description: "Monatsende", payment: "Lastschrift", amount: 7, recurrence: "monthly", recurring: true, endDate: "", status: "open", updatedAt: 83 }
    );
    assert("Wöchentlich erzeugt vier Juni-Termine", getEntriesForMonth("2026-06", getActiveContext("case-test")).filter((entry) => entry.id === "repeat-weekly").map((entry) => entryOccurrenceDate(entry, "2026-06")).join(",") === "2026-06-03,2026-06-10,2026-06-17,2026-06-24", JSON.stringify(getEntriesForMonth("2026-06", getActiveContext("case-test")).filter((entry) => entry.id === "repeat-weekly")));
    assert("Jährlich bleibt jährlich", getEntriesForMonth("2026-06", getActiveContext("case-test")).some((entry) => entry.id === "repeat-yearly") && getEntriesForMonth("2027-06", getActiveContext("case-test")).some((entry) => entry.id === "repeat-yearly") && !getEntriesForMonth("2026-07", getActiveContext("case-test")).some((entry) => entry.id === "repeat-yearly"), "yearly recurrence");
    assert("Täglich respektiert Enddatum", getEntriesForMonth("2026-06", getActiveContext("case-test")).filter((entry) => entry.id === "repeat-daily").length === 5 && !entriesForDate("2026-06-06", getActiveContext("case-test")).some((entry) => entry.id === "repeat-daily"), "daily recurrence");
    assert("Monatsende wird geklemmt", entriesForDate("2026-02-28", getActiveContext("case-test")).some((entry) => entry.id === "repeat-month-end"), "31.01. -> 28.02.");
  } finally {
    state.categories = snapshot.categories;
    state.entries = snapshot.entries;
    state.debts = snapshot.debts;
    state.assets = snapshot.assets;
    state.bankConnections = snapshot.bankConnections;
    state.people = snapshot.people;
    state.selectedPersonId = snapshot.selectedPersonId;
    elements.monthInput.value = previousMonth;
    selectedCalendarDay = previousSelectedCalendarDay;
    ensurePeople();
    render();
  }

  const passed = results.filter((result) => result.pass).length;
  const failed = results.length - passed;
  const summary = { passed, failed, results };
  return summary;
}

window.runBudgetUpSelfTest = runBudgetUpSelfTest;

if (new URLSearchParams(window.location.search).has("selftest")) {
  window.addEventListener("load", () => {
    const result = runBudgetUpSelfTest();
    const marker = document.createElement("meta");
    marker.id = "budgetup-self-test";
    marker.dataset.status = result.failed ? "failed" : "passed";
    marker.dataset.result = JSON.stringify(result);
    document.head.append(marker);
  }, { once: true });
}
