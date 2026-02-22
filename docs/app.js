const STORAGE_KEY = "cunySolutionPackHistorySession";
const RETENTION_EXPIRY_KEY = "cunySessionRetentionExpiry";
const SAMPLE_ROTATION_KEY = "cunySampleRotationNextIndex";
const ANALYTICS_STORAGE_KEY = "cunyPmAppAnalyticsEvents";
const RETENTION_MS = 15 * 60 * 1000;
const RETENTION_WARNING_MS = 60 * 1000;

const ANALYTICS_EVENT_TYPES = {
  formPageLoad: "form_page_load",
  generateClicked: "generate_solution_pack_clicked",
};

const tabs = document.querySelectorAll(".tab-btn");
const panels = {
  intake: document.getElementById("tab-intake"),
  solution: document.getElementById("tab-solution"),
  history: document.getElementById("tab-history"),
  analytics: document.getElementById("tab-analytics"),
};

const intakeForm = document.getElementById("intakeForm");
const followupQuestionsEl = document.getElementById("followupQuestions");
const followupListEl = document.getElementById("followupList");
const generationStatusEl = document.getElementById("generationStatus");
const generateSolutionBtnEl = document.getElementById("generateSolutionBtn");
const solutionMetaEl = document.getElementById("solutionMeta");
const versionListEl = document.getElementById("versionList");
const timelineVisualEl = document.getElementById("timelineVisual");
const timelineMetricsEl = document.getElementById("timelineMetrics");
const matrixBodyEl = document.getElementById("matrixBody");
const solutionSubTabs = document.querySelectorAll(".solution-tab-btn");
const solutionPanels = document.querySelectorAll(".solution-panel");
const quickJumpSelectEl = document.getElementById("quickJumpSelect");
const quickTogglePanelsEl = document.getElementById("quickTogglePanels");
const quickCopySelectEl = document.getElementById("quickCopySelect");
const quickCopyBtnEl = document.getElementById("quickCopyBtn");
const quickExportPdfBtnEl = document.getElementById("quickExportPdfBtn");
const statusPriorityBadgeEl = document.getElementById("statusPriorityBadge");
const statusComplexityBadgeEl = document.getElementById("statusComplexityBadge");
const statusPhaseLabelEl = document.getElementById("statusPhaseLabel");
const statusPercentLabelEl = document.getElementById("statusPercentLabel");
const statusProgressFillEl = document.getElementById("statusProgressFill");
const phaseStepperEl = document.getElementById("phaseStepper");
const miniMetricsEl = document.getElementById("miniMetrics");
const phaseDurationChartEl = document.getElementById("phaseDurationChart");
const effortDonutEl = document.getElementById("effortDonut");
const effortLegendEl = document.getElementById("effortLegend");
const milestoneBurnupChartEl = document.getElementById("milestoneBurnupChart");
const workstreamsListEl = document.getElementById("workstreamsList");
const copyWorkstreamsEl = document.getElementById("copyWorkstreams");
const raidSnapshotListEl = document.getElementById("raidSnapshotList");
const copyRaidSnapshotEl = document.getElementById("copyRaidSnapshot");
const retentionStatusEl = document.getElementById("retentionStatus");
const exportPreviewMetaEl = document.getElementById("exportPreviewMeta");
const exportPreviewEl = document.getElementById("exportPreview");
const exportSolutionPdfEl = document.getElementById("exportSolutionPdf");
const solutionPdfProjectNameEl = document.getElementById("solutionPdfProjectName");
const solutionPdfProjectSummaryEl = document.getElementById("solutionPdfProjectSummary");
const solutionPdfTocListEl = document.getElementById("solutionPdfTocList");
const aiModeEl = document.getElementById("aiMode");
const aiModelEl = document.getElementById("aiModel");
const aiEndpointEl = document.getElementById("aiEndpoint");
const aiApiKeyEl = document.getElementById("aiApiKey");
const aiEndpointWrapEl = document.getElementById("aiEndpointWrap");
const aiKeyWrapEl = document.getElementById("aiKeyWrap");
const aiModeHintEl = document.getElementById("aiModeHint");
const aiTestConnectionEl = document.getElementById("aiTestConnection");
const aiSuggestIntakeEl = document.getElementById("aiSuggestIntake");
const aiAssistStatusEl = document.getElementById("aiAssistStatus");
const generationLoaderEl = document.getElementById("generationLoader");
const loaderProgressFillEl = document.getElementById("loaderProgressFill");
const logoHomeTriggerEl = document.getElementById("logoHomeBtn") || document.querySelector(".header-logo");
const analyticsContentEl = document.getElementById("analyticsContent");
const analyticsRefreshEl = document.getElementById("analyticsRefresh");
const analyticsLastUpdatedEl = document.getElementById("analyticsLastUpdated");
const analyticsFormTotalEl = document.getElementById("analyticsFormTotal");
const analyticsGenerateTotalEl = document.getElementById("analyticsGenerateTotal");
const analyticsFormDailyEl = document.getElementById("analytics-form-daily");
const analyticsFormWeeklyEl = document.getElementById("analytics-form-weekly");
const analyticsFormMonthlyEl = document.getElementById("analytics-form-monthly");
const analyticsFormTotalCellEl = document.getElementById("analytics-form-total");
const analyticsGenerateDailyEl = document.getElementById("analytics-generate-daily");
const analyticsGenerateWeeklyEl = document.getElementById("analytics-generate-weekly");
const analyticsGenerateMonthlyEl = document.getElementById("analytics-generate-monthly");
const analyticsGenerateTotalCellEl = document.getElementById("analytics-generate-total");
const analyticsUsageChartEl = document.getElementById("analyticsUsageChart");
const analyticsChartEmptyEl = document.getElementById("analyticsChartEmpty");
const analyticsRangeBtns = document.querySelectorAll(".analytics-range-btn");
const requiredIntakeFields = intakeForm
  ? Array.from(intakeForm.querySelectorAll("input[required], textarea[required], select[required]"))
  : [];

const sectionEls = {
  sectionA: document.getElementById("sectionA"),
  sectionB: document.getElementById("sectionB"),
  sectionC: document.getElementById("sectionC"),
  sectionD: document.getElementById("sectionD"),
  sectionE: document.getElementById("sectionE"),
  sectionF: document.getElementById("sectionF"),
  sectionG: document.getElementById("sectionG"),
  sectionH: document.getElementById("sectionH"),
};

const state = {
  currentPack: null,
  versions: loadVersions(),
  analyticsEvents: loadAnalyticsEvents(),
  retentionExpiresAt: loadRetentionExpiry(),
  retentionWarningTimer: null,
  retentionCleanupTimer: null,
  retentionWarningShown: false,
  isGenerating: false,
  generationToken: 0,
  analyticsRange: "daily",
  aiConfig: {
    mode: "local",
    model: "local-heuristic-v2",
    endpoint: "",
    apiKey: "",
  },
};

const SIMPLE_AI_DEFAULTS = {
  localModel: "local-heuristic-v2",
  cloudModel: "openrouter/free",
  cloudEndpoint: "https://openrouter.ai/api/v1/chat/completions",
};

const STATUS_PHASES = ["Initiation", "Planning", "Execution", "M&C", "Closure"];

const SAMPLE_PROJECTS = [
  {
    projectTitle: "Student Success Early Alert Workflow Modernization",
    projectGoal: "Reduce issue-to-intervention time for at-risk students from 10 days to 3 days.",
    ownerName: "Alex Rivera",
    ownerArea: "Academic Affairs",
    problem: "Current early alert intake and triage steps are inconsistent across colleges. Advisors receive incomplete requests and follow-up actions are delayed. A standard process is needed to improve response speed and accountability.",
    outcomes: [
      "Standardize intake and triage steps across participating departments",
      "Improve advisor response time and case routing accuracy",
      "Launch weekly leadership reporting with clear ownership and due dates",
    ],
    constraints: {
      budget: "Fixed at $120,000 for implementation and training",
      staffing: "Project Manager, Business Analyst, Tech Lead, Advising SMEs, Reporting Analyst",
      compliance: "FERPA, accessibility (WCAG 2.1 AA), and records-retention policy alignment",
      tech: "Must use CUNYfirst + Microsoft 365; no new platforms this phase",
    },
    goLiveOffsetDays: 112,
    urgency: "High",
    stakeholders: [
      "Academic Affairs leadership",
      "Advising and student success teams",
      "Registrar and enrollment services",
      "IT applications and data/reporting teams",
    ],
    systems: [
      "CUNYfirst student records",
      "Microsoft Teams and SharePoint",
      "Power BI reporting workspace",
    ],
    risks: [
      "Delayed decisions on cross-campus workflow standards",
      "Advisor capacity constraints during peak registration periods",
      "Late data-quality issues in source systems",
    ],
  },
  {
    projectTitle: "Curriculum Change Approval Process Redesign",
    projectGoal: "Cut curriculum change approval cycle time from one semester to six weeks.",
    ownerName: "Maya Chen",
    ownerArea: "Provost Office",
    problem: "Curriculum proposals move through multiple committees with inconsistent handoffs and limited visibility. Departments cannot reliably predict approval timing, creating delays for catalog and advising updates.",
    outcomes: [
      "Establish a single intake and routing workflow for curriculum proposals",
      "Define clear ownership and SLAs for each approval stage",
      "Provide dashboard reporting for proposal status and bottlenecks",
    ],
    constraints: {
      budget: "Limited to existing operating funds and internal tooling",
      staffing: "Faculty governance coordinator, PM, process analyst, registrar SMEs",
      compliance: "Academic policy adherence and governance committee requirements",
      tech: "Use existing SharePoint and workflow tools; no custom app build",
    },
    goLiveOffsetDays: 98,
    urgency: "Moderate",
    stakeholders: [
      "Provost Office",
      "Faculty Senate committees",
      "Department chairs and program directors",
      "Registrar and catalog management team",
    ],
    systems: [
      "Curriculum proposal repository",
      "SharePoint workflow site",
      "Academic catalog publishing workflow",
    ],
    risks: [
      "Committee meeting cadence misaligned with delivery timeline",
      "Role ambiguity between departments and governance groups",
      "Late scope additions near policy review",
    ],
  },
  {
    projectTitle: "Campus-Wide Multi-Factor Authentication Rollout",
    projectGoal: "Achieve 95% MFA enrollment for faculty, staff, and students before term start.",
    ownerName: "Jordan Patel",
    ownerArea: "Information Security",
    problem: "The current identity security posture depends heavily on single-factor access for several systems. Security incidents and audit findings require a staged MFA rollout with clear support and communication plans.",
    outcomes: [
      "Deploy phased MFA enrollment by user segment",
      "Launch support playbooks for enrollment and account recovery",
      "Track adoption and unresolved exceptions weekly",
    ],
    constraints: {
      budget: "Fixed security program allocation with no contingency budget",
      staffing: "Security lead, IAM engineer, service desk lead, communications/training lead",
      compliance: "Security audit controls and accessibility support requirements",
      tech: "Integrate with existing SSO and identity provider",
    },
    goLiveOffsetDays: 84,
    urgency: "Critical",
    stakeholders: [
      "Information Security leadership",
      "IT support and service desk",
      "Faculty and student support offices",
      "Communications and training teams",
    ],
    systems: [
      "Identity provider and SSO portal",
      "Email and collaboration suite",
      "Student and HR enterprise systems",
    ],
    risks: [
      "Service desk overload during enrollment peak",
      "User resistance due to poor onboarding communication",
      "Integration issues with legacy applications",
    ],
  },
  {
    projectTitle: "Facilities Work Order Intake Optimization",
    projectGoal: "Improve work-order intake quality and reduce assignment delays by 40%.",
    ownerName: "Nina Alvarez",
    ownerArea: "Facilities Operations",
    problem: "Facilities work requests arrive through email, calls, and ad-hoc forms with missing details. Dispatching is delayed due to repeated clarifications and inconsistent priority assignment.",
    outcomes: [
      "Consolidate all requests into one standardized intake path",
      "Define triage priorities and assignment rules",
      "Create weekly operational reporting on backlog and turnaround time",
    ],
    constraints: {
      budget: "Limited to process and configuration updates only",
      staffing: "Operations PM, dispatcher leads, maintenance supervisors, reporting analyst",
      compliance: "Safety and emergency escalation policies must be preserved",
      tech: "Use existing work-order platform and Microsoft Forms",
    },
    goLiveOffsetDays: 70,
    urgency: "High",
    stakeholders: [
      "Facilities leadership",
      "Dispatch and maintenance teams",
      "Campus operations offices",
      "Safety and risk management",
    ],
    systems: [
      "Facilities work-order system",
      "Microsoft Forms intake",
      "Operations reporting dashboard",
    ],
    risks: [
      "Incomplete intake adoption from campuses",
      "Backlog spike during transition period",
      "Priority criteria not applied consistently",
    ],
  },
  {
    projectTitle: "Research Data Governance Launch for Grant Programs",
    projectGoal: "Stand up a governance framework for grant data intake, access, and reporting in one quarter.",
    ownerName: "Samuel Brooks",
    ownerArea: "Research Administration",
    problem: "Grant-related data definitions, access approvals, and reporting responsibilities vary by unit. Inconsistent controls create audit and reporting risk for active grant programs.",
    outcomes: [
      "Define governance roles and decision rights for grant data",
      "Implement intake and approval workflow for data requests",
      "Publish baseline reporting standards and quality checks",
    ],
    constraints: {
      budget: "Fixed program budget with limited procurement lead time",
      staffing: "Program PM, data steward, compliance lead, BI analyst, legal SME",
      compliance: "Data privacy, grant reporting, and records retention requirements",
      tech: "Leverage existing data warehouse and reporting tools",
    },
    goLiveOffsetDays: 126,
    urgency: "Moderate",
    stakeholders: [
      "Research administration leadership",
      "Principal investigators and grant managers",
      "Data governance and compliance teams",
      "Finance and reporting offices",
    ],
    systems: [
      "Grant management system",
      "Enterprise data warehouse",
      "Power BI reporting workspace",
    ],
    risks: [
      "Conflicting data definitions across units",
      "Approval bottlenecks for sensitive data access",
      "Reporting gaps during initial transition",
    ],
  },
];

init();

function init() {
  followupQuestionsEl.classList.add("hidden");
  syncAIControls(true);
  setAnalyticsRange(state.analyticsRange);
  initializeIntakeValidation();
  attachEventHandlers();
  setActiveTab("intake");
  trackAnalyticsEvent(ANALYTICS_EVENT_TYPES.formPageLoad);
  setActiveSolutionTab("summary");
  if (state.versions.length > 0) {
    state.currentPack = state.versions[0];
    renderSolutionPack(state.currentPack);
  } else {
    resetSolutionPackDisplays();
    renderExportPreview(null);
  }
  initializeRetentionLifecycle();
  renderVersionHistory();
  renderAnalyticsCounts();
}

function attachEventHandlers() {
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => setActiveTab(tab.dataset.tab));
  });

  if (logoHomeTriggerEl) {
    logoHomeTriggerEl.addEventListener("click", (event) => {
      event.preventDefault();
      returnToIntakeAndReset();
    });
  }

  if (analyticsRefreshEl) {
    analyticsRefreshEl.addEventListener("click", () => {
      state.analyticsEvents = loadAnalyticsEvents();
      renderAnalyticsCounts();
    });
  }

  if (analyticsRangeBtns.length) {
    analyticsRangeBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        setAnalyticsRange(btn.dataset.analyticsRange || "daily");
      });
    });
  }

  window.addEventListener("resize", () => {
    if (getActiveTabKey() !== "analytics") return;
    renderAnalyticsUsageChart();
  });

  solutionSubTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      setActiveSolutionTab(tab.dataset.solutionTab || "summary");
    });
  });

  intakeForm.addEventListener("input", (event) => {
    if (hasAnySessionData()) {
      ensureRetentionWindowActive();
    }
    const field = event.target;
    if (requiredIntakeFields.includes(field)) {
      field.dataset.touched = "true";
    }
    validateRequiredIntakeFields();
  });

  intakeForm.addEventListener("change", (event) => {
    if (hasAnySessionData()) {
      ensureRetentionWindowActive();
    }
    const field = event.target;
    if (requiredIntakeFields.includes(field)) {
      field.dataset.touched = "true";
    }
    validateRequiredIntakeFields();
  });

  requiredIntakeFields.forEach((field) => {
    field.addEventListener("blur", () => {
      field.dataset.touched = "true";
      validateRequiredIntakeFields();
    });
  });

  intakeForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (state.isGenerating) return;
    const validation = validateRequiredIntakeFields({ showErrors: true });
    if (!validation.isValid) {
      generationStatusEl.textContent = "Please complete the required fields to continue.";
      focusFirstInvalidIntakeField(validation.firstInvalidField);
      return;
    }

    trackAnalyticsEvent(ANALYTICS_EVENT_TYPES.generateClicked);
    const generationToken = state.generationToken + 1;
    state.generationToken = generationToken;

    state.aiConfig = getAIConfig();
    const intake = getIntakeData();
    const complexity = getSelectedComplexity();
    const followups = buildFollowUpQuestions(intake);

    renderFollowUps(followups);
    generationStatusEl.textContent = "Generating solution pack...";
    state.isGenerating = true;
    updateGenerateButtonState();
    startGenerationLoader();
    const minimumDelay = waitMs(5000);
    let generatedPack = null;

    try {
      let pack = buildSolutionPack(intake, complexity, followups, state.aiConfig);
      pack = await enhancePackWithAI(pack, intake, followups, state.aiConfig);
      await minimumDelay;
      if (generationToken !== state.generationToken) return;

      state.currentPack = pack;
      renderSolutionPack(pack);
      savePackVersion(pack);
      renderVersionHistory();
      generatedPack = pack;
    } catch {
      await minimumDelay;
      if (generationToken !== state.generationToken) return;
      generationStatusEl.textContent = "Could not generate the solution pack. Please try again.";
    } finally {
      if (generationToken === state.generationToken) {
        stopGenerationLoader();
        state.isGenerating = false;
        updateGenerateButtonState();
      }
    }

    if (generatedPack && generationToken === state.generationToken) {
      setActiveTab("solution");
      scrollToSolutionPackTop();
      generationStatusEl.textContent = generatedPack.aiMeta?.mode !== "local"
        ? `Solution pack generated with ${generatedPack.aiMeta?.provider || "AI"} enhancements and saved to version history.`
        : "Solution pack generated and saved to version history.";
    }
  });

  const loadSampleDataEl = document.getElementById("loadSampleData");
  if (loadSampleDataEl) {
    loadSampleDataEl.addEventListener("click", () => {
      const sample = getNextSampleProject();
      populateSampleIntakeForm(sample);
      const intake = getIntakeData();
      renderFollowUps(buildFollowUpQuestions(intake));
      resetIntakeValidationState();
      validateRequiredIntakeFields();
      ensureRetentionWindowActive();
      generationStatusEl.textContent = `Sample loaded: ${sample.projectTitle}. Click Generate Solution Pack to create a complete example.`;
    });
  }

  if (quickJumpSelectEl) {
    quickJumpSelectEl.addEventListener("change", () => {
      const target = normalize(quickJumpSelectEl.value) || "summary";
      const tabTarget = target === "nextsteps" ? "deliverables" : target;
      setActiveSolutionTab(tabTarget);
      const anchor = document.querySelector(`[data-section-anchor="${target}"]`) || document.querySelector(`[data-section-anchor="${tabTarget}"]`);
      if (anchor) {
        anchor.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  }

  if (solutionPdfTocListEl) {
    solutionPdfTocListEl.addEventListener("click", (event) => {
      const link = event.target.closest(".toc-link[data-toc-target]");
      if (!link) return;
      const targetId = normalize(link.getAttribute("data-toc-target"));
      if (!targetId) return;
      const moved = scrollToSolutionTocTarget(targetId);
      if (moved) {
        event.preventDefault();
      }
    });
  }

  if (quickTogglePanelsEl) {
    quickTogglePanelsEl.addEventListener("click", () => {
      const collapseAll = quickTogglePanelsEl.dataset.mode !== "collapsed";
      setSolutionPanelsCollapsed(collapseAll);
    });
  }

  if (quickCopyBtnEl) {
    quickCopyBtnEl.addEventListener("click", async () => {
      if (!ensurePackExists()) return;
      const selection = normalize(quickCopySelectEl?.value);
      await copyQuickSelection(selection);
    });
  }

  if (quickExportPdfBtnEl) {
    quickExportPdfBtnEl.addEventListener("click", () => {
      document.getElementById("printPdf")?.click();
    });
  }

  if (copyRaidSnapshotEl) {
    copyRaidSnapshotEl.addEventListener("click", async () => {
      if (!ensurePackExists()) return;
      await copyText(renderRaidSnapshotText());
      generationStatusEl.textContent = "RAID snapshot copied.";
    });
  }

  if (copyWorkstreamsEl) {
    copyWorkstreamsEl.addEventListener("click", async () => {
      if (!ensurePackExists()) return;
      await copyText(renderWorkstreamsText(state.currentPack));
      generationStatusEl.textContent = "Workstreams copied.";
    });
  }

  document.querySelectorAll(".copy-btn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      if (!ensurePackExists()) return;
      const targetId = btn.dataset.copy;
      if (!targetId) return;
      const content = sectionEls[targetId]?.textContent || "";
      await copyText(content);
      generationStatusEl.textContent = "Section copied to clipboard.";
    });
  });

  document.getElementById("copyTimeline").addEventListener("click", async () => {
    if (!ensurePackExists()) return;
    const timeline = getTimelineData(state.currentPack);
    await copyText(renderTimelineText(timeline));
    generationStatusEl.textContent = "Timeline copied.";
  });

  document.getElementById("copyMatrix").addEventListener("click", async () => {
    if (!ensurePackExists()) return;
    const timeline = getTimelineData(state.currentPack);
    const matrix = getMatrixData(state.currentPack, timeline);
    await copyText(renderMatrixText(matrix));
    generationStatusEl.textContent = "Tracker matrix copied.";
  });

  document.getElementById("copyFullPack").addEventListener("click", async () => {
    if (!ensurePackExists()) return;
    await copyText(renderFullPackText(state.currentPack));
    generationStatusEl.textContent = "Full solution pack copied to clipboard.";
  });

  if (exportSolutionPdfEl) {
    exportSolutionPdfEl.addEventListener("click", () => {
      if (!ensurePackExists()) return;
      exportSolutionPackPdf();
    });
  }

  const printPdfEl = document.getElementById("printPdf");
  if (printPdfEl) {
    printPdfEl.addEventListener("click", () => {
      if (!state.currentPack) {
        generationStatusEl.textContent = "Generate a solution pack before exporting.";
        return;
      }
      setActiveTab("export");
      window.requestAnimationFrame(() => {
        window.print();
      });
    });
  }

  document.getElementById("clearHistory").addEventListener("click", () => {
    state.versions = [];
    persistVersions();
    renderVersionHistory();
    ensureRetentionWindowActive();
    generationStatusEl.textContent = "Version history cleared.";
  });

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState !== "visible") return;
    if (state.retentionExpiresAt && Date.now() >= state.retentionExpiresAt) {
      clearSessionData("Session data expired and has been deleted.");
      return;
    }
    refreshRetentionTimers();
  });

  if (aiModeEl) {
    aiModeEl.addEventListener("change", () => {
      syncAIControls(true);
      state.aiConfig = getAIConfig();
    });
  }

  if (aiModelEl) {
    aiModelEl.addEventListener("input", () => {
      state.aiConfig = getAIConfig();
    });
  }

  if (aiApiKeyEl) {
    aiApiKeyEl.addEventListener("input", () => {
      state.aiConfig = getAIConfig();
    });
  }

  if (aiEndpointEl) {
    aiEndpointEl.addEventListener("input", () => {
      state.aiConfig = getAIConfig();
    });
  }

  if (aiTestConnectionEl) {
    aiTestConnectionEl.addEventListener("click", async () => {
      state.aiConfig = getAIConfig();
      if (aiAssistStatusEl) aiAssistStatusEl.textContent = "Testing AI connection...";
      const result = await testAIConnection(state.aiConfig);
      if (aiAssistStatusEl) aiAssistStatusEl.textContent = result.message;
    });
  }

  if (aiSuggestIntakeEl) {
    aiSuggestIntakeEl.addEventListener("click", async () => {
      state.aiConfig = getAIConfig();
      const intake = getIntakeData();
      if (aiAssistStatusEl) {
        aiAssistStatusEl.textContent = state.aiConfig.mode === "local"
          ? "Generating smart local suggestions..."
          : "Generating provider-backed suggestions...";
      }

      const suggestions = await generateIntakeSuggestions(intake, state.aiConfig);
      applyIntakeSuggestions(suggestions);
      const refreshedIntake = getIntakeData();
      renderFollowUps(buildFollowUpQuestions(refreshedIntake));
      validateRequiredIntakeFields();
      if (hasAnySessionData()) ensureRetentionWindowActive();
      if (aiAssistStatusEl) aiAssistStatusEl.textContent = suggestions.summary;
    });
  }
}

function setActiveTab(key) {
  tabs.forEach((tab) => {
    const isActive = tab.dataset.tab === key;
    tab.classList.toggle("active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
    tab.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  Object.entries(panels).forEach(([panelKey, panel]) => {
    if (!panel) return;
    const isActive = panelKey === key;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });

  if (key === "analytics") {
    state.analyticsEvents = loadAnalyticsEvents();
    renderAnalyticsCounts();
  }
}

function scrollToSolutionPackTop() {
  if (!panels.solution) return;
  window.requestAnimationFrame(() => {
    const targetTop = panels.solution.getBoundingClientRect().top + window.scrollY;
    window.scrollTo({ top: Math.max(0, targetTop - 6), left: 0, behavior: "auto" });
  });
}

function scrollToIntakeTop() {
  if (!panels.intake) return;
  window.requestAnimationFrame(() => {
    const targetTop = panels.intake.getBoundingClientRect().top + window.scrollY;
    window.scrollTo({ top: Math.max(0, targetTop - 6), left: 0, behavior: "smooth" });
  });
}

function returnToIntakeAndReset() {
  clearSessionData("Ready for a fresh intake. Previous form and solution content were cleared.");
  setActiveTab("intake");
  setActiveSolutionTab("summary");
  state.aiConfig = getAIConfig();
  syncAIControls(true);
  validateRequiredIntakeFields();
  document.querySelectorAll(".intake-section").forEach((section) => {
    section.open = true;
  });
  if (retentionStatusEl) retentionStatusEl.textContent = "";
  scrollToIntakeTop();
  const primaryField = document.getElementById("projectTitle");
  if (primaryField) {
    window.setTimeout(() => {
      primaryField.focus({ preventScroll: true });
    }, 220);
  }
}

function getActiveTabKey() {
  const activeTab = document.querySelector(".tab-btn.active");
  return normalize(activeTab?.dataset?.tab) || "intake";
}

function waitMs(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function setAnalyticsRange(range) {
  const normalizedRange = normalize(range).toLowerCase();
  const supportedRanges = ["daily", "weekly", "monthly"];
  state.analyticsRange = supportedRanges.includes(normalizedRange) ? normalizedRange : "daily";

  analyticsRangeBtns.forEach((btn) => {
    const isActive = normalize(btn.dataset.analyticsRange).toLowerCase() === state.analyticsRange;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-pressed", isActive ? "true" : "false");
  });

  renderAnalyticsUsageChart();
}

function formatAnalyticsBucketLabel(dateValue, range) {
  const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (Number.isNaN(date.getTime())) return "";

  if (range === "monthly") {
    return date.toLocaleDateString(undefined, { month: "short", year: "2-digit" });
  }

  return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function buildAnalyticsBuckets(range, nowDate = new Date()) {
  const buckets = [];
  const now = startOfDay(nowDate);

  if (range === "weekly") {
    const weekStart = getStartOfWeek(now);
    for (let offset = 11; offset >= 0; offset -= 1) {
      const start = addDays(weekStart, -7 * offset);
      const end = addDays(start, 7);
      buckets.push({
        startTs: start.getTime(),
        endTs: end.getTime(),
        label: formatAnalyticsBucketLabel(start, "weekly"),
        formCount: 0,
        generateCount: 0,
      });
    }
    return buckets;
  }

  if (range === "monthly") {
    const monthStart = getStartOfMonth(now);
    for (let offset = 11; offset >= 0; offset -= 1) {
      const start = new Date(monthStart.getFullYear(), monthStart.getMonth() - offset, 1);
      const end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
      buckets.push({
        startTs: start.getTime(),
        endTs: end.getTime(),
        label: formatAnalyticsBucketLabel(start, "monthly"),
        formCount: 0,
        generateCount: 0,
      });
    }
    return buckets;
  }

  for (let offset = 13; offset >= 0; offset -= 1) {
    const start = addDays(now, -offset);
    const end = addDays(start, 1);
    buckets.push({
      startTs: start.getTime(),
      endTs: end.getTime(),
      label: formatAnalyticsBucketLabel(start, "daily"),
      formCount: 0,
      generateCount: 0,
    });
  }

  return buckets;
}

function buildAnalyticsTimeSeries(range) {
  const buckets = buildAnalyticsBuckets(range);
  const events = state.analyticsEvents || [];

  if (!buckets.length || !events.length) {
    return {
      buckets,
      maxCount: 0,
      totalCount: 0,
    };
  }

  const firstStart = buckets[0].startTs;
  const lastEnd = buckets[buckets.length - 1].endTs;

  events.forEach((item) => {
    const ts = Number(item?.ts);
    if (!Number.isFinite(ts)) return;
    if (ts < firstStart || ts >= lastEnd) return;

    const bucket = buckets.find((point) => ts >= point.startTs && ts < point.endTs);
    if (!bucket) return;

    if (item.type === ANALYTICS_EVENT_TYPES.formPageLoad) {
      bucket.formCount += 1;
      return;
    }

    if (item.type === ANALYTICS_EVENT_TYPES.generateClicked) {
      bucket.generateCount += 1;
    }
  });

  const maxCount = buckets.reduce(
    (currentMax, point) => Math.max(currentMax, point.formCount, point.generateCount),
    0
  );
  const totalCount = buckets.reduce(
    (sum, point) => sum + point.formCount + point.generateCount,
    0
  );

  return {
    buckets,
    maxCount,
    totalCount,
  };
}

function getAnalyticsAxisScale(maxCount) {
  const safeMax = Math.max(0, Number(maxCount) || 0);
  if (safeMax <= 4) {
    return {
      axisMax: 4,
      step: 1,
    };
  }

  const step = Math.max(1, Math.ceil(safeMax / 4));
  return {
    axisMax: step * 4,
    step,
  };
}

function getAnalyticsLabelIndices(length) {
  if (length <= 0) return [];
  if (length <= 6) {
    return Array.from({ length }, (_, index) => index);
  }

  const labelSlots = 6;
  const result = [];
  for (let slot = 0; slot < labelSlots; slot += 1) {
    const index = Math.round((slot * (length - 1)) / (labelSlots - 1));
    if (!result.includes(index)) {
      result.push(index);
    }
  }
  return result;
}

function renderAnalyticsUsageChart() {
  if (!analyticsUsageChartEl) return;

  const range = state.analyticsRange || "daily";
  const { buckets, maxCount, totalCount } = buildAnalyticsTimeSeries(range);

  if (totalCount === 0) {
    analyticsUsageChartEl.innerHTML = "";
    analyticsUsageChartEl.setAttribute("aria-label", "No activity yet.");
    if (analyticsChartEmptyEl) analyticsChartEmptyEl.hidden = false;
    return;
  }

  if (analyticsChartEmptyEl) analyticsChartEmptyEl.hidden = true;

  const width = 960;
  const height = 260;
  const margin = { top: 16, right: 18, bottom: 44, left: 44 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const xStep = buckets.length > 1 ? innerWidth / (buckets.length - 1) : 0;
  const { axisMax, step } = getAnalyticsAxisScale(maxCount);
  const labelIndices = new Set(getAnalyticsLabelIndices(buckets.length));

  const xFor = (index) => margin.left + xStep * index;
  const yFor = (value) => margin.top + innerHeight - (Math.max(0, value) / axisMax) * innerHeight;
  const formatPoint = (point, index, key) => `${xFor(index).toFixed(2)},${yFor(point[key]).toFixed(2)}`;

  const yTicks = [];
  for (let value = 0; value <= axisMax; value += step) {
    yTicks.push(value);
  }

  const gridLines = yTicks.map((value) => {
    const y = yFor(value).toFixed(2);
    return `
      <line class="analytics-grid-line" x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}"></line>
      <text class="analytics-axis-label" x="${margin.left - 8}" y="${Number(y) + 4}" text-anchor="end">${value}</text>
    `;
  }).join("");

  const xLabels = buckets.map((point, index) => {
    if (!labelIndices.has(index)) return "";
    return `<text class="analytics-axis-label" x="${xFor(index).toFixed(2)}" y="${height - 14}" text-anchor="middle">${escapeHtml(point.label)}</text>`;
  }).join("");

  const formLine = buckets.map((point, index) => formatPoint(point, index, "formCount")).join(" ");
  const generateLine = buckets.map((point, index) => formatPoint(point, index, "generateCount")).join(" ");
  const formPoints = buckets.map((point, index) => {
    const x = xFor(index).toFixed(2);
    const y = yFor(point.formCount).toFixed(2);
    return `<circle class="analytics-point-form" cx="${x}" cy="${y}" r="2.8"></circle>`;
  }).join("");
  const generatePoints = buckets.map((point, index) => {
    const x = xFor(index).toFixed(2);
    const y = yFor(point.generateCount).toFixed(2);
    return `<circle class="analytics-point-generate" cx="${x}" cy="${y}" r="2.8"></circle>`;
  }).join("");

  analyticsUsageChartEl.setAttribute(
    "aria-label",
    `Usage over time (${range}) for form page loads and Generate Solution Pack clicks.`
  );
  analyticsUsageChartEl.innerHTML = `
    <svg class="analytics-usage-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMidYMid meet" aria-hidden="true">
      <line class="analytics-axis-line" x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}"></line>
      <line class="analytics-axis-line" x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}"></line>
      ${gridLines}
      <polyline class="analytics-series-form" points="${formLine}"></polyline>
      <polyline class="analytics-series-generate" points="${generateLine}"></polyline>
      ${formPoints}
      ${generatePoints}
      ${xLabels}
    </svg>
  `;
}

function loadAnalyticsEvents() {
  try {
    const raw = localStorage.getItem(ANALYTICS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return [];
    return pruneAnalyticsEvents(
      parsed.filter((item) => item && typeof item.type === "string" && Number.isFinite(Number(item.ts)))
    );
  } catch {
    return [];
  }
}

function persistAnalyticsEvents() {
  try {
    localStorage.setItem(ANALYTICS_STORAGE_KEY, JSON.stringify(state.analyticsEvents));
  } catch {
    // Ignore storage issues and keep runtime behavior intact.
  }
}

function pruneAnalyticsEvents(events) {
  return events
    .filter((item) => item && typeof item.type === "string" && Number.isFinite(Number(item.ts)))
    .map((item) => ({ type: item.type, ts: Number(item.ts) }));
}

function trackAnalyticsEvent(type) {
  const validTypes = Object.values(ANALYTICS_EVENT_TYPES);
  if (!validTypes.includes(type)) return;

  state.analyticsEvents = pruneAnalyticsEvents([
    ...(state.analyticsEvents || []),
    { type, ts: Date.now() },
  ]);
  persistAnalyticsEvents();

  if (getActiveTabKey() === "analytics") {
    renderAnalyticsCounts();
  }
}

function getStartOfWeek(dateValue) {
  const date = startOfDay(dateValue);
  const mondayOffset = (date.getDay() + 6) % 7;
  date.setDate(date.getDate() - mondayOffset);
  return date;
}

function getStartOfMonth(dateValue) {
  const date = startOfDay(dateValue);
  date.setDate(1);
  return date;
}

function countEventsSince(type, startTimestamp) {
  const events = state.analyticsEvents || [];
  return events.filter((item) => item.type === type && Number(item.ts) >= startTimestamp).length;
}

function countEventsTotal(type) {
  const events = state.analyticsEvents || [];
  return events.filter((item) => item.type === type).length;
}

function renderAnalyticsCounts() {
  if (analyticsContentEl) analyticsContentEl.hidden = false;
  const now = new Date();
  const startOfTodayTs = startOfDay(now).getTime();
  const startOfWeekTs = getStartOfWeek(now).getTime();
  const startOfMonthTs = getStartOfMonth(now).getTime();

  const formDaily = countEventsSince(ANALYTICS_EVENT_TYPES.formPageLoad, startOfTodayTs);
  const formWeekly = countEventsSince(ANALYTICS_EVENT_TYPES.formPageLoad, startOfWeekTs);
  const formMonthly = countEventsSince(ANALYTICS_EVENT_TYPES.formPageLoad, startOfMonthTs);
  const formTotal = countEventsTotal(ANALYTICS_EVENT_TYPES.formPageLoad);

  const generateDaily = countEventsSince(ANALYTICS_EVENT_TYPES.generateClicked, startOfTodayTs);
  const generateWeekly = countEventsSince(ANALYTICS_EVENT_TYPES.generateClicked, startOfWeekTs);
  const generateMonthly = countEventsSince(ANALYTICS_EVENT_TYPES.generateClicked, startOfMonthTs);
  const generateTotal = countEventsTotal(ANALYTICS_EVENT_TYPES.generateClicked);

  if (analyticsFormDailyEl) analyticsFormDailyEl.textContent = String(formDaily);
  if (analyticsFormWeeklyEl) analyticsFormWeeklyEl.textContent = String(formWeekly);
  if (analyticsFormMonthlyEl) analyticsFormMonthlyEl.textContent = String(formMonthly);
  if (analyticsFormTotalCellEl) analyticsFormTotalCellEl.textContent = String(formTotal);
  if (analyticsGenerateDailyEl) analyticsGenerateDailyEl.textContent = String(generateDaily);
  if (analyticsGenerateWeeklyEl) analyticsGenerateWeeklyEl.textContent = String(generateWeekly);
  if (analyticsGenerateMonthlyEl) analyticsGenerateMonthlyEl.textContent = String(generateMonthly);
  if (analyticsGenerateTotalCellEl) analyticsGenerateTotalCellEl.textContent = String(generateTotal);

  if (analyticsFormTotalEl) analyticsFormTotalEl.textContent = String(formTotal);
  if (analyticsGenerateTotalEl) analyticsGenerateTotalEl.textContent = String(generateTotal);
  if (analyticsLastUpdatedEl) analyticsLastUpdatedEl.textContent = `Last updated: ${formatDateTime(new Date())}`;
  renderAnalyticsUsageChart();
}

function startGenerationLoader() {
  if (!generationLoaderEl) return;
  generationLoaderEl.classList.add("active");
  generationLoaderEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("is-loading-solution");

  if (loaderProgressFillEl) {
    loaderProgressFillEl.classList.remove("animate");
    void loaderProgressFillEl.offsetWidth;
    loaderProgressFillEl.classList.add("animate");
  }
}

function stopGenerationLoader() {
  if (loaderProgressFillEl) {
    loaderProgressFillEl.classList.remove("animate");
  }
  if (generationLoaderEl) {
    generationLoaderEl.classList.remove("active");
    generationLoaderEl.setAttribute("aria-hidden", "true");
  }
  document.body.classList.remove("is-loading-solution");
}

function getPdfContentPageHeightPx() {
  const pageHeightInches = 11;
  const marginInches = 0.55;
  const pxPerInch = 96;
  return (pageHeightInches - marginInches * 2) * pxPerInch;
}

function isRenderableElement(element) {
  if (!element) return false;
  if (!element.isConnected) return false;
  const style = window.getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden") return false;
  return element.getBoundingClientRect().height > 0;
}

function isKeepTogetherBlock(element) {
  if (!element) return false;
  return element.id === "solutionPdfHeader" || element.id === "solutionPdfToc" || element.classList.contains("section-card");
}

function slugifyAnchorText(value) {
  return normalize(value)
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function ensureTocAnchorId(element, fallback, usedIds) {
  if (!element) return "";
  if (element.id) {
    usedIds.add(element.id);
    element.setAttribute("data-toc-anchor", "true");
    return element.id;
  }

  const base = slugifyAnchorText(element.textContent) || fallback || "section";
  let candidate = `solution-${base}`;
  let suffix = 2;

  while (usedIds.has(candidate) || (document.getElementById(candidate) && document.getElementById(candidate) !== element)) {
    candidate = `solution-${base}-${suffix}`;
    suffix += 1;
  }

  element.id = candidate;
  element.setAttribute("data-toc-anchor", "true");
  usedIds.add(candidate);
  return candidate;
}

function getSolutionPdfTocEntries() {
  if (!panels.solution) return [];
  const entries = [];
  const panel = panels.solution;
  const usedIds = new Set(Array.from(document.querySelectorAll("[id]")).map((node) => node.id));

  const packHeading = panel.querySelector(".solution-header h2");
  const packBlock = packHeading?.closest(".solution-header") || null;
  if (packHeading && packBlock) {
    const targetId = ensureTocAnchorId(packHeading, "project-management-solution-pack", usedIds);
    entries.push({
      title: normalize(packHeading.textContent) || "Project Management Solution Pack",
      level: 0,
      element: packHeading,
      blockElement: packBlock,
      targetId,
    });
  }

  const sectionHeadings = Array.from(panel.querySelectorAll(".section-card > .card-header > h3"));
  sectionHeadings.forEach((heading) => {
    const parentCard = heading.closest(".section-card");
    if (!parentCard) return;
    const targetId = ensureTocAnchorId(heading, "section-heading", usedIds);
    entries.push({
      title: normalize(heading.textContent) || "Section",
      level: 0,
      element: heading,
      blockElement: parentCard,
      targetId,
    });

    const subHeadings = Array.from(parentCard.querySelectorAll(".chart-card h4"));
    subHeadings.forEach((subHeading) => {
      const subTargetId = ensureTocAnchorId(subHeading, "sub-section", usedIds);
      entries.push({
        title: normalize(subHeading.textContent) || "Sub-section",
        level: 1,
        element: subHeading,
        blockElement: parentCard,
        targetId: subTargetId,
      });
    });
  });

  return entries;
}

function getSolutionPdfPaginationBlocks(panel) {
  if (!panel) return [];
  const selector = "#solutionPdfHeader, #solutionPdfToc, .solution-header, #solutionMeta, .section-card";
  return Array.from(panel.querySelectorAll(selector)).filter(isRenderableElement);
}

function getSolutionPdfBlockLayout(panel, pageHeightPx) {
  const blocks = getSolutionPdfPaginationBlocks(panel);
  const panelRect = panel.getBoundingClientRect();
  const blockLayout = new Map();
  let virtualCursorPx = 0;
  let previousBottomPx = 0;

  blocks.forEach((block) => {
    const blockRect = block.getBoundingClientRect();
    const blockTopPx = Math.max(0, blockRect.top - panelRect.top);
    const blockBottomPx = Math.max(blockTopPx, blockRect.bottom - panelRect.top);
    const blockHeightPx = Math.max(0, blockBottomPx - blockTopPx);
    const gapBeforePx = Math.max(0, blockTopPx - previousBottomPx);

    virtualCursorPx += gapBeforePx;

    const pageOffsetPx = virtualCursorPx % pageHeightPx;
    const keepTogether = isKeepTogetherBlock(block);
    const shouldAdvancePage = keepTogether && pageOffsetPx > 0 && (blockHeightPx > pageHeightPx || pageOffsetPx + blockHeightPx > pageHeightPx);

    if (shouldAdvancePage) {
      virtualCursorPx += pageHeightPx - pageOffsetPx;
    }

    blockLayout.set(block, {
      startPx: virtualCursorPx,
      topPx: blockTopPx,
    });

    virtualCursorPx += blockHeightPx;
    previousBottomPx = blockBottomPx;
  });

  return blockLayout;
}

function calculateSolutionPdfTocPageNumbers(entries) {
  if (!panels.solution || !entries.length) return [];
  const pageHeightPx = getPdfContentPageHeightPx();
  const blockLayout = getSolutionPdfBlockLayout(panels.solution, pageHeightPx);

  return entries.map((entry) => {
    if (!entry?.element || !entry?.blockElement) return 1;
    const blockInfo = blockLayout.get(entry.blockElement);
    if (!blockInfo) return 1;

    const entryRect = entry.element.getBoundingClientRect();
    const blockRect = entry.blockElement.getBoundingClientRect();
    const offsetWithinBlockPx = Math.max(0, entryRect.top - blockRect.top);
    const absolutePrintPositionPx = blockInfo.startPx + offsetWithinBlockPx;
    return Math.max(1, Math.floor(absolutePrintPositionPx / pageHeightPx) + 1);
  });
}

function renderSolutionPdfToc(entries, pageNumbers = []) {
  if (!solutionPdfTocListEl) return;
  solutionPdfTocListEl.innerHTML = entries.map((entry, index) => {
    const rowContent = `
      <span class="toc-title">${escapeHtml(entry.title)}</span>
      <span class="toc-leader" aria-hidden="true"></span>
      <span class="toc-page">${String(pageNumbers[index] || 1)}</span>
    `;
    const targetId = normalize(entry.targetId);
    const wrapped = targetId
      ? `<a class="toc-link" href="#${escapeHtml(targetId)}" data-toc-target="${escapeHtml(targetId)}">${rowContent}</a>`
      : `<span class="toc-link">${rowContent}</span>`;
    return `<li class="toc-item ${entry.level > 0 ? "toc-sub" : ""}">${wrapped}</li>`;
  }).join("");
}

function scrollToSolutionTocTarget(targetId) {
  const target = document.getElementById(targetId);
  if (!target) return false;

  if (getActiveTabKey() !== "solution") {
    setActiveTab("solution");
  }

  target.scrollIntoView({ behavior: "smooth", block: "start", inline: "nearest" });
  return true;
}

function buildSolutionPdfToc() {
  if (!solutionPdfTocListEl || !panels.solution) return;

  const entries = getSolutionPdfTocEntries();
  if (!entries.length) {
    solutionPdfTocListEl.innerHTML = `
      <li class="toc-item">
        <span class="toc-title">Project Management Solution Pack</span>
        <span class="toc-leader" aria-hidden="true"></span>
        <span class="toc-page">1</span>
      </li>
    `;
    return;
  }

  const hadMeasurementClass = document.body.classList.contains("is-preparing-pdf-layout");
  if (!hadMeasurementClass) {
    document.body.classList.add("is-preparing-pdf-layout");
  }

  try {
    let previousSignature = "";
    for (let pass = 0; pass < 4; pass += 1) {
      const pageNumbers = calculateSolutionPdfTocPageNumbers(entries);
      renderSolutionPdfToc(entries, pageNumbers);

      const signature = pageNumbers.join(",");
      if (signature === previousSignature) break;
      previousSignature = signature;
    }
  } finally {
    if (!hadMeasurementClass) {
      document.body.classList.remove("is-preparing-pdf-layout");
    }
  }
}

function escapeHtml(value) {
  return (value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function exportSolutionPackPdf() {
  if (!panels.solution) {
    generationStatusEl.textContent = "Solution Pack view is not available for export.";
    return;
  }

  const previousTab = getActiveTabKey();
  let hasCleanedUp = false;
  const cleanup = () => {
    if (hasCleanedUp) return;
    hasCleanedUp = true;
    document.body.classList.remove("print-solution-pack");
    if (previousTab && previousTab !== "solution") {
      setActiveTab(previousTab);
    }
  };

  setActiveTab("solution");
  buildSolutionPdfToc();
  document.body.classList.add("print-solution-pack");
  generationStatusEl.textContent = "Preparing Solution Pack PDF export...";

  window.addEventListener("afterprint", cleanup, { once: true });

  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      window.print();
      window.setTimeout(cleanup, 1200);
      generationStatusEl.textContent = "PDF export opened. Use Save as PDF in the print dialog.";
    });
  });
}

function setActiveSolutionTab(key) {
  const safeKey = normalize(key) || "summary";

  solutionSubTabs.forEach((tab) => {
    const isActive = tab.dataset.solutionTab === safeKey;
    tab.classList.toggle("active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
    tab.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  solutionPanels.forEach((panel) => {
    const isActive = panel.id === `solution-panel-${safeKey}`;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });

  if (quickJumpSelectEl && quickJumpSelectEl.value !== safeKey) {
    const exists = Array.from(quickJumpSelectEl.options).some((opt) => opt.value === safeKey);
    if (exists) quickJumpSelectEl.value = safeKey;
  }
}

function setSolutionPanelsCollapsed(collapsed) {
  const collapsibleCards = document.querySelectorAll(".quick-panel[data-collapsible]");
  collapsibleCards.forEach((card) => {
    card.classList.toggle("is-collapsed", collapsed);
  });

  if (quickTogglePanelsEl) {
    quickTogglePanelsEl.dataset.mode = collapsed ? "collapsed" : "expanded";
    quickTogglePanelsEl.textContent = collapsed ? "Expand All" : "Collapse All";
  }
}

async function copyQuickSelection(selection) {
  const timeline = getTimelineData(state.currentPack);
  const matrix = getMatrixData(state.currentPack, timeline);

  if (selection === "timeline") {
    await copyText(renderTimelineText(timeline));
    generationStatusEl.textContent = "Timeline copied.";
    return;
  }

  if (selection === "matrix") {
    await copyText(renderMatrixText(matrix));
    generationStatusEl.textContent = "Tracker matrix copied.";
    return;
  }

  if (selection === "raid") {
    await copyText(renderRaidSnapshotText());
    generationStatusEl.textContent = "RAID snapshot copied.";
    return;
  }

  if (selection === "full") {
    await copyText(renderFullPackText(state.currentPack));
    generationStatusEl.textContent = "Full solution pack copied.";
    return;
  }

  if (selection && sectionEls[selection]) {
    await copyText(sectionEls[selection].textContent || "");
    generationStatusEl.textContent = "Selected section copied.";
    return;
  }

  await copyText(renderFullPackText(state.currentPack));
  generationStatusEl.textContent = "Full solution pack copied.";
}

function getSelectedComplexity() {
  const checked = document.querySelector('input[name="complexity"]:checked');
  return checked ? checked.value : "Standard";
}

function getAIConfig() {
  return {
    mode: "local",
    model: SIMPLE_AI_DEFAULTS.localModel,
    endpoint: "",
    apiKey: "",
  };
}

function syncAIControls(applyDefaults = false) {
  const mode = normalize(aiModeEl?.value) || "local";
  const isCloud = mode === "cloud";

  if (applyDefaults) {
    if (aiModelEl) aiModelEl.value = isCloud ? SIMPLE_AI_DEFAULTS.cloudModel : SIMPLE_AI_DEFAULTS.localModel;
    if (aiEndpointEl && isCloud) aiEndpointEl.value = SIMPLE_AI_DEFAULTS.cloudEndpoint;
  } else if (aiModelEl && !normalize(aiModelEl.value)) {
    aiModelEl.value = isCloud ? SIMPLE_AI_DEFAULTS.cloudModel : SIMPLE_AI_DEFAULTS.localModel;
  }

  if (aiKeyWrapEl) aiKeyWrapEl.classList.toggle("is-hidden", !isCloud);
  if (aiEndpointWrapEl) aiEndpointWrapEl.classList.toggle("is-hidden", !isCloud);

  if (aiModeHintEl) {
    aiModeHintEl.textContent = isCloud
      ? "Use any OpenAI-compatible endpoint. For localhost endpoints (e.g., Ollama), key can be blank."
      : "No external setup required. Uses built-in local AI heuristics.";
  }

  if (aiAssistStatusEl && (applyDefaults || !aiAssistStatusEl.textContent)) {
    aiAssistStatusEl.textContent = isCloud ? "Cloud API mode ready." : "";
  }

  if (aiAssistStatusEl && isCloud && !normalize(aiEndpointEl?.value)) {
    aiAssistStatusEl.textContent = "Add endpoint URL to enable Cloud API mode.";
  }
  if (aiAssistStatusEl && isCloud && !isLocalEndpoint(normalize(aiEndpointEl?.value)) && !normalize(aiApiKeyEl?.value)) {
    aiAssistStatusEl.textContent = "Add API key/token for non-local Cloud API endpoints.";
  }
}

function getIntakeData() {
  const formData = new FormData(intakeForm);
  return {
    projectTitle: normalize(formData.get("projectTitle")),
    projectGoal: normalize(formData.get("projectGoal")),
    ownerName: normalize(formData.get("ownerName")),
    ownerArea: normalize(formData.get("ownerArea")),
    problem: normalize(formData.get("problem")),
    outcomes: toList(formData.get("outcomes")),
    constraints: {
      time: normalize(formData.get("constraintTime")),
      budget: normalize(formData.get("constraintBudget")),
      staffing: normalize(formData.get("constraintStaffing")),
      compliance: normalize(formData.get("constraintCompliance")),
      tech: normalize(formData.get("constraintTech")),
    },
    goLiveDate: normalize(formData.get("goLiveDate")),
    urgency: normalize(formData.get("urgency")) || "Moderate",
    stakeholders: toList(formData.get("stakeholders")),
    systems: toList(formData.get("systems")),
    risks: toList(formData.get("risks")),
  };
}

function initializeIntakeValidation() {
  requiredIntakeFields.forEach((field) => {
    if (!field.id) return;
    if (field.dataset.baseDescribedby === undefined) {
      field.dataset.baseDescribedby = normalize(field.getAttribute("aria-describedby"));
    }
  });
  validateRequiredIntakeFields();
}

function resetIntakeValidationState() {
  requiredIntakeFields.forEach((field) => {
    field.dataset.touched = "false";
    setIntakeFieldValidation(field, true, false);
  });
  updateGenerateButtonState();
}

function isRequiredFieldFilled(field) {
  if (!field) return true;
  const value = normalize(field.value);
  if (field.tagName === "SELECT") return Boolean(value);
  return Boolean(value);
}

function getRequiredFieldMessage(field) {
  const label = normalize(field?.dataset?.requiredLabel) || "this field";
  const lowerLabel = label.charAt(0).toLowerCase() + label.slice(1);
  if (field?.type === "date") {
    return `Please select ${lowerLabel}.`;
  }
  if (field?.tagName === "SELECT") {
    return `Please choose ${lowerLabel}.`;
  }
  return `Please enter ${lowerLabel}.`;
}

function getFieldErrorElement(field) {
  if (!field?.id) return null;
  return document.getElementById(`error-${field.id}`);
}

function setIntakeFieldValidation(field, isValid, showError) {
  if (!field) return;
  const shouldShowError = showError && !isValid;
  const errorEl = getFieldErrorElement(field);
  const errorId = errorEl?.id || "";
  const baseIds = normalize(field.dataset.baseDescribedby)
    .split(/\s+/)
    .filter(Boolean);
  const describedByIds = shouldShowError && errorId
    ? Array.from(new Set([...baseIds, errorId]))
    : baseIds;

  if (describedByIds.length) {
    field.setAttribute("aria-describedby", describedByIds.join(" "));
  } else {
    field.removeAttribute("aria-describedby");
  }

  field.classList.toggle("field-invalid", shouldShowError);
  field.setAttribute("aria-invalid", shouldShowError ? "true" : "false");
  if (errorEl) {
    errorEl.textContent = shouldShowError ? getRequiredFieldMessage(field) : "";
  }
}

function updateGenerateButtonState() {
  if (!generateSolutionBtnEl) return;
  const allComplete = requiredIntakeFields.every((field) => isRequiredFieldFilled(field));
  generateSolutionBtnEl.disabled = state.isGenerating || !allComplete;
  generateSolutionBtnEl.setAttribute("aria-disabled", generateSolutionBtnEl.disabled ? "true" : "false");
}

function validateRequiredIntakeFields({ showErrors = false } = {}) {
  let firstInvalidField = null;

  requiredIntakeFields.forEach((field) => {
    if (showErrors) {
      field.dataset.touched = "true";
    }
    const isValid = isRequiredFieldFilled(field);
    const touched = field.dataset.touched === "true";
    setIntakeFieldValidation(field, isValid, showErrors || touched);
    if (!isValid && !firstInvalidField) {
      firstInvalidField = field;
    }
  });

  updateGenerateButtonState();
  return {
    isValid: !firstInvalidField,
    firstInvalidField,
  };
}

function focusFirstInvalidIntakeField(field) {
  if (!field) return;
  const parentSection = field.closest("details");
  if (parentSection && !parentSection.open) parentSection.open = true;
  field.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  window.setTimeout(() => {
    field.focus({ preventScroll: true });
  }, 180);
}

function populateSampleIntakeForm(sampleProject) {
  if (!sampleProject) return;
  const sampleGoLive = toInputDate(addDays(new Date(), sampleProject.goLiveOffsetDays || 112));
  const sampleFields = {
    projectTitle: sampleProject.projectTitle,
    projectGoal: sampleProject.projectGoal,
    ownerName: sampleProject.ownerName,
    ownerArea: sampleProject.ownerArea,
    problem: sampleProject.problem,
    outcomes: (sampleProject.outcomes || []).join("\n"),
    constraintBudget: sampleProject.constraints?.budget || "",
    constraintStaffing: sampleProject.constraints?.staffing || "",
    constraintCompliance: sampleProject.constraints?.compliance || "",
    constraintTech: sampleProject.constraints?.tech || "",
    goLiveDate: sampleGoLive,
    urgency: sampleProject.urgency || "Moderate",
    stakeholders: (sampleProject.stakeholders || []).join("\n"),
    systems: (sampleProject.systems || []).join("\n"),
    risks: (sampleProject.risks || []).join("\n"),
  };

  Object.entries(sampleFields).forEach(([fieldId, value]) => {
    const field = document.getElementById(fieldId);
    if (!field) return;
    field.value = value;
  });

  const standardComplexity = document.querySelector('input[name="complexity"][value="Standard"]');
  if (standardComplexity) standardComplexity.checked = true;

  document.querySelectorAll(".intake-section").forEach((section) => {
    section.open = true;
  });
}

function getNextSampleProject() {
  const nextIndex = loadSampleRotationIndex();
  const normalizedIndex = nextIndex >= 0 ? nextIndex % SAMPLE_PROJECTS.length : 0;
  const sample = SAMPLE_PROJECTS[normalizedIndex];
  const followingIndex = (normalizedIndex + 1) % SAMPLE_PROJECTS.length;
  persistSampleRotationIndex(followingIndex);
  return sample;
}

function loadSampleRotationIndex() {
  try {
    const raw = sessionStorage.getItem(SAMPLE_ROTATION_KEY);
    const index = Number(raw);
    return Number.isInteger(index) && index >= 0 ? index : 0;
  } catch {
    return 0;
  }
}

function persistSampleRotationIndex(index) {
  try {
    sessionStorage.setItem(SAMPLE_ROTATION_KEY, String(index));
  } catch {
    // Ignore storage failures; rotation will restart from first sample.
  }
}

function normalize(value) {
  return (value || "").toString().trim();
}

function toList(value) {
  return normalize(value)
    .split(/\n|,/)
    .map((line) => line.trim().replace(/^[-*\u2022]\s*/, ""))
    .filter(Boolean);
}

function buildFollowUpQuestions(intake) {
  const questions = [];

  if (!intake.projectTitle) {
    questions.push("What project title should appear in steering committee updates and reporting?");
  }
  if (!intake.projectGoal) {
    questions.push("What is the single measurable goal this project must achieve?");
  }
  if (!intake.ownerName || !intake.ownerArea) {
    questions.push("Who is the accountable sponsor and which business area owns decisions?");
  }
  if (!intake.problem) {
    questions.push("What is the root problem or opportunity, and what happens if no action is taken?");
  }
  if (intake.outcomes.length === 0) {
    questions.push("Which 2-3 outcomes define success in business terms?");
  }
  if (!intake.goLiveDate) {
    questions.push("When is the latest acceptable go-live date, and is it fixed or flexible?");
  }
  if (intake.stakeholders.length === 0) {
    questions.push("Which stakeholder groups must approve or adopt this change?");
  }
  const hasAnyConstraint = Object.values(intake.constraints).some(Boolean);
  if (!hasAnyConstraint) {
    questions.push("Which constraint is non-negotiable right now: time, budget, staffing, compliance, or technology?");
  }

  return questions;
}

function renderFollowUps(questions) {
  followupListEl.innerHTML = "";

  if (questions.length === 0) {
    followupQuestionsEl.classList.add("hidden");
    return;
  }

  questions.forEach((q) => {
    const li = document.createElement("li");
    li.textContent = q;
    followupListEl.appendChild(li);
  });

  followupQuestionsEl.classList.remove("hidden");
}

function buildSolutionPack(intake, complexity, followups, aiConfig = { mode: "local", model: "local-heuristic-v2", endpoint: SIMPLE_AI_DEFAULTS.cloudEndpoint, apiKey: "" }) {
  const submissionTimestamp = new Date().toISOString();
  const context = buildContext(intake, complexity, submissionTimestamp);
  const timeline = buildTimelineData(context);
  const workMatrix = buildWorkMatrix(intake, context, timeline);
  const workstreams = buildWorkstreamsData(intake, context);
  const cloudReady = isCloudMode(aiConfig) && isAIProviderConfigured(aiConfig);
  const effectiveMode = cloudReady ? "cloud" : "local";
  const providerLabel = cloudReady
    ? getAIProviderLabel(aiConfig.endpoint)
    : "Local assistant";
  const aiMeta = {
    mode: effectiveMode,
    model: cloudReady ? (aiConfig.model || SIMPLE_AI_DEFAULTS.cloudModel) : SIMPLE_AI_DEFAULTS.localModel,
    provider: providerLabel,
    endpoint: cloudReady ? (aiConfig.endpoint || SIMPLE_AI_DEFAULTS.cloudEndpoint) : "",
  };
  const sections = {
    sectionA: buildExecutiveSummary(intake, context, followups),
    sectionB: buildGovernanceModel(intake, context),
    sectionC: buildProcessModel(intake, context),
    sectionD: buildPlanningDeliverables(intake, context),
    sectionE: buildExecutionModel(intake, context),
    sectionF: buildExpertCollaboration(intake, context),
    sectionG: buildNextSteps(intake, context),
    sectionH: buildAIInsights(intake, context, followups, timeline, workMatrix, aiMeta),
  };

  return {
    id: createId(),
    createdAt: submissionTimestamp,
    title: intake.projectTitle || "Untitled Project",
    complexity,
    intakeSnapshot: intake,
    context,
    timeline,
    workMatrix,
    workstreams,
    aiMeta,
    sections,
  };
}

function buildWorkstreamsData(intake, context = {}) {
  const streams = [];
  const outcomes = Array.isArray(intake?.outcomes) ? intake.outcomes : [];
  const stakeholders = Array.isArray(intake?.stakeholders) ? intake.stakeholders : [];
  const systems = Array.isArray(intake?.systems) ? intake.systems : [];
  const risks = Array.isArray(intake?.risks) ? intake.risks : [];
  const constraints = intake?.constraints || {};
  const problemText = normalize(intake?.problem);
  const goalText = normalize(intake?.projectGoal);
  const scenario = (normalize(context.scenarioType) || "Cross-functional").toLowerCase();
  const primaryOutcome = outcomes[0] || "the stated project outcomes";
  const keyStakeholders = stakeholders.slice(0, 2).join(" and ") || "key stakeholder groups";
  const techFocus = constraints.tech || systems[0] || "approved enterprise systems";
  const complianceFocus = constraints.compliance || "applicable policy, security, and accessibility requirements";
  const riskFocus = risks[0] || "cross-team dependencies and unresolved delivery blockers";
  const hasCompliance = Boolean(constraints.compliance) || /policy|compliance|legal|security|accessibility/i.test(`${problemText} ${goalText}`);
  const hasDataReporting = /data|report|dashboard|analytics|metrics?/i.test(
    `${problemText} ${goalText} ${outcomes.join(" ")} ${systems.join(" ")}`
  );

  const addWorkstream = (name, focus, suggestedRole) => {
    streams.push({
      name,
      focus,
      pendingOwner: `To be assigned (${suggestedRole})`,
    });
  };

  addWorkstream(
    "Governance and Decisioning",
    `Set decision rights, escalation path, and governance cadence for this ${scenario} initiative.`,
    "Sponsor / PM"
  );

  addWorkstream(
    "Requirements and Process Design",
    `Define requirements, workflow handoffs, and acceptance criteria needed to deliver ${primaryOutcome}.`,
    "Product Owner / Business Analyst"
  );

  addWorkstream(
    "Delivery and Implementation",
    `Coordinate build/configuration, testing, and deployment activities across ${techFocus}.`,
    "Tech Lead / Delivery Lead"
  );

  addWorkstream(
    "Communications",
    `Prepare and deliver clear project communications for ${keyStakeholders}, including updates, decisions, and rollout notices.`,
    "Communications Lead"
  );

  addWorkstream(
    "Training",
    "Design and execute role-based training, enablement materials, and readiness support before go-live.",
    "Training Lead"
  );

  addWorkstream(
    "Data Integration",
    `Coordinate data mapping, integration dependencies, and validation checkpoints across ${techFocus}.`,
    "Data Integration Lead"
  );

  addWorkstream(
    "Risk and Dependency Management",
    `Track, escalate, and mitigate risks and dependencies, including ${riskFocus}.`,
    "Project Manager"
  );

  if (hasCompliance) {
    addWorkstream(
      "Compliance and Controls",
      `Validate controls and approvals against ${complianceFocus} before go-live decisions.`,
      "Compliance / Security Lead"
    );
  }

  if (hasDataReporting) {
    addWorkstream(
      "Data and Reporting",
      "Define core metrics, reporting cadence, and data-quality checkpoints for leadership visibility.",
      "Data / Reporting Lead"
    );
  }

  return streams;
}

function buildContext(intake, complexity, submissionTimestamp) {
  const textBlob = `${intake.projectGoal} ${intake.problem} ${intake.outcomes.join(" ")} ${intake.constraints.compliance}`.toLowerCase();

  let scenarioType = "Cross-functional";
  if (/policy|regulation|compliance|governance/.test(textBlob)) {
    scenarioType = "Policy and compliance";
  } else if (/system|platform|integration|application|technology|data/.test(textBlob)) {
    scenarioType = "Technology enablement";
  } else if (/process|operations|service|workflow|intake/.test(textBlob)) {
    scenarioType = "Operations improvement";
  } else if (/student|faculty|curriculum|academic|program/.test(textBlob)) {
    scenarioType = "Academic initiative";
  }

  const governanceMode = complexity === "Light"
    ? "Lightweight governance"
    : complexity === "Enterprise"
      ? "Full governance"
      : "Balanced governance";

  const urgencyProfile = intake.urgency || "Moderate";
  const constraints = Object.entries(intake.constraints)
    .filter(([, value]) => value)
    .map(([key, value]) => `${capitalize(key)}: ${value}`);

  const topConstraints = constraints.length
    ? constraints.slice(0, 3)
    : ["Time, budget, and staffing assumptions are pending confirmation."];

  const milestones = buildMilestones(
    submissionTimestamp,
    intake.goLiveDate,
    urgencyProfile,
    intake.constraints.time
  );
  const timelineAnchorLabel = formatDate(submissionTimestamp);
  const timelineEndLabel = milestones[milestones.length - 1]?.targetLabel || "TBD";
  const checkpointPlan = buildCheckpointPlan(milestones, intake);

  return {
    scenarioType,
    governanceMode,
    urgencyProfile,
    topConstraints,
    milestones,
    timelineAnchorLabel,
    timelineEndLabel,
    checkpointPlan,
    recommendationBias: "plan",
    ownerLabel: intake.ownerName && intake.ownerArea
      ? `${intake.ownerName} (${intake.ownerArea})`
      : "Sponsor to be confirmed",
  };
}

function buildMilestones(submissionTimestamp, goLiveDate, urgency, timelineConstraint = "") {
  const phases = [
    {
      phase: "Initiation",
      milestone: "Charter, scope, governance, and success metrics approved",
      stage: "Initiation",
      weight: 0.18,
    },
    {
      phase: "Planning",
      milestone: "Requirements baseline, dependency map, and resourcing confirmed",
      stage: "Planning",
      weight: 0.14,
    },
    {
      phase: "Execution",
      milestone: "Solution build/configuration complete and quality-checked",
      stage: "Execution",
      weight: 0.38,
    },
    {
      phase: "Monitor",
      milestone: "Risk, status, and quality checkpoints passed for launch readiness",
      stage: "Monitor",
      weight: 0.2,
    },
    {
      phase: "Closing",
      milestone: "Go-live completed with hypercare and handoff to operations",
      stage: "Closing",
      weight: 0.1,
    },
  ];
  const start = startOfDay(submissionTimestamp);
  const fallbackDurationDays = urgencyToDurationDays(urgency);
  const parsedConstraintDurationDays = parseDurationFromConstraint(timelineConstraint);
  const parsedGoLive = goLiveDate ? startOfDay(goLiveDate) : null;
  const hasValidGoLive = parsedGoLive && !Number.isNaN(parsedGoLive.getTime()) && parsedGoLive >= start;
  const totalDurationDays = hasValidGoLive
    ? daysBetween(start, parsedGoLive)
    : parsedConstraintDurationDays || fallbackDurationDays;
  const phaseCount = phases.length;
  const boundaries = phases.map((item, index) => {
    if (index === phaseCount - 1) return totalDurationDays;
    const throughThisPhase = phases
      .slice(0, index + 1)
      .reduce((sum, phaseItem) => sum + (phaseItem.weight || 0), 0);
    return Math.min(totalDurationDays, Math.round(totalDurationDays * throughThisPhase));
  });

  return phases.map((item, index) => {
    const phaseStartOffset = index === 0 ? 0 : boundaries[index - 1];
    const phaseEndOffset = boundaries[index];
    const phaseStart = addDays(start, phaseStartOffset);
    const phaseEnd = addDays(start, phaseEndOffset);

    return {
      phaseNumber: index + 1,
      phase: item.phase,
      stage: item.stage,
      milestone: item.milestone,
      targetLabel: formatDate(phaseEnd),
      windowLabel: `${formatDate(phaseStart)} - ${formatDate(phaseEnd)}`,
      startISO: phaseStart.toISOString(),
      targetISO: phaseEnd.toISOString(),
    };
  });
}

function buildTimelineData(context) {
  return normalizeMilestones(context.milestones).map((milestone) => ({
    phaseNumber: milestone.phaseNumber,
    stage: milestone.stage,
    phase: milestone.phase,
    milestone: milestone.milestone,
    window: milestone.windowLabel,
    targetDate: milestone.targetLabel,
    startISO: milestone.startISO || "",
    targetISO: milestone.targetISO || "",
  }));
}

function buildWorkMatrix(intake, context, timeline) {
  const outcomes = intake.outcomes.length
    ? intake.outcomes
    : ["Define measurable success outcomes and acceptance criteria"];

  const dueFor = (index) => timeline[index]?.targetDate || timeline[timeline.length - 1]?.targetDate || "TBD";
  const itemStatus = context.urgencyProfile === "Critical"
    ? ["In Progress", "In Progress", "At Risk", "Not Started", "Not Started", "Not Started"]
    : context.urgencyProfile === "High"
      ? ["In Progress", "Not Started", "Not Started", "Not Started", "Not Started", "Not Started"]
      : ["Not Started", "Not Started", "Not Started", "Not Started", "Not Started", "Not Started"];

  return [
    {
      workItem: "Charter and scope signoff",
      owner: "Sponsor + PM",
      status: itemStatus[0],
      dueDate: dueFor(0),
      dependencies: "Project title, goal, and owner confirmed",
    },
    {
      workItem: `Requirements baseline for ${outcomes[0].toLowerCase()}`,
      owner: "Product Owner + SMEs",
      status: itemStatus[1],
      dueDate: dueFor(1),
      dependencies: "Charter and scope signoff",
    },
    {
      workItem: "Solution design and governance approval",
      owner: "PM + Tech Lead",
      status: itemStatus[2],
      dueDate: dueFor(2),
      dependencies: "Requirements baseline complete",
    },
    {
      workItem: "Build/configuration and integration testing",
      owner: "Tech Lead + Delivery Team",
      status: itemStatus[3],
      dueDate: dueFor(3),
      dependencies: "Design approval and environment readiness",
    },
    {
      workItem: "Readiness, training, and communications",
      owner: "Change/Training Lead",
      status: itemStatus[4],
      dueDate: dueFor(4),
      dependencies: "Testing evidence and stakeholder signoff",
    },
    {
      workItem: "Go-live, hypercare, and handoff",
      owner: "PM + Operations Lead",
      status: itemStatus[5],
      dueDate: dueFor(4),
      dependencies: "Readiness checklist signed and launch decision",
    },
  ];
}

function buildExecutiveSummary(intake, context, followups) {
  const projectTitle = intake.projectTitle || "this initiative";
  const goal = intake.projectGoal || "a measurable business outcome";
  const outcomes = intake.outcomes.length
    ? intake.outcomes.slice(0, 3).join("; ")
    : "Outcome measures pending sponsor confirmation.";

  const assumptions = followups.length
    ? "Assumptions to confirm: " + followups.slice(0, 2).join(" | ")
    : "Core intake details are complete; proceed to execution planning.";

  const approach = context.governanceMode === "Lightweight governance"
    ? "Use lightweight governance with fast decisions and lean artifacts."
    : context.governanceMode === "Full governance"
      ? "Use full governance with formal gates, steering approvals, and audit-ready artifacts."
      : "Use balanced governance with clear accountability and minimal overhead.";

  return [
    `- Success looks like: ${projectTitle} delivers ${goal} with adoption across ${intake.stakeholders.slice(0, 2).join(" and ") || "key stakeholder groups"}.`,
    `- Scenario type: ${context.scenarioType} with ${context.urgencyProfile.toLowerCase()} urgency and ${context.governanceMode.toLowerCase()}.`,
    `- Primary outcomes to track: ${outcomes}.`,
    `- Top constraints: ${context.topConstraints.join(" | ")}.`,
    `- Timeline anchor: start on ${context.timelineAnchorLabel}; planned completion by ${context.timelineEndLabel}.`,
    "- Phase progression: Initiation -> Planning -> Execution -> Monitor -> Closing with named checkpoints and owners.",
    `- Recommended approach: ${approach}`,
    `- Sponsor accountability: ${context.ownerLabel}.`,
    `- Assumptions and intake gaps: ${assumptions}`,
  ].join("\n");
}

function buildGovernanceModel(intake, context) {
  const cadence = context.governanceMode === "Lightweight governance"
    ? [
        "Working Group: weekly 45-minute delivery sync",
        "Sponsor touchpoint: biweekly decisions and risk review",
      ]
    : context.governanceMode === "Full governance"
      ? [
          "Steering Committee: biweekly governance, decisions, and escalations",
          "Working Group: weekly delivery and dependency management",
          "SME forum: weekly issue resolution for compliance, data, and technology",
        ]
      : [
          "Steering Committee: monthly decision and performance review",
          "Working Group: weekly planning and execution sync",
          "SME forum: biweekly advisory review",
        ];

  const artifacts = context.governanceMode === "Full governance"
    ? "Required artifacts: charter, RACI, milestone plan, RAID log, decision log, weekly status, readiness checklist, go-live signoff."
    : "Required artifacts: scope one-pager, action log, RAID log, decision log, weekly status summary.";

  return [
    "Recommended governance structure:",
    "- Steering Committee: Sponsor, product/business owner, PM, Tech Lead, and key domain leads.",
    "- Working Group: PM, delivery leads, operations owner, analytics/reporting lead.",
    "- SME Ring: security, accessibility, legal/compliance, procurement, communications, training.",
    "",
    "RACI-style accountability:",
    "Role | Accountability",
    "Sponsor | A - Approves scope, funding, and unresolved tradeoffs",
    "Product Owner | A/R - Prioritizes requirements and accepts deliverables",
    "Project Manager | R - Orchestrates plan, risks, dependencies, reporting",
    "Tech Lead | R - Designs technical approach, integration, and release controls",
    "SMEs | C/R - Provide controls, standards, and signoff inputs",
    "",
    "Decision rights and escalation path:",
    "- Working Group can decide day-to-day sequencing and delivery adjustments within approved scope.",
    "- Steering Committee decides scope changes, budget shifts, timeline resets, and policy exceptions.",
    "- Escalation path: Workstream Lead -> PM -> Sponsor -> Steering Committee within 48 hours for critical blockers.",
    "",
    "Meeting cadence:",
    ...cadence.map((line) => `- ${line}`),
    `- ${artifacts}`,
  ].join("\n");
}

function buildProcessModel(intake, context) {
  const mvpProcess = context.governanceMode === "Lightweight governance"
    ? "MVP process: one intake template, one 60-minute triage, one owner-assigned action list, one weekly checkpoint."
    : "MVP process: one intake template, triage workshop, prioritized backlog, weekly checkpoint with decision log.";

  return [
    "High-level end-to-end process model:",
    "Intake -> Initiation -> Planning -> Execution -> Monitor -> Closing",
    "",
    "Workflow mapping suggestions:",
    "- Intake handoff: Sponsor submits scope intent; PM validates clarity and dependencies within 2 business days.",
    "- Assessment approval: SMEs confirm constraints (security, compliance, data, procurement) before planning baseline.",
    "- Plan approval: Steering confirms timeline, resources, and measurable outcomes.",
    "- Execution controls: Use action log and decision log to track ownership and unresolved blockers.",
    "- Validation gate: UAT/readiness checklist signed by business owner and delivery leads before launch.",
    "- SLA guidance: Intake triage <= 2 business days, dependency resolution <= 5 business days, escalation response <= 48 hours.",
    "",
    "Minimum viable process option for small projects:",
    `- ${mvpProcess}`,
    "- Keep approval chain to Sponsor + PM + one SME reviewer; avoid unnecessary review layers.",
  ].join("\n");
}

function buildPlanningDeliverables(intake, context) {
  const outcomes = intake.outcomes.length ? intake.outcomes : ["Define measurable outcomes with sponsor"];
  const inScope = outcomes.slice(0, 3).map((item) => `- ${item}`).join("\n");

  const outScope = [
    "- Non-critical enhancements not tied to stated outcomes",
    "- New integrations outside approved timeline",
    "- Policy exceptions without steering approval",
  ].join("\n");

  const dependencyMap = [
    "- Sponsor decision on scope and success metrics",
    "- SME validation of compliance/security/data constraints",
    "- Tooling or system access readiness",
    "- Training and communications content before launch",
    "- Operational handoff ownership for post-launch support",
  ].join("\n");

  const resourcePlan = [
    "- Sponsor (decision authority)",
    "- Product Owner / Business Lead",
    "- Project Manager",
    "- Tech Lead / Systems Analyst",
    "- Operations Lead",
    "- Data/Reporting Analyst",
    "- Change, Communications, and Training Lead",
  ].join("\n");

  const topRiskInputs = intake.risks.length ? intake.risks.slice(0, 2) : ["Scope creep", "Resource contention"];

  const raidLog = [
    `1) Risk: ${topRiskInputs[0] || "Ambiguous scope"} | Mitigation: lock acceptance criteria and run weekly scope review.`,
    `2) Risk: ${topRiskInputs[1] || "Cross-team dependency delays"} | Mitigation: dependency tracker with owner and due date.`,
    "3) Risk: Slow decisions from governance groups | Mitigation: define decision SLA and escalation trigger.",
    "4) Risk: Compliance/accessibility rework near launch | Mitigation: involve SMEs in assessment phase and readiness gate.",
    "5) Risk: Adoption gap after deployment | Mitigation: publish rollout communications, training, and support plan.",
  ].join("\n");

  const checkpointPlan = (context.checkpointPlan || [])
    .map((checkpoint, idx) => `${idx + 1}. ${checkpoint}`)
    .join("\n");

  const stakeholderResponsibilities = buildStakeholderResponsibilityLines(intake);

  return [
    "Scope statement:",
    `Project: ${intake.projectTitle || "Project to be named"}`,
    `Goal: ${intake.projectGoal || "Goal pending sponsor input"}`,
    "",
    "In scope:",
    inScope,
    "",
    "Out of scope:",
    outScope,
    "",
    "Milestones and phased timeline:",
    ...normalizeMilestones(context.milestones).map((m) => `- ${formatMilestoneLine(m)}`),
    "",
    "Dependency map:",
    dependencyMap,
    "",
    "Resource plan (roles needed):",
    resourcePlan,
    "",
    "Stakeholder responsibilities (kickoff baseline):",
    stakeholderResponsibilities,
    "",
    "Checkpoint schedule (Initiation, Planning, Execution, Monitor, Closing):",
    checkpointPlan,
    "",
    "RAID log starter (top 5 risks + mitigations):",
    raidLog,
  ].join("\n");
}

function buildExecutionModel(intake, context) {
  const reportingFocus = context.recommendationBias || "plan";
  const operationalCheckpoints = (context.checkpointPlan || [])
    .slice(0, 4)
    .map((checkpoint) => `- ${checkpoint}`)
    .join("\n");
  return [
    "Recommended operating rhythm:",
    "- Weekly status review: progress, milestones, blockers, risks, and decisions.",
    "- Action log review: verify owner, due date, status, and blocker notes.",
    "- Decision log review: unresolved decisions, decision owner, due date, impact.",
    "- Monitor checkpoint review: verify phase exit criteria before moving forward.",
    "",
    "Ownership model for deliverables:",
    "- Sponsor owns strategic decisions and benefit realization.",
    "- Product Owner owns requirement quality and acceptance criteria.",
    "- PM owns timeline, dependencies, RAID management, and reporting cadence.",
    "- Tech/Operations leads own solution build, testing, and readiness outputs.",
    "- SMEs own policy/compliance/accessibility controls at each checkpoint.",
    "",
    "Operational checkpoints:",
    operationalCheckpoints,
    "",
    "Leadership-ready reporting template:",
    "- Overall status: Green / Amber / Red",
    "- This period accomplishments: top 3",
    "- Next period priorities: top 3",
    "- Milestone health: on-track / at-risk / delayed",
    "- Top risks and mitigations: top 3",
    "- Decisions needed this week: owner + due date",
    `- Reporting emphasis: ${capitalize(reportingFocus)} and accountability transparency`,
    `- Governance posture: ${context.governanceMode}`,
  ].join("\n");
}

function buildExpertCollaboration(intake, context) {
  const systemsSummary = intake.systems.length
    ? intake.systems.join(", ")
    : "Systems to be confirmed during assessment";

  return [
    "SME involvement recommendations:",
    "- Security: engage during assessment to define controls and signoff criteria.",
    "- Data/Analytics: engage during planning for data definitions, reporting metrics, and quality checks.",
    "- Procurement: engage during planning if vendors/tools/contracts are impacted.",
    "- Accessibility: engage at requirements stage and validate before launch.",
    "- Legal/Compliance: engage at assessment and any scope change with policy impact.",
    "- Communications: engage before execution to align messaging and adoption timeline.",
    "- Training: engage during execution to prepare role-based enablement before go-live.",
    `- Impacted systems/tools context: ${systemsSummary}.`,
    "",
    "Workshop structure to move fast:",
    "- Discovery session (90 minutes): clarify problem, outcomes, constraints, and success metrics.",
    "- Requirements jam (120 minutes): define must-have requirements, acceptance criteria, and priorities.",
    "- Process mapping session (90 minutes): document current-to-future state, handoffs, approvals, and SLAs.",
    "- Decision closeout (45 minutes): finalize tradeoffs, owners, dates, and escalation triggers.",
    `- Workshop operating mode: ${context.governanceMode} with clear pre-reads and owner-assigned outputs.`,
  ].join("\n");
}

function buildNextSteps(intake, context) {
  const startDate = new Date();
  const weekLabel = `week of ${formatDate(startDate)}`;

  return [
    `Do this next week checklist (${weekLabel}):`,
    "1. Confirm sponsor, product owner, and PM accountability in writing.",
    "2. Finalize scope statement with in-scope and out-of-scope boundaries.",
    "3. Run a 90-minute discovery session to align on outcomes and constraints.",
    "4. Build the initial milestone plan and dependency tracker.",
    "5. Stand up RAID log and decision log with named owners.",
    "6. Validate compliance, accessibility, security, and legal requirements with SMEs.",
    "7. Lock meeting cadence and publish required artifact templates.",
    "8. Prepare leadership-ready weekly status format and first reporting cycle.",
    "9. Define launch readiness criteria and draft rollout/training plan.",
    `10. Align governance strictness to ${context.governanceMode.toLowerCase()} and enforce escalation within 48 hours for blockers.`,
  ].join("\n");
}

function buildCheckpointPlan(milestones, intake) {
  const normalized = normalizeMilestones(milestones);
  const stakeholderGroups = intake.stakeholders.length
    ? intake.stakeholders.slice(0, 3).join(", ")
    : "key stakeholder groups";

  return normalized.map((milestone, index) => {
    const owner = index === 0
      ? "Sponsor + PM"
      : index === 1
        ? "Product Owner + SMEs"
        : index === 2
          ? "Tech Lead + PM"
          : index === 3
            ? "PM + Quality/Compliance SMEs"
            : "Sponsor + Operations Lead";

    return `Phase ${milestone.phaseNumber} (${milestone.stage}) checkpoint by ${milestone.targetLabel}: ${milestone.milestone}. Owner: ${owner}. Required signoff: ${stakeholderGroups}.`;
  });
}

function buildStakeholderResponsibilityLines(intake) {
  const groups = intake.stakeholders.length
    ? intake.stakeholders.slice(0, 5)
    : ["Sponsor office", "Business owner", "Delivery team", "SMEs", "End users"];

  return groups
    .map((group, index) => {
      const responsibility = index === 0
        ? "Approves scope, priorities, and major tradeoffs"
        : index === 1
          ? "Validates requirements and accepts deliverables"
          : index === 2
            ? "Executes plan, resolves blockers, and reports progress"
            : index === 3
              ? "Provides controls, standards, and checkpoint signoff inputs"
              : "Participates in validation, readiness, and adoption";
      return `- ${group}: ${responsibility}`;
    })
    .join("\n");
}

function buildAIInsights(intake, context, followups, timeline, workMatrix, aiMeta = { provider: "Local assistant" }) {
  const requiredChecks = [
    Boolean(intake.projectTitle),
    Boolean(intake.projectGoal),
    Boolean(intake.ownerName),
    Boolean(intake.ownerArea),
    Boolean(intake.problem),
    intake.outcomes.length > 0,
    intake.stakeholders.length > 0,
    Boolean(intake.goLiveDate),
  ];
  const completeness = Math.round((requiredChecks.filter(Boolean).length / requiredChecks.length) * 100);

  const riskScore = [
    context.urgencyProfile === "Critical" ? 2 : context.urgencyProfile === "High" ? 1 : 0,
    Object.values(intake.constraints).filter(Boolean).length >= 4 ? 1 : 0,
    followups.length > 2 ? 1 : 0,
  ].reduce((sum, value) => sum + value, 0);

  const riskSignal = riskScore >= 3 ? "High" : riskScore === 2 ? "Moderate" : "Low";
  const nextMilestone = timeline[0]?.targetDate || "TBD";
  const delayedItems = workMatrix.filter((item) => item.status === "At Risk").length;
  const constraintsList = Array.isArray(context.topConstraints) ? context.topConstraints : [];
  const constraints = constraintsList.length ? constraintsList.join("; ") : "No constraints recorded.";
  const aiLabel = aiMeta.provider || "Local assistant";

  return [
    `AI insights source: ${aiLabel}`,
    `Intake completeness score: ${completeness}%`,
    `Delivery risk signal: ${riskSignal}`,
    `Next milestone target: ${nextMilestone}`,
    "",
    "AI priority recommendations:",
    `1. Lock decision owners and acceptance criteria for ${context.scenarioType.toLowerCase()} workstreams in this cycle.`,
    `2. Triage constraints early: ${constraints}`,
    `3. Run a focused dependency review with PM + Tech Lead + SMEs before Phase 2 gate.`,
    `4. Require weekly risk aging review and close/open decisions with due dates.`,
    `5. Publish an adoption readiness mini-plan no later than ${timeline[3]?.targetDate || timeline[timeline.length - 1]?.targetDate || "the validation phase"}.`,
    "",
    `Signals monitored: follow-up gaps (${followups.length}), at-risk matrix items (${delayedItems}), urgency (${context.urgencyProfile}).`,
  ].join("\n");
}

async function generateIntakeSuggestions(intake, aiConfig) {
  if (!isCloudMode(aiConfig)) {
    return generateLocalIntakeSuggestions(intake);
  }

  if (!isAIProviderConfigured(aiConfig)) {
    const fallback = generateLocalIntakeSuggestions(intake);
    fallback.summary = `${getAIProviderLabel(aiConfig.endpoint)} is not configured. Local suggestions applied.`;
    return fallback;
  }

  try {
    const prompt = [
      "Return compact JSON only.",
      "Generate actionable intake suggestions for a project management form.",
      "Do not include markdown or code fences.",
      "Schema:",
      "{\"projectGoal\":\"string\",\"outcomes\":[\"string\"],\"stakeholders\":[\"string\"],\"risks\":[\"string\"],\"constraints\":{\"time\":\"string\",\"budget\":\"string\",\"staffing\":\"string\",\"compliance\":\"string\",\"tech\":\"string\"},\"summary\":\"string\"}",
      "Use concise enterprise-ready language.",
      `Current intake: ${JSON.stringify(intake)}`,
    ].join("\n");

    const responseText = await callAIProviderText(prompt, aiConfig);
    const parsed = parseJsonResponse(responseText);
    return sanitizeSuggestionPayload(parsed, "AI suggestions applied.");
  } catch {
    const fallback = generateLocalIntakeSuggestions(intake);
    fallback.summary = `${getAIProviderLabel(aiConfig.endpoint)} request failed. Local suggestions applied instead.`;
    return fallback;
  }
}

function generateLocalIntakeSuggestions(intake) {
  const scenarioHint = inferScenarioHint(intake);
  const suggestedGoal = intake.projectGoal || `Deliver a measurable improvement in ${scenarioHint} outcomes within the agreed timeline.`;
  const suggestedOutcomes = intake.outcomes.length ? intake.outcomes : [
    "Establish clear scope, ownership, and measurable success criteria.",
    "Reduce cycle time and unresolved dependencies across teams.",
    "Launch with documented governance, reporting, and adoption support.",
  ];
  const suggestedStakeholders = intake.stakeholders.length ? intake.stakeholders : [
    "Business leadership / sponsor",
    "Operations and process owners",
    "Technology and data teams",
    "End-user representatives",
  ];
  const suggestedRisks = intake.risks.length ? intake.risks : [
    "Decision delays due to cross-team dependencies",
    "Scope expansion without governance approval",
    "Late compliance or accessibility changes",
  ];

  return sanitizeSuggestionPayload({
    projectGoal: suggestedGoal,
    outcomes: suggestedOutcomes,
    stakeholders: suggestedStakeholders,
    risks: suggestedRisks,
    constraints: {
      time: intake.constraints.time || "Define a fixed delivery window with phased checkpoints.",
      budget: intake.constraints.budget || "Confirm approved budget ceiling and change-control threshold.",
      staffing: intake.constraints.staffing || "Identify core roles (PM, BA, Tech Lead, SMEs) and weekly capacity.",
      compliance: intake.constraints.compliance || "Confirm applicable policy, legal, and accessibility requirements early.",
      tech: intake.constraints.tech || "Confirm approved tools, integrations, and environment readiness constraints.",
    },
    summary: "AI suggestions were added to strengthen missing inputs.",
  }, "AI suggestions were added to strengthen missing inputs.");
}

function sanitizeSuggestionPayload(payload, fallbackSummary) {
  return {
    projectGoal: normalize(payload?.projectGoal),
    outcomes: Array.isArray(payload?.outcomes) ? payload.outcomes.map(normalize).filter(Boolean).slice(0, 6) : [],
    stakeholders: Array.isArray(payload?.stakeholders) ? payload.stakeholders.map(normalize).filter(Boolean).slice(0, 8) : [],
    risks: Array.isArray(payload?.risks) ? payload.risks.map(normalize).filter(Boolean).slice(0, 8) : [],
    constraints: {
      time: normalize(payload?.constraints?.time),
      budget: normalize(payload?.constraints?.budget),
      staffing: normalize(payload?.constraints?.staffing),
      compliance: normalize(payload?.constraints?.compliance),
      tech: normalize(payload?.constraints?.tech),
    },
    summary: normalize(payload?.summary) || fallbackSummary,
  };
}

function applyIntakeSuggestions(suggestions) {
  const goalField = document.getElementById("projectGoal");
  const outcomesField = document.getElementById("outcomes");
  const stakeholdersField = document.getElementById("stakeholders");
  const risksField = document.getElementById("risks");
  const timeField = document.getElementById("constraintTime");
  const budgetField = document.getElementById("constraintBudget");
  const staffingField = document.getElementById("constraintStaffing");
  const complianceField = document.getElementById("constraintCompliance");
  const techField = document.getElementById("constraintTech");

  if (goalField && !normalize(goalField.value) && suggestions.projectGoal) goalField.value = suggestions.projectGoal;
  if (outcomesField && !normalize(outcomesField.value) && suggestions.outcomes.length) outcomesField.value = suggestions.outcomes.join("\n");
  if (stakeholdersField && !normalize(stakeholdersField.value) && suggestions.stakeholders.length) stakeholdersField.value = suggestions.stakeholders.join("\n");
  if (risksField && !normalize(risksField.value) && suggestions.risks.length) risksField.value = suggestions.risks.join("\n");
  if (timeField && !normalize(timeField.value) && suggestions.constraints.time) timeField.value = suggestions.constraints.time;
  if (budgetField && !normalize(budgetField.value) && suggestions.constraints.budget) budgetField.value = suggestions.constraints.budget;
  if (staffingField && !normalize(staffingField.value) && suggestions.constraints.staffing) staffingField.value = suggestions.constraints.staffing;
  if (complianceField && !normalize(complianceField.value) && suggestions.constraints.compliance) complianceField.value = suggestions.constraints.compliance;
  if (techField && !normalize(techField.value) && suggestions.constraints.tech) techField.value = suggestions.constraints.tech;
}

async function enhancePackWithAI(pack, intake, followups, aiConfig) {
  if (!isCloudMode(aiConfig)) {
    return pack;
  }

  if (!isAIProviderConfigured(aiConfig)) {
    pack.aiMeta = {
      mode: "local",
      provider: "Local assistant",
      model: "local-heuristic-v2",
    };
    return pack;
  }

  try {
    const prompt = [
      "Return compact JSON only.",
      "You are enhancing a project management solution pack.",
      "Do not rewrite existing sections.",
      "Create only one field: sectionH (AI insights and recommendations).",
      "Keep it direct, leadership-friendly, actionable, and concise.",
      "Include: readiness signal, top 5 actions, risk signal, and any missing critical clarifications.",
      `Project intake: ${JSON.stringify(intake)}`,
      `Baseline context: ${JSON.stringify(pack.context)}`,
      `Timeline: ${JSON.stringify(pack.timeline)}`,
      `Tracker matrix: ${JSON.stringify(pack.workMatrix)}`,
      `Follow-up questions: ${JSON.stringify(followups)}`,
      "Schema: {\"sectionH\":\"string\"}",
    ].join("\n");

    const responseText = await callAIProviderText(prompt, aiConfig);
    const parsed = parseJsonResponse(responseText);
    const sectionH = normalize(parsed?.sectionH);

    if (sectionH) {
      pack.sections.sectionH = sectionH;
      pack.aiMeta = {
        mode: "cloud",
        provider: getAIProviderLabel(aiConfig.endpoint),
        model: aiConfig.model || SIMPLE_AI_DEFAULTS.cloudModel,
        endpoint: aiConfig.endpoint || SIMPLE_AI_DEFAULTS.cloudEndpoint,
      };
    }
  } catch {
    pack.aiMeta = {
      mode: "local",
      provider: "Local assistant (provider unavailable)",
      model: "local-heuristic-v2",
    };
  }

  return pack;
}

async function callAIProviderText(prompt, aiConfig) {
  const endpoint = normalize(aiConfig.endpoint) || SIMPLE_AI_DEFAULTS.cloudEndpoint;
  if (isOllamaChatEndpoint(endpoint)) {
    return callOllamaText(prompt, aiConfig, endpoint);
  }

  const headers = {
    "Content-Type": "application/json",
  };
  if (normalize(aiConfig.apiKey)) {
    headers.Authorization = `Bearer ${aiConfig.apiKey}`;
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers,
    body: JSON.stringify({
      model: aiConfig.model || SIMPLE_AI_DEFAULTS.cloudModel,
      temperature: 0.2,
      messages: [
        {
          role: "system",
          content: "You are a project management copilot. Output valid JSON only and keep recommendations practical.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    }),
  });

  if (!response.ok) {
    throw new Error(`Provider request failed: ${response.status}`);
  }

  const data = await response.json();
  return data?.choices?.[0]?.message?.content || "";
}

async function callOllamaText(prompt, aiConfig, endpoint = "http://localhost:11434/api/chat") {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: aiConfig.model || "llama3.1:8b",
      stream: false,
      messages: [
        {
          role: "system",
          content: "You are a project management copilot. Output valid JSON only and keep recommendations practical.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    }),
  });

  if (!response.ok) {
    throw new Error(`Ollama request failed: ${response.status}`);
  }

  const data = await response.json();
  return data?.message?.content || "";
}

async function testAIConnection(aiConfig) {
  if (!isCloudMode(aiConfig)) {
    return { ok: true, message: "Local assistant is active." };
  }

  if (!isAIProviderConfigured(aiConfig)) {
    return { ok: false, message: `${getAIProviderLabel(aiConfig.endpoint)} is not configured yet. Check endpoint or API key.` };
  }

  try {
    const probePrompt = [
      "Return JSON only: {\"ok\":true,\"provider\":\"string\"}",
      "Keep response to one short JSON object.",
    ].join("\n");
    const responseText = await callAIProviderText(probePrompt, aiConfig);
    const parsed = tryJsonParse((responseText || "").trim());
    const hasSignal = Boolean((responseText || "").trim()) || Boolean(parsed);
    return hasSignal
      ? { ok: true, message: `${getAIProviderLabel(aiConfig.endpoint)} connection successful.` }
      : { ok: false, message: `${getAIProviderLabel(aiConfig.endpoint)} responded, but returned empty output.` };
  } catch {
    return { ok: false, message: `${getAIProviderLabel(aiConfig.endpoint)} connection failed. Verify configuration and retry.` };
  }
}

function isAIProviderConfigured(aiConfig) {
  if (!isCloudMode(aiConfig)) return true;
  const endpoint = normalize(aiConfig.endpoint) || SIMPLE_AI_DEFAULTS.cloudEndpoint;
  if (!endpoint) return false;
  if (!isLocalEndpoint(endpoint) && !normalize(aiConfig.apiKey)) return false;
  return true;
}

function getAIProviderLabel(endpoint) {
  const url = normalize(endpoint).toLowerCase();
  if (!url) return "Cloud API";
  if (url.includes("openrouter.ai")) return "Cloud API (OpenRouter)";
  if (url.includes("huggingface.co")) return "Cloud API (Hugging Face)";
  if (isLocalEndpoint(url)) return "Cloud API (Local endpoint)";
  return "Cloud API (Custom endpoint)";
}

function isCloudMode(aiConfig) {
  return false;
}

function isLocalEndpoint(endpoint) {
  const url = normalize(endpoint).toLowerCase();
  return url.includes("localhost") || url.includes("127.0.0.1");
}

function isOllamaChatEndpoint(endpoint) {
  return /\/api\/chat\/?$/.test(normalize(endpoint).toLowerCase());
}

function parseJsonResponse(text) {
  const safeText = (text || "").toString().trim();
  const direct = tryJsonParse(safeText);
  if (direct) return direct;

  const fencedMatch = safeText.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fencedMatch) {
    const parsedFenced = tryJsonParse(fencedMatch[1]);
    if (parsedFenced) return parsedFenced;
  }

  const start = safeText.indexOf("{");
  const end = safeText.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const parsedSlice = tryJsonParse(safeText.slice(start, end + 1));
    if (parsedSlice) return parsedSlice;
  }

  throw new Error("Could not parse JSON from AI response.");
}

function tryJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function inferScenarioHint(intake) {
  const blob = `${intake.problem} ${intake.projectGoal} ${intake.outcomes.join(" ")}`.toLowerCase();
  if (/policy|compliance|legal|regulation/.test(blob)) return "policy and compliance";
  if (/system|integration|data|platform|technology/.test(blob)) return "technology delivery";
  if (/student|faculty|academic|curriculum/.test(blob)) return "academic operations";
  return "cross-functional delivery";
}

function getTimelineData(pack) {
  if (Array.isArray(pack.timeline) && pack.timeline.length) {
    return pack.timeline;
  }
  if (pack.context?.milestones) {
    return buildTimelineData({ milestones: pack.context.milestones });
  }
  return [];
}

function getMatrixData(pack, timeline) {
  if (Array.isArray(pack.workMatrix) && pack.workMatrix.length) {
    return pack.workMatrix;
  }
  const legacyIntake = pack.intakeSnapshot || { outcomes: [] };
  return buildWorkMatrix(legacyIntake, pack.context || { urgencyProfile: "Moderate" }, timeline);
}

function renderTimeline(timeline) {
  timelineVisualEl.innerHTML = "";
  renderTimelineMetrics(timeline);

  if (!timeline.length) {
    const emptyNode = document.createElement("li");
    emptyNode.className = "timeline-node timeline-empty";
    emptyNode.textContent = "No milestones available.";
    timelineVisualEl.appendChild(emptyNode);
    return;
  }

  timeline.forEach((item, index) => {
    const phaseStatus = getPhaseProgressStatus(item, index, timeline);
    const node = document.createElement("li");
    node.className = "timeline-node";

    const phase = document.createElement("span");
    phase.className = "timeline-phase";
    phase.textContent = `Phase ${item.phaseNumber || index + 1}`;

    const title = document.createElement("span");
    title.className = "timeline-title";
    title.textContent = item.milestone || item.phase || "Milestone";

    const detail = document.createElement("span");
    detail.className = "timeline-detail";
    const stageLabel = item.stage ? `${item.stage} | ` : "";
    detail.textContent = `${stageLabel}${item.window || item.phase || ""}`;

    const status = document.createElement("span");
    status.className = `phase-status ${getPhaseStatusClass(phaseStatus)}`;
    status.textContent = `${getPhaseStatusIcon(phaseStatus)} ${phaseStatus}`;

    const date = document.createElement("span");
    date.className = "timeline-date";
    date.textContent = item.targetDate || "TBD";

    node.appendChild(phase);
    node.appendChild(title);
    node.appendChild(detail);
    node.appendChild(status);
    node.appendChild(date);
    timelineVisualEl.appendChild(node);
  });
}

function renderTimelineMetrics(timeline) {
  if (!timelineMetricsEl) return;
  timelineMetricsEl.innerHTML = "";

  if (!timeline.length) return;

  const counts = { Complete: 0, "In Progress": 0, Upcoming: 0 };
  timeline.forEach((item, index) => {
    const status = getPhaseProgressStatus(item, index, timeline);
    counts[status] = (counts[status] || 0) + 1;
  });

  const chips = [
    { label: "Complete", count: counts.Complete, className: "metric-complete", icon: "✓" },
    { label: "In Progress", count: counts["In Progress"], className: "metric-progress", icon: "◔" },
    { label: "Upcoming", count: counts.Upcoming, className: "metric-upcoming", icon: "○" },
  ];

  chips.forEach((chip) => {
    const item = document.createElement("div");
    item.className = `metric-chip ${chip.className}`;
    item.innerHTML = `<span class="metric-icon">${chip.icon}</span><span>${chip.label}: ${chip.count}</span>`;
    timelineMetricsEl.appendChild(item);
  });
}

function renderMatrix(matrixRows) {
  matrixBodyEl.innerHTML = "";

  if (!matrixRows.length) {
    const emptyRow = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 5;
    td.textContent = "No work items available.";
    emptyRow.appendChild(td);
    matrixBodyEl.appendChild(emptyRow);
    return;
  }

  matrixRows.forEach((row) => {
    const tr = document.createElement("tr");

    const itemTd = document.createElement("td");
    itemTd.textContent = row.workItem;

    const ownerTd = document.createElement("td");
    ownerTd.textContent = row.owner;

    const statusTd = document.createElement("td");
    const statusPill = document.createElement("span");
    statusPill.className = `status-pill ${getStatusClass(row.status)}`;
    statusPill.textContent = row.status;
    statusTd.appendChild(statusPill);

    const dueTd = document.createElement("td");
    dueTd.textContent = row.dueDate;

    const dependencyTd = document.createElement("td");
    dependencyTd.textContent = row.dependencies;

    tr.appendChild(itemTd);
    tr.appendChild(ownerTd);
    tr.appendChild(statusTd);
    tr.appendChild(dueTd);
    tr.appendChild(dependencyTd);
    matrixBodyEl.appendChild(tr);
  });
}

function normalizePhaseLabel(label) {
  const text = normalize(label).toLowerCase();
  if (text === "monitor") return "M&C";
  if (text === "closing") return "Closure";
  if (text === "initiation") return "Initiation";
  if (text === "planning") return "Planning";
  if (text === "execution") return "Execution";
  if (text === "m&c") return "M&C";
  if (text === "closure") return "Closure";
  return capitalize(text) || "Initiation";
}

function renderProjectStatus(pack, timeline, matrixRows) {
  const statuses = timeline.map((item, index) => getPhaseProgressStatus(item, index, timeline));
  const completeCount = statuses.filter((status) => status === "Complete").length;
  const hasInProgress = statuses.includes("In Progress");
  const progressPercent = timeline.length
    ? Math.max(0, Math.min(100, Math.round(((completeCount + (hasInProgress ? 0.5 : 0)) / timeline.length) * 100)))
    : 0;

  let currentIndex = statuses.indexOf("In Progress");
  if (currentIndex < 0) {
    currentIndex = completeCount >= timeline.length && timeline.length ? timeline.length - 1 : 0;
  }
  const currentPhaseRaw = timeline[currentIndex]?.phase || "Initiation";
  const currentPhase = normalizePhaseLabel(currentPhaseRaw);

  if (statusPhaseLabelEl) statusPhaseLabelEl.textContent = `Current phase: ${currentPhase}`;
  if (statusPercentLabelEl) statusPercentLabelEl.textContent = `${progressPercent}%`;

  if (statusProgressFillEl) {
    statusProgressFillEl.style.width = `${progressPercent}%`;
    statusProgressFillEl.parentElement?.setAttribute("aria-valuenow", String(progressPercent));
  }

  const priority = pack?.intakeSnapshot?.urgency || pack?.context?.urgencyProfile || "Moderate";
  if (statusPriorityBadgeEl) {
    statusPriorityBadgeEl.textContent = `Priority: ${priority}`;
    statusPriorityBadgeEl.className = `status-kpi-badge priority-${normalize(priority).toLowerCase() || "moderate"}`;
  }

  const complexity = pack?.complexity || "Standard";
  if (statusComplexityBadgeEl) {
    statusComplexityBadgeEl.textContent = `Complexity: ${complexity}`;
    statusComplexityBadgeEl.className = `status-kpi-badge complexity-${normalize(complexity).toLowerCase() || "standard"}`;
  }

  if (phaseStepperEl) {
    phaseStepperEl.innerHTML = "";
    const currentStep = STATUS_PHASES.indexOf(currentPhase);
    STATUS_PHASES.forEach((phase, index) => {
      const li = document.createElement("li");
      li.className = "step-box";
      if (index < currentStep) li.classList.add("step-complete");
      if (index === currentStep) li.classList.add("step-active");
      li.textContent = phase;
      phaseStepperEl.appendChild(li);
    });
  }

  renderMiniMetrics(pack, timeline, matrixRows);
}

function renderMiniMetrics(pack, timeline, matrixRows) {
  if (!miniMetricsEl) return;
  const stakeholders = pack?.intakeSnapshot?.stakeholders?.length || 0;
  const risks = pack?.intakeSnapshot?.risks?.length || 0;
  const deliverables = matrixRows?.length || 0;
  const targetDate = timeline[timeline.length - 1]?.targetISO || timeline[timeline.length - 1]?.targetDate;
  const target = parseFlexibleDate(targetDate);
  const gapDays = target ? daysUntilDate(target) : null;

  const cards = [
    { label: "Stakeholders", value: stakeholders || "0" },
    { label: "Deliverables", value: deliverables || "0" },
    { label: "Risks", value: risks || "0" },
    { label: "Days to target", value: gapDays === null ? "--" : String(gapDays) },
  ];

  miniMetricsEl.innerHTML = "";
  cards.forEach((card) => {
    const div = document.createElement("div");
    div.className = "mini-metric-card";
    div.innerHTML = `<span class="mini-metric-value">${card.value}</span><span class="mini-metric-label">${card.label}</span>`;
    miniMetricsEl.appendChild(div);
  });
}

function daysUntilDate(dateValue) {
  const target = parseFlexibleDate(dateValue);
  if (!target) return null;
  const today = startOfDay(new Date());
  const diffMs = target.getTime() - today.getTime();
  return Math.round(diffMs / (24 * 60 * 60 * 1000));
}

function renderOutcomeCharts(pack, timeline, matrixRows) {
  renderPhaseDurationChart(timeline);
  renderEffortAllocationChart(pack, timeline, matrixRows);
  renderMilestoneBurnupChart(timeline);
}

function renderPhaseDurationChart(timeline) {
  if (!phaseDurationChartEl) return;

  if (!timeline.length) {
    phaseDurationChartEl.innerHTML = '<p class="chart-empty">Generate a solution pack to view phase durations.</p>';
    return;
  }

  const rows = timeline.map((item, index) => {
    const durationDays = estimatePhaseDurationDays(item, index, timeline);
    return {
      phase: normalizePhaseLabel(item.phase || item.stage || `Phase ${index + 1}`),
      days: durationDays,
    };
  });

  const maxDays = Math.max(...rows.map((row) => row.days), 1);
  phaseDurationChartEl.innerHTML = rows.map((row) => {
    const width = Math.max(6, Math.round((row.days / maxDays) * 100));
    return `
      <div class="bar-row">
        <span class="bar-label">${row.phase}</span>
        <div class="bar-track"><span class="bar-fill" style="width:${width}%"></span></div>
        <span class="bar-value">${row.days}d</span>
      </div>
    `;
  }).join("");
}

function renderEffortAllocationChart(pack, timeline, matrixRows) {
  if (!effortDonutEl || !effortLegendEl) return;

  const buckets = [
    { key: "Initiation", label: "Initiation", color: "var(--accent-summary)" },
    { key: "Planning", label: "Planning", color: "var(--accent-governance)" },
    { key: "Execution", label: "Execution", color: "var(--accent-deliverables)" },
    { key: "M&C", label: "Monitoring & Controlling", color: "var(--accent-risks)" },
    { key: "Closure", label: "Closure", color: "var(--accent-next)" },
  ];

  const totals = {};
  buckets.forEach((bucket) => { totals[bucket.key] = 0; });

  timeline.forEach((item, index) => {
    const phase = normalizePhaseLabel(item.phase || item.stage || `Phase ${index + 1}`);
    const days = estimatePhaseDurationDays(item, index, timeline);
    totals[phase] = (totals[phase] || 0) + days;
  });

  const totalDays = Object.values(totals).reduce((sum, value) => sum + value, 0) || 1;
  let cursor = 0;
  const gradientStops = [];
  const legendItems = [];

  buckets.forEach((bucket) => {
    const value = totals[bucket.key] || 0;
    const percent = Math.round((value / totalDays) * 100);
    const start = cursor;
    const end = cursor + (value / totalDays) * 100;
    cursor = end;

    gradientStops.push(`${bucket.color} ${start.toFixed(2)}% ${end.toFixed(2)}%`);
    legendItems.push(`<li><span class="legend-swatch" style="background:${bucket.color}"></span><span>${bucket.label}: ${percent}%</span></li>`);
  });

  effortDonutEl.style.background = `conic-gradient(${gradientStops.join(", ")})`;
  const workstreams = new Set((matrixRows || []).map((row) => normalize(row.owner))).size;
  effortDonutEl.innerHTML = `<span>${workstreams || 0} owners</span>`;
  effortLegendEl.innerHTML = legendItems.join("");
}

function renderMilestoneBurnupChart(timeline) {
  if (!milestoneBurnupChartEl) return;

  if (!timeline.length) {
    milestoneBurnupChartEl.innerHTML = '<p class="chart-empty">Generate a solution pack to view burnup progress.</p>';
    return;
  }

  const phaseRows = timeline.map((item, index) => ({
    label: normalizePhaseLabel(item.phase || item.stage || `Phase ${index + 1}`),
    durationDays: estimatePhaseDurationDays(item, index, timeline),
    status: getPhaseProgressStatus(item, index, timeline),
  }));

  const totalDays = phaseRows.reduce((sum, row) => sum + row.durationDays, 0) || 1;
  let plannedCumulative = 0;
  let actualCumulative = 0;

  const points = phaseRows.map((row) => {
    const phaseWeight = row.durationDays / totalDays;
    plannedCumulative += phaseWeight * 100;
    const completionFactor = row.status === "Complete"
      ? 1
      : row.status === "In Progress"
        ? 0.55
        : 0;
    actualCumulative += phaseWeight * 100 * completionFactor;

    return {
      label: row.label,
      planned: Math.max(0, Math.min(100, plannedCumulative)),
      actual: Math.max(0, Math.min(100, actualCumulative)),
    };
  });

  const width = 460;
  const height = 210;
  const margin = { top: 14, right: 16, bottom: 46, left: 36 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const denominator = Math.max(points.length - 1, 1);
  const xFor = (index) => margin.left + (index / denominator) * plotWidth;
  const yFor = (value) => margin.top + ((100 - value) / 100) * plotHeight;

  const plannedPath = points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${xFor(index).toFixed(2)} ${yFor(point.planned).toFixed(2)}`)
    .join(" ");

  const actualPath = points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${xFor(index).toFixed(2)} ${yFor(point.actual).toFixed(2)}`)
    .join(" ");

  const yTicks = [0, 25, 50, 75, 100];
  const yGrid = yTicks.map((tick) => `
    <line x1="${margin.left}" y1="${yFor(tick).toFixed(2)}" x2="${width - margin.right}" y2="${yFor(tick).toFixed(2)}" />
    <text x="${(margin.left - 6).toFixed(2)}" y="${(yFor(tick) + 3).toFixed(2)}" text-anchor="end">${tick}%</text>
  `).join("");

  const xLabels = points.map((point, index) => `
    <text x="${xFor(index).toFixed(2)}" y="${(height - 16).toFixed(2)}" text-anchor="middle">${point.label}</text>
  `).join("");

  const plannedDots = points.map((point, index) => `
    <circle cx="${xFor(index).toFixed(2)}" cy="${yFor(point.planned).toFixed(2)}" r="3.2" class="burnup-dot planned" />
  `).join("");

  const actualDots = points.map((point, index) => `
    <circle cx="${xFor(index).toFixed(2)}" cy="${yFor(point.actual).toFixed(2)}" r="3.2" class="burnup-dot actual" />
  `).join("");

  const lastPoint = points[points.length - 1];
  const summaryText = `Planned ${Math.round(lastPoint.planned)}% | Current ${Math.round(lastPoint.actual)}%`;

  milestoneBurnupChartEl.innerHTML = `
    <div class="burnup-chart">
      <svg class="burnup-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Burnup chart with planned and current cumulative completion by phase">
        <g class="burnup-grid">${yGrid}</g>
        <g class="burnup-xlabels">${xLabels}</g>
        <path d="${plannedPath}" class="burnup-line planned" />
        <path d="${actualPath}" class="burnup-line actual" />
        <g>${plannedDots}${actualDots}</g>
      </svg>
      <div class="burnup-legend">
        <span><i class="legend-line planned"></i>Planned cumulative</span>
        <span><i class="legend-line actual"></i>Current cumulative</span>
        <span class="burnup-summary">${summaryText}</span>
      </div>
    </div>
  `;
}

function estimatePhaseDurationDays(item, index, timeline) {
  const start = parseFlexibleDate(item.startISO || item.window?.split("-")[0] || "");
  const end = parseFlexibleDate(item.targetISO || item.targetDate || "");
  if (start && end) {
    return Math.max(1, daysBetween(start, end));
  }
  if (index === 0) return 14;
  const prevEnd = parseFlexibleDate(timeline[index - 1]?.targetISO || timeline[index - 1]?.targetDate || "");
  if (prevEnd && end) return Math.max(1, daysBetween(prevEnd, end));
  return 14;
}

function getWorkstreamsData(pack) {
  const fromPack = Array.isArray(pack?.workstreams) ? pack.workstreams : [];
  const sanitized = fromPack
    .map((stream) => ({
      name: normalize(stream?.name),
      focus: normalize(stream?.focus),
      pendingOwner: normalize(stream?.pendingOwner),
    }))
    .filter((stream) => stream.name || stream.focus || stream.pendingOwner);

  if (sanitized.length) return sanitized;

  const intake = pack?.intakeSnapshot || {
    projectTitle: "",
    projectGoal: "",
    problem: "",
    outcomes: [],
    stakeholders: [],
    systems: [],
    risks: [],
    constraints: { time: "", budget: "", staffing: "", compliance: "", tech: "" },
  };
  const context = pack?.context || {};
  return buildWorkstreamsData(intake, context);
}

function renderWorkstreams(pack) {
  if (!workstreamsListEl) return;

  const streams = getWorkstreamsData(pack);
  workstreamsListEl.innerHTML = "";

  if (!streams.length) {
    const empty = document.createElement("li");
    empty.textContent = "No workstreams available. Add more intake details and regenerate.";
    workstreamsListEl.appendChild(empty);
    return;
  }

  streams.forEach((stream, index) => {
    const item = document.createElement("li");
    item.className = "workstream-item";

    const title = document.createElement("p");
    title.className = "workstream-title";
    title.textContent = `${index + 1}. ${stream.name || `Workstream ${index + 1}`}`;

    const focus = document.createElement("p");
    focus.className = "workstream-focus";
    focus.textContent = stream.focus || "Focus to be defined.";

    const owner = document.createElement("p");
    owner.className = "workstream-owner";
    owner.textContent = `Pending owner: ${stream.pendingOwner || "To be assigned"}`;

    item.appendChild(title);
    item.appendChild(focus);
    item.appendChild(owner);
    workstreamsListEl.appendChild(item);
  });
}

function renderWorkstreamsText(pack) {
  const streams = getWorkstreamsData(pack);
  if (!streams.length) return "- No workstreams available.";

  return streams
    .map((stream, index) => {
      const name = stream.name || `Workstream ${index + 1}`;
      const focus = stream.focus || "Focus to be defined.";
      const owner = stream.pendingOwner || "To be assigned";
      return `${index + 1}. ${name}\n   Focus: ${focus}\n   Pending owner: ${owner}`;
    })
    .join("\n");
}

function renderRaidSnapshot(pack) {
  if (!raidSnapshotListEl) return;
  const items = getRaidSnapshotItems(pack);
  raidSnapshotListEl.innerHTML = "";
  if (!items.length) {
    const li = document.createElement("li");
    li.textContent = "No risks/issues listed yet. Add risks in Intake to populate this module.";
    raidSnapshotListEl.appendChild(li);
    return;
  }

  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    raidSnapshotListEl.appendChild(li);
  });
}

function getRaidSnapshotItems(pack) {
  const sectionD = getPackSection(pack, "sectionD");
  const raidLines = sectionD
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => /^\d+\)\s*Risk:/i.test(line))
    .slice(0, 5);

  if (raidLines.length) return raidLines;

  const risks = pack?.intakeSnapshot?.risks || [];
  return risks.slice(0, 5).map((risk, index) => `${index + 1}) Risk: ${risk} | Mitigation: assign owner, due date, and weekly follow-up.`);
}

function renderRaidSnapshotText() {
  if (!raidSnapshotListEl) return "- No RAID snapshot available.";
  const lines = Array.from(raidSnapshotListEl.querySelectorAll("li")).map((li) => `- ${li.textContent}`);
  return lines.length ? lines.join("\n") : "- No RAID snapshot available.";
}

function renderTimelineText(timeline) {
  if (!timeline.length) return "- No milestones available.";
  return timeline
    .map((item, index) => {
      const status = getPhaseProgressStatus(item, index, timeline);
      return `- Phase ${item.phaseNumber || index + 1} [${item.stage || "Execution"} | ${status}]: ${item.milestone || "Milestone"} | ${item.phase || "Phase"} | Window ${item.window || "TBD"} | Target ${item.targetDate || "TBD"}`;
    })
    .join("\n");
}

function renderMatrixText(matrixRows) {
  if (!matrixRows.length) return "- No work items available.";
  return matrixRows
    .map((row, index) => `${index + 1}. ${row.workItem} | Owner: ${row.owner} | Status: ${row.status} | Due: ${row.dueDate} | Dependencies: ${row.dependencies}`)
    .join("\n");
}

function normalizeMilestones(milestones = []) {
  if (!Array.isArray(milestones)) return [];

  return milestones.map((item, index) => {
    if (typeof item === "string") {
      const match = item.match(/^Phase\s+(\d+):\s*(.+?)(?:\s*\((.+)\))?$/i);
      return {
        phaseNumber: match ? Number(match[1]) : index + 1,
        stage: "Execution",
        phase: match ? match[2] : item,
        milestone: match ? match[2] : "Milestone",
        windowLabel: "Window pending",
        targetLabel: match?.[3] || "TBD",
      };
    }

    return {
      phaseNumber: item.phaseNumber || index + 1,
      stage: item.stage || "Execution",
      phase: item.phase || "Phase",
      milestone: item.milestone || item.phase || "Milestone",
      windowLabel: item.windowLabel || item.window || "Window pending",
      targetLabel: item.targetLabel || item.targetDate || "TBD",
      startISO: item.startISO || "",
      targetISO: item.targetISO || "",
    };
  });
}

function formatMilestoneLine(milestone) {
  return `Phase ${milestone.phaseNumber} [${milestone.stage}]: ${milestone.phase} - ${milestone.milestone} (${milestone.windowLabel}; target ${milestone.targetLabel})`;
}

function getStatusClass(status) {
  const normalized = (status || "").toLowerCase().trim();
  if (normalized === "in progress") return "status-in-progress";
  if (normalized === "at risk") return "status-at-risk";
  if (normalized === "complete") return "status-complete";
  return "status-not-started";
}

function getPhaseProgressStatus(item, index, timeline) {
  const now = startOfDay(new Date());
  const target = parseFlexibleDate(item.targetISO || item.targetDate);
  const prevTarget = index > 0
    ? parseFlexibleDate(timeline[index - 1]?.targetISO || timeline[index - 1]?.targetDate)
    : null;

  if (target && now > target) return "Complete";
  if (index === 0) return "In Progress";
  if (prevTarget && now >= prevTarget && (!target || now <= target)) return "In Progress";
  return "Upcoming";
}

function getPhaseStatusClass(status) {
  if (status === "Complete") return "phase-complete";
  if (status === "In Progress") return "phase-progress";
  return "phase-upcoming";
}

function getPhaseStatusIcon(status) {
  if (status === "Complete") return "✓";
  if (status === "In Progress") return "◔";
  return "○";
}

function parseFlexibleDate(value) {
  if (!value) return null;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function getAISectionText(pack) {
  if (pack?.sections?.sectionH) return pack.sections.sectionH;

  const fallbackIntake = pack?.intakeSnapshot || {
    projectTitle: "",
    projectGoal: "",
    ownerName: "",
    ownerArea: "",
    problem: "",
    outcomes: [],
    stakeholders: [],
    systems: [],
    risks: [],
    constraints: { time: "", budget: "", staffing: "", compliance: "", tech: "" },
    goLiveDate: "",
    urgency: pack?.context?.urgencyProfile || "Moderate",
  };
  const fallbackContext = pack?.context || {
    scenarioType: "Cross-functional",
    urgencyProfile: "Moderate",
    topConstraints: [],
    governanceMode: "Balanced governance",
  };
  const fallbackFollowups = buildFollowUpQuestions(fallbackIntake);
  const timeline = getTimelineData(pack || {});
  const matrix = getMatrixData(pack || {}, timeline);
  return buildAIInsights(fallbackIntake, fallbackContext, fallbackFollowups, timeline, matrix, pack?.aiMeta || { provider: "Local assistant" });
}

function getPackSection(pack, key) {
  return normalize(pack?.sections?.[key]);
}

function renderSolutionPdfHeader(pack) {
  if (!solutionPdfProjectNameEl || !solutionPdfProjectSummaryEl) return;

  const projectName = normalize(pack?.intakeSnapshot?.projectTitle) || normalize(pack?.title) || "Untitled Project";
  const projectGoal = normalize(pack?.intakeSnapshot?.projectGoal) || "deliver defined outcomes and measurable value";
  const scenario = (normalize(pack?.context?.scenarioType) || "cross-functional").toLowerCase();
  const owner = normalize(pack?.context?.ownerLabel) || "the sponsor and delivery team";

  solutionPdfProjectNameEl.textContent = projectName;
  solutionPdfProjectSummaryEl.textContent = `${projectName} is a ${scenario} initiative focused on ${projectGoal}. This PDF contains the complete Project Management Solution Pack prepared for kickoff and execution planning. It includes the executive summary, delivery timeline, governance model, role ownership, planning deliverables, and sequenced next steps for ${owner}.`;
}

function renderSolutionPack(pack) {
  const scenario = pack.context?.scenarioType || "Cross-functional";
  const aiLabel = pack.aiMeta?.provider ? ` | AI: ${pack.aiMeta.provider}` : " | AI: Local assistant";
  solutionMetaEl.textContent = `${pack.title} | ${pack.complexity} complexity | ${scenario} | Generated ${formatDateTime(pack.createdAt)}${aiLabel}`;
  renderSolutionPdfHeader(pack);
  const timeline = getTimelineData(pack);
  const matrix = getMatrixData(pack, timeline);

  renderTimeline(timeline);
  renderMatrix(matrix);
  renderProjectStatus(pack, timeline, matrix);
  renderOutcomeCharts(pack, timeline, matrix);
  renderWorkstreams(pack);
  renderRaidSnapshot(pack);
  setSolutionPanelsCollapsed(false);

  Object.entries(pack.sections || {}).forEach(([key, text]) => {
    if (sectionEls[key]) {
      sectionEls[key].textContent = text;
    }
  });
  if (sectionEls.sectionH) {
    sectionEls.sectionH.textContent = getAISectionText(pack);
  }
  buildSolutionPdfToc();
  renderExportPreview(pack);
}

function renderExportPreview(pack) {
  if (!exportPreviewEl || !exportPreviewMetaEl) return;

  if (!pack) {
    exportPreviewMetaEl.textContent = "";
    exportPreviewEl.innerHTML = '<div class="export-empty">Generate a solution pack from Intake to preview the full plan here.</div>';
    return;
  }

  exportPreviewMetaEl.textContent = `${pack.title} | ${pack.complexity} complexity | Generated ${formatDateTime(pack.createdAt)}`;
  exportPreviewEl.innerHTML = "";

  const timeline = getTimelineData(pack);
  const matrix = getMatrixData(pack, timeline);
  const aiSection = getAISectionText(pack);

  const addSection = (heading, bodyText) => {
    const block = document.createElement("section");
    block.className = "export-block";

    const h = document.createElement("h3");
    h.className = "export-heading";
    h.textContent = heading;

    const body = document.createElement("pre");
    body.className = "export-body";
    body.textContent = bodyText || "No content.";

    block.appendChild(h);
    block.appendChild(body);
    exportPreviewEl.appendChild(block);
  };

  const timelineBlock = document.createElement("section");
  timelineBlock.className = "export-block";
  const timelineHeading = document.createElement("h3");
  timelineHeading.className = "export-heading";
  timelineHeading.textContent = "Delivery timeline";
  timelineBlock.appendChild(timelineHeading);
  const timelineList = document.createElement("ol");
  timelineList.className = "export-timeline";
  timeline.forEach((item, index) => {
    const li = document.createElement("li");
    li.textContent = `Phase ${item.phaseNumber || index + 1} [${item.stage || "Execution"}]: ${item.milestone || "Milestone"} | ${item.window || "Window TBD"} | Target ${item.targetDate || "TBD"}`;
    timelineList.appendChild(li);
  });
  if (!timeline.length) {
    const li = document.createElement("li");
    li.textContent = "No milestones available.";
    timelineList.appendChild(li);
  }
  timelineBlock.appendChild(timelineList);
  exportPreviewEl.appendChild(timelineBlock);

  const matrixBlock = document.createElement("section");
  matrixBlock.className = "export-block";
  const matrixHeading = document.createElement("h3");
  matrixHeading.className = "export-heading";
  matrixHeading.textContent = "Execution tracker matrix";
  matrixBlock.appendChild(matrixHeading);

  const matrixWrap = document.createElement("div");
  matrixWrap.className = "export-matrix-wrap";
  const table = document.createElement("table");
  table.className = "export-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th scope="col">Work item</th>
        <th scope="col">Owner</th>
        <th scope="col">Status</th>
        <th scope="col">Due date</th>
        <th scope="col">Dependencies</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  const tbody = table.querySelector("tbody");
  matrix.forEach((row) => {
    const tr = document.createElement("tr");
    const cells = [row.workItem, row.owner, row.status, row.dueDate, row.dependencies];
    cells.forEach((cellText) => {
      const td = document.createElement("td");
      td.textContent = cellText || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  if (!matrix.length && tbody) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 5;
    td.textContent = "No work items available.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
  matrixWrap.appendChild(table);
  matrixBlock.appendChild(matrixWrap);
  exportPreviewEl.appendChild(matrixBlock);

  addSection("Executive summary", getPackSection(pack, "sectionA"));
  addSection("Governance + decision model", getPackSection(pack, "sectionB"));
  addSection("Process model + mapping", getPackSection(pack, "sectionC"));
  addSection("Planning deliverables", getPackSection(pack, "sectionD"));
  addSection("Execution + accountability", getPackSection(pack, "sectionE"));
  addSection("Expert collaboration recommendations", getPackSection(pack, "sectionF"));
  addSection("Next steps", getPackSection(pack, "sectionG"));
  addSection("AI insights and recommendations", aiSection);
}

function renderFullPackText(pack) {
  const timeline = getTimelineData(pack);
  const matrix = getMatrixData(pack, timeline);
  const aiSection = getAISectionText(pack);
  const workstreams = renderWorkstreamsText(pack);

  return [
    "Project Management Solution Pack",
    `${pack.title} | ${pack.complexity} complexity | Generated ${formatDateTime(pack.createdAt)}`,
    `AI enhancement mode: ${pack.aiMeta?.provider || "Local assistant"}`,
    "",
    "Delivery timeline",
    renderTimelineText(timeline),
    "",
    "Execution tracker matrix",
    renderMatrixText(matrix),
    "",
    "Executive summary",
    getPackSection(pack, "sectionA"),
    "",
    "Governance + decision model",
    getPackSection(pack, "sectionB"),
    "",
    "Process model + mapping",
    getPackSection(pack, "sectionC"),
    "",
    "Planning deliverables",
    getPackSection(pack, "sectionD"),
    "",
    "Workstreams",
    workstreams,
    "",
    "Execution + accountability",
    getPackSection(pack, "sectionE"),
    "",
    "Expert collaboration recommendations",
    getPackSection(pack, "sectionF"),
    "",
    "Next steps",
    getPackSection(pack, "sectionG"),
    "",
    "AI insights and recommendations",
    aiSection,
  ].join("\n");
}

async function copyText(text) {
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "absolute";
    area.style.left = "-9999px";
    document.body.appendChild(area);
    area.select();
    document.execCommand("copy");
    document.body.removeChild(area);
  }
}

function savePackVersion(pack) {
  state.versions.unshift(pack);
  state.versions = state.versions.slice(0, 25);
  startRetentionWindow(true);
  persistVersions();
}

function loadVersions() {
  try {
    const expiresAt = loadRetentionExpiry();
    if (expiresAt && Date.now() >= expiresAt) {
      sessionStorage.removeItem(STORAGE_KEY);
      sessionStorage.removeItem(RETENTION_EXPIRY_KEY);
      return [];
    }

    const raw = sessionStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistVersions() {
  try {
    if (state.versions.length === 0) {
      sessionStorage.removeItem(STORAGE_KEY);
    } else {
      sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state.versions));
    }
    persistRetentionExpiry();
  } catch {
    generationStatusEl.textContent = "Could not save history in this browser session.";
  }
}

function loadRetentionExpiry() {
  try {
    const raw = sessionStorage.getItem(RETENTION_EXPIRY_KEY);
    const expiresAt = Number(raw);
    return Number.isFinite(expiresAt) && expiresAt > 0 ? expiresAt : null;
  } catch {
    return null;
  }
}

function persistRetentionExpiry() {
  try {
    if (state.retentionExpiresAt) {
      sessionStorage.setItem(RETENTION_EXPIRY_KEY, String(state.retentionExpiresAt));
    } else {
      sessionStorage.removeItem(RETENTION_EXPIRY_KEY);
    }
  } catch {
    // Ignore storage failures; in-memory session still works.
  }
}

function initializeRetentionLifecycle() {
  if (state.retentionExpiresAt && Date.now() >= state.retentionExpiresAt) {
    clearSessionData("Session data expired and has been deleted.");
    return;
  }
  if (!hasAnySessionData()) {
    state.retentionExpiresAt = null;
    state.retentionWarningShown = false;
    persistRetentionExpiry();
    refreshRetentionTimers();
    return;
  }
  if (!state.retentionExpiresAt && hasAnySessionData()) {
    startRetentionWindow(false);
    return;
  }
  refreshRetentionTimers();
}

function hasAnyIntakeData() {
  if (!intakeForm) return false;
  const intake = getIntakeData();
  const hasConstraint = Object.values(intake.constraints || {}).some((value) => Boolean(normalize(value)));

  return Boolean(
    intake.projectTitle ||
    intake.projectGoal ||
    intake.ownerName ||
    intake.ownerArea ||
    intake.problem ||
    intake.outcomes.length ||
    hasConstraint ||
    intake.goLiveDate ||
    intake.stakeholders.length ||
    intake.systems.length ||
    intake.risks.length ||
    (intake.urgency && intake.urgency !== "Moderate")
  );
}

function hasAnySessionData() {
  return hasAnyIntakeData() || Boolean(state.currentPack) || state.versions.length > 0;
}

function ensureRetentionWindowActive() {
  if (!hasAnySessionData()) {
    state.retentionExpiresAt = null;
    state.retentionWarningShown = false;
    persistRetentionExpiry();
    refreshRetentionTimers();
    return;
  }
  if (!state.retentionExpiresAt) {
    startRetentionWindow(false);
    return;
  }
  refreshRetentionTimers();
}

function startRetentionWindow(reset = false) {
  if (!hasAnySessionData()) return;
  if (!state.retentionExpiresAt || reset) {
    state.retentionExpiresAt = Date.now() + RETENTION_MS;
    state.retentionWarningShown = false;
    persistRetentionExpiry();
  }
  refreshRetentionTimers();
}

function clearRetentionTimers() {
  if (state.retentionWarningTimer) {
    window.clearTimeout(state.retentionWarningTimer);
    state.retentionWarningTimer = null;
  }
  if (state.retentionCleanupTimer) {
    window.clearTimeout(state.retentionCleanupTimer);
    state.retentionCleanupTimer = null;
  }
}

function refreshRetentionTimers() {
  clearRetentionTimers();

  if (!state.retentionExpiresAt || !hasAnySessionData()) {
    if (retentionStatusEl) retentionStatusEl.textContent = "";
    return;
  }

  const now = Date.now();
  if (now >= state.retentionExpiresAt) {
    clearSessionData("Session data expired and has been deleted.");
    return;
  }

  if (retentionStatusEl) {
    retentionStatusEl.textContent = `Session data auto-deletes at ${formatDateTime(state.retentionExpiresAt)}. Export before expiration.`;
  }

  const warningDelay = Math.max(0, state.retentionExpiresAt - RETENTION_WARNING_MS - now);
  const cleanupDelay = Math.max(0, state.retentionExpiresAt - now);

  state.retentionWarningTimer = window.setTimeout(() => {
    if (!hasAnySessionData()) return;
    if (state.retentionWarningShown) return;
    if (document.visibilityState === "visible" && retentionStatusEl) {
      state.retentionWarningShown = true;
      const warning = "Privacy reminder: Session data will be deleted in about 1 minute. Export now to keep a copy.";
      retentionStatusEl.textContent = warning;
      if (generationStatusEl) generationStatusEl.textContent = warning;
      window.alert(warning);
    }
  }, warningDelay);

  state.retentionCleanupTimer = window.setTimeout(() => {
    clearSessionData("Session data expired and has been deleted.");
  }, cleanupDelay);
}

function clearSessionData(statusMessage = "Session data cleared.") {
  state.generationToken += 1;
  clearRetentionTimers();
  stopGenerationLoader();
  state.isGenerating = false;
  state.retentionExpiresAt = null;
  state.retentionWarningShown = false;
  persistRetentionExpiry();
  try {
    sessionStorage.removeItem(SAMPLE_ROTATION_KEY);
  } catch {
    // Ignore storage failures.
  }

  state.currentPack = null;
  state.versions = [];
  persistVersions();

  if (intakeForm) intakeForm.reset();
  resetIntakeValidationState();
  const defaultComplexity = document.querySelector('input[name="complexity"][value="Standard"]');
  if (defaultComplexity) defaultComplexity.checked = true;

  followupListEl.innerHTML = "";
  followupQuestionsEl.classList.add("hidden");
  if (aiAssistStatusEl) aiAssistStatusEl.textContent = "";

  resetSolutionPackDisplays();
  renderExportPreview(null);
  renderVersionHistory();

  if (generationStatusEl) generationStatusEl.textContent = statusMessage;
  if (retentionStatusEl) retentionStatusEl.textContent = "Session cache cleared for privacy.";
}

function resetSolutionPackDisplays() {
  solutionMetaEl.textContent = "";
  if (solutionPdfProjectNameEl) solutionPdfProjectNameEl.textContent = "Project Name";
  if (solutionPdfProjectSummaryEl) {
    solutionPdfProjectSummaryEl.textContent = "Generate a solution pack to prepare a distribution-ready PDF summary.";
  }
  if (solutionPdfTocListEl) {
    solutionPdfTocListEl.innerHTML = `
      <li class="toc-item">
        <span class="toc-title">Generate a solution pack to build the table of contents.</span>
        <span class="toc-leader" aria-hidden="true"></span>
        <span class="toc-page">1</span>
      </li>
    `;
  }
  timelineVisualEl.innerHTML = '<li class="timeline-node timeline-empty">Generate a solution pack from Intake to view the timeline.</li>';
  if (timelineMetricsEl) timelineMetricsEl.innerHTML = "";
  matrixBodyEl.innerHTML = "<tr><td colspan=\"5\">Generate a solution pack from Intake to view the tracker matrix.</td></tr>";
  if (phaseDurationChartEl) phaseDurationChartEl.innerHTML = '<p class="chart-empty">Generate a solution pack to view phase durations.</p>';
  if (effortDonutEl) {
    effortDonutEl.style.background = "var(--neutral-300)";
    effortDonutEl.innerHTML = "<span>0 owners</span>";
  }
  if (effortLegendEl) effortLegendEl.innerHTML = "";
  if (milestoneBurnupChartEl) milestoneBurnupChartEl.innerHTML = '<p class="chart-empty">Generate a solution pack to view burnup progress.</p>';
  if (workstreamsListEl) workstreamsListEl.innerHTML = "<li>Generate a solution pack from Intake to view workstreams.</li>";
  if (raidSnapshotListEl) raidSnapshotListEl.innerHTML = "<li>Generate a solution pack from Intake to view the RAID snapshot.</li>";

  if (statusPriorityBadgeEl) statusPriorityBadgeEl.textContent = "Priority: --";
  if (statusPriorityBadgeEl) statusPriorityBadgeEl.className = "status-kpi-badge";
  if (statusComplexityBadgeEl) statusComplexityBadgeEl.textContent = "Complexity: --";
  if (statusComplexityBadgeEl) statusComplexityBadgeEl.className = "status-kpi-badge";
  if (statusPhaseLabelEl) statusPhaseLabelEl.textContent = "Current phase: --";
  if (statusPercentLabelEl) statusPercentLabelEl.textContent = "0%";
  if (statusProgressFillEl) {
    statusProgressFillEl.style.width = "0%";
    statusProgressFillEl.parentElement?.setAttribute("aria-valuenow", "0");
  }
  if (phaseStepperEl) {
    phaseStepperEl.innerHTML = STATUS_PHASES.map((phase) => `<li class="step-box">${phase}</li>`).join("");
  }
  if (miniMetricsEl) miniMetricsEl.innerHTML = "";
  setSolutionPanelsCollapsed(false);

  Object.values(sectionEls).forEach((el) => {
    if (el) {
      el.textContent = "Generate a solution pack from Intake to view this section.";
    }
  });
}

function renderVersionHistory() {
  versionListEl.innerHTML = "";

  if (state.versions.length === 0) {
    const empty = document.createElement("li");
    empty.className = "version-item";
    empty.textContent = "No saved versions yet. Generate your first solution pack from Intake.";
    versionListEl.appendChild(empty);
    return;
  }

  state.versions.forEach((version) => {
    const li = document.createElement("li");
    li.className = "version-item";

    const title = document.createElement("strong");
    title.textContent = version.title || "Untitled Project";

    const meta = document.createElement("div");
    meta.className = "version-meta";
    meta.textContent = `${formatDateTime(version.createdAt)} | ${version.complexity} | ${version.context?.scenarioType || "General"} | AI: ${version.aiMeta?.provider || "Local assistant"}`;

    const actions = document.createElement("div");
    actions.className = "version-actions";

    const loadBtn = document.createElement("button");
    loadBtn.className = "secondary-btn";
    loadBtn.type = "button";
    loadBtn.textContent = "Load";
    loadBtn.addEventListener("click", () => {
      state.currentPack = version;
      renderSolutionPack(version);
      setActiveTab("solution");
      ensureRetentionWindowActive();
      generationStatusEl.textContent = "Selected version loaded.";
    });

    const copyBtn = document.createElement("button");
    copyBtn.className = "secondary-btn";
    copyBtn.type = "button";
    copyBtn.textContent = "Copy";
    copyBtn.addEventListener("click", async () => {
      await copyText(renderFullPackText(version));
      generationStatusEl.textContent = "Saved version copied.";
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "secondary-btn";
    deleteBtn.type = "button";
    deleteBtn.textContent = "Delete";
    deleteBtn.addEventListener("click", () => {
      state.versions = state.versions.filter((item) => item.id !== version.id);
      persistVersions();
      renderVersionHistory();
      ensureRetentionWindowActive();
      generationStatusEl.textContent = "Version removed.";
    });

    actions.appendChild(loadBtn);
    actions.appendChild(copyBtn);
    actions.appendChild(deleteBtn);

    li.appendChild(title);
    li.appendChild(meta);
    li.appendChild(actions);
    versionListEl.appendChild(li);
  });
}

function parseDurationFromConstraint(timelineText) {
  const text = normalize(timelineText).toLowerCase();
  if (!text) return null;

  const weekRangeMatch = text.match(/(\d+)\s*-\s*(\d+)\s*weeks?/);
  if (weekRangeMatch) return Number(weekRangeMatch[2]) * 7;

  const monthRangeMatch = text.match(/(\d+)\s*-\s*(\d+)\s*months?/);
  if (monthRangeMatch) return Number(monthRangeMatch[2]) * 30;

  const weekMatch = text.match(/(\d+)\s*weeks?/);
  if (weekMatch) return Number(weekMatch[1]) * 7;

  const dayMatch = text.match(/(\d+)\s*days?/);
  if (dayMatch) return Number(dayMatch[1]);

  const monthMatch = text.match(/(\d+)\s*months?/);
  if (monthMatch) return Number(monthMatch[1]) * 30;

  return null;
}

function urgencyToDurationDays(urgency) {
  if (urgency === "Critical") return 56;
  if (urgency === "High") return 84;
  if (urgency === "Low") return 140;
  return 112;
}

function startOfDay(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    const fallback = new Date();
    fallback.setHours(0, 0, 0, 0);
    return fallback;
  }
  date.setHours(0, 0, 0, 0);
  return date;
}

function addDays(baseDate, days) {
  const date = new Date(baseDate);
  date.setDate(date.getDate() + days);
  return date;
}

function daysBetween(startDate, endDate) {
  const start = startOfDay(startDate);
  const end = startOfDay(endDate);
  const diffMs = end.getTime() - start.getTime();
  const dayMs = 24 * 60 * 60 * 1000;
  return Math.max(1, Math.round(diffMs / dayMs));
}

function formatDate(dateValue) {
  const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (Number.isNaN(date.getTime())) return "TBD";
  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function toInputDate(dateValue) {
  const date = dateValue instanceof Date ? new Date(dateValue) : new Date(dateValue);
  if (Number.isNaN(date.getTime())) return "";
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function formatDateTime(dateValue) {
  const date = new Date(dateValue);
  if (Number.isNaN(date.getTime())) return "Unknown time";
  return date.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function capitalize(text) {
  if (!text) return "";
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function ensurePackExists() {
  if (state.currentPack) return true;
  generationStatusEl.textContent = "Generate a solution pack before using copy or export actions.";
  return false;
}
