const STORAGE_KEY = "cunySolutionPackHistory";

const tabs = document.querySelectorAll(".tab-btn");
const panels = {
  intake: document.getElementById("tab-intake"),
  solution: document.getElementById("tab-solution"),
  export: document.getElementById("tab-export"),
  history: document.getElementById("tab-history"),
};

const intakeForm = document.getElementById("intakeForm");
const followupQuestionsEl = document.getElementById("followupQuestions");
const followupListEl = document.getElementById("followupList");
const generationStatusEl = document.getElementById("generationStatus");
const solutionMetaEl = document.getElementById("solutionMeta");
const versionListEl = document.getElementById("versionList");
const timelineVisualEl = document.getElementById("timelineVisual");
const matrixBodyEl = document.getElementById("matrixBody");
const exportPreviewMetaEl = document.getElementById("exportPreviewMeta");
const exportPreviewEl = document.getElementById("exportPreview");
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
  aiConfig: {
    mode: "local",
    model: "local-heuristic-v2",
    endpoint: "https://openrouter.ai/api/v1/chat/completions",
    apiKey: "",
  },
};

const SIMPLE_AI_DEFAULTS = {
  localModel: "local-heuristic-v2",
  cloudModel: "openrouter/free",
  cloudEndpoint: "https://openrouter.ai/api/v1/chat/completions",
};

init();

function init() {
  followupQuestionsEl.classList.add("hidden");
  syncAIControls(true);
  attachEventHandlers();
  setActiveTab("intake");
  renderExportPreview(null);
  renderVersionHistory();
}

function attachEventHandlers() {
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => setActiveTab(tab.dataset.tab));
  });

  intakeForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    state.aiConfig = getAIConfig();
    const intake = getIntakeData();
    const complexity = getSelectedComplexity();
    const followups = buildFollowUpQuestions(intake);

    renderFollowUps(followups);
    generationStatusEl.textContent = "Generating solution pack...";
    let pack = buildSolutionPack(intake, complexity, followups, state.aiConfig);
    pack = await enhancePackWithAI(pack, intake, followups, state.aiConfig);
    state.currentPack = pack;
    renderSolutionPack(pack);
    savePackVersion(pack);
    renderVersionHistory();
    setActiveTab("solution");
    generationStatusEl.textContent = pack.aiMeta?.mode !== "local"
      ? `Solution pack generated with ${pack.aiMeta?.provider || "AI"} enhancements and saved to version history.`
      : "Solution pack generated and saved to version history.";
  });

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

  document.getElementById("printPdf").addEventListener("click", () => {
    if (!state.currentPack) {
      generationStatusEl.textContent = "Generate a solution pack before exporting.";
      return;
    }
    setActiveTab("export");
    window.requestAnimationFrame(() => {
      window.print();
    });
  });

  document.getElementById("clearHistory").addEventListener("click", () => {
    state.versions = [];
    persistVersions();
    renderVersionHistory();
    generationStatusEl.textContent = "Version history cleared.";
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
    const isActive = panelKey === key;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });
}

function getSelectedComplexity() {
  const checked = document.querySelector('input[name="complexity"]:checked');
  return checked ? checked.value : "Standard";
}

function getAIConfig() {
  const mode = normalize(aiModeEl?.value) || "local";
  const model = normalize(aiModelEl?.value) || (mode === "cloud" ? SIMPLE_AI_DEFAULTS.cloudModel : SIMPLE_AI_DEFAULTS.localModel);
  const endpoint = normalize(aiEndpointEl?.value) || SIMPLE_AI_DEFAULTS.cloudEndpoint;
  const apiKey = normalize(aiApiKeyEl?.value);
  return { mode, model, endpoint, apiKey };
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
    aiMeta,
    sections,
  };
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
      phase: "Planning and mobilization",
      milestone: "Charter, scope, governance, and success metrics approved",
      stage: "Planning",
      weight: 0.18,
    },
    {
      phase: "Planning to execution handoff",
      milestone: "Requirements baseline, dependency map, and resourcing confirmed",
      stage: "Planning",
      weight: 0.14,
    },
    {
      phase: "Execution wave",
      milestone: "Solution build/configuration complete and quality-checked",
      stage: "Execution",
      weight: 0.38,
    },
    {
      phase: "Monitoring and control",
      milestone: "Risk, status, and quality checkpoints passed for launch readiness",
      stage: "Monitoring",
      weight: 0.2,
    },
    {
      phase: "Completion and stabilization",
      milestone: "Go-live completed with hypercare and handoff to operations",
      stage: "Monitoring",
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
    "- Phase progression: Planning -> Execution -> Monitoring -> Completion with named checkpoints and owners.",
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
    "Intake -> Plan -> Execute -> Monitor -> Validate -> Launch",
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
    "Checkpoint schedule (planning, execution, monitoring):",
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
    "- Monitoring checkpoint review: verify phase exit criteria before moving forward.",
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
      "Do not rewrite sections A-G.",
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
  return normalize(aiConfig?.mode) === "cloud";
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

  if (!timeline.length) {
    const emptyNode = document.createElement("li");
    emptyNode.className = "timeline-node timeline-empty";
    emptyNode.textContent = "No milestones available.";
    timelineVisualEl.appendChild(emptyNode);
    return;
  }

  timeline.forEach((item, index) => {
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

    const date = document.createElement("span");
    date.className = "timeline-date";
    date.textContent = item.targetDate || "TBD";

    node.appendChild(phase);
    node.appendChild(title);
    node.appendChild(detail);
    node.appendChild(date);
    timelineVisualEl.appendChild(node);
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

function renderTimelineText(timeline) {
  if (!timeline.length) return "- No milestones available.";
  return timeline
    .map((item, index) => `- Phase ${item.phaseNumber || index + 1} [${item.stage || "Execution"}]: ${item.milestone || "Milestone"} | ${item.phase || "Phase"} | Window ${item.window || "TBD"} | Target ${item.targetDate || "TBD"}`)
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

function renderSolutionPack(pack) {
  const scenario = pack.context?.scenarioType || "Cross-functional";
  const aiLabel = pack.aiMeta?.provider ? ` | AI: ${pack.aiMeta.provider}` : " | AI: Local assistant";
  solutionMetaEl.textContent = `${pack.title} | ${pack.complexity} complexity | ${scenario} | Generated ${formatDateTime(pack.createdAt)}${aiLabel}`;
  const timeline = getTimelineData(pack);
  const matrix = getMatrixData(pack, timeline);

  renderTimeline(timeline);
  renderMatrix(matrix);

  Object.entries(pack.sections || {}).forEach(([key, text]) => {
    if (sectionEls[key]) {
      sectionEls[key].textContent = text;
    }
  });
  if (sectionEls.sectionH) {
    sectionEls.sectionH.textContent = getAISectionText(pack);
  }
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

  addSection("A) Executive summary", getPackSection(pack, "sectionA"));
  addSection("B) Governance + decision model", getPackSection(pack, "sectionB"));
  addSection("C) Process model + mapping", getPackSection(pack, "sectionC"));
  addSection("D) Planning deliverables", getPackSection(pack, "sectionD"));
  addSection("E) Execution + accountability", getPackSection(pack, "sectionE"));
  addSection("F) Expert collaboration recommendations", getPackSection(pack, "sectionF"));
  addSection("G) Next steps", getPackSection(pack, "sectionG"));
  addSection("H) AI insights and recommendations", aiSection);
}

function renderFullPackText(pack) {
  const timeline = getTimelineData(pack);
  const matrix = getMatrixData(pack, timeline);
  const aiSection = getAISectionText(pack);

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
    "A) Executive summary",
    getPackSection(pack, "sectionA"),
    "",
    "B) Governance + decision model",
    getPackSection(pack, "sectionB"),
    "",
    "C) Process model + mapping",
    getPackSection(pack, "sectionC"),
    "",
    "D) Planning deliverables",
    getPackSection(pack, "sectionD"),
    "",
    "E) Execution + accountability",
    getPackSection(pack, "sectionE"),
    "",
    "F) Expert collaboration recommendations",
    getPackSection(pack, "sectionF"),
    "",
    "G) Next steps",
    getPackSection(pack, "sectionG"),
    "",
    "H) AI insights and recommendations",
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
  persistVersions();
}

function loadVersions() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistVersions() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.versions));
  } catch {
    generationStatusEl.textContent = "Could not save history in this browser session.";
  }
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
