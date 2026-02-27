// blocks.js - Standard language library for Minerva Planning Group engagement letters
//
// STRUCTURE:
//   Goals (Section I)      - Always exactly 3. Goal 1 states the primary focus.
//                            Goals 2 and 3 are identical across all letters.
//   Objectives (Section II)- Restate primary focus in detail + client-specific scenarios
//                            + 3 standard closers (insurance, estate, portfolio).
//   Steps (Section III)    - Meeting-by-meeting implementation detail.
//
// {{CLIENTS}} is replaced at runtime with the actual client name(s).

// ── GOALS 2 AND 3 - always the same ───────────────────────────────────────────
const GOAL_PORTFOLIO = 'Formulate a target portfolio for financial investments (e.g. stocks, bonds, mutual funds) to support {{CLIENTS}} in retirement';
const GOAL_ADVISOR   = 'Establish a long-term relationship with a financial advisor with whom {{CLIENTS}} can consult on any personal financial and investment issues that should arise';

// ── STANDARD OBJECTIVE CLOSERS - always last 3 objectives ─────────────────────
const OBJ_INSURANCE  = 'Review insurance coverage to confirm that it is cost effective and adequately covers risk to plan';
const OBJ_ESTATE     = 'Discuss role of standard estate documents and provide referral to estate attorney if needed';
const OBJ_PORTFOLIO  = 'Review portfolio at a high level, identify discrepancies between target portfolio and current investments and identify any red flags';

// ── STANDARD STEPS - appear in every engagement ───────────────────────────────
const STEPS_CORE_M1 = [
  'Obtain data and prepare accurate Base Plan (current financial position)',
  'Clarify goals and issues to be addressed in rank order of priority (Meeting 1)',
];
const STEPS_FINALIZE = [
  'Finalize plan to reflect any changes discussed in first meeting, and run an alternate scenario if necessary (Meeting 2)',
];
const STEPS_CORE_M2 = [
  'Review insurance coverage, identify any potential savings, and discuss any risks that may exist (Meeting 2)',
  'Discuss role of standard estate documents and provide referral to estate attorney if needed (Meeting 2)',
  'Outline general investing strategy and allocation recommended for retirement, compare that allocation to current investments, and identify any changes that should be made (Meeting 2)',
];

// ── PLANNING BLOCKS ────────────────────────────────────────────────────────────

const BLOCKS = {

  retirement: {
    id: 'retirement',
    label: 'Retirement Planning',
    default: true,
    goal1: 'Establish a financial plan with primary goal of funding retirement for {{CLIENTS}}',
    objectives: [
      'Establish a plan with primary goal of funding retirement for {{CLIENTS}}',
      'Planning engagement will include reviewing three potential scenarios: two in the first meeting and a final scenario for the second meeting, to account for any changes that need to be made to the plan after the first meeting',
    ],
    stepsM1: [
      'Estimate cash needs and time horizons for funding retirement (Meeting 1)',
      'Prepare summary of goals, status of current plan relative to goals, and any shortfalls that exist (Meeting 1)',
      'Run one alternate plan scenario based on initial plan results (Meeting 1)',
    ],
    stepsM2: [],
  },

  insurance: {
    id: 'insurance',
    label: 'Insurance Review',
    default: true,
    goal1: null,
    objectives: [],
    stepsM1: [],
    stepsM2: [
      'Run life insurance needs analysis as part of insurance review (Meeting 2)',
    ],
  },

  estate: {
    id: 'estate',
    label: 'Estate Planning Review',
    default: true,
    goal1: null,
    objectives: [],
    stepsM1: [],
    stepsM2: [],
  },

  investment: {
    id: 'investment',
    label: 'Investment Portfolio Review',
    default: true,
    goal1: null,
    objectives: [],
    stepsM1: [],
    stepsM2: [],
  },

  college: {
    id: 'college',
    label: 'College Funding Analysis',
    default: false,
    goal1: 'Establish a financial plan with primary goals of potentially funding college for their children and funding retirement for {{CLIENTS}}',
    objectives: [
      'Establish a plan with primary goals of potentially funding, or at least subsidizing, college expenses for children and funding retirement for {{CLIENTS}}',
      'Planning engagement will include reviewing three potential scenarios: a base scenario ({{CLIENTS}} retire at planned ages with no college funding), an alternate scenario that looks at the impact of college funding, and one additional alternate scenario for the second meeting',
    ],
    stepsM1: [
      'Estimate cash needs and time horizons for funding college and retirement (Meeting 1)',
      'Prepare summary of goals, status of current plan relative to goals, and any shortfalls that exist (Meeting 1)',
      'Analyze impact on plan of college funding alongside retirement and other goals (Meeting 1)',
    ],
    stepsM2: [
      'Recommend college savings approach (e.g. 529 plans) and contribution levels (Meeting 2)',
    ],
  },

  scenarios: {
    id: 'scenarios',
    label: 'Scenario Analysis',
    default: false,
    goal1: null,
    objectives: [
      'Examine impact on plan of [describe scenario - e.g. retiring at 62 vs 65, Social Security timing, etc.]',
    ],
    stepsM1: [
      'Examine impact on plan of [describe scenario] (Meeting 1)',
    ],
    stepsM2: [],
  },

  homePurchase: {
    id: 'homePurchase',
    label: 'Home Purchase Analysis',
    default: false,
    goal1: null,
    objectives: [
      'Examine impact on plan of home purchase, including effect on savings rate and retirement timeline',
    ],
    stepsM1: [
      'Analyze impact of home purchase on overall financial plan and retirement goals (Meeting 1)',
    ],
    stepsM2: [],
  },

  taxDistribution: {
    id: 'taxDistribution',
    label: 'Tax-Efficient Distribution Strategy',
    default: false,
    goal1: null,
    objectives: [
      'Develop tax-efficient distribution strategy to minimize taxes in retirement',
    ],
    stepsM1: [],
    stepsM2: [
      'Review current account structure and develop tax-efficient withdrawal sequencing strategy (Meeting 2)',
    ],
  },
};

// ── ASSEMBLER ──────────────────────────────────────────────────────────────────
// Returns { goals: [], objectives: [], steps: [{label, steps}] }

function assembleContent(selectedIds, clientName) {
  const selected = selectedIds.map(id => BLOCKS[id]).filter(Boolean);
  const sub = s => s.replace(/\{\{CLIENTS\}\}/g, clientName);

  // Determine which key blocks are selected (used in goals, objectives, steps)
  const hasCollege = selected.some(b => b.id === 'college');
  const hasRetirement = selected.some(b => b.id === 'retirement');

  // Goals - always exactly 3
  // Priority for Goal 1: college overrides retirement overrides others
  const priorityOrder = ['college', 'retirement'];
  const goal1Block = priorityOrder.map(id => selected.find(b => b.id === id && b.goal1))
                                   .find(Boolean) || selected.find(b => b.goal1);
  const goal1 = sub(goal1Block ? goal1Block.goal1 : BLOCKS.retirement.goal1);

  const goals = [goal1, sub(GOAL_PORTFOLIO), sub(GOAL_ADVISOR)];

  // Objectives - if college + retirement both selected, college supersedes retirement
  let blockObjs;
  if (hasCollege && hasRetirement) {
    const others = selected.filter(b => b.id !== 'retirement' && b.id !== 'college');
    blockObjs = [
      ...BLOCKS.college.objectives.map(sub),
      ...others.flatMap(b => (b.objectives || []).map(sub)),
    ];
  } else {
    blockObjs = selected.flatMap(b => (b.objectives || []).map(sub));
  }
  const objectives = [...blockObjs, OBJ_INSURANCE, OBJ_ESTATE, OBJ_PORTFOLIO];

  // Steps - if college selected, it supersedes retirement's M1 steps (more specific)

  let blockStepsM1, blockStepsM2;
  if (hasCollege && hasRetirement) {
    // College steps fully cover retirement steps - use college only for M1
    const others = selected.filter(b => b.id !== 'retirement' && b.id !== 'college');
    blockStepsM1 = [
      ...BLOCKS.college.stepsM1.map(sub),
      ...others.flatMap(b => (b.stepsM1 || []).map(sub)),
    ];
    blockStepsM2 = selected.flatMap(b => (b.stepsM2 || []).map(sub));
  } else {
    blockStepsM1 = selected.flatMap(b => (b.stepsM1 || []).map(sub));
    blockStepsM2 = selected.flatMap(b => (b.stepsM2 || []).map(sub));
  }

  const allSteps = [
    ...STEPS_CORE_M1.map(sub),
    ...blockStepsM1,
    ...STEPS_FINALIZE.map(sub),
    ...blockStepsM2,
    ...STEPS_CORE_M2.map(sub),
  ];

  const steps = [{ label: 'Phase 1: Strategic Plan', steps: allSteps }];

  return { goals, objectives, steps };
}
