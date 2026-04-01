/* ── IPS Task Pane ─────────────────────────────────────────────────────────
   Screens: 1) Search  2) Preview/edit  3) Signers  4) Done
   Data source: MinervaBridge at localhost:3001
   ───────────────────────────────────────────────────────────────────────── */

'use strict';

const MGP_CPI = 0.0395; // From risk/return spreadsheet (3/28/96 to 3/27/26)

const PORTFOLIOS = {
  'Equity Tilted Balanced': { grossReturn: 0.0808, maxDrawdown: -0.458, recoveryDays: 1139 },
  'Balanced':               { grossReturn: 0.0765, maxDrawdown: -0.390, recoveryDays: 896  },
  'Conservative':           { grossReturn: 0.0638, maxDrawdown: -0.254, recoveryDays: 485  },
};

function pct(val) {
  return (val * 100).toFixed(1) + '%';
}

function getPortfolioStats(allocation, estimatedFeeBps) {
  const p = PORTFOLIOS[allocation];
  if (!p) return null;
  const feePct = (parseFloat(estimatedFeeBps) || 0) / 10000;
  const allocationReturn = p.grossReturn - feePct - MGP_CPI;
  return {
    allocationReturn: pct(allocationReturn),
    maxDrawdown:      pct(p.maxDrawdown),
    recoveryDays:     p.recoveryDays.toString(),
  };
}

// ── Screen management ───────────────────────────────────────────────────────
let clientData = {};

function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.style.display = 'none');
  document.getElementById(id).style.display = 'block';
}

function setStatus(id, msg, isError = false) {
  const el = document.getElementById(id);
  if (el) {
    el.textContent = msg;
    el.style.color = isError ? '#c00' : '#555';
  }
}

// ── Screen 1: Search ────────────────────────────────────────────────────────
async function searchClient() {
  const query = document.getElementById('searchInput').value.trim();
  if (!query) return;

  setStatus('status', 'Loading staged client...');
  try {
    const resp = await fetch('http://localhost:3001/api/ips-client');
    if (!resp.ok) throw new Error('MinervaBridge not reachable');
    const data = await resp.json();

    if (data.error) {
      setStatus('status', data.error, true);
      return;
    }

    if (!data.name.toLowerCase().includes(query.toLowerCase())) {
      setStatus('status', `Staged client is "${data.name}" — does not match "${query}". Stage the correct client in MinervaBridge.`, true);
      return;
    }

    loadClient(data);
  } catch (e) {
    setStatus('status', 'Could not reach MinervaBridge. Is it running?', true);
  }
}

function loadClient(data) {
  clientData = data;
  const stats = getPortfolioStats(data.allocation, data.estimated_fee);

  document.getElementById('prevName').textContent       = data.name || '';
  document.getElementById('prevAllocation').textContent = data.allocation || '';
  document.getElementById('prevBackground').value       = data.background_info || '';
  document.getElementById('prevCashNeeds').value        = data.liquidity_needs || '';
  document.getElementById('prevTimeFrame').value        = data.time_frame || '';
  document.getElementById('prevUnique').value           = data.unique_considerations || '';

  if (stats) {
    document.getElementById('prevReturn').textContent   = stats.allocationReturn;
    document.getElementById('prevDrawdown').textContent = stats.maxDrawdown;
    document.getElementById('prevRecovery').textContent = stats.recoveryDays;
  } else {
    document.getElementById('prevReturn').textContent   = `Unknown allocation: "${data.allocation}"`;
    document.getElementById('prevDrawdown').textContent = '-';
    document.getElementById('prevRecovery').textContent = '-';
  }

  setStatus('status', '');
  showScreen('screen2');
}

// ── Screen 2: Preview ───────────────────────────────────────────────────────
function goToSigners() {
  const allocation = clientData.allocation || '';
  if (!PORTFOLIOS[allocation]) {
    setStatus('status2', `Unknown allocation: "${allocation}". Check Wealthbox field.`, true);
    return;
  }
  setStatus('status2', '');
  showScreen('screen3');
}

// ── Screen 3: Signers ───────────────────────────────────────────────────────
async function generateIPS() {
  const signer1 = document.getElementById('signer1').value.trim();
  if (!signer1) {
    setStatus('status3', 'Please enter at least one signer name.', true);
    return;
  }

  const signer2 = document.getElementById('signer2').value.trim();
  const allocation = clientData.allocation || '';
  const stats = getPortfolioStats(allocation, clientData.estimated_fee);

  setStatus('status3', 'Generating IPS...');

  const replacements = {
    '{{BACKGROUND}}':            document.getElementById('prevBackground').value,
    '{{ALLOCATION}}':            allocation,
    '{{ALLOCATION_RETURN}}':     stats.allocationReturn,
    '{{MAX_DRAWDOWN}}':          stats.maxDrawdown,
    '{{RECOVERY_TIME}}':         stats.recoveryDays,
    '{{CASH_NEEDS}}':            document.getElementById('prevCashNeeds').value,
    '{{TIME_FRAME}}':            document.getElementById('prevTimeFrame').value,
    '{{UNIQUE_CONSIDERATIONS}}': document.getElementById('prevUnique').value,
    '{{SIGNER_1}}':              signer1,
    '{{SIGNER_2}}':              signer2,
  };

  try {
    const resp = await fetch(
      'https://q2wmpbgtvk-jpg.github.io/minerva-addin/ips/ips-template.docx'
    );
    if (!resp.ok) throw new Error('Could not load IPS template');
    const buf = await resp.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    async function replaceInEntry(filename) {
      const file = zip.file(filename);
      if (!file) return;
      let text = await file.async('string');
      for (const [find, replace] of Object.entries(replacements)) {
        text = text.split(find).join(x(replace));
      }
      zip.file(filename, text);
    }

    await replaceInEntry('word/document.xml');

    const modifiedBase64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });

    await Word.run(async context => {
      const doc = context.application.createDocument(modifiedBase64);
      context.load(doc);
      await context.sync();
      doc.open();
      await context.sync();
    });

    setStatus('status3', '');
    showScreen('screen4');
  } catch (e) {
    setStatus('status3', 'Error generating IPS: ' + e.message, true);
  }
}

// ── XML escape ──────────────────────────────────────────────────────────────
function x(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ── Init ────────────────────────────────────────────────────────────────────
Office.onReady(() => {
  showScreen('screen1');
  document.getElementById('searchBtn').addEventListener('click', searchClient);
  document.getElementById('searchInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') searchClient();
  });
  document.getElementById('backBtn').addEventListener('click', () => showScreen('screen1'));
  document.getElementById('nextBtn').addEventListener('click', goToSigners);
  document.getElementById('backBtn2').addEventListener('click', () => showScreen('screen2'));
  document.getElementById('generateBtn').addEventListener('click', generateIPS);
});
