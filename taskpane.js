'use strict';

Office.onReady(function() {
  checkBridgeStatus();
});

const HOURLY_RATE = 300;

// ── CRM Integration ───────────────────────────────────────
async function pullFromCRM() {
  const btn = document.getElementById('crmPullBtn');
  const status = document.getElementById('crmStatus');
  
  btn.disabled = true;
  btn.textContent = 'Connecting…';
  status.textContent = '';
  status.className = 'crm-status';
  
  try {
    const response = await fetch('https://localhost:3001/api/client');
    if (!response.ok) throw new Error('Server returned ' + response.status);
    const data = await response.json();
    if (data.error) throw new Error(data.error);
    
    if (data.primary_name) {
      document.getElementById('primaryName').value = data.primary_name;
    }
    
    if (data.spouse_name) {
      document.getElementById('hasSpouse').checked = true;
      toggleSpouse();
      document.getElementById('spouseName').value = data.spouse_name;
    } else {
      document.getElementById('hasSpouse').checked = false;
      toggleSpouse();
    }
    
    status.textContent = '✓ Loaded: ' + (data.household_name || data.primary_name);
    status.className = 'crm-status success';
    
  } catch (err) {
    if (err.name === 'TypeError' && err.message.includes('Load failed')) {
      status.textContent = '✗ MinervaBridge not running.';
    } else {
      status.textContent = '✗ ' + err.message;
    }
    status.className = 'crm-status error';
  } finally {
    btn.disabled = false;
    btn.textContent = 'Pull from CRM';
  }
}

async function checkBridgeStatus() {
  const status = document.getElementById('crmStatus');
  if (!status) return;
  try {
    const resp = await fetch('https://localhost:3001/api/status');
    if (resp.ok) {
      status.textContent = 'MinervaBridge connected';
      status.className = 'crm-status success';
    }
  } catch {
    status.textContent = 'MinervaBridge not detected';
    status.className = 'crm-status';
  }
}

// ── Navigation ─────────────────────────────────────────────
function goToScreen(n) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + n).classList.add('active');
  for (let i = 1; i <= 4; i++) {
    const dot = document.getElementById('step-dot-' + i);
    if (!dot) continue;
    dot.classList.remove('active', 'done');
    if (i < n)  dot.classList.add('done');
    if (i === n) dot.classList.add('active');
  }
}

function toggleSpouse() {
  const show = document.getElementById('hasSpouse').checked;
  document.getElementById('spouseField').style.display = show ? 'block' : 'none';
  if (!show) document.getElementById('spouseName').value = '';
}

function updateFeeDisplay() {
  const h = parseFloat(document.getElementById('hours').value) || 0;
  document.getElementById('feeDisplay').textContent = '$' + (h * HOURLY_RATE).toLocaleString('en-US');
}

// ── Screen 1 → 2 ──────────────────────────────────────────
function goToScreen2() {
  const name      = document.getElementById('primaryName').value.trim();
  const hours     = parseFloat(document.getElementById('hours').value);
  const hasSpouse = document.getElementById('hasSpouse').checked;
  const spouse    = document.getElementById('spouseName').value.trim();
  const err = document.getElementById('error-1');
  if (!name)               { err.textContent = "Please enter the client's name."; return; }
  if (!hours || hours <= 0){ err.textContent = 'Please enter a valid number of hours.'; return; }
  if (hasSpouse && !spouse){ err.textContent = 'Please enter the spouse/partner name, or uncheck the box.'; return; }
  err.textContent = '';
  goToScreen(2);
}

// ── Screen 2 → 3 ──────────────────────────────────────────
function goToScreen3() {
  const checkboxMap = {
    retirement: 'blkRetirement', insurance: 'blkInsurance',
    estate: 'blkEstate', investment: 'blkInvestment',
    college: 'blkCollege', scenarios: 'blkScenarios',
    homePurchase: 'blkHome', taxDistribution: 'blkTax'
  };
  const any = Object.values(checkboxMap).some(id => {
    const el = document.getElementById(id);
    return el && el.checked;
  });
  const err = document.getElementById('error-2');
  if (!any) { err.textContent = 'Please select at least one planning block.'; return; }
  err.textContent = '';
  buildEditScreen();
  goToScreen(3);
}

// ── Tab switching ─────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  document.getElementById('panel-' + name).classList.add('active');
}

// ── Build editable Screen 3 ───────────────────────────────
function buildEditScreen() {
  const clients = getClients();
  const checkboxMap = {
    retirement: 'blkRetirement', insurance: 'blkInsurance',
    estate: 'blkEstate', investment: 'blkInvestment',
    college: 'blkCollege', scenarios: 'blkScenarios',
    homePurchase: 'blkHome', taxDistribution: 'blkTax'
  };
  const selectedIds = Object.keys(checkboxMap).filter(blockId => {
    const el = document.getElementById(checkboxMap[blockId]);
    return el && el.checked;
  });

  const { goals, objectives, steps } = assembleContent(selectedIds, clients);

  // ── Goals ─────────────────────────────────────────────
  const goalsList = document.getElementById('goals-list');
  goalsList.innerHTML = '';
  goals.forEach(text => addItemRow(goalsList, text));
  addAddButton(goalsList);

  // ── Objectives ────────────────────────────────────────
  const objList = document.getElementById('objectives-list');
  objList.innerHTML = '';
  objectives.forEach(text => addItemRow(objList, text));
  addAddButton(objList);

  // ── Steps (single Phase 1 group) ──────────────────────
  const stepsList = document.getElementById('steps-list');
  stepsList.innerHTML = '';
  steps.forEach(group => {
    const lbl = document.createElement('div');
    lbl.className = 'item-group-label';
    lbl.textContent = group.label;
    lbl.dataset.blockId = 'phase1';
    stepsList.appendChild(lbl);
    group.steps.forEach(text => addItemRow(stepsList, text));
  });
}

function addItemRow(container, text) {
  const row = document.createElement('div');
  row.className = 'item-row';

  const bullet = document.createElement('span');
  bullet.className = 'item-bullet';
  bullet.textContent = '•';

  const ta = document.createElement('textarea');
  ta.className = 'item-text';
  ta.value = text;
  ta.rows = 1;
  ta.addEventListener('input', autoResize);
  ta.addEventListener('focus', autoResize);

  const del = document.createElement('button');
  del.className = 'item-delete';
  del.textContent = '×';
  del.title = 'Remove item';
  del.onclick = () => row.remove();

  row.appendChild(bullet);
  row.appendChild(ta);
  row.appendChild(del);
  container.appendChild(row);

  // Initial resize
  setTimeout(() => autoResize.call(ta), 0);
  return row;
}

function addAddButton(container) {
  const btn = document.createElement('button');
  btn.className = 'add-item-btn';
  btn.textContent = '+ Add item';
  btn.onclick = () => {
    const row = addItemRow(container, '');
    container.insertBefore(row, btn);
    row.querySelector('textarea').focus();
  };
  container.appendChild(btn);
}

function autoResize() {
  this.style.height = 'auto';
  this.style.height = this.scrollHeight + 'px';
}

// ── Collect edited content ────────────────────────────────
function collectItems(listId) {
  return Array.from(
    document.querySelectorAll('#' + listId + ' .item-text')
  ).map(ta => ta.value.trim()).filter(Boolean);
}

function collectSteps() {
  const result = [];
  const stepsList = document.getElementById('steps-list');
  let current = null;

  stepsList.childNodes.forEach(node => {
    if (node.classList && node.classList.contains('item-group-label')) {
      if (current) result.push(current);
      current = { blockId: node.dataset.blockId, label: node.textContent, steps: [] };
    } else if (node.classList && node.classList.contains('item-row')) {
      const val = node.querySelector('textarea').value.trim();
      if (val && current) current.steps.push(val);
    }
  });
  if (current) result.push(current);
  return result;
}

// ── Screen 3 → 4 ──────────────────────────────────────────
function goToScreen4() {
  buildPreview();
  goToScreen(4);
}

function buildPreview() {
  const clients = getClients();
  const hours   = parseFloat(document.getElementById('hours').value);
  const goals   = collectItems('goals-list');
  const objs    = collectItems('objectives-list');
  const steps   = collectSteps();

  let html = `
    <div class="preview-section"><div class="preview-label">Client</div><div class="preview-value">${clients}</div></div>
    <div class="preview-section"><div class="preview-label">Estimated Fee</div><div class="preview-value">$${(hours*HOURLY_RATE).toLocaleString('en-US')} (${hours} hrs × $300/hr)</div></div>
    <hr class="preview-divider"/>
    <div class="preview-section"><div class="preview-label">Goals (${goals.length})</div>
      ${goals.map(g => `<div class="preview-block"><span class="preview-check">✓</span> ${g}</div>`).join('')}
    </div>
    <hr class="preview-divider"/>
    <div class="preview-section"><div class="preview-label">Objectives (${objs.length})</div>
      ${objs.map(o => `<div class="preview-block"><span class="preview-check">✓</span> ${o}</div>`).join('')}
    </div>
    <hr class="preview-divider"/>
    <div class="preview-section"><div class="preview-label">Steps</div>
      ${steps.map(g => `<div style="margin-top:4px"><strong style="font-size:11px">${g.label}</strong>${g.steps.map(s=>`<div class="preview-block"><span class="preview-check">✓</span> ${s}</div>`).join('')}</div>`).join('')}
    </div>`;
  document.getElementById('preview-box').innerHTML = html;
}

// ── Generate Letter ────────────────────────────────────────
async function generateLetter() {
  const btn = document.getElementById('generateBtn');
  btn.textContent = 'Generating…';
  btn.disabled = true;
  document.getElementById('error-4').textContent = '';

  try {
    const primary   = document.getElementById('primaryName').value.trim();
    const hasSpouse = document.getElementById('hasSpouse').checked;
    const spouse    = document.getElementById('spouseName').value.trim();
    const hours     = parseFloat(document.getElementById('hours').value);
    const fee       = hours * HOURLY_RATE;
    const clients   = getClients();
    const today     = formatDate(new Date());
    const goals     = collectItems('goals-list');
    const objs      = collectItems('objectives-list');
    const steps     = collectSteps();

    if (typeof JSZip === 'undefined') throw new Error('JSZip library not loaded. Check network connection.');

    const resp = await fetch('https://q2wmpbgtvk-jpg.github.io/minerva-addin/template.docx');
    if (!resp.ok) throw new Error('Could not load template. Is the server running?');
    const templateBuffer = await resp.arrayBuffer();

    const zip = await JSZip.loadAsync(templateBuffer);
    if (!zip.file('word/document.xml')) throw new Error('Template zip missing document.xml');

    async function replaceInEntry(filename, replacements) {
      const file = zip.file(filename);
      if (!file) return;
      let text = await file.async('string');
      for (const [find, replace] of replacements) {
        if (typeof find === 'string' && find.startsWith('PARA:')) {
          const marker = find.slice(5);
          const markerPos = text.indexOf(marker);
          if (markerPos !== -1) {
            const pStart = text.lastIndexOf('<w:p ', markerPos);
            const pEnd = text.indexOf('</w:p>', markerPos) + '</w:p>'.length;
            if (pStart !== -1 && pEnd > pStart) {
              text = text.slice(0, pStart) + replace + text.slice(pEnd);
            }
          }
        } else {
          text = text.split(find).join(replace);
        }
      }
      zip.file(filename, text);
    }

    const goalsXml   = buildListXmlRuns(goals, 11);
    const objsXml    = buildListXmlRuns(objs, 20);
    const stepsXml   = buildStepsXmlRuns(steps);

    const textReplacements = [
      ['{{CLIENT_NAME}}', x(clients)],
      ['{{DATE}}',        x(today)],
      ['{{TOTAL_FEE}}',   x(fee.toLocaleString('en-US'))],
      ['{{HOURS}}',       x(String(hours))],
    ];

    const sectionReplacements = [
      ['PARA:{{GOALS}}',            goalsXml],
      ['PARA:{{OBJECTIVES}}',       objsXml],
      ['PARA:{{PLANNING_BLOCKS}}',  stepsXml],
    ];

    // Spouse signature - if no spouse, we'll remove the entire table row after other replacements
    const spouseReplacement = hasSpouse
      ? [['{{SPOUSE_SIGNATURE}}', x(spouse)]]
      : [['{{SPOUSE_SIGNATURE}}', '']];

    const pronounReplacements = !hasSpouse ? [
      ['We have read', 'I have read'],
      ['We will pay', 'I will pay'],
      ['We understand', 'I understand'],
      ['We acknowledge', 'I acknowledge'],
      ['we will receive', 'I will receive'],
      ['we receive', 'I receive'],
    ] : [];

    await replaceInEntry('word/document.xml', [
      ...textReplacements,
      ...sectionReplacements,
      ...spouseReplacement,
      ...pronounReplacements,
    ]);

    // Remove spouse signature row if no spouse
    if (!hasSpouse) {
      const docFile = zip.file('word/document.xml');
      let docText = await docFile.async('string');
      const spousePos = docText.indexOf('{{SPOUSE_SIGNATURE}}');
      if (spousePos !== -1) {
        const trStart = docText.lastIndexOf('<w:tr ', spousePos);
        const trEnd = docText.indexOf('</w:tr>', spousePos) + '</w:tr>'.length;
        if (trStart !== -1 && trEnd > trStart) {
          docText = docText.slice(0, trStart) + docText.slice(trEnd);
          zip.file('word/document.xml', docText);
        }
      }
    }

    for (const hFile of ['word/header1.xml', 'word/header2.xml']) {
      await replaceInEntry(hFile, textReplacements);
    }

    const modifiedBase64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });

    if (typeof Word === 'undefined') throw new Error('Word API not available.');
    await Word.run(async ctx => {
      if (!Office.context.requirements.isSetSupported('WordApiHiddenDocument', '1.3')) {
        throw new Error('WordApiHiddenDocument 1.3 not supported on this version of Word. Please update Office.');
      }
      const newDoc = ctx.application.createDocument(modifiedBase64);
      await ctx.sync();
      newDoc.open();
      await ctx.sync();
    });

    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById('screen-success').classList.add('active');

  } catch(err) {
    document.getElementById('error-4').textContent = 'Error: ' + err.message;
    btn.textContent = 'Generate Letter';
    btn.disabled = false;
  }
}


// ── Raw XML builders (JSZip direct insertion) ─────────────

function buildListXmlRuns(items, numId = 11) {
  return items.map(text =>
    `<w:p><w:pPr>` +
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>` +
    `<w:tabs><w:tab w:val="clear" w:pos="360"/></w:tabs>` +
    `<w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/>` +
    `<w:ind w:left="900" w:hanging="360"/>` +
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>` +
    `</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>` +
    `<w:t xml:space="preserve">${x(text)}</w:t></w:r></w:p>`
  ).join('');
}

function buildStepsXmlRuns(groups) {
  const blockParas = groups.flatMap(g =>
    g.steps.map(text =>
      `<w:p><w:pPr>` +
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>` +
      `<w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/>` +
      `<w:ind w:left="1353" w:hanging="446"/>` +
      `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>` +
      `<w:t xml:space="preserve">${x(text)}</w:t></w:r></w:p>`
    )
  );
  return blockParas.join('');
}

// ── OOXML builders ────────────────────────────────────────
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function listPara(text) {
  return `<w:p ${W}>
    <w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr>
      <w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/>
      <w:ind w:left="900" w:hanging="360"/>
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>
      <w:t>${x(text)}</w:t></w:r></w:p>`;
}

function buildListOoxml(items) {
  return wrap(items.map(listPara).join(''));
}

function stepHeading(text) {
  return `<w:p ${W}>
    <w:pPr><w:spacing w:before="180" w:after="60"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:u w:val="single"/><w:sz w:val="22"/></w:rPr>
      <w:t>${x(text)}</w:t></w:r></w:p>`;
}

function stepPara(text) {
  return `<w:p ${W}>
    <w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      <w:spacing w:before="60" w:after="60" w:line="276" w:lineRule="auto"/>
      <w:ind w:left="720" w:hanging="360"/>
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>
      <w:t>${x(text)}</w:t></w:r></w:p>`;
}

function buildStepsOoxml(groups) {
  const phase = `<w:p ${W}><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="22"/></w:rPr>
      <w:t>Phase 1: Strategic Plan</w:t></w:r></w:p>`;
  const paras = groups.flatMap(g => [
    stepHeading(g.label),
    ...g.steps.map(stepPara),
  ]);
  return wrap(phase + paras.join(''));
}

function wrap(inner) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:wordDocument ${W}><w:body>${inner}</w:body></w:wordDocument>`;
}

// ── Utilities ─────────────────────────────────────────────
function getClients() {
  const primary   = document.getElementById('primaryName').value.trim();
  const hasSpouse = document.getElementById('hasSpouse').checked;
  const spouse    = document.getElementById('spouseName').value.trim();
  return hasSpouse ? primary + ' and ' + spouse : primary;
}

function x(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function arrayBufferToBase64(buf) {
  let b = '';
  new Uint8Array(buf).forEach(v => b += String.fromCharCode(v));
  return btoa(b);
}

function formatDate(d) {
  return d.toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });
}

function resetForm() {
  ['primaryName','spouseName','hours'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('hasSpouse').checked = false;
  document.getElementById('spouseField').style.display = 'none';
  document.getElementById('feeDisplay').textContent = '$0';
  ['blkRetirement','blkInsurance','blkEstate','blkInvestment'].forEach(id => document.getElementById(id).checked = true);
  ['blkCollege','blkHome','blkTax'].forEach(id => document.getElementById(id).checked = false);
  goToScreen(1);
}
