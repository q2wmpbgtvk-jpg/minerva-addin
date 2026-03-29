'use strict';

Office.onReady(function() {});

// ── Navigation ─────────────────────────────────────────────
function goToScreen(n) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + n).classList.add('active');
  for (let i = 1; i <= 3; i++) {
    const dot = document.getElementById('step-dot-' + i);
    if (!dot) continue;
    dot.classList.remove('active', 'done');
    if (i < n)  dot.classList.add('done');
    if (i === n) dot.classList.add('active');
  }
}

function toggleClient2() {
  const show = document.getElementById('hasClient2').checked;
  document.getElementById('client2Field').style.display = show ? 'block' : 'none';
  if (!show) document.getElementById('client2Name').value = '';
  updateCombinedName();
}

// Auto-suggest combined name when individual names change
function updateCombinedName() {
  const c1 = document.getElementById('client1Name').value.trim();
  const c2 = document.getElementById('client2Name').value.trim();
  const combined = document.getElementById('combinedNames');

  // Only auto-fill if user hasn't manually edited it
  if (combined.dataset.manual === 'true') return;

  if (c1 && c2) {
    // Extract first names and shared last name if applicable
    const parts1 = c1.split(' ');
    const parts2 = c2.split(' ');
    const last1 = parts1[parts1.length - 1];
    const last2 = parts2[parts2.length - 1];

    if (last1 === last2) {
      combined.value = parts1[0] + ' and ' + c2;
    } else {
      combined.value = c1 + ' and ' + c2;
    }
  } else if (c1) {
    combined.value = c1;
  }
}

// ── Screen 1 → 2 ──────────────────────────────────────────
function goToScreen2() {
  const c1        = document.getElementById('client1Name').value.trim();
  const hasC2     = document.getElementById('hasClient2').checked;
  const c2        = document.getElementById('client2Name').value.trim();
  const combined  = document.getElementById('combinedNames').value.trim();
  const err       = document.getElementById('error-1');

  if (!c1)              { err.textContent = "Please enter Client 1's name."; return; }
  if (hasC2 && !c2)     { err.textContent = "Please enter Client 2's name, or uncheck the box."; return; }
  if (!combined)        { err.textContent = "Please enter the combined name."; return; }
  err.textContent = '';

  buildPreview();
  goToScreen(2);
}

// ── Preview ────────────────────────────────────────────────
function buildPreview() {
  const c1       = document.getElementById('client1Name').value.trim();
  const hasC2    = document.getElementById('hasClient2').checked;
  const c2       = document.getElementById('client2Name').value.trim();
  const combined = document.getElementById('combinedNames').value.trim();
  const today    = formatDate(new Date());

  let html = `
    <div class="preview-section">
      <div class="preview-label">Combined Name (opening paragraph)</div>
      <div class="preview-value">${combined}</div>
    </div>
    <hr class="preview-divider"/>
    <div class="preview-section">
      <div class="preview-label">Client 1 Signature Line</div>
      <div class="preview-value">${c1}</div>
    </div>`;

  if (hasC2) {
    html += `
    <hr class="preview-divider"/>
    <div class="preview-section">
      <div class="preview-label">Client 2 Signature Line</div>
      <div class="preview-value">${c2}</div>
    </div>`;
  }

  html += `
    <hr class="preview-divider"/>
    <div class="preview-section">
      <div class="preview-label">Date</div>
      <div class="preview-value">${today}</div>
    </div>
    <hr class="preview-divider"/>
    <div class="preview-section">
      <div class="preview-label">Document</div>
      <div class="preview-value">Investment Advisory Agreement with standard fee schedule, Schedules A & B, and signature blocks.</div>
    </div>`;

  document.getElementById('preview-box').innerHTML = html;
}

// ── Generate IAA ───────────────────────────────────────────
async function generateIAA() {
  const btn = document.getElementById('generateBtn');
  btn.textContent = 'Generating…';
  btn.disabled = true;
  document.getElementById('error-2').textContent = '';

  try {
    const c1       = document.getElementById('client1Name').value.trim();
    const hasC2    = document.getElementById('hasClient2').checked;
    const c2       = document.getElementById('client2Name').value.trim();
    const combined = document.getElementById('combinedNames').value.trim();

    // Step 1: Check JSZip
    if (typeof JSZip === 'undefined') throw new Error('JSZip library not loaded.');

    // Step 2: Load template
    const resp = await fetch('template.docx');
    if (!resp.ok) throw new Error('Could not load IAA template.');
    const templateBuffer = await resp.arrayBuffer();

    // Step 3: Open with JSZip
    const zip = await JSZip.loadAsync(templateBuffer);
    if (!zip.file('word/document.xml')) throw new Error('Template missing document.xml');

    // Step 4: Build replacements
    const textReplacements = [
      ['{{CLIENT_NAMES}}', x(combined)],
      ['{{CLIENT1_NAME}}', x(c1)],
    ];

    // Handle second client: replace placeholder or remove it
    if (hasC2 && c2) {
      textReplacements.push(['{{CLIENT2_NAME}}', x(c2)]);
    } else {
      textReplacements.push(['{{CLIENT2_NAME}}', '']);
    }

    // Step 5: Apply to document.xml
    await replaceInEntry(zip, 'word/document.xml', textReplacements);

    // Step 6: Generate modified docx
    const modifiedBase64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' });

    // Step 7: Open as new document
    if (typeof Word === 'undefined') throw new Error('Word API not available.');
    await Word.run(async ctx => {
      if (!Office.context.requirements.isSetSupported('WordApiHiddenDocument', '1.3')) {
        throw new Error('WordApiHiddenDocument 1.3 not supported. Please update Office.');
      }
      const newDoc = ctx.application.createDocument(modifiedBase64);
      await ctx.sync();
      newDoc.open();
      await ctx.sync();
    });

    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById('screen-success').classList.add('active');

  } catch(err) {
    document.getElementById('error-2').textContent = 'Error: ' + err.message;
    btn.textContent = 'Generate IAA';
    btn.disabled = false;
  }
}

// ── Helper: replace text in a zip entry ────────────────────
async function replaceInEntry(zip, filename, replacements) {
  const file = zip.file(filename);
  if (!file) return;
  let text = await file.async('string');
  for (const [find, replace] of replacements) {
    text = text.split(find).join(replace);
  }
  zip.file(filename, text);
}

// ── Utilities ──────────────────────────────────────────────
function x(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function formatDate(d) {
  return d.toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });
}

function resetForm() {
  ['client1Name','client2Name','combinedNames'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('hasClient2').checked = false;
  document.getElementById('client2Field').style.display = 'none';
  document.getElementById('combinedNames').dataset.manual = 'false';
  goToScreen(1);
}

// ── Event listeners for auto-combined-name ─────────────────
document.addEventListener('DOMContentLoaded', function() {
  document.getElementById('client1Name').addEventListener('input', updateCombinedName);
  document.getElementById('client2Name').addEventListener('input', updateCombinedName);

  // Track if user manually edits the combined name field
  document.getElementById('combinedNames').addEventListener('input', function() {
    this.dataset.manual = 'true';
  });
  document.getElementById('combinedNames').addEventListener('focus', function() {
    // Allow auto-fill to resume if field is cleared
    if (!this.value.trim()) this.dataset.manual = 'false';
  });
});
