const toggleButtons = document.querySelectorAll('.toggle-btn');
const modeSections = document.querySelectorAll('.mode-section');
const subjectRowsContainer = document.getElementById('subject-rows');
const addSubjectBtn = document.getElementById('add-subject');
const singleForm = document.getElementById('single-form');
const singleErrors = document.getElementById('single-errors');
const singleProcessing = document.getElementById('single-processing');
const singleSteps = document.getElementById('single-steps');
const singleResult = document.getElementById('single-result');
const singleDetails = document.getElementById('single-details');
const singleTableWrapper = document.getElementById('single-table-wrapper');
const singleSummary = document.getElementById('single-summary');
const resetSingleBtn = document.getElementById('reset-single');
const statusMessage = document.getElementById('status-message');
const exportCsvBtn = document.getElementById('export-csv');
const printBtn = document.getElementById('print-btn');
const studentPhotoInput = document.getElementById('student-photo');

const schoolNameField = document.getElementById('school-name');
const schoolAddressField = document.getElementById('school-address');
const schoolSessionField = document.getElementById('school-session');
const schoolLogoInput = document.getElementById('school-logo-input');
const schoolLogoPreview = document.getElementById('school-logo-preview');
const accentColorInput = document.getElementById('accent-color');

const bulkForm = document.getElementById('bulk-form');
const bulkErrors = document.getElementById('bulk-errors');
const setSubjectsBtn = document.getElementById('set-subjects');
const bulkHeader = document.getElementById('bulk-header');
const bulkRows = document.getElementById('bulk-rows');
const addBulkRowBtn = document.getElementById('add-bulk-row');
const bulkProcessing = document.getElementById('bulk-processing');
const bulkSteps = document.getElementById('bulk-steps');
const bulkResult = document.getElementById('bulk-result');
const bulkDetails = document.getElementById('bulk-details');
const bulkTableWrapper = document.getElementById('bulk-table-wrapper');
const bulkSummary = document.getElementById('bulk-summary');

let currentExportData = [];
let currentExportHeaders = [];
let bulkSubjects = [];
let bulkFullMark = 100;
let bulkCredits = [];
let studentPhotoData = '';
let schoolLogoData = schoolLogoPreview.getAttribute('src');
let accentColor = accentColorInput.value;

// Mode switching
ToggleModes();
function ToggleModes() {
  toggleButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      toggleButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      modeSections.forEach(sec => sec.classList.remove('active'));
      modeSections.forEach(sec => sec.hidden = true);
      document.getElementById(btn.dataset.target).classList.add('active');
      document.getElementById(btn.dataset.target).hidden = false;
      statusMessage.textContent = '';
    });
  });
}

const overlineEl = document.querySelector('.overline');
const subtitleEl = document.querySelector('.subtitle');

function getBranding() {
  return {
    name: schoolNameField.value.trim() || 'Shree Janta Secondary School',
    address: schoolAddressField.value.trim() || 'Gauradaha-1, Jhapa',
    session: schoolSessionField.value.trim(),
    logo: schoolLogoData
  };
}

function refreshBranding() {
  const brand = getBranding();
  overlineEl.textContent = brand.name;
  subtitleEl.textContent = `${brand.address} • For +2 Level – School Use`;
  schoolLogoPreview.src = brand.logo;
}

function setAccent(color) {
  if (!color) return;
  accentColor = color;
  const root = document.documentElement.style;
  root.setProperty('--accent', accentColor);
  root.setProperty('--accent-2', lightenColor(accentColor, 20));
}

function lightenColor(hex, percent) {
  const num = parseInt(hex.replace('#', ''), 16);
  const amt = Math.round(2.55 * percent);
  const r = (num >> 16) + amt;
  const g = (num >> 8 & 0x00ff) + amt;
  const b = (num & 0x0000ff) + amt;
  return '#' + (
    0x1000000 +
    (r < 255 ? (r < 0 ? 0 : r) : 255) * 0x10000 +
    (g < 255 ? (g < 0 ? 0 : g) : 255) * 0x100 +
    (b < 255 ? (b < 0 ? 0 : b) : 255)
  ).toString(16).slice(1);
}

function readImageFile(input, callback) {
  const file = input.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => callback(e.target.result);
  reader.readAsDataURL(file);
}

schoolNameField.addEventListener('input', refreshBranding);
schoolAddressField.addEventListener('input', refreshBranding);
schoolSessionField.addEventListener('input', refreshBranding);
schoolLogoInput.addEventListener('change', () => {
  readImageFile(schoolLogoInput, (data) => {
    schoolLogoData = data;
    refreshBranding();
  });
});
accentColorInput.addEventListener('input', () => setAccent(accentColorInput.value));
studentPhotoInput.addEventListener('change', () => {
  readImageFile(studentPhotoInput, (data) => {
    studentPhotoData = data;
  });
});
refreshBranding();
setAccent(accentColor);

// Subject rows setup
function createSubjectRow(subject = '', full = 100, credit = 3, obtained = '') {
// Subject rows setup
function createSubjectRow(subject = '', full = 100, obtained = '') {
  const row = document.createElement('div');
  row.className = 'row';
  row.innerHTML = `
    <input type="text" class="sub-name" placeholder="Physics" value="${subject}">
    <input type="number" class="sub-full" min="1" value="${full}">
    <input type="number" class="sub-credit" min="0.5" step="0.5" value="${credit}">
    <input type="number" class="sub-obtained" min="0" value="${obtained}">
    <button type="button" class="secondary remove-row">Remove</button>
  `;
  row.querySelector('.remove-row').addEventListener('click', () => row.remove());
  return row;
}

function seedSubjectRows() {
  subjectRowsContainer.innerHTML = '';
  const defaults = [
    { name: 'English', credit: 3 },
    { name: 'Nepali', credit: 3 },
    { name: 'Math', credit: 4 },
    { name: 'Science', credit: 4 },
    { name: 'Social', credit: 2 }
  ];
  defaults.forEach(sub => subjectRowsContainer.appendChild(createSubjectRow(sub.name, 100, sub.credit)));
  const defaults = ['English', 'Nepali', 'Math', 'Science', 'Social'];
  defaults.forEach(sub => subjectRowsContainer.appendChild(createSubjectRow(sub)));
}
seedSubjectRows();
addSubjectBtn.addEventListener('click', () => subjectRowsContainer.appendChild(createSubjectRow()));
resetSingleBtn.addEventListener('click', () => {
  singleForm.reset();
  seedSubjectRows();
  singleErrors.textContent = '';
  singleResult.hidden = true;
  statusMessage.textContent = '';
  studentPhotoData = '';
});

// Grade mapping
function gradeFromPercentage(pct) {
  if (pct >= 90) return { grade: 'A+', gp: 4.0, remark: 'Outstanding' };
  if (pct >= 80) return { grade: 'A', gp: 3.6, remark: 'Excellent' };
  if (pct >= 70) return { grade: 'B+', gp: 3.2, remark: 'Very Good' };
  if (pct >= 60) return { grade: 'B', gp: 2.8, remark: 'Good' };
  if (pct >= 50) return { grade: 'C+', gp: 2.4, remark: 'Satisfactory' };
  if (pct >= 40) return { grade: 'C', gp: 2.0, remark: 'Acceptable' };
  if (pct >= 30) return { grade: 'D', gp: 1.6, remark: 'Basic' };
  return { grade: 'E', gp: 0.8, remark: 'Insufficient' };
}

function calculateSingle(data) {
  const { student, subjects } = data;
  let totalFull = 0, totalObtained = 0, totalCredit = 0, totalPoints = 0;
  let totalFull = 0, totalObtained = 0;
  const rows = subjects.map(sub => {
    const pct = (sub.obtained / sub.full) * 100;
    const grade = gradeFromPercentage(pct);
    totalFull += sub.full;
    totalObtained += sub.obtained;
    totalCredit += sub.credit;
    totalPoints += grade.gp * sub.credit;
    return { ...sub, pct: pct.toFixed(2), grade: grade.grade, gp: grade.gp };
  });
  const overallPct = ((totalObtained / totalFull) * 100) || 0;
  const overallGrade = gradeFromPercentage(overallPct);
  const gpa = totalCredit ? (totalPoints / totalCredit) : 0;
  return {
    student,
    rows,
    totals: {
      totalFull,
      totalObtained,
      pct: overallPct.toFixed(2),
      grade: overallGrade.grade,
      remark: overallGrade.remark,
      gpa: gpa.toFixed(2),
      credits: totalCredit
    }
  };
    return { ...sub, pct: pct.toFixed(2), grade: grade.grade };
  });
  const overallPct = ((totalObtained / totalFull) * 100) || 0;
  const overallGrade = gradeFromPercentage(overallPct);
  return { student, rows, totals: { totalFull, totalObtained, pct: overallPct.toFixed(2), grade: overallGrade.grade, remark: overallGrade.remark } };
}

function validateSingle() {
  const name = document.getElementById('student-name').value.trim();
  const cls = document.getElementById('student-class').value.trim();
  const roll = document.getElementById('student-roll').value.trim();
  if (!name || !cls || !roll) return 'Name, Class, and Roll are required.';
  const subjects = [];
  const rows = [...subjectRowsContainer.querySelectorAll('.row')];
  if (rows.length === 0) return 'Please add at least one subject.';
  for (const row of rows) {
    const subName = row.querySelector('.sub-name').value.trim();
    const full = Number(row.querySelector('.sub-full').value);
    const credit = Number(row.querySelector('.sub-credit').value);
    const obtained = Number(row.querySelector('.sub-obtained').value);
    if (!subName) return 'Subject name cannot be empty.';
    if (Number.isNaN(full) || full <= 0) return 'Full marks must be valid numbers.';
    if (Number.isNaN(credit) || credit <= 0) return 'Credit hours must be greater than zero.';
    if (Number.isNaN(obtained) || obtained < 0) return 'Obtained marks must be a positive number.';
    if (obtained > full) return `Obtained marks for ${subName} cannot exceed full marks.`;
    subjects.push({ name: subName, full, obtained, credit });
  }
  return {
    student: {
      name,
      className: cls,
      faculty: document.getElementById('student-faculty').value.trim(),
      section: document.getElementById('student-section').value.trim(),
      roll,
      exam: document.getElementById('student-exam').value.trim(),
    },
    subjects
  };
    const obtained = Number(row.querySelector('.sub-obtained').value);
    if (!subName) return 'Subject name cannot be empty.';
    if (Number.isNaN(full) || full <= 0) return 'Full marks must be valid numbers.';
    if (Number.isNaN(obtained) || obtained < 0) return 'Obtained marks must be a positive number.';
    if (obtained > full) return `Obtained marks for ${subName} cannot exceed full marks.`;
    subjects.push({ name: subName, full, obtained });
  }
  return { student: { name, className: cls, faculty: document.getElementById('student-faculty').value.trim(), section: document.getElementById('student-section').value.trim(), roll }, subjects };
}

function showProcessing(panel, stepsEl, steps) {
  panel.hidden = false;
  stepsEl.innerHTML = '';
  return new Promise(resolve => {
    steps.forEach((text, idx) => {
      setTimeout(() => {
        const li = document.createElement('li');
        li.textContent = text;
        stepsEl.appendChild(li);
        if (idx === steps.length - 1) setTimeout(resolve, 400);
      }, 400 * idx);
    });
  });
}

function renderSingleResult(result) {
  const brand = getBranding();
  singleDetails.innerHTML = `
    <div class="marksheet-header">
      <div class="marksheet-brand">
        <img src="${brand.logo}" alt="School logo">
        <div>
          <h3>${brand.name}</h3>
          <p>${brand.address}</p>
          <p class="muted">${brand.session || 'NEB Academic Session'}</p>
        </div>
      </div>
      ${studentPhotoData ? `<div class="photo-frame"><img src="${studentPhotoData}" alt="Student photo"></div>` : ''}
    </div>
    <div class="student-meta">
      <div><strong>Name:</strong> ${result.student.name}</div>
      <div><strong>Class:</strong> ${result.student.className}</div>
      <div><strong>Faculty:</strong> ${result.student.faculty || '-'} </div>
      <div><strong>Section:</strong> ${result.student.section || '-'} </div>
      <div><strong>Roll No.:</strong> ${result.student.roll}</div>
      <div><strong>Exam:</strong> ${result.student.exam || brand.session || '-'}</div>
    </div>
  `;
  const table = document.createElement('div');
  table.innerHTML = `
    <div class="table-header"><div>Subject</div><div>Full</div><div>Credit</div><div>Obtained</div><div>Grade (GP)</div></div>
  singleDetails.innerHTML = `
    <div><strong>Name:</strong> ${result.student.name}</div>
    <div><strong>Class:</strong> ${result.student.className}</div>
    <div><strong>Faculty:</strong> ${result.student.faculty || '-'} </div>
    <div><strong>Section:</strong> ${result.student.section || '-'} </div>
    <div><strong>Roll No.:</strong> ${result.student.roll}</div>
  `;
  const table = document.createElement('div');
  table.innerHTML = `
    <div class="table-header"><div>Subject</div><div>Full Marks</div><div>Obtained</div><div>Grade</div></div>
  `;
  result.rows.forEach(r => {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = '2fr 1fr 1fr 1fr 1.2fr';
    row.innerHTML = `<div>${r.name}</div><div>${r.full}</div><div>${r.credit}</div><div>${r.obtained}</div><div>${r.grade} (${r.gp.toFixed(2)})</div>`;
    row.style.gridTemplateColumns = '2fr 1fr 1fr 1fr';
    row.innerHTML = `<div>${r.name}</div><div>${r.full}</div><div>${r.obtained}</div><div>${r.grade}</div>`;
    table.appendChild(row);
  });
  singleTableWrapper.innerHTML = '';
  singleTableWrapper.appendChild(table);
  singleSummary.innerHTML = `
    <span>Total: ${result.totals.totalObtained} / ${result.totals.totalFull}</span>
    <span>Percentage: ${result.totals.pct}%</span>
    <span>Grade: ${result.totals.grade}</span>
    <span>Remark: ${result.totals.remark}</span>
    <span>GPA: ${result.totals.gpa} (Credits: ${result.totals.credits})</span>
  `;
  singleResult.hidden = false;
  statusMessage.textContent = 'Result generated successfully.';
  currentExportHeaders = ['Student', 'Class', 'Roll', 'Subject', 'Full Marks', 'Credit Hrs', 'Obtained', 'Grade', 'GP'];
  currentExportData = result.rows.map(r => [result.student.name, result.student.className, result.student.roll, r.name, r.full, r.credit, r.obtained, r.grade, r.gp.toFixed(2)]);
  currentExportData.push(['', '', '', 'TOTAL', result.totals.totalFull, result.totals.credits, result.totals.totalObtained, result.totals.grade, result.totals.gpa]);
  `;
  singleResult.hidden = false;
  statusMessage.textContent = 'Result generated successfully.';
  currentExportHeaders = ['Student', 'Class', 'Roll', 'Subject', 'Full Marks', 'Obtained', 'Grade'];
  currentExportData = result.rows.map(r => [result.student.name, result.student.className, result.student.roll, r.name, r.full, r.obtained, r.grade]);
}

singleForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  singleErrors.textContent = '';
  singleResult.hidden = true;
  if (!studentPhotoInput.files?.length) studentPhotoData = '';
  const data = validateSingle();
  if (typeof data === 'string') {
    singleErrors.textContent = data;
    return;
  }
  disableActions(true);
  await showProcessing(singleProcessing, singleSteps, [
    'Validating input data…',
    'Calculating total and percentage…',
    'Assigning grades as per NEB scale…',
    'Preparing result sheet…'
  ]);
  singleProcessing.hidden = true;
  const result = calculateSingle(data);
  renderSingleResult(result);
  disableActions(false);
});

function disableActions(disabled) {
  document.querySelectorAll('button').forEach(btn => {
    if (btn.id === 'print-btn' || btn.id === 'export-csv') return;
    btn.disabled = disabled;
  });
}

// Bulk mode
function buildBulkHeader() {
  bulkHeader.innerHTML = '';
  if (!bulkSubjects.length) {
    bulkHeader.textContent = 'Set subjects to start adding students';
    bulkHeader.style.padding = '12px 16px';
    return;
  }
  bulkHeader.style.padding = '12px 16px';
  bulkHeader.style.display = 'grid';
  bulkHeader.style.gridTemplateColumns = `1.5fr 0.8fr ${'1fr '.repeat(bulkSubjects.length)} 70px`;
  bulkHeader.innerHTML = `<div>Student Name</div><div>Roll</div>${bulkSubjects.map((s, idx) => `<div>${s}<span class="pill">${bulkCredits[idx] || 1} cr</span></div>`).join('')}<div>Action</div>`;
  bulkHeader.style.gridTemplateColumns = `1.5fr 0.8fr ${'1fr '.repeat(bulkSubjects.length)}`;
  bulkHeader.innerHTML = `<div>Student Name</div><div>Roll</div>${bulkSubjects.map(s => `<div>${s}</div>`).join('')}`;
}

function createBulkRow(name = '', roll = '', marks = []) {
  const row = document.createElement('div');
  row.className = 'row';
  row.style.gridTemplateColumns = `1.5fr 0.8fr ${'1fr '.repeat(bulkSubjects.length)} 70px`;
  row.innerHTML = `
    <input type="text" class="bulk-name" value="${name}">
    <input type="text" class="bulk-roll" value="${roll}">
    ${bulkSubjects.map((_, idx) => `<input type="number" class="bulk-mark" data-index="${idx}" min="0" max="${bulkFullMark}" value="${marks[idx] ?? ''}">`).join('')}
    <button type="button" class="secondary remove-bulk">Remove</button>
  `;
  row.querySelector('.remove-bulk').addEventListener('click', () => row.remove());
  return row;
}

function resetBulkTable() {
  bulkRows.innerHTML = '';
  if (bulkSubjects.length) bulkRows.appendChild(createBulkRow());
}

setSubjectsBtn.addEventListener('click', () => {
  const list = document.getElementById('bulk-subjects').value.split(',').map(s => s.trim()).filter(Boolean);
  bulkFullMark = Number(document.getElementById('bulk-fullmarks').value) || 100;
  const creditsInput = document.getElementById('bulk-credits').value.trim();
  if (!list.length) {
    bulkErrors.textContent = 'Please enter at least one subject.';
    return;
  }
  if (creditsInput) {
    const creditsList = creditsInput.split(',').map(c => Number(c.trim())).filter(v => !Number.isNaN(v));
    if (creditsList.length !== list.length) {
      bulkErrors.textContent = 'Credit hours count must match subject count.';
      return;
    }
    if (creditsList.some(c => c <= 0)) {
      bulkErrors.textContent = 'Credit hours must be greater than zero.';
      return;
    }
    bulkCredits = creditsList;
  } else {
    bulkCredits = Array(list.length).fill(1);
  }
  bulkSubjects = list;
  bulkErrors.textContent = '';
  buildBulkHeader();
  resetBulkTable();
});

addBulkRowBtn.addEventListener('click', () => {
  if (!bulkSubjects.length) {
    bulkErrors.textContent = 'Set subjects first.';
    return;
  }
  bulkRows.appendChild(createBulkRow());
});

function validateBulk() {
  if (!bulkSubjects.length) return 'Please set subjects first.';
  const rows = [...bulkRows.querySelectorAll('.row')];
  if (!rows.length) return 'Add at least one student row.';
  const students = [];
  for (const row of rows) {
    const name = row.querySelector('.bulk-name').value.trim();
    const roll = row.querySelector('.bulk-roll').value.trim();
    if (!name || !roll) return 'Student name and roll are required for all rows.';
    const marks = [...row.querySelectorAll('.bulk-mark')].map((input, idx) => {
      const val = Number(input.value);
      if (Number.isNaN(val) || val < 0) throw new Error(`Invalid marks for ${bulkSubjects[idx]}`);
      if (val > bulkFullMark) throw new Error(`Marks for ${bulkSubjects[idx]} exceed full marks.`);
      return val;
    });
    students.push({ name, roll, marks });
  }
  return students;
}

function calculateBulk(students) {
  return students.map(student => {
    const subjectResults = student.marks.map((m, idx) => {
      const pct = (m / bulkFullMark) * 100;
      const g = gradeFromPercentage(pct);
      return { subject: bulkSubjects[idx], mark: m, grade: g.grade, gp: g.gp, credit: bulkCredits[idx] || 1 };
      return { subject: bulkSubjects[idx], mark: m, grade: gradeFromPercentage(pct).grade };
    });
    const total = student.marks.reduce((a, b) => a + b, 0);
    const fullTotal = bulkFullMark * bulkSubjects.length;
    const pct = ((total / fullTotal) * 100) || 0;
    const overall = gradeFromPercentage(pct);
    const totalPoints = subjectResults.reduce((sum, r) => sum + (r.gp * r.credit), 0);
    const totalCredits = subjectResults.reduce((sum, r) => sum + r.credit, 0);
    const gpa = totalCredits ? (totalPoints / totalCredits) : 0;
    return { ...student, subjectResults, total, pct: pct.toFixed(2), grade: overall.grade, gpa: gpa.toFixed(2), credits: totalCredits };
    return { ...student, subjectResults, total, pct: pct.toFixed(2), grade: overall.grade };
  }).sort((a, b) => b.pct - a.pct);
}

function renderBulk(results) {
  const info = {
    className: document.getElementById('bulk-class').value.trim(),
    faculty: document.getElementById('bulk-faculty').value.trim(),
    section: document.getElementById('bulk-section').value.trim(),
    exam: document.getElementById('bulk-exam').value.trim(),
  };
  const brand = getBranding();
  bulkDetails.innerHTML = `
    <div class="marksheet-header">
      <div class="marksheet-brand">
        <img src="${brand.logo}" alt="School logo">
        <div>
          <h3>${brand.name}</h3>
          <p>${brand.address}</p>
          <p class="muted">${brand.session || info.exam || 'Academic Session'}</p>
        </div>
      </div>
    </div>
    <div class="student-meta">
      <div><strong>Class:</strong> ${info.className || '-'}</div>
      <div><strong>Faculty:</strong> ${info.faculty || '-'}</div>
      <div><strong>Section:</strong> ${info.section || '-'}</div>
      <div><strong>Exam:</strong> ${info.exam || brand.session || '-'}</div>
      <div><strong>Full Marks per subject:</strong> ${bulkFullMark}</div>
      <div><strong>Credits:</strong> ${bulkCredits.join(', ') || '-'}</div>
    </div>
  `;
  const table = document.createElement('div');
  table.innerHTML = `<div class="table-header">${['Roll', 'Name', ...bulkSubjects, 'Total', '%', 'Grade', 'GPA'].map(h => `<div>${h}</div>`).join('')}</div>`;
  results.forEach((res, idx) => {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = `0.8fr 1.5fr ${'1fr '.repeat(bulkSubjects.length + 4)}`;
  bulkDetails.innerHTML = `
    <div><strong>Class:</strong> ${info.className || '-'}</div>
    <div><strong>Faculty:</strong> ${info.faculty || '-'}</div>
    <div><strong>Section:</strong> ${info.section || '-'}</div>
    <div><strong>Exam:</strong> ${info.exam || '-'}</div>
    <div><strong>Full Marks per subject:</strong> ${bulkFullMark}</div>
  `;
  const table = document.createElement('div');
  table.innerHTML = `<div class="table-header">${['Roll', 'Name', ...bulkSubjects, 'Total', '%', 'Grade'].map(h => `<div>${h}</div>`).join('')}</div>`;
  results.forEach((res, idx) => {
    const row = document.createElement('div');
    row.className = 'row';
    row.style.gridTemplateColumns = `0.8fr 1.5fr ${'1fr '.repeat(bulkSubjects.length + 3)}`;
    row.innerHTML = `
      <div>${res.roll}</div>
      <div>${res.name}</div>
      ${res.subjectResults.map(r => `<div>${r.mark} (${r.grade})</div>`).join('')}
      <div>${res.total}</div>
      <div>${res.pct}%</div>
      <div>${res.grade}</div>
      <div>${res.gpa}</div>
    `;
    if (idx < 3) row.classList.add('highlight-top');
    if (Number(res.pct) < 40) row.classList.add('highlight-low');
    table.appendChild(row);
  });
  bulkTableWrapper.innerHTML = '';
  bulkTableWrapper.appendChild(table);
  if (results.length) {
    const avgPct = (results.reduce((sum, r) => sum + Number(r.pct), 0) / results.length).toFixed(2);
    const avgGpa = (results.reduce((sum, r) => sum + Number(r.gpa), 0) / results.length).toFixed(2);
    const toppers = results.slice(0, Math.min(3, results.length)).map(r => `${r.name} (${r.pct}%)`).join(', ');
    const lowCount = results.filter(r => Number(r.pct) < 40).length;
    bulkSummary.innerHTML = `
      <span>Class Avg: ${avgPct}%</span>
      <span>Avg GPA: ${avgGpa}</span>
      <span>Top performers: ${toppers || '-'}</span>
      <span>Students below 40%: ${lowCount}</span>
    `;
  } else {
    bulkSummary.textContent = '';
  }
  bulkResult.hidden = false;
  statusMessage.textContent = 'Class grade sheet ready.';
  currentExportHeaders = ['Roll', 'Name', ...bulkSubjects, 'Total', 'Percentage', 'Grade', 'GPA'];
  currentExportData = results.map(r => [r.roll, r.name, ...r.subjectResults.map(sr => sr.mark), r.total, r.pct, r.grade, r.gpa]);
  bulkResult.hidden = false;
  statusMessage.textContent = 'Class grade sheet ready.';
  currentExportHeaders = ['Roll', 'Name', ...bulkSubjects, 'Total', 'Percentage', 'Grade'];
  currentExportData = results.map(r => [r.roll, r.name, ...r.subjectResults.map(sr => sr.mark), r.total, r.pct, r.grade]);
}

bulkForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  bulkErrors.textContent = '';
  bulkResult.hidden = true;
  let students;
  try {
    const validated = validateBulk();
    if (typeof validated === 'string') {
      bulkErrors.textContent = validated;
      return;
    }
    students = validated;
  } catch (err) {
    bulkErrors.textContent = err.message;
    return;
  }
  disableActions(true);
  await showProcessing(bulkProcessing, bulkSteps, [
    'Reading student records…',
    'Calculating totals and percentages…',
    'Assigning grades for each student…',
    'Preparing class grade sheet…'
  ]);
  bulkProcessing.hidden = true;
  const results = calculateBulk(students);
  renderBulk(results);
  disableActions(false);
});

function downloadCSV(headers, rows) {
  if (!headers.length || !rows.length) {
    statusMessage.textContent = 'No data to export yet.';
    return;
  }
  const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'grade-sheet.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  statusMessage.textContent = 'CSV downloaded successfully.';
}

exportCsvBtn.addEventListener('click', () => downloadCSV(currentExportHeaders, currentExportData));
printBtn.addEventListener('click', () => window.print());

// Initial header state
buildBulkHeader();
