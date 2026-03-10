import { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

const METADATA_KEYS = new Set([
  'name', 'id', 'sis_id', 'section', 'section_id', 'section_sis_id',
  'submitted', 'attempt', 'n correct', 'n incorrect', 'score',
  'student', 'student_id', 'student_name', 'login_id',
]);

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return { headers: [], rows: [] };

  function parseLine(line) {
    const cells = [];
    let cell = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQ && line[i + 1] === '"') { cell += '"'; i++; }
        else inQ = !inQ;
      } else if (c === ',' && !inQ) {
        cells.push(cell.trim().replace(/^"|"$/g, ''));
        cell = '';
      } else cell += c;
    }
    cells.push(cell.trim().replace(/^"|"$/g, ''));
    return cells;
  }

  const headers = parseLine(lines[0]);
  const rows = lines
    .slice(1)
    .map(line => parseLine(line))
    .filter(cells => cells.some(v => v !== ''));

  return { headers, rows };
}

function extractCourseNumber(filename) {
  const match = filename.match(/[A-Z]+[_\-\s]?(\d{3,4})/i);
  return match ? match[1] : '';
}

function cleanEMPLID(id) {
  return (id || '').replace(/^OSH/i, '').trim();
}

function identifyQuestionColumns(headers, rows) {
  const answerColIdxs = [];
  const scoreColIdxs  = [];

  for (let i = 0; i < headers.length; i++) {
    if (METADATA_KEYS.has(headers[i].toLowerCase().trim())) continue;

    const vals = rows.map(r => (r[i] || '').trim()).filter(v => v !== '');
    if (!vals.length) continue;

    if (vals.every(v => /^[A-Ea-e]$/.test(v))) {
      answerColIdxs.push(i);
    } else if (vals.every(v => /^[01]$/.test(v))) {
      scoreColIdxs.push(i);
    }
  }

  return { answerColIdxs, scoreColIdxs };
}

function toExcelDateSerial(dateStr) {
  const [year, month, day] = dateStr.split('-').map(Number);
  const d = new Date(Date.UTC(year, month - 1, day));
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return (d - epoch) / (24 * 60 * 60 * 1000);
}

export default function App() {
  const [csvFile, setCsvFile]         = useState(null);
  const [testIdsFile, setTestIdsFile] = useState(null);
  const [parsedCSV, setParsedCSV]     = useState(null);
  const [testIdsMap, setTestIdsMap]   = useState(null);
  const [term, setTerm]               = useState('');
  const [classNumber, setClassNumber] = useState('');
  const [effectiveDate, setEffectiveDate] = useState('');
  const [detectedTestId, setDetectedTestId] = useState('');
  const [preview, setPreview]         = useState(null);
  const [error, setError]             = useState('');
  const [downloaded, setDownloaded]   = useState(false);
  const [csvDragActive, setCsvDragActive] = useState(false);
  const [testIdsDragActive, setTestIdsDragActive] = useState(false);

  function autoDetect(courseNum, tidsMap) {
    if (courseNum) {
      setClassNumber(prev => {
        if (prev && prev !== courseNum + '-') return prev;
        return courseNum + '-';
      });
    }
    if (courseNum && tidsMap && tidsMap[courseNum]) {
      setDetectedTestId(String(tidsMap[courseNum]));
    }
  }

  function handleCsvFile(file) {
    setCsvFile(file);
    setPreview(null);
    setDownloaded(false);
    setError('');
    const reader = new FileReader();
    reader.onload = (e) => {
      const { headers, rows } = parseCSV(e.target.result);
      setParsedCSV({ headers, rows });
      const courseNum = extractCourseNumber(file.name);
      autoDetect(courseNum, testIdsMap);
    };
    reader.readAsText(file);
  }

  function handleTestIdsFile(file) {
    setTestIdsFile(file);
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const map = {};
      let courseCol = -1, testIdCol = -1;
      for (let i = 0; i < Math.min(5, data.length); i++) {
        const row = (data[i] || []).map(c => String(c || '').toUpperCase());
        const ci = row.findIndex(c => c.includes('COURSE'));
        const ti = row.findIndex(c => c.includes('TEST') && c.includes('ID'));
        if (ci >= 0 && ti >= 0) {
          courseCol = ci; testIdCol = ti;
          for (let j = i + 1; j < data.length; j++) {
            const r = data[j];
            if (!r || r[courseCol] == null || r[testIdCol] == null) continue;
            const courseMatch = String(r[courseCol]).match(/\d+/);
            if (courseMatch) map[courseMatch[0]] = r[testIdCol];
          }
          break;
        }
      }
      setTestIdsMap(map);
      if (csvFile) {
        const courseNum = extractCourseNumber(csvFile.name);
        if (courseNum && map[courseNum]) setDetectedTestId(String(map[courseNum]));
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleTransform() {
    setError('');
    if (!parsedCSV)          { setError('Please upload a Canvas CSV file first.'); return; }
    if (!term.trim())        { setError('Please enter a Term (e.g. S26).'); return; }
    if (!classNumber.trim()) { setError('Please enter a Class Number (e.g. 320-004).'); return; }
    if (!effectiveDate)      { setError('Please enter an Effective Date.'); return; }
    if (!detectedTestId)     { setError('No Test ID found — please upload the Test IDs reference file.'); return; }

    const { headers, rows } = parsedCSV;
    const { answerColIdxs, scoreColIdxs } = identifyQuestionColumns(headers, rows);

    if (!answerColIdxs.length) {
      setError('Could not identify question columns. Make sure this is a Canvas Student Analysis Report CSV.');
      return;
    }

    const sisIdIdx = headers.indexOf('sis_id');
    const idIdx    = headers.indexOf('id');
    const outputRows = [];

    for (let qi = 0; qi < answerColIdxs.length; qi++) {
      for (const studentRow of rows) {
        const rawId = studentRow[sisIdIdx] ?? studentRow[idIdx] ?? '';
        const emplid = cleanEMPLID(rawId);
        if (!emplid) continue;

        const answer  = (studentRow[answerColIdxs[qi]] || '').trim().toUpperCase();
        const correct = scoreColIdxs[qi] !== undefined
          ? parseInt(studentRow[scoreColIdxs[qi]] || '0')
          : 0;

        outputRows.push({
          EMPLID:         emplid,
          TERM:           term.trim().toUpperCase(),
          CLASSNUMBER:    classNumber.trim().toUpperCase(),
          TESTID:         parseInt(detectedTestId),
          TYPE:           'Post',
          QUESTIONID:     qi + 1,
          EFFECTIVEDATE:  effectiveDate,
          ANSWEROPTION:   answer,
          STUDENTCORRECT: isNaN(correct) ? 0 : correct,
        });
      }
    }

    if (!outputRows.length) {
      setError('No valid student data found. Check that the CSV contains student rows.');
      return;
    }

    const students = new Set(outputRows.map(r => r.EMPLID)).size;
    setPreview({ rows: outputRows, students, questions: answerColIdxs.length, total: outputRows.length });
    setDownloaded(false);
  }

  function handleDownload() {
    if (!preview) return;
    const wb = XLSX.utils.book_new();
    const headerRow = ['EMPLID', 'TERM', 'CLASSNUMBER', 'TESTID', 'TYPE', 'QUESTIONID', 'EFFECTIVEDATE', 'ANSWEROPTION', 'STUDENTCORRECT'];
    const aoa = [headerRow];
    const dateSerial = toExcelDateSerial(effectiveDate);

    for (const row of preview.rows) {
      aoa.push([
        row.EMPLID, row.TERM, row.CLASSNUMBER, row.TESTID, row.TYPE,
        row.QUESTIONID, dateSerial, row.ANSWEROPTION, row.STUDENTCORRECT,
      ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = 1; r < aoa.length; r++) {
      const emplCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
      if (emplCell) { emplCell.t = 's'; emplCell.z = '@'; }
      const dateCell = ws[XLSX.utils.encode_cell({ r, c: 6 })];
      if (dateCell) { dateCell.t = 'n'; dateCell.z = 'mm-dd-yy'; }
      for (const c of [3, 5, 8]) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell) cell.t = 'n';
      }
    }
    ws['!cols'] = [
      { wch: 14 }, { wch: 8 }, { wch: 14 }, { wch: 10 }, { wch: 8 },
      { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 16 },
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Attempt Details');
    const filename = `${term.trim().toUpperCase()}_${classNumber.trim().replace(/[/\\]/g, '-')}_Post.xlsx`;
    XLSX.writeFile(wb, filename);
    setDownloaded(true);
  }

  function makeDragHandlers(setActive, handler) {
    return {
      onDragEnter: (e) => { e.preventDefault(); setActive(true); },
      onDragOver:  (e) => { e.preventDefault(); setActive(true); },
      onDragLeave: (e) => { e.preventDefault(); setActive(false); },
      onDrop: (e) => {
        e.preventDefault(); setActive(false);
        const file = e.dataTransfer.files[0];
        if (file) handler(file);
      },
    };
  }

  const csvDragHandlers     = makeDragHandlers(setCsvDragActive, handleCsvFile);
  const testIdsDragHandlers = makeDragHandlers(setTestIdsDragActive, handleTestIdsFile);
  const detectedCourse = csvFile ? extractCourseNumber(csvFile.name) : '';
  const outputFilename = term && classNumber
    ? `${term.toUpperCase()}_${classNumber.toUpperCase()}_Post.xlsx`
    : 'output.xlsx';

  return (
    <div className="app">
      <header className="app-header">
        <span className="header-badge">POST TEST</span>
        <h1 className="header-title">Data Transformation Tool</h1>
        <p className="header-sub">Convert Canvas exports to the standardized database format</p>
      </header>
      <main className="main">
        <section className="card">
          <div className="step-label"><span className="step-num">01</span>Upload Files</div>
          <div className="upload-grid">
            <div className={`drop-zone ${csvDragActive ? 'drag-over' : ''} ${csvFile ? 'has-file' : ''}`}
              {...csvDragHandlers} onClick={() => document.getElementById('csv-input').click()}>
              <input id="csv-input" type="file" accept=".csv" style={{ display: 'none' }}
                onChange={e => e.target.files[0] && handleCsvFile(e.target.files[0])} />
              <div className="drop-icon">{csvFile ? '✓' : '↑'}</div>
              <div className="drop-label">{csvFile ? csvFile.name : 'Canvas Export CSV'}</div>
              <div className="drop-sub">{csvFile ? `${parsedCSV?.rows?.length || 0} student rows found` : 'Drop or click to upload'}</div>
            </div>
            <div className={`drop-zone ${testIdsDragActive ? 'drag-over' : ''} ${testIdsFile ? 'has-file' : ''}`}
              {...testIdsDragHandlers} onClick={() => document.getElementById('testids-input').click()}>
              <input id="testids-input" type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                onChange={e => e.target.files[0] && handleTestIdsFile(e.target.files[0])} />
              <div className="drop-icon">{testIdsFile ? '✓' : '↑'}</div>
              <div className="drop-label">{testIdsFile ? testIdsFile.name : 'Test IDs Reference'}</div>
              <div className="drop-sub">{testIdsFile ? (detectedTestId ? `Test ID ${detectedTestId} matched` : 'Loaded — no match yet') : 'Drop or click to upload (.xlsx)'}</div>
            </div>
          </div>
          {(detectedCourse || detectedTestId) && (
            <div className="auto-detect-bar">
              {detectedCourse && <span className="detect-pill"><span className="detect-dot" />Course {detectedCourse} detected from filename</span>}
              {detectedTestId && <span className="detect-pill"><span className="detect-dot" />Test ID {detectedTestId} auto-filled</span>}
            </div>
          )}
        </section>
        <section className="card">
          <div className="step-label"><span className="step-num">02</span>Configure Output</div>
          <div className="fields-grid">
            <div className="field">
              <label className="field-label">TERM</label>
              <input className="field-input" type="text" placeholder="e.g. S26" value={term} onChange={e => setTerm(e.target.value)} />
            </div>
            <div className="field">
              <label className="field-label">CLASS NUMBER</label>
              <input className="field-input" type="text" placeholder="e.g. 320-004" value={classNumber} onChange={e => setClassNumber(e.target.value)} />
              {classNumber.endsWith('-') && <div className="field-hint">↑ Type the section after the dash</div>}
            </div>
            <div className="field">
              <label className="field-label">TEST ID<span className="auto-tag">auto</span></label>
              <input className={`field-input ${detectedTestId ? 'auto-filled' : ''}`} type="text"
                placeholder="Upload Test IDs file above" value={detectedTestId} onChange={e => setDetectedTestId(e.target.value)} />
            </div>
            <div className="field">
              <label className="field-label">EFFECTIVE DATE</label>
              <input className="field-input" type="date" value={effectiveDate} onChange={e => setEffectiveDate(e.target.value)} />
            </div>
          </div>
        </section>
        <section className="card">
          <div className="step-label"><span className="step-num">03</span>Transform & Download</div>
          {error && <div className="error-msg" role="alert">{error}</div>}
          <button className="btn-transform" onClick={handleTransform}>Run Transformation</button>
          {preview && (
            <div className="preview-box">
              <div className="preview-stats">
                <div className="stat"><div className="stat-val">{preview.students}</div><div className="stat-key">Students</div></div>
                <div className="stat-divider" />
                <div className="stat"><div className="stat-val">{preview.questions}</div><div className="stat-key">Questions</div></div>
                <div className="stat-divider" />
                <div className="stat"><div className="stat-val">{preview.total.toLocaleString()}</div><div className="stat-key">Output Rows</div></div>
              </div>
              <button className="btn-download" onClick={handleDownload}>
                {downloaded ? '✓  File downloaded' : `↓  Download ${outputFilename}`}
              </button>
            </div>
          )}
        </section>
      </main>
      <footer className="app-footer">
        POST TEST Data Transformation · UW Oshkosh College of Business · Assessment
      </footer>
    </div>
  );
}
