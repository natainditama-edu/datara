/**
 * Kalkulator Data Mining - Universitas Primakara
 * Dosen: Nengah Widya Utami, S.Pd., M.Kom.
 * Core Logic Engine (app.js)
 */

// State Application
const state = {
  currentDatasetName: "Dataset 03 Prediksi Tingkat Stress Mahasiswa_original.xlsx",
  dataTraining: [],
  dataTesting: [],
  headers: [],
  targetColIndex: -1,
  idColIndex: 0,
  featureIndices: [],
  mode: "klasifikasi", // 'klasifikasi' or 'regresi'
  smoothingMethod: "laplace", // 'laplace' (Orange standard) or 'pure' (Slide Pertemuan 7)
  analysisResults: null
};

// DOM Ready
document.addEventListener("DOMContentLoaded", () => {
  initFileUpload();
  initPresetControls();
  initMethodToggle();
  initExportButton();

  // Load default dataset (Dataset 03 Original)
  loadPresetDataset("Dataset 03 Prediksi Tingkat Stress Mahasiswa_original.xlsx");
});

/* ==========================================================================
   Presets & File Upload
   ========================================================================== */
function initPresetControls() {
  const btnSample03 = document.getElementById("btnLoadSample03");
  if (btnSample03) {
    btnSample03.addEventListener("click", () => {
      loadPresetDataset("Dataset 03 Prediksi Tingkat Stress Mahasiswa_original.xlsx");
      const dropdown = document.getElementById("presetDropdown");
      if (dropdown) dropdown.value = "Dataset 03 Prediksi Tingkat Stress Mahasiswa_original.xlsx";
    });
  }

  const dropdown = document.getElementById("presetDropdown");
  if (dropdown) {
    dropdown.addEventListener("change", (e) => {
      if (e.target.value) {
        loadPresetDataset(e.target.value);
      }
    });
  }
}

function loadPresetDataset(presetKey) {
  if (!window.SAMPLE_DATASETS || !window.SAMPLE_DATASETS[presetKey]) {
    alert("Dataset contoh tidak ditemukan: " + presetKey);
    return;
  }

  const ds = window.SAMPLE_DATASETS[presetKey];
  state.currentDatasetName = presetKey;

  processParsedWorkbook({
    data_training: ds.data_training,
    data_testing: ds.data_testing
  }, presetKey);
}

function initFileUpload() {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");

  if (!dropzone || !fileInput) return;

  dropzone.addEventListener("click", () => fileInput.click());

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  });
}

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      
      const sheets = {};
      workbook.SheetNames.forEach(sheetName => {
        const ws = workbook.Sheets[sheetName];
        sheets[sheetName.toLowerCase().trim()] = XLSX.utils.sheet_to_json(ws, { header: 1 });
      });

      let trainRows = sheets["data_training"] || sheets["training"] || sheets["train"] || sheets[Object.keys(sheets)[0]];
      let testRows = sheets["data_testing"] || sheets["testing"] || sheets["test"] || [];

      trainRows = (trainRows || []).filter(r => r && r.some(cell => cell !== null && cell !== undefined && cell !== ""));
      testRows = (testRows || []).filter(r => r && r.some(cell => cell !== null && cell !== undefined && cell !== ""));

      state.currentDatasetName = file.name;
      processParsedWorkbook({
        data_training: trainRows,
        data_testing: testRows
      }, file.name);

    } catch (err) {
      console.error(err);
      alert("Gagal membaca file Excel. Pastikan format file .xlsx valid.");
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ==========================================================================
   Workbook Processing & Schema Detection
   ========================================================================== */
function processParsedWorkbook(data, sourceName) {
  const trainRows = data.data_training;
  const testRows = data.data_testing;

  if (!trainRows || trainRows.length < 2) {
    alert("Format data training tidak valid (minimal 1 baris header dan 1 baris data).");
    return;
  }

  const rawHeaders = trainRows[0].map(h => String(h || "").trim());
  state.headers = rawHeaders;

  state.dataTraining = trainRows.slice(1);
  state.dataTesting = testRows.length > 1 ? testRows.slice(1) : [];

  let targetIdx = rawHeaders.length - 1;
  state.targetColIndex = targetIdx;

  state.featureIndices = [];
  for (let i = 1; i < rawHeaders.length; i++) {
    if (i !== targetIdx) state.featureIndices.push(i);
  }

  const targetSample = state.dataTraining.map(r => r[targetIdx]).filter(v => v !== null && v !== undefined);
  const isNumericTarget = targetSample.length > 0 && targetSample.every(v => !isNaN(Number(v)) && typeof v !== "boolean");

  if (isNumericTarget) {
    state.mode = "regresi";
  } else {
    state.mode = "klasifikasi";
  }

  updateDatasetMetaUI(sourceName);
  runCalculation();
}

function updateDatasetMetaUI(sourceName) {
  const infoBar = document.getElementById("fileInfoBar");
  if (infoBar) {
    infoBar.style.display = "flex";
    document.getElementById("currentFileName").innerText = sourceName;
    document.getElementById("trainRowCount").innerText = `${state.dataTraining.length} baris training`;
    document.getElementById("testRowCount").innerText = `${state.dataTesting.length} baris testing`;
    
    const rolePill = document.getElementById("detectedRolePill");
    if (state.mode === "klasifikasi") {
      rolePill.className = "role-pill role-klasifikasi";
      rolePill.innerText = "Klasifikasi (Naive Bayes)";
    } else {
      rolePill.className = "role-pill role-estimasi";
      rolePill.innerText = "Estimasi / Peramalan (Regresi Linear)";
    }
  }

  const modeSelector = document.getElementById("modeSelector");
  if (modeSelector) {
    modeSelector.value = state.mode;
  }
}

function initMethodToggle() {
  const methodSelect = document.getElementById("smoothingSelector");
  if (methodSelect) {
    methodSelect.addEventListener("change", (e) => {
      state.smoothingMethod = e.target.value;
      runCalculation();
    });
  }

  const modeSelect = document.getElementById("modeSelector");
  if (modeSelect) {
    modeSelect.addEventListener("change", (e) => {
      state.mode = e.target.value;
      runCalculation();
    });
  }
}

/* ==========================================================================
   Master Calculation Dispatcher
   ========================================================================== */
function runCalculation() {
  if (state.mode === "klasifikasi") {
    calculateNaiveBayes();
  } else {
    calculateRegression();
  }
}

/* ==========================================================================
   NAIVE BAYES CLASSIFIER ENGINE
   ========================================================================== */
function calculateNaiveBayes() {
  const trainData = state.dataTraining;
  const testData = state.dataTesting;
  const targetIdx = state.targetColIndex;
  const featureIdxs = state.featureIndices;
  const headers = state.headers;

  const totalN = trainData.length;

  // 1. Prior Probability P(C)
  const classCounts = {};
  trainData.forEach(row => {
    const c = String(row[targetIdx]).trim();
    classCounts[c] = (classCounts[c] || 0) + 1;
  });

  const classes = Object.keys(classCounts);
  const priors = {};
  classes.forEach(c => {
    priors[c] = classCounts[c] / totalN;
  });

  // 2. Feature Distinct Values & Conditional Counts
  const featureValues = {};
  const featureCounts = {};
  let hasZeroProbability = false;

  featureIdxs.forEach(fIdx => {
    featureValues[fIdx] = new Set();
    featureCounts[fIdx] = {};

    trainData.forEach(row => {
      const val = String(row[fIdx] || "").trim();
      if (val) featureValues[fIdx].add(val);
    });

    testData.forEach(row => {
      const val = String(row[fIdx] || "").trim();
      if (val) featureValues[fIdx].add(val);
    });

    const valList = Array.from(featureValues[fIdx]);
    valList.forEach(val => {
      featureCounts[fIdx][val] = {};
      classes.forEach(c => {
        featureCounts[fIdx][val][c] = 0;
      });
    });

    trainData.forEach(row => {
      const val = String(row[fIdx] || "").trim();
      const c = String(row[targetIdx]).trim();
      if (featureCounts[fIdx][val]) {
        featureCounts[fIdx][val][c] = (featureCounts[fIdx][val][c] || 0) + 1;
      }
    });
  });

  featureIdxs.forEach(fIdx => {
    Object.keys(featureCounts[fIdx]).forEach(val => {
      classes.forEach(c => {
        if (featureCounts[fIdx][val][c] === 0) {
          hasZeroProbability = true;
        }
      });
    });
  });

  // 3. Evaluate Data Testing
  const testPredictions = testData.map((testRow, testIndex) => {
    const idVal = testRow[0] || (testIndex + 9);
    const rowValues = {};
    featureIdxs.forEach(fIdx => {
      rowValues[headers[fIdx]] = String(testRow[fIdx] || "").trim();
    });

    const classResults = {};
    let totalScoreLaplace = 0;
    let totalScorePure = 0;

    classes.forEach(c => {
      const Nc = classCounts[c];
      const prior = priors[c];

      const pureMults = [];
      const laplaceMults = [];

      let pureCondProd = 1;
      let laplaceCondProd = 1;

      featureIdxs.forEach(fIdx => {
        const val = String(testRow[fIdx] || "").trim();
        const k = featureValues[fIdx].size;
        const count = (featureCounts[fIdx][val] && featureCounts[fIdx][val][c]) ? featureCounts[fIdx][val][c] : 0;

        const pPure = Nc > 0 ? (count / Nc) : 0;
        pureMults.push({
          feature: headers[fIdx],
          val: val,
          count: count,
          denom: Nc,
          p: pPure,
          expr: `${count}/${Nc}`
        });
        pureCondProd *= pPure;

        const pLaplace = (count + 1) / (Nc + k);
        laplaceMults.push({
          feature: headers[fIdx],
          val: val,
          count: count,
          denom: Nc + k,
          k: k,
          p: pLaplace,
          expr: `(${count}+1)/(${Nc}+${k})`
        });
        laplaceCondProd *= pLaplace;
      });

      const pureScore = prior * pureCondProd;
      const laplaceScore = prior * laplaceCondProd;

      totalScorePure += pureScore;
      totalScoreLaplace += laplaceScore;

      classResults[c] = {
        prior: prior,
        pureMults: pureMults,
        pureCondProd: pureCondProd,
        pureScore: pureScore,
        laplaceMults: laplaceMults,
        laplaceCondProd: laplaceCondProd,
        laplaceScore: laplaceScore
      };
    });

    classes.forEach(c => {
      classResults[c].normProbPure = totalScorePure > 0 ? (classResults[c].pureScore / totalScorePure) : (1 / classes.length);
      classResults[c].normProbLaplace = totalScoreLaplace > 0 ? (classResults[c].laplaceScore / totalScoreLaplace) : (1 / classes.length);
    });

    let bestClassLaplace = classes[0];
    let maxScoreLaplace = -1;
    let bestClassPure = classes[0];
    let maxScorePure = -1;

    classes.forEach(c => {
      if (classResults[c].laplaceScore > maxScoreLaplace) {
        maxScoreLaplace = classResults[c].laplaceScore;
        bestClassLaplace = c;
      }
      if (classResults[c].pureScore > maxScorePure) {
        maxScorePure = classResults[c].pureScore;
        bestClassPure = c;
      }
    });

    return {
      id: idVal,
      rowValues: rowValues,
      rawRow: testRow,
      classResults: classResults,
      bestClassLaplace: bestClassLaplace,
      bestClassPure: bestClassPure,
      isZeroPure: totalScorePure === 0
    };
  });

  state.analysisResults = {
    totalN: totalN,
    classCounts: classCounts,
    classes: classes,
    priors: priors,
    featureValues: featureValues,
    featureCounts: featureCounts,
    hasZeroProbability: hasZeroProbability,
    testPredictions: testPredictions
  };

  renderNaiveBayesUI();
}

/* ==========================================================================
   RENDER NAIVE BAYES RESULTS TO UI (Clean, Minimal, No Emojis)
   ========================================================================== */
function renderNaiveBayesUI() {
  const container = document.getElementById("resultsArea");
  if (!container) return;

  const res = state.analysisResults;
  const isLaplace = state.smoothingMethod === "laplace";

  let html = `
    <!-- Stats Row -->
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-value">${res.totalN}</div>
        <div class="stat-label">Data Training</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${state.featureIndices.length}</div>
        <div class="stat-label">Fitur (X)</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${res.classes.length}</div>
        <div class="stat-label">Kelas Target (${res.classes.join(", ")})</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${res.testPredictions.length}</div>
        <div class="stat-label">Data Testing</div>
      </div>
    </div>
  `;

  // Zero Probability Alert (Concise, no icon)
  if (res.hasZeroProbability && !isLaplace) {
    html += `
      <div class="alert alert-warning" style="padding: 12px 16px;">
        <div>
          <strong>Perhatian:</strong> Ditemukan frekuensi 0 pada data training untuk kombinasi fitur data uji. 
          Pada metode murni slide dosen, seluruh skor posterior menjadi 0. 
          Gunakan <strong>Laplace Smoothing</strong> untuk hasil stabil sesuai standar Orange Data Mining.
        </div>
      </div>
    `;
  }

  // Section 1: HASIL PREDIKSI DATA TESTING (Langsung ditampilkan di awal)
  html += `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Hasil Prediksi Data Testing</div>
          <div class="card-desc">Keputusan klasifikasi untuk data uji berdasarkan nilai probabilitas posterior tertinggi.</div>
        </div>
      </div>

      <div style="display: flex; flex-direction: column; gap: 16px;">
        ${res.testPredictions.map(testItem => {
          const chosenClass = isLaplace ? testItem.bestClassLaplace : testItem.bestClassPure;
          const cardClass = `pred-${String(chosenClass).toLowerCase()}`;

          return `
            <div class="prediction-card ${cardClass}" style="padding: 16px;">
              <div class="prediction-header" style="margin-bottom: 12px;">
                <div>
                  <div class="pred-title" style="font-size: 15px;">
                    Data Testing No. ${testItem.id}
                  </div>
                  <div style="font-size: 12px; color: var(--gray-600); margin-top: 2px;">
                    Input: ${Object.entries(testItem.rowValues).map(([k, v]) => `${k}: <strong>${v}</strong>`).join(", ")}
                  </div>
                </div>

                <div>
                  <span class="badge-decision badge-${String(chosenClass).toLowerCase()}" style="font-size: 13px; padding: 4px 12px;">
                    Prediksi: ${chosenClass}
                  </span>
                </div>
              </div>

              <!-- Probability Bar -->
              <div class="prob-bar-container" style="margin: 10px 0;">
                <div class="prob-labels">
                  ${res.classes.map(c => {
                    const prob = isLaplace ? testItem.classResults[c].normProbLaplace : testItem.classResults[c].normProbPure;
                    return `<span>${c}: ${(prob * 100).toFixed(2)}%</span>`;
                  }).join("")}
                </div>
                <div class="prob-bar" style="height: 18px;">
                  ${res.classes.map((c, cIdx) => {
                    const prob = isLaplace ? testItem.classResults[c].normProbLaplace : testItem.classResults[c].normProbPure;
                    const colorClass = cIdx === 0 ? 'bg-blue' : (cIdx === 1 ? 'bg-red' : 'bg-green');
                    return `
                      <div class="prob-bar-segment ${colorClass}" style="width: ${(prob * 100).toFixed(2)}%;">
                        ${(prob * 100) >= 15 ? `${(prob * 100).toFixed(1)}%` : ''}
                      </div>
                    `;
                  }).join("")}
                </div>
              </div>

              <!-- Step Calculation Numbers -->
              <div class="step-breakdown" style="padding: 10px 12px; margin-top: 8px;">
                <div style="font-size: 12px; font-weight: 600; color: var(--gray-700); margin-bottom: 6px;">Perkalian Probabilitas:</div>
                ${res.classes.map(c => {
                  const cr = testItem.classResults[c];
                  const mults = isLaplace ? cr.laplaceMults : cr.pureMults;
                  const finalScore = isLaplace ? cr.laplaceScore : cr.pureScore;
                  const normP = isLaplace ? cr.normProbLaplace : cr.normProbPure;
                  const numStr = mults.map(m => m.p.toFixed(3)).join(" × ");

                  return `
                    <div style="font-size: 12px; margin-bottom: 4px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 4px;">
                      <span style="color: var(--gray-600);">${c}:</span>
                      <span class="font-mono">
                        ${cr.prior.toFixed(2)} × ${numStr} = <strong>${finalScore.toFixed(6)}</strong> (${(normP * 100).toFixed(2)}%)
                      </span>
                    </div>
                  `;
                }).join("")}
              </div>
            </div>
          `;
        }).join("")}
      </div>
    </div>
  `;

  // Section 2: Data Masukan (Training & Testing)
  html += `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Tabel Data Masukan</div>
          <div class="card-desc">Data training (${res.totalN} baris) dan data testing (${res.testPredictions.length} baris) dari file Excel.</div>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px;">
        <div>
          <h4 style="margin-bottom: 8px; font-size: 13px; font-weight: 700; color: var(--gray-700);">Data Training (N = ${res.totalN})</h4>
          <div class="table-responsive" style="max-height: 240px;">
            <table>
              <thead>
                <tr>${state.headers.map(h => `<th>${h}</th>`).join("")}</tr>
              </thead>
              <tbody>
                ${state.dataTraining.map(r => `
                  <tr>
                    ${r.map((cell, idx) => `
                      <td class="${idx === state.targetColIndex ? 'font-bold' : ''}">
                        ${cell !== null && cell !== undefined ? cell : '-'}
                      </td>
                    `).join("")}
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>

        <div>
          <h4 style="margin-bottom: 8px; font-size: 13px; font-weight: 700; color: var(--gray-700);">Data Testing (${res.testPredictions.length} Baris)</h4>
          <div class="table-responsive" style="max-height: 240px;">
            <table>
              <thead>
                <tr>${state.headers.map(h => `<th>${h}</th>`).join("")}</tr>
              </thead>
              <tbody>
                ${res.testPredictions.map(p => `
                  <tr style="background: #f0fdf4;">
                    <td><strong>${p.id}</strong></td>
                    ${state.featureIndices.map(fIdx => `<td>${p.rawRow[fIdx] || '-'}</td>`).join("")}
                    <td><span class="badge-decision badge-${String(isLaplace ? p.bestClassLaplace : p.bestClassPure).toLowerCase()}" style="font-size: 11px; padding: 2px 8px;">${isLaplace ? p.bestClassLaplace : (p.isZeroPure ? 'Tidak Terdefinisi' : p.bestClassPure)}</span></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  `;

  // Section 3: Probabilitas Prior P(Kelas)
  html += `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Peluang Awal Kelas: Prior Probability P(Kelas)</div>
          <div class="card-desc">P(Kelas) = Jumlah Kasus Kelas / Total Seluruh Kasus (N = ${res.totalN})</div>
        </div>
      </div>

      <div class="table-responsive">
        <table>
          <thead>
            <tr>
              <th>Kelas Target</th>
              <th class="text-center">Jumlah Kasus</th>
              <th class="text-center">Total N</th>
              <th class="text-center">Pecahan</th>
              <th class="text-center">Desimal P(Kelas)</th>
              <th class="text-center">Persentase</th>
            </tr>
          </thead>
          <tbody>
            ${res.classes.map(c => {
              const count = res.classCounts[c];
              const prior = res.priors[c];
              return `
                <tr>
                  <td><strong>${c}</strong></td>
                  <td class="text-center font-bold">${count}</td>
                  <td class="text-center">${res.totalN}</td>
                  <td class="text-center font-mono">${count} / ${res.totalN}</td>
                  <td class="text-center font-mono font-bold" style="color: var(--primary);">${prior.toFixed(4)}</td>
                  <td class="text-center font-bold">${(prior * 100).toFixed(1)}%</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;

  // Section 4: Probabilitas Bersyarat
  html += `
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Tabel Probabilitas Bersyarat: P(Fitur | Kelas)</div>
          <div class="card-desc">
            ${isLaplace 
              ? 'Menggunakan rumus Laplace: (count + 1) / (N_kelas + k)' 
              : 'Menggunakan rumus klasik: count / N_kelas'}
          </div>
        </div>
      </div>

      <div class="table-responsive">
        <table>
          <thead>
            <tr>
              <th>Fitur</th>
              <th>Nilai Kategori</th>
              <th class="text-center">Variasi (k)</th>
              ${res.classes.map(c => `<th class="text-center">Kasus (${c})</th>`).join("")}
              ${res.classes.map(c => `<th class="text-center">${isLaplace ? `P_Laplace(X|${c})` : `P(X|${c})`}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${state.featureIndices.map(fIdx => {
              const fname = state.headers[fIdx];
              const vals = Array.from(res.featureValues[fIdx]);
              const k = vals.length;

              return vals.map((v, vIdx) => {
                return `
                  <tr>
                    ${vIdx === 0 ? `<td rowspan="${k}" style="vertical-align: top; font-weight: 700; background: var(--gray-50);">${fname}</td>` : ''}
                    <td><strong>${v}</strong></td>
                    <td class="text-center"><span class="badge-tag">k = ${k}</span></td>
                    ${res.classes.map(c => {
                      const cnt = (res.featureCounts[fIdx][v] && res.featureCounts[fIdx][v][c]) || 0;
                      return `<td class="text-center font-mono">${cnt}</td>`;
                    }).join("")}
                    ${res.classes.map(c => {
                      const cnt = (res.featureCounts[fIdx][v] && res.featureCounts[fIdx][v][c]) || 0;
                      const Nc = res.classCounts[c];
                      if (isLaplace) {
                        const pLp = (cnt + 1) / (Nc + k);
                        return `
                          <td class="text-center font-mono font-bold" style="color: var(--primary);">
                            ${pLp.toFixed(4)} <span style="font-size: 11px; color: var(--gray-500); font-weight: normal;">(${cnt}+1)/(${Nc}+${k})</span>
                          </td>
                        `;
                      } else {
                        const pPure = Nc > 0 ? (cnt / Nc) : 0;
                        return `
                          <td class="text-center font-mono font-bold ${pPure === 0 ? 'font-bold' : ''}" style="${pPure === 0 ? 'color: var(--danger); background: #fef2f2;' : 'color: var(--primary);'}">
                            ${pPure.toFixed(4)} <span style="font-size: 11px; color: var(--gray-500); font-weight: normal;">${cnt}/${Nc}</span>
                          </td>
                        `;
                      }
                    }).join("")}
                  </tr>
                `;
              }).join("");
            }).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;

  container.innerHTML = html;
}

/* ==========================================================================
   MULTIPLE LINEAR REGRESSION ENGINE
   ========================================================================== */
function calculateRegression() {
  const trainData = state.dataTraining;
  const testData = state.dataTesting;
  const targetIdx = state.targetColIndex;
  const featureIdxs = state.featureIndices;
  const headers = state.headers;

  const N = trainData.length;
  const p = featureIdxs.length;

  const X = [];
  const Y = [];

  trainData.forEach(row => {
    const xRow = [1];
    featureIdxs.forEach(fIdx => {
      xRow.push(parseFloat(row[fIdx]) || 0);
    });
    X.push(xRow);
    Y.push(parseFloat(row[targetIdx]) || 0);
  });

  const XT = transposeMatrix(X);
  const XTX = multiplyMatrices(XT, X);
  const invXTX = invertMatrix(XTX);

  if (!invXTX) {
    alert("Matriks X^T * X bersifat singular / tidak dapat diinversikan.");
    return;
  }

  const XTY = multiplyMatrixVector(XT, Y);
  const beta = multiplyMatrixVector(invXTX, XTY);

  let ssTot = 0;
  let ssRes = 0;
  let sumAbsErr = 0;
  let sumAbsPctErr = 0;

  const meanY = Y.reduce((acc, v) => acc + v, 0) / N;

  const trainPreds = [];
  for (let i = 0; i < N; i++) {
    let yHat = 0;
    for (let j = 0; j <= p; j++) {
      yHat += X[i][j] * beta[j];
    }
    trainPreds.push(yHat);

    const err = Y[i] - yHat;
    ssRes += err * err;
    ssTot += (Y[i] - meanY) * (Y[i] - meanY);
    sumAbsErr += Math.abs(err);
    if (Y[i] !== 0) {
      sumAbsPctErr += Math.abs(err / Y[i]);
    }
  }

  const r2 = ssTot !== 0 ? Math.max(0, 1 - (ssRes / ssTot)) : 1;
  const rmse = Math.sqrt(ssRes / N);
  const mae = sumAbsErr / N;
  const mape = (sumAbsPctErr / N) * 100;

  const testPredictions = testData.map((testRow, testIndex) => {
    const idVal = testRow[0] || (testIndex + 9);
    let predY = beta[0];
    const termBreakdown = [`b0 (${beta[0].toFixed(4)})`];

    featureIdxs.forEach((fIdx, colIdx) => {
      const xVal = parseFloat(testRow[fIdx]) || 0;
      const bCoeff = beta[colIdx + 1];
      predY += bCoeff * xVal;
      termBreakdown.push(`(${bCoeff.toFixed(4)} × ${xVal})`);
    });

    return {
      id: idVal,
      rawRow: testRow,
      predY: predY,
      termBreakdown: termBreakdown
    };
  });

  state.analysisResults = {
    N: N,
    p: p,
    beta: beta,
    r2: r2,
    rmse: rmse,
    mae: mae,
    mape: mape,
    meanY: meanY,
    testPredictions: testPredictions
  };

  renderRegressionUI();
}

/* ==========================================================================
   RENDER REGRESSION RESULTS TO UI (Clean, Minimal, No Emojis)
   ========================================================================== */
function renderRegressionUI() {
  const container = document.getElementById("resultsArea");
  if (!container) return;

  const res = state.analysisResults;
  const headers = state.headers;
  const featureIdxs = state.featureIndices;
  const targetIdx = state.targetColIndex;

  let formulaStr = `Y = ${res.beta[0].toFixed(4)}`;
  featureIdxs.forEach((fIdx, i) => {
    const b = res.beta[i + 1];
    const sign = b >= 0 ? "+" : "-";
    formulaStr += ` ${sign} ${Math.abs(b).toFixed(4)} · (${headers[fIdx]})`;
  });

  let html = `
    <!-- Stats Row -->
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-value">${(res.r2 * 100).toFixed(2)}%</div>
        <div class="stat-label">R² (Koefisien Determinasi)</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${res.rmse.toFixed(4)}</div>
        <div class="stat-label">RMSE</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${res.mae.toFixed(4)}</div>
        <div class="stat-label">MAE</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${res.mape.toFixed(2)}%</div>
        <div class="stat-label">MAPE</div>
      </div>
    </div>

    <!-- Section 1: HASIL PREDIKSI DATA TESTING -->
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Hasil Prediksi Data Testing</div>
          <div class="card-desc">Substitusi nilai atribut uji ke dalam persamaan model regresi linear berganda.</div>
        </div>
      </div>

      <div style="display: flex; flex-direction: column; gap: 16px;">
        ${res.testPredictions.map(testItem => {
          return `
            <div class="prediction-card pred-rendah" style="border-color: #6366f1; padding: 16px;">
              <div class="prediction-header" style="margin-bottom: 12px;">
                <div>
                  <div class="pred-title" style="font-size: 15px;">
                    Data Testing No. ${testItem.id}
                  </div>
                  <div style="font-size: 12px; color: var(--gray-600); margin-top: 2px;">
                    Input: ${featureIdxs.map(fIdx => `${headers[fIdx]}: <strong>${testItem.rawRow[fIdx]}</strong>`).join(", ")}
                  </div>
                </div>

                <div>
                  <span class="badge-decision badge-normal" style="font-size: 14px; padding: 4px 14px;">
                    Prediksi Y: ${testItem.predY.toFixed(4)}
                  </span>
                </div>
              </div>

              <div class="formula-box" style="background: #f8fafc; color: var(--gray-800); border: 1px solid var(--gray-300); border-left: 4px solid #6366f1; font-size: 12px; padding: 10px 14px;">
                Y = ${testItem.termBreakdown.join(" + ")}<br>
                Y = <strong>${testItem.predY.toFixed(4)}</strong>
              </div>
            </div>
          `;
        }).join("")}
      </div>
    </div>

    <!-- Section 2: PERSAMAAN REGRESI -->
    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">Persamaan Model Regresi Linear</div>
          <div class="card-desc">Model dibentuk dengan metode Ordinary Least Squares (OLS): β = (X^T · X)^(-1) · X^T · Y</div>
        </div>
      </div>

      <div class="formula-box" style="margin-bottom: 16px;">
        <strong>${formulaStr}</strong>
        <div class="formula-note">Target Y: ${headers[targetIdx]}</div>
      </div>

      <div class="table-responsive">
        <table>
          <thead>
            <tr>
              <th>Variabel</th>
              <th>Simbol</th>
              <th class="text-right">Nilai Koefisien</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Konstanta (Intercept)</strong></td>
              <td class="font-mono">b0</td>
              <td class="text-right font-mono font-bold" style="color: var(--primary);">${res.beta[0].toFixed(6)}</td>
            </tr>
            ${featureIdxs.map((fIdx, i) => {
              const b = res.beta[i + 1];
              return `
                <tr>
                  <td><strong>${headers[fIdx]}</strong></td>
                  <td class="font-mono">b${i+1}</td>
                  <td class="text-right font-mono font-bold" style="color: ${b >= 0 ? 'var(--success-text)' : 'var(--danger-text)'};">${b.toFixed(6)}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;

  container.innerHTML = html;
}

/* ==========================================================================
   EXCEL EXPORT
   ========================================================================== */
function initExportButton() {
  const exportBtn = document.getElementById("exportExcelBtn");
  if (!exportBtn) return;

  exportBtn.addEventListener("click", () => {
    if (!state.analysisResults) {
      alert("Belum ada kalkulasi yang selesai.");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      const trainSheetData = [state.headers, ...state.dataTraining];
      const wsTrain = XLSX.utils.aoa_to_sheet(trainSheetData);
      XLSX.utils.book_append_sheet(wb, wsTrain, "data_training");

      const testSheetData = [state.headers];
      const preds = state.analysisResults.testPredictions;
      
      state.dataTesting.forEach((row, idx) => {
        const rowCopy = [...row];
        if (state.mode === "klasifikasi") {
          const predClass = state.smoothingMethod === "laplace" 
            ? preds[idx].bestClassLaplace 
            : preds[idx].bestClassPure;
          rowCopy[state.targetColIndex] = predClass;
        } else {
          rowCopy[state.targetColIndex] = Number(preds[idx].predY.toFixed(4));
        }
        testSheetData.push(rowCopy);
      });

      const wsTest = XLSX.utils.aoa_to_sheet(testSheetData);
      XLSX.utils.book_append_sheet(wb, wsTest, "data_testing");

      const hitungData = [];
      if (state.mode === "klasifikasi") {
        const res = state.analysisResults;
        hitungData.push(["LEMBAR HITUNG NAIVE BAYES"]);
        hitungData.push(["Total Data Training (N)", res.totalN]);
        hitungData.push([]);
        hitungData.push(["PRIOR PROBABILITY P(KELAS)"]);
        hitungData.push(["Kelas Target", "Jumlah Kasus", "Total N", "Prior P(C)", "Persentase"]);
        res.classes.forEach(c => {
          hitungData.push([c, res.classCounts[c], res.totalN, res.priors[c], `${(res.priors[c] * 100).toFixed(2)}%`]);
        });
        hitungData.push([]);
        hitungData.push(["HASIL PREDIKSI DATA TESTING"]);
        hitungData.push(["No.", ...state.featureIndices.map(f => state.headers[f]), "Skor Rendah", "Skor Tinggi", "Prob Rendah", "Prob Tinggi", "Prediksi Kelas"]);
        res.testPredictions.forEach(tp => {
          const cr = tp.classResults;
          const isLp = state.smoothingMethod === "laplace";
          const rScore = isLp ? cr["Rendah"]?.laplaceScore : cr["Rendah"]?.pureScore;
          const tScore = isLp ? cr["Tinggi"]?.laplaceScore : cr["Tinggi"]?.pureScore;
          const rProb = isLp ? cr["Rendah"]?.normProbLaplace : cr["Rendah"]?.normProbPure;
          const tProb = isLp ? cr["Tinggi"]?.normProbLaplace : cr["Tinggi"]?.normProbPure;
          const finalClass = isLp ? tp.bestClassLaplace : tp.bestClassPure;

          hitungData.push([
            tp.id,
            ...state.featureIndices.map(f => tp.rawRow[f]),
            rScore || 0,
            tScore || 0,
            `${((rProb || 0) * 100).toFixed(2)}%`,
            `${((tProb || 0) * 100).toFixed(2)}%`,
            finalClass
          ]);
        });
      }

      const wsHitung = XLSX.utils.aoa_to_sheet(hitungData);
      XLSX.utils.book_append_sheet(wb, wsHitung, "hitung");

      const exportFileName = `HASIL_${state.currentDatasetName.replace('.xlsx', '')}_SELESAI.xlsx`;
      XLSX.writeFile(wb, exportFileName);

    } catch (e) {
      console.error(e);
      alert("Gagal mengunduh Excel: " + e.message);
    }
  });
}

/* ==========================================================================
   LINEAR ALGEBRA MATRIX HELPERS
   ========================================================================= */
function transposeMatrix(A) {
  const m = A.length;
  const n = A[0].length;
  const AT = [];
  for (let j = 0; j < n; j++) {
    AT[j] = [];
    for (let i = 0; i < m; i++) {
      AT[j][i] = A[i][j];
    }
  }
  return AT;
}

function multiplyMatrices(A, B) {
  const rowsA = A.length;
  const colsA = A[0].length;
  const colsB = B[0].length;
  const C = [];
  for (let i = 0; i < rowsA; i++) {
    C[i] = [];
    for (let j = 0; j < colsB; j++) {
      let sum = 0;
      for (let k = 0; k < colsA; k++) {
        sum += A[i][k] * B[k][j];
      }
      C[i][j] = sum;
    }
  }
  return C;
}

function multiplyMatrixVector(A, v) {
  const m = A.length;
  const n = A[0].length;
  const result = [];
  for (let i = 0; i < m; i++) {
    let sum = 0;
    for (let j = 0; j < n; j++) {
      sum += A[i][j] * v[j];
    }
    result.push(sum);
  }
  return result;
}

function invertMatrix(M) {
  const n = M.length;
  const A = [];
  for (let i = 0; i < n; i++) {
    A[i] = [];
    for (let j = 0; j < n; j++) {
      A[i][j] = M[i][j];
    }
    for (let j = 0; j < n; j++) {
      A[i][n + j] = (i === j) ? 1 : 0;
    }
  }

  for (let i = 0; i < n; i++) {
    let maxRow = i;
    for (let k = i + 1; k < n; k++) {
      if (Math.abs(A[k][i]) > Math.abs(A[maxRow][i])) {
        maxRow = k;
      }
    }

    if (Math.abs(A[maxRow][i]) < 1e-12) {
      return null;
    }

    const temp = A[i];
    A[i] = A[maxRow];
    A[maxRow] = temp;

    const pivot = A[i][i];
    for (let j = 0; j < 2 * n; j++) {
      A[i][j] /= pivot;
    }

    for (let k = 0; k < n; k++) {
      if (k !== i) {
        const factor = A[k][i];
        for (let j = 0; j < 2 * n; j++) {
          A[k][j] -= factor * A[i][j];
        }
      }
    }
  }

  const inv = [];
  for (let i = 0; i < n; i++) {
    inv[i] = [];
    for (let j = 0; j < n; j++) {
      inv[i][j] = A[i][n + j];
    }
  }
  return inv;
}
