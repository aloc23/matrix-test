document.addEventListener('DOMContentLoaded', function() {
  // -------------------- State --------------------
  let rawData = [];
  let mappedData = [];
  let mappingConfigured = false;
  let config = {
    weekLabelRow: 0,
    weekColStart: 0,
    weekColEnd: 0,
    firstDataRow: 1,
    lastDataRow: 1
  };
  let weekLabels = [];
  let weekCheckboxStates = [];
  let repaymentRows = [];
  let openingBalance = 0;
  let loanOutstanding = 0;
  let roiInvestment = 120000;
  let roiInterest = 0.0;

  // ROI week/date mapping
  let weekStartDates = [];
  let investmentWeekIndex = 0;

  // --- Chart.js chart instances for destroy ---
  let mainChart = null;
  let roiPieChart = null;
  let roiLineChart = null;
  window.tornadoChartObj = null; // global for tornado chart
  let summaryChart = null;

  // -------------------- Tabs & UI Interactions --------------------
  function setupTabs() {
    document.querySelectorAll('.tabs button').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.tabs button').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        document.querySelectorAll('.tab-content').forEach(sec => sec.classList.remove('active'));
        const tabId = btn.getAttribute('data-tab');
        const panel = document.getElementById(tabId);
        if (panel) panel.classList.add('active');
        setTimeout(() => {
          updateAllTabs();
        }, 50);
      });
    });
    document.querySelectorAll('.subtabs button').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.subtabs button').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        document.querySelectorAll('.subtab-panel').forEach(sec => sec.classList.remove('active'));
        const subtabId = 'subtab-' + btn.getAttribute('data-subtab');
        const subpanel = document.getElementById(subtabId);
        if (subpanel) subpanel.classList.add('active');
        setTimeout(updateAllTabs, 50);
      });
    });
    document.querySelectorAll('.collapsible-header').forEach(btn => {
      btn.addEventListener('click', function() {
        const content = btn.nextElementSibling;
        const caret = btn.querySelector('.caret');
        if (content && content.classList.contains('active')) {
          content.classList.remove('active');
          if (caret) caret.style.transform = 'rotate(-90deg)';
        } else if (content) {
          content.classList.add('active');
          if (caret) caret.style.transform = 'none';
        }
      });
    });
  }
  setupTabs();

  // -------------------- Spreadsheet Upload & Mapping --------------------
  function setupSpreadsheetUpload() {
    var spreadsheetUpload = document.getElementById('spreadsheetUpload');
    if (spreadsheetUpload) {
      spreadsheetUpload.addEventListener('change', function(event) {
        const reader = new FileReader();
        reader.onload = function (e) {
          const dataArr = new Uint8Array(e.target.result);
          const workbook = XLSX.read(dataArr, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          if (!json.length) return;
          rawData = json;
          mappedData = json;
          autoDetectMapping(mappedData);
          mappingConfigured = false;
          renderMappingPanel(mappedData);
          updateWeekLabels();
          updateAllTabs();
        };
        reader.readAsArrayBuffer(event.target.files[0]);
      });
    }
  }
  setupSpreadsheetUpload();

  function renderMappingPanel(allRows) {
    const panel = document.getElementById('mappingPanel');
    if (!panel) return;
    panel.innerHTML = '';

    function drop(label, id, max, sel, onChange, items) {
      let lab = document.createElement('label');
      lab.textContent = label;
      let selElem = document.createElement('select');
      selElem.className = 'mapping-dropdown';
      for (let i = 0; i < max; i++) {
        let opt = document.createElement('option');
        opt.value = i;
        let textVal = items && items[i] ? items[i] : (allRows[i] ? allRows[i].slice(0,8).join(',').slice(0,32) : '');
        opt.textContent = `${id==='row'?'Row':'Col'} ${i+1}: ${textVal}`;
        selElem.appendChild(opt);
      }
      selElem.value = sel;
      selElem.onchange = function() { onChange(parseInt(this.value,10)); };
      lab.appendChild(selElem);
      panel.appendChild(lab);
    }

    drop('Which row contains week labels? ', 'row', Math.min(allRows.length, 30), config.weekLabelRow, v => { config.weekLabelRow = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); });
    panel.appendChild(document.createElement('br'));

    let weekRow = allRows[config.weekLabelRow] || [];
    drop('First week column: ', 'col', weekRow.length, config.weekColStart, v => { config.weekColStart = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); }, weekRow);
    drop('Last week column: ', 'col', weekRow.length, config.weekColEnd, v => { config.weekColEnd = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); }, weekRow);
    panel.appendChild(document.createElement('br'));

    drop('First data row: ', 'row', allRows.length, config.firstDataRow, v => { config.firstDataRow = v; renderMappingPanel(allRows); updateAllTabs(); });
    drop('Last data row: ', 'row', allRows.length, config.lastDataRow, v => { config.lastDataRow = v; renderMappingPanel(allRows); updateAllTabs(); });
    panel.appendChild(document.createElement('br'));

    // Opening balance input
    let obDiv = document.createElement('div');
    obDiv.innerHTML = `Opening Balance: <input type="number" id="openingBalanceInput" value="${openingBalance}" style="width:120px;">`;
    panel.appendChild(obDiv);
    setTimeout(() => {
      let obInput = document.getElementById('openingBalanceInput');
      if (obInput) obInput.oninput = function() {
        openingBalance = parseFloat(obInput.value) || 0;
        updateAllTabs();
        renderMappingPanel(allRows);
      };
    }, 0);

    // Reset button for mapping
    const resetBtn = document.createElement('button');
    resetBtn.textContent = "Reset Mapping";
    resetBtn.style.marginLeft = '10px';
    resetBtn.onclick = function() {
      autoDetectMapping(allRows);
      weekCheckboxStates = weekLabels.map(()=>true);
      openingBalance = 0;
      renderMappingPanel(allRows);
      updateWeekLabels();
      updateAllTabs();
    };
    panel.appendChild(resetBtn);

    // Collapsible Week filter UI
    if (weekLabels.length) {
      const weekFilterDiv = document.createElement('div');
      weekFilterDiv.className = "collapsible-week-filter";

      // Collapsible header
      const collapseBtn = document.createElement('button');
      collapseBtn.type = 'button';
      collapseBtn.className = 'collapse-toggle';
      collapseBtn.innerHTML = `<span class="caret" style="display:inline-block;transition:transform 0.2s;margin-right:6px;">&#9654;</span>Filter week columns to include:`;
      collapseBtn.style.marginBottom = '10px';
      collapseBtn.style.background = 'none';
      collapseBtn.style.color = '#1976d2';
      collapseBtn.style.fontWeight = 'bold';
      collapseBtn.style.fontSize = '1.06em';
      collapseBtn.style.border = 'none';
      collapseBtn.style.cursor = 'pointer';
      collapseBtn.style.outline = 'none';
      collapseBtn.style.padding = '4px 0';

      // Collapsible content
      const collapsibleContent = document.createElement('div');
      collapsibleContent.className = "week-checkbox-collapsible-content";
      collapsibleContent.style.display = 'none';
      collapsibleContent.style.margin = '14px 0 4px 0';

      // Buttons
      const selectAllBtn = document.createElement('button');
      selectAllBtn.textContent = "Select All";
      selectAllBtn.type = 'button';
      selectAllBtn.style.marginRight = '8px';
      selectAllBtn.onclick = function() {
        weekCheckboxStates = weekCheckboxStates.map(()=>true);
        updateAllTabs();
        renderMappingPanel(allRows);
      };
      const deselectAllBtn = document.createElement('button');
      deselectAllBtn.textContent = "Deselect All";
      deselectAllBtn.type = 'button';
      deselectAllBtn.onclick = function() {
        weekCheckboxStates = weekCheckboxStates.map(()=>false);
        updateAllTabs();
        renderMappingPanel(allRows);
      };
      collapsibleContent.appendChild(selectAllBtn);
      collapsibleContent.appendChild(deselectAllBtn);

      // Checkbox group
      const groupDiv = document.createElement('div');
      groupDiv.className = 'week-checkbox-group';
      groupDiv.style.marginTop = '8px';
      weekLabels.forEach((label, idx) => {
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = weekCheckboxStates[idx] !== false;
        cb.id = 'weekcol_cb_' + idx;
        cb.onchange = function() {
          weekCheckboxStates[idx] = cb.checked;
          updateAllTabs();
          renderMappingPanel(allRows);
        };
        const lab = document.createElement('label');
        lab.htmlFor = cb.id;
        lab.textContent = label;
        lab.style.marginRight = '13px';
        groupDiv.appendChild(cb);
        groupDiv.appendChild(lab);
      });
      collapsibleContent.appendChild(groupDiv);

      // Collapsible logic
      collapseBtn.addEventListener('click', function() {
        const isOpen = collapsibleContent.style.display !== 'none';
        collapsibleContent.style.display = isOpen ? 'none' : 'block';
        const caret = collapseBtn.querySelector('.caret');
        caret.style.transform = isOpen ? 'rotate(0)' : 'rotate(90deg)';
      });

      weekFilterDiv.appendChild(collapseBtn);
      weekFilterDiv.appendChild(collapsibleContent);
      panel.appendChild(weekFilterDiv);
    }

    // Save Mapping Button
    const saveBtn = document.createElement('button');
    saveBtn.textContent = "Save Mapping";
    saveBtn.style.margin = "10px 0";
    saveBtn.onclick = function() {
      mappingConfigured = true;
      updateWeekLabels();
      updateAllTabs();
      renderMappingPanel(allRows);
    };
    panel.appendChild(saveBtn);

    // Compact preview
    if (weekLabels.length && mappingConfigured) {
      const previewWrap = document.createElement('div');
      const compactTable = document.createElement('table');
      compactTable.className = "compact-preview-table";
      const tr1 = document.createElement('tr');
      tr1.appendChild(document.createElement('th'));
      getFilteredWeekIndices().forEach(fi => {
        const th = document.createElement('th');
        th.textContent = weekLabels[fi];
        tr1.appendChild(th);
      });
      compactTable.appendChild(tr1);
      const tr2 = document.createElement('tr');
      const lbl = document.createElement('td');
      lbl.textContent = "Bank Balance (rolling)";
      tr2.appendChild(lbl);
      let rolling = getRollingBankBalanceArr();
      getFilteredWeekIndices().forEach((fi, i) => {
        let bal = rolling[i];
        let td = document.createElement('td');
        td.textContent = isNaN(bal) ? '' : `€${Math.round(bal)}`;
        if (bal < 0) td.style.background = "#ffeaea";
        tr2.appendChild(td);
      });
      compactTable.appendChild(tr2);
      previewWrap.style.overflowX = "auto";
      previewWrap.appendChild(compactTable);
      panel.appendChild(previewWrap);
    }
  }

  function autoDetectMapping(sheet) {
    for (let r = 0; r < Math.min(sheet.length, 10); r++) {
      for (let c = 0; c < Math.min(sheet[r].length, 30); c++) {
        const val = (sheet[r][c] || '').toString().toLowerCase();
        if (/week\s*\d+/.test(val) || /week\s*\d+\/\d+/.test(val)) {
          config.weekLabelRow = r;
          config.weekColStart = c;
          let lastCol = c;
          while (
            lastCol < sheet[r].length &&
            ((sheet[r][lastCol] || '').toLowerCase().indexOf('week') >= 0 ||
            /^\d{1,2}\/\d{1,2}/.test(sheet[r][lastCol] || ''))
          ) {
            lastCol++;
          }
          config.weekColEnd = lastCol - 1;
          config.firstDataRow = r + 1;
          config.lastDataRow = sheet.length-1;
          return;
        }
      }
    }
    config.weekLabelRow = 4;
    config.weekColStart = 5;
    config.weekColEnd = Math.max(5, (sheet[4]||[]).length-1);
    config.firstDataRow = 6;
    config.lastDataRow = sheet.length-1;
  }

  function extractWeekStartDates(weekLabels, baseYear) {
    let currentYear = baseYear;
    let lastMonthIdx = -1;
    const months = [
      "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"
    ];
    return weekLabels.map(label => {
      let match = label.match(/(\d{1,2})\s*([A-Za-z]{3,})/);
      if (!match) return null;
      let [_, day, monthStr] = match;
      let monthIdx = months.findIndex(m =>
        monthStr.toLowerCase().startsWith(m)
      );
      if (monthIdx === -1) return null;
      if (lastMonthIdx !== -1 && monthIdx < lastMonthIdx) currentYear++;
      lastMonthIdx = monthIdx;
      let date = new Date(currentYear, monthIdx, parseInt(day, 10));
      return date;
    });
  }

  function populateInvestmentWeekDropdown() {
    const dropdown = document.getElementById('investmentWeek');
    if (!dropdown) return;
    dropdown.innerHTML = '';
    weekLabels.forEach((label, i) => {
      const opt = document.createElement('option');
      let dateStr = weekStartDates[i]
        ? weekStartDates[i].toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
        : 'N/A';
      opt.value = i;
      opt.textContent = `${label} (${dateStr})`;
      dropdown.appendChild(opt);
    });
    dropdown.value = investmentWeekIndex;
  }

  function updateWeekLabels() {
    let weekRow = mappedData[config.weekLabelRow] || [];
    weekLabels = weekRow.slice(config.weekColStart, config.weekColEnd+1).map(x => x || '');
    window.weekLabels = weekLabels; // make global for charts
    if (!weekCheckboxStates || weekCheckboxStates.length !== weekLabels.length) {
      // Preserve previous checkbox states by matching labels
      const prevLabels = window.weekLabels || [];
      const prevStates = weekCheckboxStates || [];
      weekCheckboxStates = weekLabels.map((label, idx) => {
        const prevIdx = prevLabels.indexOf(label);
        return prevIdx !== -1 ? prevStates[prevIdx] : true;
      });
    }
    populateWeekDropdown(weekLabels);

    // ROI week start date integration. Use a default base year (2025) or prompt user for year.
    weekStartDates = extractWeekStartDates(weekLabels, 2025);
    populateInvestmentWeekDropdown();
  }

  function getFilteredWeekIndices() {
    return weekCheckboxStates.map((checked, idx) => checked ? idx : null).filter(idx => idx !== null);
  }

  // -------------------- Calculation Helpers --------------------
  function getIncomeArr() {
    if (!mappedData || !mappingConfigured) return [];
    let arr = [];
    for (let w = 0; w < weekLabels.length; w++) {
      if (!weekCheckboxStates[w]) continue;
      let absCol = config.weekColStart + w;
      let sum = 0;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let val = mappedData[r][absCol];
        if (typeof val === "string") val = val.replace(/,/g, '').replace(/€|\s/g,'');
        let num = parseFloat(val);
        if (!isNaN(num) && num > 0) sum += num;
      }
      arr[w] = sum;
    }
    return arr;
  }
  function getExpenditureArr() {
    if (!mappedData || !mappingConfigured) return [];
    let arr = [];
    for (let w = 0; w < weekLabels.length; w++) {
      if (!weekCheckboxStates[w]) continue;
      let absCol = config.weekColStart + w;
      let sum = 0;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let val = mappedData[r][absCol];
        if (typeof val === "string") val = val.replace(/,/g, '').replace(/€|\s/g,'');
        let num = parseFloat(val);
        if (!isNaN(num) && num < 0) sum += Math.abs(num);
      }
      arr[w] = sum;
    }
    return arr;
  }
  function getRepaymentArr() {
    if (!mappingConfigured || !weekLabels.length) return [];
    let arr = Array(weekLabels.length).fill(0);
    repaymentRows.forEach(r => {
      if (r.type === "week") {
        let weekIdx = weekLabels.indexOf(r.week);
        if (weekIdx === -1) weekIdx = 0;
        arr[weekIdx] += r.amount;
      } else {
        if (r.frequency === "monthly") {
          let perMonth = Math.ceil(arr.length/12);
          for (let m=0; m<12; m++) {
            for (let w=m*perMonth; w<(m+1)*perMonth && w<arr.length; w++) arr[w] += r.amount;
          }
        }
        if (r.frequency === "quarterly") {
          let perQuarter = Math.ceil(arr.length/4);
          for (let q=0;q<4;q++) {
            for (let w=q*perQuarter; w<(q+1)*perQuarter && w<arr.length; w++) arr[w] += r.amount;
          }
        }
        if (r.frequency === "one-off") { arr[0] += r.amount; }
      }
    });
    return getFilteredWeekIndices().map(idx => arr[idx]);
  }
  function getNetProfitArr(incomeArr, expenditureArr, repaymentArr) {
    return incomeArr.map((inc, i) => (inc || 0) - (expenditureArr[i] || 0) - (repaymentArr[i] || 0));
  }
  function getRollingBankBalanceArr() {
    let incomeArr = getIncomeArr();
    let expenditureArr = getExpenditureArr();
    let repaymentArr = getRepaymentArr();
    let rolling = [];
    let ob = openingBalance;
    getFilteredWeekIndices().forEach((fi, i) => {
      let income = incomeArr[fi] || 0;
      let out = expenditureArr[fi] || 0;
      let repay = repaymentArr[i] || 0;
      let prev = (i === 0 ? ob : rolling[i-1]);
      rolling[i] = prev + income - out - repay;
    });
    return rolling;
  }
  function getMonthAgg(arr, months=12) {
    let filtered = arr.filter((_,i)=>getFilteredWeekIndices().includes(i));
    let perMonth = Math.ceil(filtered.length/months);
    let out = [];
    for(let m=0;m<months;m++) {
      let sum=0;
      for(let w=m*perMonth;w<(m+1)*perMonth && w<filtered.length;w++) sum += filtered[w];
      out.push(sum);
    }
    return out;
  }

  // -------------------- Repayments UI --------------------
  const weekSelect = document.getElementById('weekSelect');
  const repaymentFrequency = document.getElementById('repaymentFrequency');
  function populateWeekDropdown(labels) {
    if (!weekSelect) return;
    weekSelect.innerHTML = '';
    (labels && labels.length ? labels : Array.from({length: 52}, (_, i) => `Week ${i+1}`)).forEach(label => {
      const opt = document.createElement('option');
      opt.value = label;
      opt.textContent = label;
      weekSelect.appendChild(opt);
    });
  }

  function setupRepaymentForm() {
    if (!weekSelect || !repaymentFrequency) return;
    document.querySelectorAll('input[name="repaymentType"]').forEach(radio => {
      radio.addEventListener('change', function() {
        if (this.value === "week") {
          weekSelect.disabled = false;
          repaymentFrequency.disabled = true;
        } else {
          weekSelect.disabled = true;
          repaymentFrequency.disabled = false;
        }
      });
    });

    let addRepaymentForm = document.getElementById('addRepaymentForm');
    if (addRepaymentForm) {
      addRepaymentForm.onsubmit = function(e) {
        e.preventDefault();
        const type = document.querySelector('input[name="repaymentType"]:checked').value;
        let week = null, frequency = null;
        if (type === "week") {
          week = weekSelect.value;
        } else {
          frequency = repaymentFrequency.value;
        }
        const amount = document.getElementById('repaymentAmount').value;
        if (!amount) return;
        repaymentRows.push({ type, week, frequency, amount: parseFloat(amount), editing: false });
        renderRepaymentRows();
        this.reset();
        populateWeekDropdown(weekLabels);
        document.getElementById('weekSelect').selectedIndex = 0;
        document.getElementById('repaymentFrequency').selectedIndex = 0;
        document.querySelector('input[name="repaymentType"][value="week"]').checked = true;
        weekSelect.disabled = false;
        repaymentFrequency.disabled = true;
        updateAllTabs();
      };
    }
  }
  setupRepaymentForm();

  function renderRepaymentRows() {
    const container = document.getElementById('repaymentRows');
    if (!container) return;
    container.innerHTML = "";
    repaymentRows.forEach((row, i) => {
      const div = document.createElement('div');
      div.className = 'repayment-row';
      const weekSelectElem = document.createElement('select');
      (weekLabels.length ? weekLabels : Array.from({length:52}, (_,i)=>`Week ${i+1}`)).forEach(label => {
        const opt = document.createElement('option');
        opt.value = label;
        opt.textContent = label;
        weekSelectElem.appendChild(opt);
      });
      weekSelectElem.value = row.week || "";
      weekSelectElem.disabled = !row.editing || row.type !== "week";

      const freqSelect = document.createElement('select');
      ["monthly", "quarterly", "one-off"].forEach(f => {
        const opt = document.createElement('option');
        opt.value = f;
        opt.textContent = f.charAt(0).toUpperCase() + f.slice(1);
        freqSelect.appendChild(opt);
      });
      freqSelect.value = row.frequency || "monthly";
      freqSelect.disabled = !row.editing || row.type !== "frequency";

      const amountInput = document.createElement('input');
      amountInput.type = 'number';
      amountInput.value = row.amount;
      amountInput.placeholder = 'Repayment €';
      amountInput.disabled = !row.editing;

      const editBtn = document.createElement('button');
      editBtn.textContent = row.editing ? 'Save' : 'Edit';
      editBtn.onclick = function() {
        if (row.editing) {
          if (row.type === "week") {
            row.week = weekSelectElem.value;
          } else {
            row.frequency = freqSelect.value;
          }
          row.amount = parseFloat(amountInput.value);
        }
        row.editing = !row.editing;
        renderRepaymentRows();
        updateAllTabs();
      };

      const removeBtn = document.createElement('button');
      removeBtn.textContent = 'Remove';
      removeBtn.onclick = function() {
        repaymentRows.splice(i, 1);
        renderRepaymentRows();
        updateAllTabs();
      };

      const modeLabel = document.createElement('span');
      modeLabel.style.marginRight = "10px";
      modeLabel.textContent = row.type === "week" ? "Week" : "Frequency";

      if (row.type === "week") {
        div.appendChild(modeLabel);
        div.appendChild(weekSelectElem);
      } else {
        div.appendChild(modeLabel);
        div.appendChild(freqSelect);
      }
      div.appendChild(amountInput);
      div.appendChild(editBtn);
      div.appendChild(removeBtn);

      container.appendChild(div);
    });
  }

  function updateLoanSummary() {
    const totalRepaid = getRepaymentArr().reduce((a,b)=>a+b,0);
    let totalRepaidBox = document.getElementById('totalRepaidBox');
    let remainingBox = document.getElementById('remainingBox');
    if (totalRepaidBox) totalRepaidBox.textContent = "Total Repaid: €" + totalRepaid.toLocaleString();
    if (remainingBox) remainingBox.textContent = "Remaining: €" + (loanOutstanding-totalRepaid).toLocaleString();
  }
  let loanOutstandingInput = document.getElementById('loanOutstandingInput');
  if (loanOutstandingInput) {
    loanOutstandingInput.oninput = function() {
      loanOutstanding = parseFloat(this.value) || 0;
      updateLoanSummary();
    };
  }

  // -------------------- Main Chart & Summary --------------------
  function updateChartAndSummary() {
    let mainChartElem = document.getElementById('mainChart');
    let mainChartSummaryElem = document.getElementById('mainChartSummary');
    let mainChartNoDataElem = document.getElementById('mainChartNoData');
    if (!mainChartElem || !mainChartSummaryElem || !mainChartNoDataElem) return;

    if (!mappingConfigured || !weekLabels.length || getFilteredWeekIndices().length === 0) {
      if (mainChartNoDataElem) mainChartNoDataElem.style.display = "";
      if (mainChartSummaryElem) mainChartSummaryElem.innerHTML = "";
      if (mainChart && typeof mainChart.destroy === "function") mainChart.destroy();
      return;
    } else {
      if (mainChartNoDataElem) mainChartNoDataElem.style.display = "none";
    }

    const filteredWeeks = getFilteredWeekIndices();
    const labels = filteredWeeks.map(idx => weekLabels[idx]);
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const repaymentArr = getRepaymentArr();
    const rollingArr = getRollingBankBalanceArr();
    const netProfitArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);

    const data = {
      labels: labels,
      datasets: [
        {
          label: "Income",
          data: filteredWeeks.map(idx => incomeArr[idx] || 0),
          backgroundColor: "rgba(76,175,80,0.6)",
          borderColor: "#388e3c",
          fill: false,
          type: "bar"
        },
        {
          label: "Expenditure",
          data: filteredWeeks.map(idx => expenditureArr[idx] || 0),
          backgroundColor: "rgba(244,67,54,0.6)",
          borderColor: "#c62828",
          fill: false,
          type: "bar"
        },
        {
          label: "Repayment",
          data: filteredWeeks.map((_, i) => repaymentArr[i] || 0),
          backgroundColor: "rgba(255,193,7,0.6)",
          borderColor: "#ff9800",
          fill: false,
          type: "bar"
        },
        {
          label: "Net Profit",
          data: filteredWeeks.map((_, i) => netProfitArr[i] || 0),
          backgroundColor: "rgba(33,150,243,0.3)",
          borderColor: "#1976d2",
          type: "line",
          fill: false,
          yAxisID: "y"
        },
        {
          label: "Rolling Bank Balance",
          data: rollingArr,
          backgroundColor: "rgba(156,39,176,0.2)",
          borderColor: "#8e24aa",
          type: "line",
          fill: true,
          yAxisID: "y"
        }
      ]
    };

    if (mainChart && typeof mainChart.destroy === "function") mainChart.destroy();

    mainChart = new Chart(mainChartElem.getContext('2d'), {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        plugins: {
          legend: { display: true },
          tooltip: { mode: "index", intersect: false }
        },
        scales: {
          x: { stacked: true },
          y: {
            beginAtZero: true,
            title: { display: true, text: "€" }
          }
        }
      }
    });

    let totalIncome = incomeArr.reduce((a,b)=>a+(b||0), 0);
    let totalExpenditure = expenditureArr.reduce((a,b)=>a+(b||0), 0);
    let totalRepayment = repaymentArr.reduce((a,b)=>a+(b||0), 0);
    let finalBalance = rollingArr[rollingArr.length - 1] || 0;
    let lowestBalance = Math.min(...rollingArr);

    mainChartSummaryElem.innerHTML = `
      <b>Total Income:</b> €${Math.round(totalIncome).toLocaleString()}<br>
      <b>Total Expenditure:</b> €${Math.round(totalExpenditure).toLocaleString()}<br>
      <b>Total Repayments:</b> €${Math.round(totalRepayment).toLocaleString()}<br>
      <b>Final Bank Balance:</b> <span style="color:${finalBalance<0?'#c00':'#388e3c'}">€${Math.round(finalBalance).toLocaleString()}</span><br>
      <b>Lowest Bank Balance:</b> <span style="color:${lowestBalance<0?'#c00':'#388e3c'}">€${Math.round(lowestBalance).toLocaleString()}</span>
    `;
  }

  // ---------- P&L Tab Functions ----------
  function renderSectionSummary(headerId, text, arr) {
    const headerElem = document.getElementById(headerId);
    if (!headerElem) return;
    headerElem.innerHTML = text;
  }
  function renderPnlTables() {
  // Weekly Breakdown
  const weeklyTable = document.getElementById('pnlWeeklyBreakdown');
  const monthlyTable = document.getElementById('pnlMonthlyBreakdown');
  const cashFlowTable = document.getElementById('pnlCashFlow');
  const pnlSummary = document.getElementById('pnlSummary');
  if (!weeklyTable || !monthlyTable || !cashFlowTable) return;

  // ---- Weekly table ----
  let tbody = weeklyTable.querySelector('tbody');
  if (tbody) tbody.innerHTML = '';
  let incomeArr = getIncomeArr();
  let expenditureArr = getExpenditureArr();
  let repaymentArr = getRepaymentArr();
  let rollingArr = getRollingBankBalanceArr();
  let netArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);
  let weekIdxs = getFilteredWeekIndices();
  let rows = '';
  let minBal = null, minBalWeek = null;

  weekIdxs.forEach((idx, i) => {
    const net = (incomeArr[idx] || 0) - (expenditureArr[idx] || 0) - (repaymentArr[i] || 0);
    const netTooltip = `Income - Expenditure - Repayment\n${incomeArr[idx]||0} - ${expenditureArr[idx]||0} - ${repaymentArr[i]||0} = ${net}`;
    const balTooltip = `Prev Bal + Income - Expenditure - Repayment\n${i===0?openingBalance:rollingArr[i-1]} + ${incomeArr[idx]||0} - ${expenditureArr[idx]||0} - ${repaymentArr[i]||0} = ${rollingArr[i]||0}`;
    let row = `<tr${rollingArr[i]<0?' class="negative-balance-row"':''}>` +
      `<td>${weekLabels[idx]}</td>` +
      `<td${incomeArr[idx]<0?' class="negative-number"':''}>€${Math.round(incomeArr[idx]||0).toLocaleString()}</td>` +
      `<td${expenditureArr[idx]<0?' class="negative-number"':''}>€${Math.round(expenditureArr[idx]||0).toLocaleString()}</td>` +
      `<td${repaymentArr[i]<0?' class="negative-number"':''}>€${Math.round(repaymentArr[i]||0).toLocaleString()}</td>` +
      `<td class="${net<0?'negative-number':''}" data-tooltip="${netTooltip}">€${Math.round(net||0).toLocaleString()}</td>` +
      `<td${rollingArr[i]<0?' class="negative-number"':''} data-tooltip="${balTooltip}">€${Math.round(rollingArr[i]||0).toLocaleString()}</td></tr>`;
    rows += row;
    if (minBal===null||rollingArr[i]<minBal) {minBal=rollingArr[i];minBalWeek=weekLabels[idx];}
  });
  if (tbody) tbody.innerHTML = rows;
  renderSectionSummary('weekly-breakdown-header', `Total Net: €${netArr.reduce((a,b)=>a+(b||0),0).toLocaleString()}`, netArr);

  // ---- Monthly Breakdown ----
  let months = 12;
  let incomeMonth = getMonthAgg(incomeArr, months);
  let expMonth = getMonthAgg(expenditureArr, months);
  let repayMonth = getMonthAgg(repaymentArr, months);
  let netMonth = incomeMonth.map((inc, i) => inc - (expMonth[i]||0) - (repayMonth[i]||0));
  let mtbody = monthlyTable.querySelector('tbody');
  if (mtbody) {
    mtbody.innerHTML = '';
    for (let m=0; m<months; m++) {
      const netTooltip = `Income - Expenditure - Repayment\n${incomeMonth[m]||0} - ${expMonth[m]||0} - ${repayMonth[m]||0} = ${netMonth[m]||0}`;
      mtbody.innerHTML += `<tr>
        <td>Month ${m+1}</td>
        <td${incomeMonth[m]<0?' class="negative-number"':''}>€${Math.round(incomeMonth[m]||0).toLocaleString()}</td>
        <td${expMonth[m]<0?' class="negative-number"':''}>€${Math.round(expMonth[m]||0).toLocaleString()}</td>
        <td class="${netMonth[m]<0?'negative-number':''}" data-tooltip="${netTooltip}">€${Math.round(netMonth[m]||0).toLocaleString()}</td>
        <td${repayMonth[m]<0?' class="negative-number"':''}>€${Math.round(repayMonth[m]||0).toLocaleString()}</td>
      </tr>`;
    }
  }
  renderSectionSummary('monthly-breakdown-header', `Total Net: €${netMonth.reduce((a,b)=>a+(b||0),0).toLocaleString()}`, netMonth);

  // ---- Cash Flow Table ----
  let ctbody = cashFlowTable.querySelector('tbody');
  let closingArr = [];
  if (ctbody) {
    ctbody.innerHTML = '';
    let opening = openingBalance;
    let closing = opening;
    for (let m=0; m<months; m++) {
      let inflow = incomeMonth[m] || 0;
      let outflow = (expMonth[m] || 0) + (repayMonth[m] || 0);
      closing = opening + inflow - outflow;
      closingArr.push(closing);
      const closingTooltip = `Opening + Inflow - Outflow\n${opening} + ${inflow} - ${outflow} = ${closing}`;
      ctbody.innerHTML += `<tr>
        <td>Month ${m+1}</td>
        <td>€${Math.round(opening).toLocaleString()}</td>
        <td>€${Math.round(inflow).toLocaleString()}</td>
        <td>€${Math.round(outflow).toLocaleString()}</td>
        <td${closing<0?' class="negative-number"':''} data-tooltip="${closingTooltip}">€${Math.round(closing).toLocaleString()}</td>
      </tr>`;
      opening = closing;
    }
  }
  renderSectionSummary('cashflow-header', `Closing Bal: €${Math.round(closingArr[closingArr.length-1]||0).toLocaleString()}`, closingArr);

  // ---- P&L Summary ----
  if (pnlSummary) {
    pnlSummary.innerHTML = `
      <b>Total Income:</b> €${Math.round(incomeArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Total Expenditure:</b> €${Math.round(expenditureArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Total Repayments:</b> €${Math.round(repaymentArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Final Bank Balance:</b> <span style="color:${rollingArr[rollingArr.length-1]<0?'#c00':'#388e3c'}">€${Math.round(rollingArr[rollingArr.length-1]||0).toLocaleString()}</span><br>
      <b>Lowest Bank Balance:</b> <span style="color:${minBal<0?'#c00':'#388e3c'}">${minBalWeek?minBalWeek+': ':''}€${Math.round(minBal||0).toLocaleString()}</span>
    `;
  }
}
  // ---------- Summary Tab Functions ----------
  function renderSummaryTab() {
    // Key Financials
    let incomeArr = getIncomeArr();
    let expenditureArr = getExpenditureArr();
    let repaymentArr = getRepaymentArr();
    let rollingArr = getRollingBankBalanceArr();
    let netArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);
    let totalIncome = incomeArr.reduce((a,b)=>a+(b||0),0);
    let totalExpenditure = expenditureArr.reduce((a,b)=>a+(b||0),0);
    let totalRepayment = repaymentArr.reduce((a,b)=>a+(b||0),0);
    let finalBal = rollingArr[rollingArr.length-1]||0;
    let minBal = Math.min(...rollingArr);

    // Update KPI cards if present
    if (document.getElementById('kpiTotalIncome')) {
      document.getElementById('kpiTotalIncome').textContent = '€' + totalIncome.toLocaleString();
      document.getElementById('kpiTotalExpenditure').textContent = '€' + totalExpenditure.toLocaleString();
      document.getElementById('kpiTotalRepayments').textContent = '€' + totalRepayment.toLocaleString();
      document.getElementById('kpiFinalBank').textContent = '€' + Math.round(finalBal).toLocaleString();
      document.getElementById('kpiLowestBank').textContent = '€' + Math.round(minBal).toLocaleString();
    }

    let summaryElem = document.getElementById('summaryKeyFinancials');
    if (summaryElem) {
      summaryElem.innerHTML = `
        <b>Total Income:</b> €${Math.round(totalIncome).toLocaleString()}<br>
        <b>Total Expenditure:</b> €${Math.round(totalExpenditure).toLocaleString()}<br>
        <b>Total Repayments:</b> €${Math.round(totalRepayment).toLocaleString()}<br>
        <b>Final Bank Balance:</b> <span style="color:${finalBal<0?'#c00':'#388e3c'}">€${Math.round(finalBal).toLocaleString()}</span><br>
        <b>Lowest Bank Balance:</b> <span style="color:${minBal<0?'#c00':'#388e3c'}">€${Math.round(minBal).toLocaleString()}</span>
      `;
    }
    // Summary Chart
    let summaryChartElem = document.getElementById('summaryChart');
    if (summaryChart && typeof summaryChart.destroy === "function") summaryChart.destroy();
    if (summaryChartElem) {
      summaryChart = new Chart(summaryChartElem.getContext('2d'), {
        type: 'bar',
        data: {
          labels: ["Income", "Expenditure", "Repayment", "Final Bank", "Lowest Bank"],
          datasets: [{
            label: "Totals",
            data: [
              Math.round(totalIncome),
              -Math.round(totalExpenditure),
              -Math.round(totalRepayment),
              Math.round(finalBal),
              Math.round(minBal)
            ],
            backgroundColor: [
              "#4caf50","#f44336","#ffc107","#2196f3","#9c27b0"
            ]
          }]
        },
        options: {
          responsive:true,
          plugins:{legend:{display:false}},
          scales: { y: { beginAtZero: true } }
        }
      });
    }

    // Tornado Chart logic
    function renderTornadoChart() {
      // Calculate row impact by "sum of absolute values" for each data row
      let impact = [];
      if (!mappedData || !mappingConfigured) return;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let label = mappedData[r][0] || `Row ${r + 1}`;
        let vals = [];
        for (let w = 0; w < weekLabels.length; w++) {
          if (!weekCheckboxStates[w]) continue;
          let absCol = config.weekColStart + w;
          let val = mappedData[r][absCol];
          if (typeof val === "string") val = val.replace(/,/g,'').replace(/€|\s/g,'');
          let num = parseFloat(val);
          if (!isNaN(num)) vals.push(num);
        }
        let total = vals.reduce((a,b)=>a+Math.abs(b),0);
        if (total > 0) impact.push({label, total});
      }
      impact.sort((a,b)=>b.total-a.total);
      impact = impact.slice(0, 10);

      let ctx = document.getElementById('tornadoChart').getContext('2d');
      if (window.tornadoChartObj && typeof window.tornadoChartObj.destroy === "function") window.tornadoChartObj.destroy();
      window.tornadoChartObj = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: impact.map(x=>x.label),
          datasets: [{ label: "Total Impact (€)", data: impact.map(x=>x.total), backgroundColor: '#1976d2' }]
        },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }
      });
    }
    renderTornadoChart();
  }

  // -------------------- ROI/Payback Section --------------------
function renderRoiSection() {
  const dropdown = document.getElementById('investmentWeek');
  if (!dropdown || !weekStartDates.length) return;
  investmentWeekIndex = parseInt(dropdown.value, 10) || 0;
  const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
  const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
  const investmentWeek = investmentWeekIndex;
  const investmentDate = weekStartDates[investmentWeek] || null;

  const repaymentsFull = getRepaymentArr ? getRepaymentArr() : [];
  const repayments = repaymentsFull.slice(investmentWeek);

  const cashflows = [-investment, ...repayments];

  let cashflowDates = [investmentDate];
  for (let i = 1; i < cashflows.length; i++) {
    let idx = investmentWeek + i;
    cashflowDates[i] = weekStartDates[idx] || null;
  }
  // --- ROI Suggestions Integration with new HTML ---
// Event listeners moved to DOMContentLoaded to avoid duplicate registration.

// Also call updateSuggestedRepaymentsOverlay once on ROI tab load for initial display.
updateSuggestedRepaymentsOverlay();

  function npv(rate, cashflows) {
    if (!cashflows.length) return 0;
    return cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate, i), 0);
  }
  function irr(cashflows, guess=0.1) {
    if (!cashflows || cashflows.length < 2) return NaN;
    
    // Check if all cashflows are the same sign (no return possible)
    const hasPositive = cashflows.some(cf => cf > 0);
    const hasNegative = cashflows.some(cf => cf < 0);
    if (!hasPositive || !hasNegative) return NaN;
    
    let rate = guess, epsilon = 1e-8, maxIter = 200;
    let bestRate = rate, bestNPV = Infinity;
    
    for (let iter = 0; iter < maxIter; iter++) {
      let npv0 = npv(rate, cashflows);
      let npv1 = npv(rate + epsilon, cashflows);
      
      // Track best solution
      if (Math.abs(npv0) < Math.abs(bestNPV)) {
        bestNPV = npv0;
        bestRate = rate;
      }
      
      if (Math.abs(npv0) < 1e-10) return rate; // Converged
      
      let deriv = (npv1 - npv0) / epsilon;
      if (!isFinite(deriv) || Math.abs(deriv) < 1e-15) break;
      
      let newRate = rate - npv0 / deriv;
      
      // Bounds checking and step limiting
      if (!isFinite(newRate)) break;
      if (newRate < -0.99) newRate = -0.99; // Prevent extreme negative rates
      if (newRate > 10) newRate = 10; // Prevent extreme positive rates
      
      // Step size limiting for stability
      const maxStep = 0.5;
      if (Math.abs(newRate - rate) > maxStep) {
        newRate = rate + Math.sign(newRate - rate) * maxStep;
      }
      
      if (Math.abs(newRate - rate) < 1e-12) break;
      rate = newRate;
    }
    
    // Return best found solution if within reasonable bounds
    return Math.abs(bestNPV) < 0.01 ? bestRate : NaN;
  }
  function npv_date(rate, cashflows, dateArr) {
    const msPerDay = 24 * 3600 * 1000;
    const baseDate = dateArr[0];
    return cashflows.reduce((acc, val, i) => {
      if (!dateArr[i]) return acc;
      let days = (dateArr[i] - baseDate) / msPerDay;
      let years = days / 365.25;
      return acc + val / Math.pow(1 + rate, years);
    }, 0);
  }

  let npvVal = (discountRate && cashflows.length > 1 && cashflowDates[0]) ?
    npv_date(discountRate / 100, cashflows, cashflowDates) : null;
  let irrVal = (cashflows.length > 1) ? irr(cashflows) : NaN;

  let discCum = 0, payback = null;
  for (let i = 1; i < cashflows.length; i++) {
    let discounted = repayments[i - 1] / Math.pow(1 + discountRate / 100, i);
    discCum += discounted;
    if (payback === null && discCum >= investment) payback = i;
  }

  let tableHtml = `
    <table class="table table-sm">
      <thead>
        <tr>
          <th>Period</th>
          <th>Date</th>
          <th>Repayment</th>
          <th>Cumulative</th>
          <th>Discounted Cumulative</th>
        </tr>
      </thead>
      <tbody>
  `;
  let cum = 0, discCum2 = 0;
  for (let i = 0; i < repayments.length; i++) {
    cum += repayments[i];
    // Discounted cumulative only increases if repayment > 0
    if (repayments[i] > 0) {
      discCum2 += repayments[i] / Math.pow(1 + discountRate / 100, i + 1);
    }
    tableHtml += `
      <tr>
        <td>${weekLabels[investmentWeek + i] || (i + 1)}</td>
        <td>${weekStartDates[investmentWeek + i] ? weekStartDates[investmentWeek + i].toLocaleDateString('en-GB') : '-'}</td>
        <td>€${repayments[i].toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
        <td>€${cum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
        <td>€${discCum2.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
      </tr>
    `;
  }
  tableHtml += `</tbody></table>`;

  let summary = `<div style="background: #f8f9fa; padding: 1.5em; border-radius: 8px; margin: 1em 0;">
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1em; margin-bottom: 1em;">
      <div><b>Total Investment:</b> €${investment.toLocaleString()}</div>
      <div><b>Investment Week:</b> ${weekLabels[investmentWeek] || `Week ${investmentWeek + 1}`}</div>
      <div><b>Total Repayments:</b> €${repayments.reduce((a, b) => a + b, 0).toLocaleString()}</div>
      <div><b>Discount Rate:</b> ${discountRate}%</div>
    </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1em; text-align: center;">
      <div style="background: white; padding: 1em; border-radius: 6px; border-left: 4px solid #2196f3;">
        <div style="font-size: 0.9em; color: #666;">NPV</div>
        <div style="font-size: 1.2em; font-weight: bold; color: ${npvVal > 0 ? '#4caf50' : '#f44336'};">
          ${typeof npvVal === "number" ? "€" + npvVal.toLocaleString(undefined, { maximumFractionDigits: 2 }) : "N/A"}
        </div>
      </div>
      <div style="background: white; padding: 1em; border-radius: 6px; border-left: 4px solid ${irrVal > 0.15 ? '#4caf50' : irrVal > 0.08 ? '#ff9800' : '#f44336'};">
        <div style="font-size: 0.9em; color: #666;">IRR</div>
        <div style="font-size: 1.2em; font-weight: bold;">
          ${isFinite(irrVal) && !isNaN(irrVal) ? (irrVal * 100).toFixed(2) + '%' : 'N/A'}
        </div>
      </div>
      <div style="background: white; padding: 1em; border-radius: 6px; border-left: 4px solid #9c27b0;">
        <div style="font-size: 0.9em; color: #666;">Payback Period</div>
        <div style="font-size: 1.2em; font-weight: bold;">
          ${payback ? `${payback} periods` : 'N/A'}
        </div>
      </div>
    </div>
  </div>`;

  let badge = '';
  if (irrVal > 0.15) badge = '<div style="display: inline-block; background: #4caf50; color: white; padding: 6px 12px; border-radius: 20px; font-size: 0.9em; margin-top: 0.5em;">🎯 Attractive ROI</div>';
  else if (irrVal > 0.08) badge = '<div style="display: inline-block; background: #ff9800; color: white; padding: 6px 12px; border-radius: 20px; font-size: 0.9em; margin-top: 0.5em;">⚠️ Moderate ROI</div>';
  else if (!isNaN(irrVal) && irrVal > 0) badge = '<div style="display: inline-block; background: #f44336; color: white; padding: 6px 12px; border-radius: 20px; font-size: 0.9em; margin-top: 0.5em;">❌ Low ROI</div>';
  else badge = '<div style="display: inline-block; background: #757575; color: white; padding: 6px 12px; border-radius: 20px; font-size: 0.9em; margin-top: 0.5em;">❓ Insufficient Data</div>';

  document.getElementById('roiSummary').innerHTML = summary + badge;
  document.getElementById('roiPaybackTableWrap').innerHTML = tableHtml;

  // Update Actual IRR and Total Repayments display
  let actualIRRDisplay = document.getElementById('actualIrrResult');
  let actualRepaymentsDisplay = document.getElementById('actualRepaymentsResult');
  
  if (actualIRRDisplay) {
    if (isFinite(irrVal) && !isNaN(irrVal)) {
      actualIRRDisplay.innerHTML = `<span style="color: ${irrVal > 0.12 ? '#4caf50' : irrVal > 0.08 ? '#ff9800' : '#f44336'}; font-weight: bold;">${(irrVal * 100).toFixed(2)}%</span>`;
    } else {
      actualIRRDisplay.textContent = '--';
    }
  }
  
  if (actualRepaymentsDisplay) {
    const totalRepayments = repayments.reduce((a, b) => a + b, 0);
    actualRepaymentsDisplay.innerHTML = `<span style="font-weight: bold;">€${totalRepayments.toLocaleString()}</span>`;
  }

  // Charts
  renderRoiCharts(investment, repayments);

  if (!repayments.length || repayments.reduce((a, b) => a + b, 0) === 0) {
    document.getElementById('roiSummary').innerHTML += '<div class="alert alert-warning">No repayments scheduled. ROI cannot be calculated.</div>';
  }
}

// ROI Performance Chart (line) + Pie chart
function renderRoiCharts(investment, repayments) {
  if (!Array.isArray(repayments) || repayments.length === 0) return;

  // Build cumulative and discounted cumulative arrays
  let cumArr = [];
  let discCumArr = [];
  let cum = 0, discCum = 0;
  const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
  for (let i = 0; i < repayments.length; i++) {
    cum += repayments[i] || 0;
    cumArr.push(cum);

    // Discounted only if repayment > 0
    if (repayments[i] > 0) {
      discCum += repayments[i] / Math.pow(1 + discountRate / 100, i + 1);
    }
    discCumArr.push(discCum);
  }

  // Build X labels
  const weekLabels = window.weekLabels || repayments.map((_, i) => `Week ${i + 1}`);

  // ROI Performance Chart (Line)
  let roiLineElem = document.getElementById('roiLineChart');
  if (roiLineElem) {
    const roiLineCtx = roiLineElem.getContext('2d');
    if (window.roiLineChart && typeof window.roiLineChart.destroy === "function") window.roiLineChart.destroy();
    window.roiLineChart = new Chart(roiLineCtx, {
      type: 'line',
      data: {
        labels: weekLabels.slice(0, repayments.length),
        datasets: [
          {
            label: "Cumulative Repayments",
            data: cumArr,
            borderColor: "#4caf50",
            backgroundColor: "#4caf5040",
            fill: false,
            tension: 0.15
          },
          {
            label: "Discounted Cumulative",
            data: discCumArr,
            borderColor: "#1976d2",
            backgroundColor: "#1976d240",
            borderDash: [6,4],
            fill: false,
            tension: 0.15
          },
          {
            label: "Initial Investment",
            data: Array(repayments.length).fill(investment),
            borderColor: "#f44336",
            borderDash: [3,3],
            borderWidth: 1,
            pointRadius: 0,
            fill: false
          }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: true } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "€" } }
        }
      }
    });
  }

  // Pie chart (optional)
  let roiPieElem = document.getElementById('roiPieChart');
  if (roiPieElem) {
    const roiPieCtx = roiPieElem.getContext('2d');
    if (window.roiPieChart && typeof window.roiPieChart.destroy === "function") window.roiPieChart.destroy();
    window.roiPieChart = new Chart(roiPieCtx, {
      type: 'pie',
      data: {
        labels: ["Total Repayments", "Unrecouped"],
        datasets: [{
          data: [
            cumArr[cumArr.length - 1] || 0,
            Math.max(investment - (cumArr[cumArr.length - 1] || 0), 0)
          ],
          backgroundColor: ["#4caf50", "#f3b200"]
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}

// --- ROI input events ---
document.getElementById('roiInvestmentInput').addEventListener('input', renderRoiSection);
document.getElementById('roiInterestInput').addEventListener('input', renderRoiSection);
document.getElementById('refreshRoiBtn').addEventListener('click', renderRoiSection);
document.getElementById('investmentWeek').addEventListener('change', renderRoiSection);
  
function suggestOptimalRepayments({
    investmentAmount,
    investmentWeekIndex,
    weekLabels,
    cashflow,
    openingBalance,
    targetIRR
  }) {
    let cfs = cashflow.slice();
    cfs[investmentWeekIndex] = (cfs[investmentWeekIndex] || 0) - investmentAmount;

    let repaymentWeeks = [];
    for (let i = investmentWeekIndex + 1; i < weekLabels.length; i++) repaymentWeeks.push(i);

    let suggestedRepayments = Array(weekLabels.length).fill(null);

    function computeIRR(cf, guess = 0.1) {
      if (!cf || cf.length < 2) return NaN;
      
      const hasPositive = cf.some(val => val > 0);
      const hasNegative = cf.some(val => val < 0);
      if (!hasPositive || !hasNegative) return NaN;
      
      let rate = guess, epsilon = 1e-8, maxIter = 150;
      let bestRate = rate, bestNPV = Infinity;
      
      for (let k = 0; k < maxIter; k++) {
        let npv = 0, d_npv = 0;
        for (let j = 0; j < cf.length; j++) {
          const factor = Math.pow(1 + rate, j);
          npv += cf[j] / factor;
          if (j > 0) d_npv -= j * cf[j] / Math.pow(1 + rate, j + 1);
        }
        
        // Track best solution
        if (Math.abs(npv) < Math.abs(bestNPV)) {
          bestNPV = npv;
          bestRate = rate;
        }
        
        if (Math.abs(npv) < 1e-10) return rate;
        if (!isFinite(d_npv) || Math.abs(d_npv) < 1e-15) break;
        
        let newRate = rate - npv / d_npv;
        
        // Bounds and step limiting
        if (!isFinite(newRate)) break;
        if (newRate < -0.99) newRate = -0.99;
        if (newRate > 10) newRate = 10;
        
        const maxStep = 0.3;
        if (Math.abs(newRate - rate) > maxStep) {
          newRate = rate + Math.sign(newRate - rate) * maxStep;
        }
        
        if (Math.abs(newRate - rate) < 1e-12) break;
        rate = newRate;
      }
      
      return Math.abs(bestNPV) < 0.01 ? bestRate : NaN;
    }

    // Enhanced repayment strategy with multiple approaches
    let totalToRepay = investmentAmount;
    let simulatedBank = openingBalance;
    let tempCF = cfs.slice();
    let repayments = Array(weekLabels.length).fill(0);
    let remaining = totalToRepay;

    // Strategy 1: Prioritize weeks with highest cashflow availability
    let weekPriorities = repaymentWeeks.map(w => ({
      week: w,
      availableCash: Math.max(0, simulatedBank + tempCF[w]),
      cashflowStrength: tempCF[w] || 0
    })).sort((a, b) => b.availableCash - a.availableCash);

    // Strategy 2: Balanced approach considering time value and cash availability
    let totalAvailable = weekPriorities.reduce((sum, wp) => sum + wp.availableCash, 0);
    
    if (totalAvailable >= remaining) {
      // Distribute proportionally based on availability and time preference
      for (let wp of weekPriorities) {
        if (remaining <= 1e-6) break;
        
        // Weight earlier payments slightly higher for better IRR
        let timeWeight = 1 + (weekLabels.length - wp.week) / weekLabels.length * 0.2;
        let allocation = Math.min(
          wp.availableCash * timeWeight,
          remaining * (wp.availableCash / totalAvailable) * 1.2
        );
        
        allocation = Math.min(allocation, remaining);
        if (allocation > 1) {
          repayments[wp.week] = allocation;
          remaining -= allocation;
        }
      }
    } else {
      // Insufficient funds - use all available cash
      for (let wp of weekPriorities) {
        if (remaining <= 1e-6) break;
        let allocation = Math.min(wp.availableCash, remaining);
        if (allocation > 1) {
          repayments[wp.week] = allocation;
          remaining -= allocation;
        }
      }
    }

    // Strategy 3: Try to optimize for target IRR if possible
    if (targetIRR && targetIRR > 0) {
      let attempts = 0;
      while (attempts < 5) {
        let cfWithRepayments = cfs.slice();
        repayments.forEach((amt, idx) => { 
          cfWithRepayments[idx] = (cfWithRepayments[idx] || 0) - amt; 
        });
        
        let currentIRR = computeIRR(cfWithRepayments);
        if (!isFinite(currentIRR)) break;
        
        let irrDiff = currentIRR - targetIRR;
        if (Math.abs(irrDiff) < 0.01) break; // Close enough
        
        // Adjust payments to move towards target
        if (irrDiff < 0) {
          // Need higher IRR - try to front-load payments
          for (let i = 0; i < repaymentWeeks.length - 2; i++) {
            let w1 = repaymentWeeks[i];
            let w2 = repaymentWeeks[i + 1];
            if (repayments[w2] > 10) {
              let shift = Math.min(repayments[w2] * 0.1, 1000);
              repayments[w1] += shift;
              repayments[w2] -= shift;
            }
          }
        } else {
          // IRR too high - try to back-load payments
          for (let i = repaymentWeeks.length - 1; i > 1; i--) {
            let w1 = repaymentWeeks[i - 1];
            let w2 = repaymentWeeks[i];
            if (repayments[w1] > 10) {
              let shift = Math.min(repayments[w1] * 0.1, 1000);
              repayments[w2] += shift;
              repayments[w1] -= shift;
            }
          }
        }
        attempts++;
      }
    }

    let cfWithRepayments = cfs.slice();
    repayments.forEach((amt, idx) => { cfWithRepayments[idx] = (cfWithRepayments[idx] || 0) - amt; });
    let achievedIRR = computeIRR(cfWithRepayments);

    suggestedRepayments = repayments.map(r => r > 0 ? Math.round(r) : null);

    return { suggestedRepayments, achievedIRR: achievedIRR || 0 };
  }

  function renderPaybackTableRows({repayments, suggestedRepayments, weekLabels, weekStartDates, tableBodyId}) {
    let html = '';
    for (let i = 0; i < weekLabels.length; i++) {
      const actual = repayments[i] || 0;
      const suggested = suggestedRepayments && suggestedRepayments[i] ? suggestedRepayments[i] : null;
      let cellHtml = '';
      if (suggested !== null && Math.abs(suggested - actual) < 0.01 && suggested !== 0) {
        cellHtml = `<span style="color:#219653; font-weight:bold;">€${actual.toLocaleString()}</span>`;
      } else if (suggested !== null && suggested !== 0) {
        cellHtml = `<span style="color:#219653; font-weight:bold;">€${suggested.toLocaleString()}</span>`;
        if (actual > 0) {
          cellHtml += `<br><span style="color:#888;font-size:90%;">(Actual: €${actual.toLocaleString()})</span>`;
        }
      } else if (actual > 0) {
        cellHtml = `€${actual.toLocaleString()}`;
      } else {
        cellHtml = '';
      }
      html += `
        <tr>
          <td>${weekLabels[i]}</td>
          <td>${weekStartDates && weekStartDates[i] ? weekStartDates[i].toLocaleDateString('en-GB') : '-'}</td>
          <td style="text-align:right;">${cellHtml}</td>
        </tr>
      `;
    }
    document.getElementById(tableBodyId).innerHTML = html;
  }

  // --- Handler for "Show Suggested Repayments" ---
  function updateSuggestedRepaymentsOverlay() {
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const cashflow = weekLabels.map((_, i) => (incomeArr[i] || 0) - (expenditureArr[i] || 0));
    const investmentWeekIndex = parseInt(document.getElementById('investmentWeek').value, 10) || 0;
    const targetIRR = parseFloat(document.getElementById('roiTargetIrrInput').value) / 100;
    const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
    const { suggestedRepayments, achievedIRR } = suggestOptimalRepayments({
      investmentAmount: investment,
      investmentWeekIndex,
      weekLabels,
      cashflow,
      openingBalance,
      targetIRR
    });

    const repayments = getRepaymentArr();

    renderPaybackTableRows({
      repayments,
      suggestedRepayments,
      weekLabels,
      weekStartDates,
      tableBodyId: 'roiPaybackTableBody'
    });

    const irrDisplay = document.getElementById('suggestedIrrResult');
    if (irrDisplay)
      irrDisplay.innerHTML = `Achievable IRR: <b>${(achievedIRR*100).toFixed(2)}%</b> ${Math.abs(achievedIRR - targetIRR) < 0.005 ? '<span class="badge badge-success">Target Met</span>' : '<span class="badge badge-warning">Best possible</span>'}`;
  }

  // --- Modal installment selection ---
  function createInstallmentSelectionModal() {
    const modal = document.createElement('div');
    modal.id = 'installmentModal';
    modal.className = 'modal-overlay';
    modal.style.cssText = `
      position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
      background: rgba(0,0,0,0.5); z-index: 1000; display: none;
      justify-content: center; align-items: center;
    `;
    
    const modalContent = document.createElement('div');
    modalContent.className = 'modal-content';
    modalContent.style.cssText = `
      background: white; padding: 2em; border-radius: 8px; max-width: 600px; width: 90%;
      max-height: 80vh; overflow-y: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    `;
    
    modalContent.innerHTML = `
      <h3>Select Investment & Repayment Schedule</h3>
      <div style="margin: 1em 0;">
        <label style="display: block; margin-bottom: 0.5em;"><b>Investment Week:</b></label>
        <select id="modalInvestmentWeek" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
        </select>
      </div>
      <div style="margin: 1em 0;">
        <label style="display: block; margin-bottom: 0.5em;"><b>Investment Amount (€):</b></label>
        <input type="number" id="modalInvestmentAmount" value="300000" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
      </div>
      <div style="margin: 1em 0;">
        <label style="display: block; margin-bottom: 0.5em;"><b>Target IRR (%):</b></label>
        <input type="number" id="modalTargetIRR" value="12" min="0" max="100" step="0.1" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
      </div>
      <div style="margin: 1em 0;">
        <h4>Suggested Repayment Schedule:</h4>
        <div id="modalRepaymentSuggestions" style="max-height: 200px; overflow-y: auto; border: 1px solid #eee; padding: 1em; margin: 0.5em 0;"></div>
        <div id="modalIrrResult" style="margin: 0.5em 0; font-weight: bold;"></div>
      </div>
      <div style="text-align: right; margin-top: 1.5em;">
        <button id="modalApplyBtn" style="background: #4caf50; color: white; border: none; padding: 8px 16px; border-radius: 4px; margin-right: 8px; cursor: pointer;">Apply Schedule</button>
        <button id="modalCancelBtn" style="background: #757575; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">Cancel</button>
      </div>
    `;
    
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
    
    // Modal event handlers
    document.getElementById('modalCancelBtn').addEventListener('click', () => {
      modal.style.display = 'none';
    });
    
    document.getElementById('modalApplyBtn').addEventListener('click', () => {
      const investmentWeek = parseInt(document.getElementById('modalInvestmentWeek').value, 10) || 0;
      const investmentAmount = parseFloat(document.getElementById('modalInvestmentAmount').value) || 0;
      const targetIRR = parseFloat(document.getElementById('modalTargetIRR').value) || 0;
      
      // Apply to main form
      document.getElementById('investmentWeek').value = investmentWeek;
      document.getElementById('roiInvestmentInput').value = investmentAmount;
      document.getElementById('roiTargetIrrInput').value = targetIRR;
      
      modal.style.display = 'none';
      renderRoiSection();
      updateSuggestedRepaymentsOverlay();
    });
    
    // Close modal on background click
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.style.display = 'none';
    });
    
    return modal;
  }
  
  // Create modal on first load
  const installmentModal = createInstallmentSelectionModal();
  
  function showInstallmentModal() {
    const modal = document.getElementById('installmentModal');
    
    // Populate investment week options
    const modalWeekSelect = document.getElementById('modalInvestmentWeek');
    modalWeekSelect.innerHTML = '';
    weekLabels.forEach((label, i) => {
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = label || `Week ${i + 1}`;
      modalWeekSelect.appendChild(opt);
    });
    
    // Set current values
    modalWeekSelect.value = document.getElementById('investmentWeek').value || 0;
    document.getElementById('modalInvestmentAmount').value = document.getElementById('roiInvestmentInput').value || 300000;
    document.getElementById('modalTargetIRR').value = document.getElementById('roiTargetIrrInput').value || 12;
    
    // Update suggestions
    updateModalSuggestions();
    
    modal.style.display = 'flex';
  }
  
  function updateModalSuggestions() {
    const investmentAmount = parseFloat(document.getElementById('modalInvestmentAmount').value) || 0;
    const investmentWeekIndex = parseInt(document.getElementById('modalInvestmentWeek').value, 10) || 0;
    const targetIRR = parseFloat(document.getElementById('modalTargetIRR').value) / 100 || 0.12;
    
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const cashflow = weekLabels.map((_, i) => (incomeArr[i] || 0) - (expenditureArr[i] || 0));
    
    const { suggestedRepayments, achievedIRR } = suggestOptimalRepayments({
      investmentAmount,
      investmentWeekIndex,
      weekLabels,
      cashflow,
      openingBalance,
      targetIRR
    });
    
    // Display suggestions
    const suggestionsDiv = document.getElementById('modalRepaymentSuggestions');
    let html = '<table style="width: 100%; font-size: 0.9em;"><thead><tr><th>Week</th><th>Suggested Repayment</th></tr></thead><tbody>';
    
    suggestedRepayments.forEach((amount, i) => {
      if (amount && amount > 0) {
        html += `<tr><td>${weekLabels[i] || `Week ${i + 1}`}</td><td style="text-align: right;">€${amount.toLocaleString()}</td></tr>`;
      }
    });
    
    html += '</tbody></table>';
    suggestionsDiv.innerHTML = html;
    
    const irrResult = document.getElementById('modalIrrResult');
    const irrDiff = Math.abs(achievedIRR - targetIRR);
    const badge = irrDiff < 0.005 ? '<span style="background: #4caf50; color: white; padding: 2px 6px; border-radius: 3px; font-size: 0.8em;">Target Achieved</span>' 
                                   : '<span style="background: #ff9800; color: white; padding: 2px 6px; border-radius: 3px; font-size: 0.8em;">Best Possible</span>';
    irrResult.innerHTML = `Achievable IRR: <b>${(achievedIRR * 100).toFixed(2)}%</b> ${badge}`;
  }
  
  // Event listeners for modal inputs
  document.addEventListener('change', (e) => {
    if (e.target.id === 'modalInvestmentAmount' || e.target.id === 'modalInvestmentWeek' || e.target.id === 'modalTargetIRR') {
      updateModalSuggestions();
    }
  });
  
  // --- Wire up to button/input ---
  const showBtn = document.getElementById('showSuggestedRepaymentsBtn');
  if (showBtn) {
    showBtn.textContent = 'Advanced Repayment Planner';
    showBtn.addEventListener('click', showInstallmentModal);
  }
  const irrInput = document.getElementById('roiTargetIrrInput');
  if (irrInput) irrInput.addEventListener('change', updateSuggestedRepaymentsOverlay);

  // --- Optionally: render suggestions on initial load for convenience ---
  updateSuggestedRepaymentsOverlay();

  // ... rest of your script (updateAllTabs, etc) ...
  function updateAllTabs() {
    try {
      renderRepaymentRows();
    } catch (e) {
      console.error('Error in renderRepaymentRows:', e);
    }
    try {
      updateLoanSummary();
    } catch (e) {
      console.error('Error in updateLoanSummary:', e);
    }
    try {
      updateChartAndSummary();
    } catch (e) {
      console.error('Error in updateChartAndSummary:', e);
    }
    try {
      renderPnlTables();
    } catch (e) {
      console.error('Error in renderPnlTables:', e);
    }
    try {
      renderSummaryTab();
    } catch (e) {
      console.error('Error in renderSummaryTab:', e);
    }
    try {
      renderRoiSection();
    } catch (e) {
      console.error('Error in renderRoiSection:', e);
    }
    try {
      renderTornadoChart();
    } catch (e) {
      console.error('Error in renderTornadoChart:', e);
    }
    // Optionally show suggestions each time:
    try {
      updateSuggestedRepaymentsOverlay();
    } catch (e) {
      console.error('Error in updateSuggestedRepaymentsOverlay:', e);
    }
  }
  updateAllTabs();
});
