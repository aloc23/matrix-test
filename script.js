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

  // --- DEMO: Suggested Repayments for Overlay ---
  window.suggestedRepayments = []; // Array parallel to weekLabels

  function computeSuggestedRepayments() {
    // Demo: Suggest repayments for weeks 5, 12, 20 (indexes in weekLabels)
    // Replace with your real IRR/NPV logic later!
    window.suggestedRepayments = Array(window.weekLabels.length).fill(null);
    window.suggestedRepayments[5] = 10000;
    window.suggestedRepayments[12] = 40000;
    window.suggestedRepayments[20] = 10000;
  }

  // -------------------- Tabs & UI Interactions --------------------
  function setupTabs() {
    document.querySelectorAll('.tabs button').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.tabs button').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        document.querySelectorAll('.tab-content').forEach(sec => sec.classList.remove('active'));
        var tabId = btn.getAttribute('data-tab');
        var panel = document.getElementById(tabId);
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
        var subtabId = 'subtab-' + btn.getAttribute('data-subtab');
        var subpanel = document.getElementById(subtabId);
        if (subpanel) subpanel.classList.add('active');
        setTimeout(updateAllTabs, 50);
      });
    });
    document.querySelectorAll('.collapsible-header').forEach(btn => {
      btn.addEventListener('click', function() {
        var content = btn.nextElementSibling;
        var caret = btn.querySelector('.caret');
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

  function renderMappingPanel(allRows) { /* unchanged */ }
  function autoDetectMapping(sheet) { /* unchanged */ }
  function extractWeekStartDates(weekLabels, baseYear) { /* unchanged */ }
  function populateInvestmentWeekDropdown() { /* unchanged */ }

  function updateWeekLabels() {
    let weekRow = mappedData[config.weekLabelRow] || [];
    weekLabels = weekRow.slice(config.weekColStart, config.weekColEnd+1).map(x => x || '');
    window.weekLabels = weekLabels; // make global for charts
    if (!weekCheckboxStates || weekCheckboxStates.length !== weekLabels.length) {
      weekCheckboxStates = weekLabels.map(() => true);
    }
    populateWeekDropdown(weekLabels);

    // ROI week start date integration. Use a default base year (2025) or prompt user for year.
    weekStartDates = extractWeekStartDates(weekLabels, 2025);
    populateInvestmentWeekDropdown();
  }

  function getFilteredWeekIndices() { /* unchanged */ }
  function getIncomeArr() { /* unchanged */ }
  function getExpenditureArr() { /* unchanged */ }
  function getRepaymentArr() { /* unchanged */ }
  function getNetProfitArr(incomeArr, expenditureArr, repaymentArr) { /* unchanged */ }
  function getRollingBankBalanceArr() { /* unchanged */ }
  function getMonthAgg(arr, months=12) { /* unchanged */ }

  // -------------------- Repayments UI --------------------
  const weekSelect = document.getElementById('weekSelect');
  const repaymentFrequency = document.getElementById('repaymentFrequency');
  function populateWeekDropdown(labels) { /* unchanged */ }
  function setupRepaymentForm() { /* unchanged */ }
  function renderRepaymentRows() { /* unchanged */ }
  function updateLoanSummary() { /* unchanged */ }

  let loanOutstandingInput = document.getElementById('loanOutstandingInput');
  if (loanOutstandingInput) {
    loanOutstandingInput.oninput = function() {
      loanOutstanding = parseFloat(this.value) || 0;
      updateLoanSummary();
    };
  }

  // -------------------- Main Chart & Summary --------------------
  function updateChartAndSummary() { /* unchanged */ }

  // ---------- P&L Tab Functions ----------
  function renderSectionSummary(headerId, text, arr) { /* unchanged */ }
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

      // Overlay actual & suggested repayment
      let actual = repaymentArr[i] || 0;
      let suggested = window.suggestedRepayments[weekIdxs[i]];

      let repaymentCell = '';
      if (suggested !== null && suggested !== undefined && suggested !== actual) {
        repaymentCell = `<div>€${Math.round(actual).toLocaleString()}</div>
          <div class="suggested-repayment">Suggested: €${Math.round(suggested).toLocaleString()}</div>`;
      } else if (suggested !== null && suggested !== undefined) {
        repaymentCell = `<div class="suggested-repayment">Suggested: €${Math.round(suggested).toLocaleString()}</div>`;
      } else {
        repaymentCell = `€${Math.round(actual).toLocaleString()}`;
      }

      let row = `<tr${rollingArr[i]<0?' class="negative-balance-row"':''}>` +
        `<td>${weekLabels[idx]}</td>` +
        `<td${incomeArr[idx]<0?' class="negative-number"':''}>€${Math.round(incomeArr[idx]||0).toLocaleString()}</td>` +
        `<td${expenditureArr[idx]<0?' class="negative-number"':''}>€${Math.round(expenditureArr[idx]||0).toLocaleString()}</td>` +
        `<td>${repaymentCell}</td>` +
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
      let closing = opening = openingBalance;
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
  function renderSummaryTab() { /* unchanged */ }
  function renderRoiSection() { /* unchanged */ }
  function renderRoiCharts(investment, repayments) { /* unchanged */ }

  document.getElementById('roiInvestmentInput').addEventListener('input', renderRoiSection);
  document.getElementById('roiInterestInput').addEventListener('input', renderRoiSection);
  document.getElementById('refreshRoiBtn').addEventListener('click', renderRoiSection);
  document.getElementById('investmentWeek').addEventListener('change', renderRoiSection);

  // -------------------- Update All Tabs --------------------
  function updateAllTabs() {
    computeSuggestedRepayments(); // <-- Add here for testing!
    renderRepaymentRows();
    updateLoanSummary();
    updateChartAndSummary();
    renderPnlTables();
    renderSummaryTab();
    renderRoiSection();
    renderTornadoChart();
  }
  updateAllTabs();
});
