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

  // Chart.js chart instances for destroy
  let mainChart = null;
  let roiPieChart = null;
  let roiLineChart = null;
  window.tornadoChartObj = null;
  let summaryChart = null;

  // Suggestions state
  let showSuggestions = false;
  let lastSuggestedRepayments = null;
  let lastAchievedIRR = null;

  // -------------------- Utility & Mapping Functions --------------------
  function getFilteredWeekIndices() {
    return weekCheckboxStates.map((checked, idx) => checked ? idx : null).filter(idx => idx !== null);
  }
  function updateWeekLabels() {
    let weekRow = mappedData[config.weekLabelRow] || [];
    weekLabels = weekRow.slice(config.weekColStart, config.weekColEnd+1).map(x => x || '');
    window.weekLabels = weekLabels;
    if (!weekCheckboxStates || weekCheckboxStates.length !== weekLabels.length) {
      weekCheckboxStates = weekLabels.map(() => true);
    }
    weekStartDates = extractWeekStartDates(weekLabels, 2025);
    populateWeekDropdown(weekLabels);
    populateInvestmentWeekDropdown();
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
  function populateWeekDropdown(labels) {
    // Used for repayment row UI, not needed for main ROI logic here
    const weekSelect = document.getElementById('weekSelect');
    if (!weekSelect) return;
    weekSelect.innerHTML = '';
    (labels && labels.length ? labels : Array.from({length: 52}, (_, i) => `Week ${i+1}`)).forEach(label => {
      const opt = document.createElement('option');
      opt.value = label;
      opt.textContent = label;
      weekSelect.appendChild(opt);
    });
  }

  // -------------------- Data Calculation Functions --------------------
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

  // -------------------- Repayment Suggestion Logic --------------------
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
      let maxIter = 30, tol = 1e-7;
      let rate = guess;
      for (let k = 0; k < maxIter; k++) {
        let npv = 0, d_npv = 0;
        for (let j = 0; j < cf.length; j++) {
          npv += cf[j] / Math.pow(1 + rate, j);
          if (j > 0) d_npv -= j * cf[j] / Math.pow(1 + rate, j + 1);
        }
        if (Math.abs(npv) < tol) return rate;
        if (!isFinite(d_npv) || d_npv === 0) break;
        rate = rate - npv / d_npv;
      }
      return rate;
    }

    let totalToRepay = investmentAmount;
    let simulatedBank = openingBalance;
    let tempCF = cfs.slice();
    let repayments = Array(weekLabels.length).fill(0);
    let remaining = totalToRepay;

    for (let w of repaymentWeeks) {
      let maxRepay = Math.max(0, simulatedBank + tempCF[w]);
      let pay = Math.min(maxRepay, remaining);
      repayments[w] = pay;
      tempCF[w] -= pay;
      simulatedBank += tempCF[w];
      remaining -= pay;
      if (remaining <= 1e-6) break;
    }
    let cfWithRepayments = cfs.slice();
    repayments.forEach((amt, idx) => { cfWithRepayments[idx] = (cfWithRepayments[idx] || 0) - amt; });
    let achievedIRR = computeIRR(cfWithRepayments);
    suggestedRepayments = repayments.map(r => r > 0 ? r : null);

    return { suggestedRepayments, achievedIRR };
  }

  // -------------------- Table Rendering (Single Table) --------------------
  function renderRepaymentTable({ actualRepayments, suggestedRepayments, weekLabels, weekStartDates }) {
    if (!mappingConfigured || !weekLabels.length) {
      document.getElementById('roiPaybackTableBody').innerHTML = '';
      return;
    }
    let cum = 0, discCum = 0;
    const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
    let html = '';
    for (let i = 0; i < weekLabels.length; i++) {
      const actual = actualRepayments[i] || 0;
      const suggested = (suggestedRepayments && suggestedRepayments[i] != null) ? suggestedRepayments[i] : null;
      cum += actual;
      if (actual > 0) {
        discCum += actual / Math.pow(1 + discountRate / 100, i + 1);
      }
      let cellHtml = '';
      if (suggested !== null) {
        if (Math.abs(actual - suggested) < 0.01) {
          cellHtml = `<span style="color:#219653; font-weight:bold;">€${actual.toLocaleString()}</span>`;
        } else {
          cellHtml = `<span style="color:#219653; font-weight:bold;">€${suggested.toLocaleString()}</span>`;
          if (actual > 0) {
            cellHtml += `<br><span style="color:#888;font-size:90%;">(Actual: €${actual.toLocaleString()})</span>`;
          }
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
          <td>€${cum.toLocaleString()}</td>
          <td>€${discCum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
        </tr>
      `;
    }
    document.getElementById('roiPaybackTableBody').innerHTML = html;
  }

  // -------------------- IRR Calculation and Summary Display --------------------
  function calculateAndShowROIHeader() {
    const repayments = getRepaymentArr();
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const cashflow = weekLabels.map((_, i) => (incomeArr[i] || 0) - (expenditureArr[i] || 0));
    const investmentWeekIndex = parseInt(document.getElementById('investmentWeek').value, 10) || 0;
    const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
    const cashflows = [-investment, ...repayments];

    function irr(cashflows, guess=0.1) {
      let rate = guess, epsilon = 1e-6, maxIter = 100;
      for (let iter=0; iter<maxIter; iter++) {
        let npv0 = cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate, i), 0);
        let npv1 = cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate+epsilon, i), 0);
        let deriv = (npv1-npv0)/epsilon;
        let newRate = rate - npv0/deriv;
        if (!isFinite(newRate)) break;
        if (Math.abs(newRate-rate) < 1e-7) return newRate;
        rate = newRate;
      }
      return NaN;
    }
    const actualIrr = irr(cashflows);
    document.getElementById('actualIrrResult').innerText = (isFinite(actualIrr) && !isNaN(actualIrr)) ? (actualIrr*100).toFixed(2) + "%" : "n/a";
    document.getElementById('actualRepaymentsResult').innerText = "€" + repayments.reduce((a,b)=>a+b,0).toLocaleString();

    // Show suggestions IRR if present
    if (showSuggestions && lastAchievedIRR != null) {
      const targetIRR = parseFloat(document.getElementById('roiTargetIrrInput').value) / 100;
      document.getElementById('suggestedIrrResult').innerHTML = `Achievable IRR: <b>${(lastAchievedIRR*100).toFixed(2)}%</b> ${Math.abs(lastAchievedIRR - targetIRR) < 0.005 ? '<span class="badge badge-success">Target Met</span>' : '<span class="badge badge-warning">Best possible</span>'}`;
    } else {
      document.getElementById('suggestedIrrResult').innerHTML = '';
    }
  }

  // -------------------- Event Handlers for Suggestions --------------------
  document.getElementById('showSuggestedRepaymentsBtn').addEventListener('click', function() {
    const repayments = getRepaymentArr();
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const cashflow = weekLabels.map((_, i) => (incomeArr[i] || 0) - (expenditureArr[i] || 0));
    const investmentWeekIndex = parseInt(document.getElementById('investmentWeek').value, 10) || 0;
    const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
    const targetIRR = parseFloat(document.getElementById('roiTargetIrrInput').value) / 100;

    const out = suggestOptimalRepayments({
      investmentAmount: investment,
      investmentWeekIndex,
      weekLabels,
      cashflow,
      openingBalance,
      targetIRR
    });

    showSuggestions = true;
    lastSuggestedRepayments = out.suggestedRepayments;
    lastAchievedIRR = out.achievedIRR;

    renderRepaymentTable({
      actualRepayments: repayments,
      suggestedRepayments: lastSuggestedRepayments,
      weekLabels,
      weekStartDates
    });
    calculateAndShowROIHeader();
  });

  document.getElementById('roiTargetIrrInput').addEventListener('change', function() {
    if (showSuggestions) {
      document.getElementById('showSuggestedRepaymentsBtn').click();
    }
  });

  function resetSuggestions() {
    showSuggestions = false;
    lastSuggestedRepayments = null;
    lastAchievedIRR = null;
  }

  // -------------------- Integration with Mapping/Upload --------------------
  function afterMappingOrUpload() {
    resetSuggestions();
    renderRepaymentTable({
      actualRepayments: getRepaymentArr(),
      suggestedRepayments: null,
      weekLabels,
      weekStartDates
    });
    calculateAndShowROIHeader();
  }

  // -------------------- Main Update Trigger --------------------
  function updateAllTabs() {
    // (call all your other update/render functions as needed)
    renderRepaymentTable({
      actualRepayments: getRepaymentArr(),
      suggestedRepayments: showSuggestions ? lastSuggestedRepayments : null,
      weekLabels,
      weekStartDates
    });
    calculateAndShowROIHeader();
  }

  // For demo, also call at startup:
  afterMappingOrUpload();

  // If you want to also update table on tab click, wire up tab logic as needed

  // Reminder: after mapping, after uploading, etc, call afterMappingOrUpload()
});
