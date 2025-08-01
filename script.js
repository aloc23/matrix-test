// ROI/NPV/IRR Financial Tool - FULL SCRIPT

document.addEventListener('DOMContentLoaded', function() {
  // --- State ---
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
  let weekStartDates = [];
  let investmentWeekIndex = 0;
  let mainChart = null;
  let roiPieChart = null;
  let roiLineChart = null;
  window.tornadoChartObj = null;
  let summaryChart = null;
  window.suggestedRepayments = [];

  // --- Remove legacy ROI edit controls if present ---
  function removeLegacyRoiEditButtons() {
    const legacyEdit = document.getElementById('roiEditControls');
    if (legacyEdit) legacyEdit.remove();
  }

  // --- Setup ROI IRR/NPV slider overlay UI ---
  function setupRoiIrrSlider() {
    if (document.getElementById('roiIrrSliderRow')) return;
    const paybackTableWrap = document.getElementById('roiPaybackTableWrap');
    if (!paybackTableWrap) return;
    const sliderRow = document.createElement('div');
    sliderRow.id = "roiIrrSliderRow";
    sliderRow.style.marginBottom = "10px";
    sliderRow.innerHTML = `
      <button id="toggleIrrSliderBtn">Edit IRR/NPV</button>
      <span id="sliderContainer" style="display:none; margin-left:20px;">
        <label for="irrSlider">Test IRR target: </label>
        <input id="irrSlider" type="range" min="0" max="0.5" step="0.01" value="0.15" style="vertical-align:middle;">
        <span id="irrSliderValue">0.15</span>
        <button id="applySuggestedRepaymentsBtn" style="margin-left:14px;">Apply Suggestions</button>
        <div id="overlayAlertBox" style="margin-top:5px;"></div>
      </span>
    `;
    paybackTableWrap.parentNode.insertBefore(sliderRow, paybackTableWrap);

    document.getElementById('toggleIrrSliderBtn').onclick = function() {
      const sliderContainer = document.getElementById('sliderContainer');
      sliderContainer.style.display = sliderContainer.style.display === 'none' ? '' : 'none';
      if (sliderContainer.style.display === '') {
        computeSuggestedRepaymentsForIrr(parseFloat(document.getElementById('irrSlider').value));
        renderRoiSection(true);
      } else {
        window.suggestedRepayments = Array(window.weekLabels.length).fill(null);
        renderRoiSection(false);
      }
    };

    document.getElementById('irrSlider').oninput = function() {
      document.getElementById('irrSliderValue').textContent = this.value;
      computeSuggestedRepaymentsForIrr(parseFloat(this.value));
      renderRoiSection(true);
    };

    document.getElementById('applySuggestedRepaymentsBtn').onclick = function() {
      applySuggestedRepayments();
      window.suggestedRepayments = Array(window.weekLabels.length).fill(null);
      renderRoiSection(false);
      document.getElementById('sliderContainer').style.display = "none";
    };
  }

  function computeSuggestedRepaymentsForIrr(targetIrr) {
    window.suggestedRepayments = Array(window.weekLabels.length).fill(null);
    const count = Math.floor(Math.random() * 4) + 3;
    const used = new Set();
    while (used.size < count && weekLabels.length > 0) {
      used.add(Math.floor(Math.random() * weekLabels.length));
    }
    used.forEach(idx => {
      window.suggestedRepayments[idx] = Math.round(Math.random() * 50000) + 5000;
    });
    document.getElementById('overlayAlertBox').innerHTML =
      `<div class="alert alert-success">Test overlay: <b>${count}</b> suggested repayments for IRR ${targetIrr} shown (green) in Payback Table.</div>`;
  }

  function applySuggestedRepayments() {
    let suggestions = window.suggestedRepayments;
    for (let i = 0; i < suggestions.length; i++) {
      if (suggestions[i] != null) {
        repaymentRows = repaymentRows.filter(
          r => !(r.type === "week" && weekLabels.indexOf(r.week) === i)
        );
        repaymentRows.push({ type: "week", week: weekLabels[i], amount: suggestions[i], editing: false });
      }
    }
    window.suggestedRepayments = Array(window.weekLabels.length).fill(null);
  }

  function updateWeekLabels() {
    let weekRow = mappedData[config.weekLabelRow] || [];
    weekLabels = weekRow.slice(config.weekColStart, config.weekColEnd+1).map(x => x || '');
    window.weekLabels = weekLabels;
    if (!weekCheckboxStates || weekCheckboxStates.length !== weekLabels.length) {
      weekCheckboxStates = weekLabels.map(() => true);
    }
    weekStartDates = extractWeekStartDates(weekLabels, 2025);
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
    return arr;
  }

  function getIncomeArr() {
    if (!mappedData || !mappingConfigured) return [];
    let arr = [];
    for (let w = 0; w < weekLabels.length; w++) {
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

  function getRollingBankBalanceArr() {
    let incomeArr = getIncomeArr();
    let expenditureArr = getExpenditureArr();
    let repaymentArr = getRepaymentArr();
    let rolling = [];
    let ob = openingBalance;
    for (let i = 0; i < weekLabels.length; i++) {
      let income = incomeArr[i] || 0;
      let out = expenditureArr[i] || 0;
      let repay = repaymentArr[i] || 0;
      let prev = (i === 0 ? ob : rolling[i-1]);
      rolling[i] = prev + income - out - repay;
    }
    return rolling;
  }

  function getNetProfitArr(incomeArr, expenditureArr, repaymentArr) {
    return incomeArr.map((inc, i) => (inc || 0) - (expenditureArr[i] || 0) - (repaymentArr[i] || 0));
  }

  function getFilteredWeekIndices() {
    // TODO: Replace the filter condition below with the actual filter logic as needed.
    // Example: Only include indices where the week label is not empty.
    return weekLabels
      .map((label, idx) => ({ label, idx }))
      .filter(obj => obj.label && obj.label.trim() !== "")
      .map(obj => obj.idx);
  }

  function renderRepaymentRows() {
    const container = document.getElementById('repaymentRows');
    if (!container) return;
    container.innerHTML = "";
    repaymentRows.forEach((row, i) => {
      const div = document.createElement('div');
      div.className = 'repayment-row';
      div.textContent = row.type === "week" ? `${row.week}: €${row.amount}` : `${row.frequency}: €${row.amount}`;
      const removeBtn = document.createElement('button');
      removeBtn.textContent = 'Remove';
      removeBtn.onclick = function() {
        repaymentRows.splice(i, 1);
        renderRepaymentRows();
        updateAllTabs();
      };
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

  function updateChartAndSummary() {
    const mainChartElem = document.getElementById('mainChart');
    if (!mainChartElem) return;
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const repaymentArr = getRepaymentArr();
    const rollingArr = getRollingBankBalanceArr();
    const netProfitArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);

    if (mainChart && typeof mainChart.destroy === "function") mainChart.destroy();
    mainChart = new Chart(mainChartElem.getContext('2d'), {
      type: 'bar',
      data: {
        labels: weekLabels,
        datasets: [
          {
            label: "Income",
            data: incomeArr,
            backgroundColor: "rgba(76,175,80,0.6)",
            borderColor: "#388e3c",
            fill: false,
            type: "bar"
          },
          {
            label: "Expenditure",
            data: expenditureArr,
            backgroundColor: "rgba(244,67,54,0.6)",
            borderColor: "#c62828",
            fill: false,
            type: "bar"
          },
          {
            label: "Repayment",
            data: repaymentArr,
            backgroundColor: "rgba(255,193,7,0.6)",
            borderColor: "#ff9800",
            fill: false,
            type: "bar"
          },
          {
            label: "Net Profit",
            data: netProfitArr,
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
      },
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
  }

  function renderRoiCharts(investment, repayments) {
    if (!Array.isArray(repayments) || repayments.length === 0) return;
    let cumArr = [];
    let discCumArr = [];
    let cum = 0, discCum = 0;
    const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
    for (let i = 0; i < repayments.length; i++) {
      cum += repayments[i] || 0;
      cumArr.push(cum);
      if (repayments[i] > 0) {
        discCum += repayments[i] / Math.pow(1 + discountRate / 100, i + 1);
      }
      discCumArr.push(discCum);
    }
    const roiWeekLabels = window.weekLabels || repayments.map((_, i) => `Week ${i + 1}`);
    let roiLineElem = document.getElementById('roiLineChart');
    if (roiLineElem) {
      const roiLineCtx = roiLineElem.getContext('2d');
      if (window.roiLineChart && typeof window.roiLineChart.destroy === "function") window.roiLineChart.destroy();
      window.roiLineChart = new Chart(roiLineCtx, {
        type: 'line',
        data: {
          labels: roiWeekLabels.slice(0, repayments.length),
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

  function renderTornadoChart() {
    let impact = [];
    if (!mappedData || !mappingConfigured) return;
    for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
      let label = mappedData[r][0] || `Row ${r + 1}`;
      let vals = [];
      for (let w = 0; w < weekLabels.length; w++) {
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
    let ctx = document.getElementById('tornadoChart');
    if (!ctx) return;
    ctx = ctx.getContext('2d');
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

  function renderSummaryTab() {
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

    if (document.getElementById('summaryKeyFinancials')) {
      document.getElementById('summaryKeyFinancials').innerHTML = `
        <b>Total Income:</b> €${Math.round(totalIncome).toLocaleString()}<br>
        <b>Total Expenditure:</b> €${Math.round(totalExpenditure).toLocaleString()}<br>
        <b>Total Repayments:</b> €${Math.round(totalRepayment).toLocaleString()}<br>
        <b>Final Bank Balance:</b> <span style="color:${finalBal<0?'#c00':'#388e3c'}">€${Math.round(finalBal).toLocaleString()}</span><br>
        <b>Lowest Bank Balance:</b> <span style="color:${minBal<0?'#c00':'#388e3c'}">€${Math.round(minBal).toLocaleString()}</span>
      `;
    }
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
  }

  function renderPnlTables() {
    // Stub: Implement weekly/monthly/cashflow tables as needed
    // For now, we just call updateChartAndSummary
    updateChartAndSummary();
  }

  // --- ROI Section (with showSuggestions param) ---
  window.renderRoiSection = function(showSuggestions) {
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
    function npv(rate, cashflows) {
      if (!cashflows.length) return 0;
      return cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate, i), 0);
    }
    function irr(cashflows, guess=0.1) {
      let rate = guess, epsilon = 1e-6, maxIter = 100;
      for (let iter=0; iter<maxIter; iter++) {
        let npv0 = npv(rate, cashflows);
        let npv1 = npv(rate+epsilon, cashflows);
        let deriv = (npv1-npv0)/epsilon;
        let newRate = rate - npv0/deriv;
        if (!isFinite(newRate)) break;
        if (Math.abs(newRate-rate) < 1e-7) return newRate;
        rate = newRate;
      }
      return NaN;
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
      if (repayments[i] > 0) {
        discCum2 += repayments[i] / Math.pow(1 + discountRate / 100, i + 1);
      }
      let suggested = showSuggestions ? (window.suggestedRepayments[investmentWeek + i] || null) : null;
      let repaymentCell = "";
      if (suggested !== null && suggested !== undefined && suggested !== repayments[i]) {
        repaymentCell = `<div>€${repayments[i].toLocaleString(undefined, {maximumFractionDigits: 2})}</div>
          <div class="suggested-repayment">Suggested: €${Math.round(suggested).toLocaleString()}</div>`;
      } else if (suggested !== null && suggested !== undefined) {
        repaymentCell = `<div class="suggested-repayment">Suggested: €${Math.round(suggested).toLocaleString()}</div>`;
      } else {
        repaymentCell = `€${repayments[i].toLocaleString(undefined, {maximumFractionDigits: 2})}`;
      }
      tableHtml += `
        <tr>
          <td>${weekLabels[investmentWeek + i] || (i + 1)}</td>
          <td>${weekStartDates[investmentWeek + i] ? weekStartDates[investmentWeek + i].toLocaleDateString('en-GB') : '-'}</td>
          <td>${repaymentCell}</td>
          <td>€${cum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
          <td>€${discCum2.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
        </tr>
      `;
    }
    tableHtml += `</tbody></table>`;
    let summary = `<b>Total Investment:</b> €${investment.toLocaleString()}<br>
      <b>Total Repayments:</b> €${repayments.reduce((a, b) => a + b, 0).toLocaleString()}<br>
      <b>NPV (${discountRate}%):</b> ${typeof npvVal === "number" ? "€" + npvVal.toLocaleString(undefined, { maximumFractionDigits: 2 }) : "n/a"}<br>
      <b>IRR:</b> ${isFinite(irrVal) && !isNaN(irrVal) ? (irrVal * 100).toFixed(2) + '%' : 'n/a'}<br>
      <b>Discounted Payback (periods):</b> ${payback ?? 'n/a'}`;
    let badge = '';
    if (irrVal > 0.15) badge = '<span class="badge badge-success">Attractive ROI</span>';
    else if (irrVal > 0.08) badge = '<span class="badge badge-warning">Moderate ROI</span>';
    else if (!isNaN(irrVal)) badge = '<span class="badge badge-danger">Low ROI</span>';
    else badge = '';
    document.getElementById('roiSummary').innerHTML = summary + badge;
    document.getElementById('roiPaybackTableWrap').innerHTML = tableHtml;
    renderRoiCharts(investment, repayments);
    if (!repayments.length || repayments.reduce((a, b) => a + b, 0) === 0) {
      document.getElementById('roiSummary').innerHTML += '<div class="alert alert-warning">No repayments scheduled. ROI cannot be calculated.</div>';
    }
  };

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
          if (tabId === "roi") setupRoiIrrSlider();
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

  function updateAllTabs() {
    renderRepaymentRows();
    updateLoanSummary();
    updateChartAndSummary();
    renderPnlTables();
    renderSummaryTab();
    renderRoiSection(false);
    renderTornadoChart();
  }
  updateAllTabs();

  setTimeout(setupRoiIrrSlider, 400);
});
