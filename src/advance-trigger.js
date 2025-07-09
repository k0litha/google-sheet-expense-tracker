function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = sheet.getSheetByName("Dashboard");
  const master = sheet.getSheetByName("Master");
  const config = sheet.getSheetByName("Config");

  const row = e.values;
  const date = new Date(row[0]); // Timestamp
  const monthName = Utilities.formatDate(
    date,
    sheet.getSpreadsheetTimeZone(),
    "MMMM yyyy"
  );

  let monthlySheet = sheet.getSheetByName(monthName);
  if (!monthlySheet) {
    monthlySheet = sheet.insertSheet(monthName);
    const headers = master
      .getRange(1, 1, 1, master.getLastColumn())
      .getValues();
    monthlySheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    monthlySheet
      .getRange(1, 1, 1, headers[0].length)
      .setFontWeight("bold")
      .setBackground("#d9ead3");
  }

  const newRow = master.getLastRow();
  const data = master
    .getRange(newRow, 1, 1, master.getLastColumn())
    .getValues();
  monthlySheet.appendRow(data[0]);

  updateAdvancedDashboard();
}

function updateAdvancedDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = sheet.getSheetByName("Dashboard");
  const master = sheet.getSheetByName("Master");
  const config = sheet.getSheetByName("Config");

  const budget = parseFloat(config.getRange("B5").getValue()) || 100000;
  const data = master.getDataRange().getValues().slice(1); // Skip header

  // Initialize analytics objects
  const analytics = initializeAnalytics();

  // Process all data
  data.forEach((row) => {
    processTransactionData(row, analytics, sheet);
  });

  // Calculate advanced metrics
  const advancedMetrics = calculateAdvancedMetrics(analytics, budget);

  // Clear and rebuild dashboard
  dashboard.clear();

  // Create comprehensive dashboard layout
  createDashboardHeader(dashboard);
  createKPISection(dashboard, advancedMetrics);
  createBudgetAnalysis(dashboard, advancedMetrics, budget);
  createTrendAnalysis(dashboard, analytics);
  createCategoryAnalysis(dashboard, analytics);
  createTimeAnalysis(dashboard, analytics);
  createPredictiveAnalysis(dashboard, analytics);
  createAnomalyDetection(dashboard, analytics);
  createExpenseBreakdown(dashboard, analytics);

  // Create all visualizations
  createAllCharts(dashboard, analytics);

  // Apply conditional formatting
  applyConditionalFormatting(dashboard);

  // Auto-resize columns
  dashboard.autoResizeColumns(1, 20);
}

function initializeAnalytics() {
  return {
    total: 0,
    days: new Set(),
    byCategory: {},
    byMonth: {},
    byDate: {},
    byWeek: {},
    byDayOfWeek: {},
    byHour: {},
    transactions: [],
    dailySpends: [],
    weeklySpends: [],
    monthlySpends: [],
    categoryTrends: {},
    recurringExpenses: {},
    expenseSize: { small: 0, medium: 0, large: 0 },
    paymentMethods: {},
    vendors: {},
    locations: {},
  };
}

function processTransactionData(row, analytics, sheet) {
  const timestamp = new Date(row[0]);
  const price = parseFloat(row[1]);
  const description = row[2] || "Unknown";
  const category = row[3] || "Other";
  const paymentMethod = row[4] || "Cash";
  const vendor = row[5] || "Unknown";
  const location = row[6] || "Unknown";

  if (!isNaN(price) && price > 0) {
    const dateKey = Utilities.formatDate(
      timestamp,
      sheet.getSpreadsheetTimeZone(),
      "yyyy-MM-dd"
    );
    const monthKey = Utilities.formatDate(
      timestamp,
      sheet.getSpreadsheetTimeZone(),
      "MMM yyyy"
    );
    const weekKey = getWeekKey(timestamp);
    const dayOfWeek = timestamp.getDay();
    const hour = timestamp.getHours();

    // Basic aggregations
    analytics.total += price;
    analytics.days.add(dateKey);

    // Category analysis
    analytics.byCategory[category] =
      (analytics.byCategory[category] || 0) + price;

    // Time-based analysis
    analytics.byMonth[monthKey] = (analytics.byMonth[monthKey] || 0) + price;
    analytics.byDate[dateKey] = (analytics.byDate[dateKey] || 0) + price;
    analytics.byWeek[weekKey] = (analytics.byWeek[weekKey] || 0) + price;
    analytics.byDayOfWeek[dayOfWeek] =
      (analytics.byDayOfWeek[dayOfWeek] || 0) + price;
    analytics.byHour[hour] = (analytics.byHour[hour] || 0) + price;

    // Payment method analysis
    analytics.paymentMethods[paymentMethod] =
      (analytics.paymentMethods[paymentMethod] || 0) + price;

    // Vendor analysis
    analytics.vendors[vendor] = (analytics.vendors[vendor] || 0) + price;

    // Location analysis
    analytics.locations[location] =
      (analytics.locations[location] || 0) + price;

    // Expense size categorization
    if (price < 50) analytics.expenseSize.small += price;
    else if (price < 200) analytics.expenseSize.medium += price;
    else analytics.expenseSize.large += price;

    // Store transaction details
    analytics.transactions.push({
      date: timestamp,
      price,
      description,
      category,
      paymentMethod,
      vendor,
      location,
      dateKey,
      monthKey,
      weekKey,
    });
  }
}

function calculateAdvancedMetrics(analytics, budget) {
  const totalDays = analytics.days.size || 1;
  const avgDailySpend = analytics.total / totalDays;
  const daysInMonth = 30;
  const projectedMonthlySpend = avgDailySpend * daysInMonth;

  // Calculate volatility (standard deviation of daily spends)
  const dailyAmounts = Object.values(analytics.byDate);
  const volatility = calculateStandardDeviation(dailyAmounts);

  // Calculate growth rate
  const monthlyData = Object.entries(analytics.byMonth).sort();
  const growthRate = calculateGrowthRate(monthlyData);

  // Budget utilization
  const budgetUtilization = (analytics.total / budget) * 100;

  // Spending efficiency (based on category distribution)
  const efficiency = calculateSpendingEfficiency(analytics.byCategory);

  // Top spending day
  const topSpendingDay = Object.entries(analytics.byDate).reduce(
    (a, b) => (a[1] > b[1] ? a : b),
    ["", 0]
  );

  return {
    total: analytics.total,
    avgDailySpend,
    projectedMonthlySpend,
    volatility,
    growthRate,
    budgetUtilization,
    efficiency,
    topSpendingDay,
    totalDays,
    avgTransactionSize: analytics.total / analytics.transactions.length,
    mostExpensiveTransaction: Math.max(
      ...analytics.transactions.map((t) => t.price)
    ),
    totalTransactions: analytics.transactions.length,
  };
}

function createDashboardHeader(dashboard) {
  // Main header
  dashboard
    .getRange("A1:T1")
    .merge()
    .setValue("🏦 ADVANCED EXPENSE ANALYTICS DASHBOARD")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Sub-header with current date
  dashboard
    .getRange("A2:T2")
    .merge()
    .setValue(
      `Generated on: ${new Date().toLocaleDateString()} | Real-time Analytics`
    )
    .setFontSize(12)
    .setBackground("#e3f2fd")
    .setHorizontalAlignment("center");
}

function createKPISection(dashboard, metrics) {
  const kpiRow = 4;

  // KPI Headers
  const kpiHeaders = [
    "💰 Total Spend",
    "📈 Daily Avg",
    "🎯 Projected Monthly",
    "📊 Volatility",
    "🔄 Growth Rate",
    "⚡ Efficiency",
    "🏆 Top Day",
    "📝 Transactions",
  ];

  const kpiValues = [
    `$${metrics.total.toFixed(2)}`,
    `$${metrics.avgDailySpend.toFixed(2)}`,
    `$${metrics.projectedMonthlySpend.toFixed(2)}`,
    `${metrics.volatility.toFixed(1)}%`,
    `${metrics.growthRate.toFixed(1)}%`,
    `${metrics.efficiency.toFixed(1)}%`,
    `$${metrics.topSpendingDay[1].toFixed(2)}`,
    metrics.totalTransactions,
  ];

  // Create KPI cards
  for (let i = 0; i < kpiHeaders.length; i++) {
    const col = i * 2 + 1;
    dashboard
      .getRange(kpiRow, col, 1, 2)
      .merge()
      .setValue(kpiHeaders[i])
      .setFontWeight("bold")
      .setBackground("#f8f9fa")
      .setHorizontalAlignment("center");

    dashboard
      .getRange(kpiRow + 1, col, 1, 2)
      .merge()
      .setValue(kpiValues[i])
      .setFontSize(14)
      .setFontWeight("bold")
      .setBackground("#e8f5e8")
      .setHorizontalAlignment("center");
  }
}

function createBudgetAnalysis(dashboard, metrics, budget) {
  const startRow = 7;

  dashboard
    .getRange(startRow, 1, 1, 4)
    .merge()
    .setValue("💳 BUDGET ANALYSIS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#ff9800")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const budgetData = [
    ["Budget Amount", `$${budget.toFixed(2)}`],
    ["Current Spend", `$${metrics.total.toFixed(2)}`],
    ["Remaining", `$${(budget - metrics.total).toFixed(2)}`],
    ["Utilization", `${metrics.budgetUtilization.toFixed(1)}%`],
    [
      "Days to Budget",
      Math.ceil((budget - metrics.total) / metrics.avgDailySpend),
    ],
    [
      "Projected vs Budget",
      `${((metrics.projectedMonthlySpend / budget) * 100).toFixed(1)}%`,
    ],
  ];

  for (let i = 0; i < budgetData.length; i++) {
    dashboard
      .getRange(startRow + 1 + i, 1)
      .setValue(budgetData[i][0])
      .setFontWeight("bold");
    dashboard.getRange(startRow + 1 + i, 2).setValue(budgetData[i][1]);
  }
}

function createTrendAnalysis(dashboard, analytics) {
  const startRow = 7;
  const startCol = 6;

  dashboard
    .getRange(startRow, startCol, 1, 4)
    .merge()
    .setValue("📈 TREND ANALYSIS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#4caf50")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Weekly trends
  const weeklyData = Object.entries(analytics.byWeek).sort();
  const weeklyTrend = calculateTrend(weeklyData);

  dashboard
    .getRange(startRow + 1, startCol)
    .setValue("Weekly Trend")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, startCol + 1)
    .setValue(weeklyTrend > 0 ? "↗️ Increasing" : "↘️ Decreasing");

  // Category trends
  const topCategory = Object.entries(analytics.byCategory).sort(
    (a, b) => b[1] - a[1]
  )[0];

  dashboard
    .getRange(startRow + 2, startCol)
    .setValue("Top Category")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 2, startCol + 1)
    .setValue(`${topCategory[0]} ($${topCategory[1].toFixed(2)})`);

  // Day of week analysis
  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const maxDaySpend = Object.entries(analytics.byDayOfWeek).reduce(
    (a, b) => (a[1] > b[1] ? a : b),
    [0, 0]
  );

  dashboard
    .getRange(startRow + 3, startCol)
    .setValue("Peak Spending Day")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 3, startCol + 1)
    .setValue(`${dayNames[maxDaySpend[0]]} ($${maxDaySpend[1].toFixed(2)})`);
}

function createCategoryAnalysis(dashboard, analytics) {
  const startRow = 7;
  const startCol = 11;

  dashboard
    .getRange(startRow, startCol, 1, 4)
    .merge()
    .setValue("🏷️ CATEGORY INSIGHTS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#9c27b0")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const categoryData = Object.entries(analytics.byCategory)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  dashboard
    .getRange(startRow + 1, startCol)
    .setValue("Top Categories")
    .setFontWeight("bold");

  for (let i = 0; i < categoryData.length; i++) {
    const percentage = ((categoryData[i][1] / analytics.total) * 100).toFixed(
      1
    );
    dashboard
      .getRange(startRow + 2 + i, startCol)
      .setValue(`${i + 1}. ${categoryData[i][0]}`);
    dashboard
      .getRange(startRow + 2 + i, startCol + 1)
      .setValue(`$${categoryData[i][1].toFixed(2)} (${percentage}%)`);
  }
}

function createTimeAnalysis(dashboard, analytics) {
  const startRow = 7;
  const startCol = 16;

  dashboard
    .getRange(startRow, startCol, 1, 4)
    .merge()
    .setValue("⏰ TIME PATTERNS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#f44336")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Peak spending hour
  const peakHour = Object.entries(analytics.byHour).reduce(
    (a, b) => (a[1] > b[1] ? a : b),
    [0, 0]
  );

  dashboard
    .getRange(startRow + 1, startCol)
    .setValue("Peak Hour")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, startCol + 1)
    .setValue(`${peakHour[0]}:00 ($${peakHour[1].toFixed(2)})`);

  // Expense size distribution
  const sizeData = analytics.expenseSize;
  const sizeTotal = sizeData.small + sizeData.medium + sizeData.large;

  dashboard
    .getRange(startRow + 2, startCol)
    .setValue("Small Expenses (<$50)")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 2, startCol + 1)
    .setValue(`${((sizeData.small / sizeTotal) * 100).toFixed(1)}%`);

  dashboard
    .getRange(startRow + 3, startCol)
    .setValue("Medium ($50-200)")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 3, startCol + 1)
    .setValue(`${((sizeData.medium / sizeTotal) * 100).toFixed(1)}%`);

  dashboard
    .getRange(startRow + 4, startCol)
    .setValue("Large (>$200)")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 4, startCol + 1)
    .setValue(`${((sizeData.large / sizeTotal) * 100).toFixed(1)}%`);
}

function createPredictiveAnalysis(dashboard, analytics) {
  const startRow = 15;

  dashboard
    .getRange(startRow, 1, 1, 8)
    .merge()
    .setValue("🔮 PREDICTIVE ANALYTICS & FORECASTING")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#607d8b")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Simple linear regression for next month prediction
  const monthlyData = Object.entries(analytics.byMonth)
    .sort()
    .map(([month, amount], index) => ({ x: index, y: amount }));

  const forecast = calculateLinearRegression(monthlyData);
  const nextMonthPrediction =
    forecast.slope * monthlyData.length + forecast.intercept;

  dashboard
    .getRange(startRow + 1, 1)
    .setValue("Next Month Forecast")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 2)
    .setValue(`$${nextMonthPrediction.toFixed(2)}`);

  // Seasonal patterns
  const seasonalSpending = calculateSeasonalPatterns(analytics.byMonth);
  dashboard
    .getRange(startRow + 2, 1)
    .setValue("Seasonal Pattern")
    .setFontWeight("bold");
  dashboard.getRange(startRow + 2, 2).setValue(seasonalSpending);

  // Spending momentum
  const recentTrend = calculateRecentTrend(analytics.byDate);
  dashboard
    .getRange(startRow + 3, 1)
    .setValue("Recent Momentum")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 3, 2)
    .setValue(recentTrend > 0 ? "📈 Increasing" : "📉 Decreasing");
}

function createAnomalyDetection(dashboard, analytics) {
  const startRow = 15;
  const startCol = 10;

  dashboard
    .getRange(startRow, startCol, 1, 6)
    .merge()
    .setValue("🚨 ANOMALY DETECTION")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#e91e63")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Detect unusual spending days
  const dailyAmounts = Object.entries(analytics.byDate);
  const avgDaily =
    dailyAmounts.reduce((sum, [, amount]) => sum + amount, 0) /
    dailyAmounts.length;
  const stdDev = calculateStandardDeviation(
    dailyAmounts.map(([, amount]) => amount)
  );

  const anomalies = dailyAmounts
    .filter(([, amount]) => Math.abs(amount - avgDaily) > 2 * stdDev)
    .sort((a, b) => b[1] - a[1]);

  dashboard
    .getRange(startRow + 1, startCol)
    .setValue("Unusual Spending Days")
    .setFontWeight("bold");

  for (let i = 0; i < Math.min(3, anomalies.length); i++) {
    dashboard
      .getRange(startRow + 2 + i, startCol)
      .setValue(`${anomalies[i][0]}: $${anomalies[i][1].toFixed(2)}`);
  }
}

function createExpenseBreakdown(dashboard, analytics) {
  const startRow = 20;

  dashboard
    .getRange(startRow, 1, 1, 20)
    .merge()
    .setValue("💼 DETAILED EXPENSE BREAKDOWN")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#3f51b5")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Top vendors
  const topVendors = Object.entries(analytics.vendors)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  dashboard
    .getRange(startRow + 1, 1)
    .setValue("Top Vendors")
    .setFontWeight("bold");
  for (let i = 0; i < topVendors.length; i++) {
    dashboard
      .getRange(startRow + 2 + i, 1)
      .setValue(`${i + 1}. ${topVendors[i][0]}`);
    dashboard
      .getRange(startRow + 2 + i, 2)
      .setValue(`$${topVendors[i][1].toFixed(2)}`);
  }

  // Payment methods
  const paymentMethods = Object.entries(analytics.paymentMethods).sort(
    (a, b) => b[1] - a[1]
  );

  dashboard
    .getRange(startRow + 1, 4)
    .setValue("Payment Methods")
    .setFontWeight("bold");
  for (let i = 0; i < paymentMethods.length && i < 5; i++) {
    dashboard.getRange(startRow + 2 + i, 4).setValue(`${paymentMethods[i][0]}`);
    dashboard
      .getRange(startRow + 2 + i, 5)
      .setValue(`$${paymentMethods[i][1].toFixed(2)}`);
  }

  // Recent transactions
  const recentTransactions = analytics.transactions
    .sort((a, b) => b.date - a.date)
    .slice(0, 5);

  dashboard
    .getRange(startRow + 1, 7)
    .setValue("Recent Transactions")
    .setFontWeight("bold");
  for (let i = 0; i < recentTransactions.length; i++) {
    dashboard
      .getRange(startRow + 2 + i, 7)
      .setValue(recentTransactions[i].description);
    dashboard
      .getRange(startRow + 2 + i, 8)
      .setValue(`$${recentTransactions[i].price.toFixed(2)}`);
  }
}

function createAllCharts(dashboard, analytics) {
  // Create multiple charts with different data
  createCategoryPieChart(dashboard, analytics.byCategory);
  createMonthlyTrendChart(dashboard, analytics.byMonth);
  createDailySpendingChart(dashboard, analytics.byDate);
  createWeeklyAnalysisChart(dashboard, analytics.byWeek);
  createHourlyPatternChart(dashboard, analytics.byHour);
  createPaymentMethodChart(dashboard, analytics.paymentMethods);
  createExpenseSizeChart(dashboard, analytics.expenseSize);
  createDayOfWeekChart(dashboard, analytics.byDayOfWeek);
}

function createCategoryPieChart(dashboard, categoryData) {
  const chartData = Object.entries(categoryData)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const range = dashboard.getRange(30, 1, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Category", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(30, 1, 0, 0)
    .setOption("title", "📊 Spending by Category")
    .setOption("pieHole", 0.3)
    .setOption("colors", [
      "#FF6B6B",
      "#4ECDC4",
      "#45B7D1",
      "#96CEB4",
      "#FECA57",
      "#FF9FF3",
      "#54A0FF",
      "#5F27CD",
    ])
    .build();

  dashboard.insertChart(chart);
}

function createMonthlyTrendChart(dashboard, monthlyData) {
  const chartData = Object.entries(monthlyData).sort();

  const range = dashboard.getRange(30, 4, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Month", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(30, 4, 0, 0)
    .setOption("title", "📈 Monthly Spending Trend")
    .setOption("curveType", "function")
    .setOption("colors", ["#FF6B6B"])
    .build();

  dashboard.insertChart(chart);
}

function createDailySpendingChart(dashboard, dailyData) {
  const chartData = Object.entries(dailyData).sort().slice(-30); // Last 30 days

  const range = dashboard.getRange(30, 7, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Date", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(range)
    .setPosition(30, 7, 0, 0)
    .setOption("title", "📅 Daily Spending (Last 30 Days)")
    .setOption("colors", ["#4ECDC4"])
    .build();

  dashboard.insertChart(chart);
}

function createWeeklyAnalysisChart(dashboard, weeklyData) {
  const chartData = Object.entries(weeklyData).sort();

  const range = dashboard.getRange(45, 1, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Week", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(45, 1, 0, 0)
    .setOption("title", "📊 Weekly Spending Analysis")
    .setOption("colors", ["#45B7D1"])
    .build();

  dashboard.insertChart(chart);
}

function createHourlyPatternChart(dashboard, hourlyData) {
  const chartData = [];
  for (let i = 0; i < 24; i++) {
    chartData.push([`${i}:00`, hourlyData[i] || 0]);
  }

  const range = dashboard.getRange(45, 4, 25, 2);
  range.clearContent();
  range.setValues([["Hour", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(range)
    .setPosition(45, 4, 0, 0)
    .setOption("title", "⏰ Hourly Spending Pattern")
    .setOption("colors", ["#96CEB4"])
    .build();

  dashboard.insertChart(chart);
}

function createPaymentMethodChart(dashboard, paymentData) {
  const chartData = Object.entries(paymentData);

  const range = dashboard.getRange(45, 7, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Payment Method", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.DOUGHNUT)
    .addRange(range)
    .setPosition(45, 7, 0, 0)
    .setOption("title", "💳 Payment Methods")
    .setOption("colors", ["#FECA57", "#FF9FF3", "#54A0FF"])
    .build();

  dashboard.insertChart(chart);
}

function createExpenseSizeChart(dashboard, sizeData) {
  const chartData = [
    ["Small (<$50)", sizeData.small],
    ["Medium ($50-200)", sizeData.medium],
    ["Large (>$200)", sizeData.large],
  ];

  const range = dashboard.getRange(60, 1, 4, 2);
  range.clearContent();
  range.setValues([["Size", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(range)
    .setPosition(60, 1, 0, 0)
    .setOption("title", "💰 Expense Size Distribution")
    .setOption("colors", ["#5F27CD"])
    .build();

  dashboard.insertChart(chart);
}

function createDayOfWeekChart(dashboard, dayData) {
  const dayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];
  const chartData = dayNames.map((day, index) => [day, dayData[index] || 0]);

  const range = dashboard.getRange(60, 4, 8, 2);
  range.clearContent();
  range.setValues([["Day", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(60, 4, 0, 0)
    .setOption("title", "📅 Spending by Day of Week")
    .setOption("colors", ["#FF6B6B"])
    .build();

  dashboard.insertChart(chart);
}

function applyConditionalFormatting(dashboard) {
  // Apply color coding for budget status
  const budgetRange = dashboard.getRange(8, 2, 6, 2);
  const budgetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Over Budget")
    .setBackground("#ffebee")
    .setFontColor("#c62828")
    .setRanges([budgetRange])
    .build();

  dashboard.setConditionalFormatRules([budgetRule]);
}

// Helper functions
function getWeekKey(date) {
  const week = getWeekNumber(date);
  return `${date.getFullYear()}-W${week}`;
}

function getWeekNumber(date) {
  const d = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  );
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function calculateStandardDeviation(values) {
  if (values.length === 0) return 0;
  const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
  const variance =
    values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) /
    values.length;
  return Math.sqrt(variance);
}

function calculateGrowthRate(monthlyData) {
  if (monthlyData.length < 2) return 0;
  const firstMonth = monthlyData[0][1];
  const lastMonth = monthlyData[monthlyData.length - 1][1];
  return ((lastMonth - firstMonth) / firstMonth) * 100;
}

function calculateSpendingEfficiency(categoryData) {
  // Simple efficiency metric based on category distribution
  const essential = [
    "Food",
    "Transportation",
    "Utilities",
    "Healthcare",
    "Housing",
  ];
  const totalSpend = Object.values(categoryData).reduce(
    (sum, val) => sum + val,
    0
  );
  const essentialSpend = essential.reduce(
    (sum, cat) => sum + (categoryData[cat] || 0),
    0
  );
  return (essentialSpend / totalSpend) * 100;
}

function calculateTrend(data) {
  if (data.length < 2) return 0;
  const values = data.map(([, value]) => value);
  const n = values.length;
  const sumX = (n * (n - 1)) / 2;
  const sumY = values.reduce((sum, val) => sum + val, 0);
  const sumXY = values.reduce((sum, val, i) => sum + val * i, 0);
  const sumXX = (n * (n - 1) * (2 * n - 1)) / 6;

  return (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
}

function calculateLinearRegression(data) {
  if (data.length === 0) return { slope: 0, intercept: 0 };

  const n = data.length;
  const sumX = data.reduce((sum, point) => sum + point.x, 0);
  const sumY = data.reduce((sum, point) => sum + point.y, 0);
  const sumXY = data.reduce((sum, point) => sum + point.x * point.y, 0);
  const sumXX = data.reduce((sum, point) => sum + point.x * point.x, 0);

  const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  return { slope, intercept };
}

function calculateSeasonalPatterns(monthlyData) {
  const seasons = {
    Winter: ["Dec", "Jan", "Feb"],
    Spring: ["Mar", "Apr", "May"],
    Summer: ["Jun", "Jul", "Aug"],
    Fall: ["Sep", "Oct", "Nov"],
  };

  const seasonalSpending = {};

  Object.entries(monthlyData).forEach(([month, amount]) => {
    const monthAbbr = month.split(" ")[0];
    for (const [season, months] of Object.entries(seasons)) {
      if (months.includes(monthAbbr)) {
        seasonalSpending[season] = (seasonalSpending[season] || 0) + amount;
        break;
      }
    }
  });

  const maxSeason = Object.entries(seasonalSpending).reduce(
    (a, b) => (a[1] > b[1] ? a : b),
    ["", 0]
  );

  return maxSeason[0] || "Unknown";
}

function calculateRecentTrend(dailyData) {
  const recentDays = Object.entries(dailyData)
    .sort()
    .slice(-7) // Last 7 days
    .map(([, amount]) => amount);

  if (recentDays.length < 2) return 0;

  const firstHalf = recentDays.slice(0, Math.floor(recentDays.length / 2));
  const secondHalf = recentDays.slice(Math.floor(recentDays.length / 2));

  const firstAvg =
    firstHalf.reduce((sum, val) => sum + val, 0) / firstHalf.length;
  const secondAvg =
    secondHalf.reduce((sum, val) => sum + val, 0) / secondHalf.length;

  return secondAvg - firstAvg;
}

// Additional advanced analytics functions
function createSpendingGoalsTracker(dashboard, analytics, budget) {
  const startRow = 75;

  dashboard
    .getRange(startRow, 1, 1, 10)
    .merge()
    .setValue("🎯 SPENDING GOALS & TARGETS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#2196f3")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Calculate progress towards different goals
  const monthlyBudget = budget;
  const weeklyBudget = budget / 4.33;
  const dailyBudget = budget / 30;

  const currentMonth = new Date().getMonth();
  const currentMonthData = Object.entries(analytics.byMonth)
    .filter(([month]) => month.includes(new Date().getFullYear().toString()))
    .reduce((sum, [, amount]) => sum + amount, 0);

  const goals = [
    {
      name: "Monthly Budget",
      target: monthlyBudget,
      current: currentMonthData,
      progress: (currentMonthData / monthlyBudget) * 100,
    },
    {
      name: "Weekly Average",
      target: weeklyBudget,
      current:
        Object.values(analytics.byWeek).reduce((sum, val) => sum + val, 0) /
        Object.keys(analytics.byWeek).length,
      progress:
        (Object.values(analytics.byWeek).reduce((sum, val) => sum + val, 0) /
          Object.keys(analytics.byWeek).length /
          weeklyBudget) *
        100,
    },
    {
      name: "Daily Target",
      target: dailyBudget,
      current: analytics.total / analytics.days.size,
      progress: (analytics.total / analytics.days.size / dailyBudget) * 100,
    },
  ];

  goals.forEach((goal, index) => {
    const row = startRow + 2 + index;
    dashboard.getRange(row, 1).setValue(goal.name).setFontWeight("bold");
    dashboard.getRange(row, 2).setValue(`${goal.target.toFixed(2)}`);
    dashboard.getRange(row, 3).setValue(`${goal.current.toFixed(2)}`);
    dashboard.getRange(row, 4).setValue(`${goal.progress.toFixed(1)}%`);

    // Color coding based on progress
    const color =
      goal.progress > 100
        ? "#ffcdd2"
        : goal.progress > 80
        ? "#fff3e0"
        : "#e8f5e8";
    dashboard.getRange(row, 1, 1, 4).setBackground(color);
  });
}

function createRecurringExpenseAnalysis(dashboard, analytics) {
  const startRow = 80;

  dashboard
    .getRange(startRow, 1, 1, 10)
    .merge()
    .setValue("🔄 RECURRING EXPENSE ANALYSIS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#795548")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Identify potential recurring expenses
  const vendorFrequency = {};
  const categoryFrequency = {};

  analytics.transactions.forEach((transaction) => {
    const vendor = transaction.vendor;
    const category = transaction.category;

    if (!vendorFrequency[vendor]) {
      vendorFrequency[vendor] = { count: 0, total: 0, amounts: [] };
    }
    vendorFrequency[vendor].count++;
    vendorFrequency[vendor].total += transaction.price;
    vendorFrequency[vendor].amounts.push(transaction.price);

    if (!categoryFrequency[category]) {
      categoryFrequency[category] = { count: 0, total: 0 };
    }
    categoryFrequency[category].count++;
    categoryFrequency[category].total += transaction.price;
  });

  // Find recurring vendors (appeared more than 3 times)
  const recurringVendors = Object.entries(vendorFrequency)
    .filter(([vendor, data]) => data.count >= 3)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 5);

  dashboard
    .getRange(startRow + 1, 1)
    .setValue("Recurring Vendors")
    .setFontWeight("bold");

  recurringVendors.forEach(([vendor, data], index) => {
    const avgAmount = data.total / data.count;
    const row = startRow + 2 + index;
    dashboard.getRange(row, 1).setValue(`${vendor} (${data.count}x)`);
    dashboard.getRange(row, 2).setValue(`${data.total.toFixed(2)}`);
    dashboard.getRange(row, 3).setValue(`${avgAmount.toFixed(2)} avg`);
  });
}

function createExpenseVariabilityAnalysis(dashboard, analytics) {
  const startRow = 85;

  dashboard
    .getRange(startRow, 1, 1, 10)
    .merge()
    .setValue("📊 EXPENSE VARIABILITY ANALYSIS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#009688")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Calculate variability for each category
  const categoryVariability = {};

  Object.keys(analytics.byCategory).forEach((category) => {
    const categoryTransactions = analytics.transactions
      .filter((t) => t.category === category)
      .map((t) => t.price);

    if (categoryTransactions.length > 1) {
      const stdDev = calculateStandardDeviation(categoryTransactions);
      const mean =
        categoryTransactions.reduce((sum, val) => sum + val, 0) /
        categoryTransactions.length;
      const coefficientOfVariation = (stdDev / mean) * 100;

      categoryVariability[category] = {
        mean,
        stdDev,
        cv: coefficientOfVariation,
        count: categoryTransactions.length,
      };
    }
  });

  // Sort by coefficient of variation (most variable first)
  const sortedVariability = Object.entries(categoryVariability)
    .sort((a, b) => b[1].cv - a[1].cv)
    .slice(0, 5);

  dashboard
    .getRange(startRow + 1, 1)
    .setValue("Most Variable Categories")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 2)
    .setValue("Avg Amount")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 3)
    .setValue("Variability")
    .setFontWeight("bold");

  sortedVariability.forEach(([category, data], index) => {
    const row = startRow + 2 + index;
    dashboard.getRange(row, 1).setValue(category);
    dashboard.getRange(row, 2).setValue(`${data.mean.toFixed(2)}`);
    dashboard.getRange(row, 3).setValue(`${data.cv.toFixed(1)}%`);
  });
}

function createMonthlyComparison(dashboard, analytics) {
  const startRow = 90;

  dashboard
    .getRange(startRow, 1, 1, 15)
    .merge()
    .setValue("📈 MONTH-OVER-MONTH COMPARISON")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#ff5722")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const monthlyData = Object.entries(analytics.byMonth).sort().slice(-6); // Last 6 months

  dashboard
    .getRange(startRow + 1, 1)
    .setValue("Month")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 2)
    .setValue("Amount")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 3)
    .setValue("Change")
    .setFontWeight("bold");
  dashboard
    .getRange(startRow + 1, 4)
    .setValue("% Change")
    .setFontWeight("bold");

  monthlyData.forEach(([month, amount], index) => {
    const row = startRow + 2 + index;
    dashboard.getRange(row, 1).setValue(month);
    dashboard.getRange(row, 2).setValue(`${amount.toFixed(2)}`);

    if (index > 0) {
      const prevAmount = monthlyData[index - 1][1];
      const change = amount - prevAmount;
      const percentChange = (change / prevAmount) * 100;

      dashboard.getRange(row, 3).setValue(`${change.toFixed(2)}`);
      dashboard.getRange(row, 4).setValue(`${percentChange.toFixed(1)}%`);

      // Color coding for changes
      const changeColor = change > 0 ? "#ffcdd2" : "#c8e6c9";
      dashboard.getRange(row, 3, 1, 2).setBackground(changeColor);
    }
  });
}

// Enhanced main function call
function updateAdvancedDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = sheet.getSheetByName("Dashboard");
  const master = sheet.getSheetByName("Master");
  const config = sheet.getSheetByName("Config");

  const budget = parseFloat(config.getRange("B4").getValue()) || 1000;
  const data = master.getDataRange().getValues().slice(1);

  const analytics = initializeAnalytics();

  data.forEach((row) => {
    processTransactionData(row, analytics, sheet);
  });

  const advancedMetrics = calculateAdvancedMetrics(analytics, budget);

  dashboard.clear();

  // Create all dashboard sections
  createDashboardHeader(dashboard);
  createKPISection(dashboard, advancedMetrics);
  createBudgetAnalysis(dashboard, advancedMetrics, budget);
  createTrendAnalysis(dashboard, analytics);
  createCategoryAnalysis(dashboard, analytics);
  createTimeAnalysis(dashboard, analytics);
  createPredictiveAnalysis(dashboard, analytics);
  createAnomalyDetection(dashboard, analytics);
  createExpenseBreakdown(dashboard, analytics);
  createSpendingGoalsTracker(dashboard, analytics, budget);
  createRecurringExpenseAnalysis(dashboard, analytics);
  createExpenseVariabilityAnalysis(dashboard, analytics);
  createMonthlyComparison(dashboard, analytics);

  // Create all visualizations
  createAllCharts(dashboard, analytics);

  // Apply formatting
  applyConditionalFormatting(dashboard);
  dashboard.autoResizeColumns(1, 20);

  // Add timestamp
  dashboard
    .getRange(1, 21)
    .setValue(`Last Updated: ${new Date().toLocaleString()}`);
}
