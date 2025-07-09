function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const master = sheet.getSheetByName("Master");

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
      .setBackground("#e8f5e8");
  }

  const newRow = master.getLastRow();
  const data = master
    .getRange(newRow, 1, 1, master.getLastColumn())
    .getValues();
  monthlySheet.appendRow(data[0]);

  updatePersonalDashboard();
}

function updatePersonalDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = sheet.getSheetByName("Dashboard");
  const master = sheet.getSheetByName("Master");
  const config = sheet.getSheetByName("Config");

  const budget = parseFloat(config.getRange("B5").getValue());
  const data = master.getDataRange().getValues().slice(1); // Skip header

  // Process data with enhanced analytics
  const analytics = processExpenseData(data, sheet);

  // Clear and rebuild dashboard
  dashboard.clear();

  // Create comprehensive dashboard
  createHeader(dashboard);
  createQuickStats(dashboard, analytics, budget);
  createBudgetOverview(dashboard, analytics, budget);
  createTopCategories(dashboard, analytics);
  createRecentActivity(dashboard, analytics);
  createAdvancedCharts(dashboard, analytics);
  createSpendingPatterns(dashboard, analytics);
  createMonthlySnapshot(dashboard, analytics);
  createPredictiveAnalytics(dashboard, analytics, budget);
  createVendorAnalysis(dashboard, analytics);

  // Format the dashboard
  formatDashboard(dashboard);
}

function processExpenseData(data, sheet) {
  const analytics = {
    total: 0,
    totalTransactions: 0,
    byCategory: {},
    byMonth: {},
    byDate: {},
    byWeekday: { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }, // Sunday = 0
    byPaymentMethod: {},
    byVendor: {},
    byLocation: {},
    byHour: {},
    transactions: [],
    avgDaily: 0,
    avgWeekly: 0,
    avgTransaction: 0,
    thisMonth: 0,
    lastMonth: 0,
    thisWeek: 0,
    lastWeek: 0,
    topSpendingDay: { date: "", amount: 0 },
    daysActive: new Set(),
    weeklyTrend: [],
    monthlyGrowth: 0,
    spendingVelocity: 0,
    largestTransaction: { amount: 0, description: "", date: null },
    smallestTransaction: { amount: Infinity, description: "", date: null },
  };

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const lastMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  const lastMonthYear = currentMonth === 0 ? currentYear - 1 : currentYear;

  // Week calculations
  const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const twoWeeksAgo = new Date(now.getTime() - 14 * 24 * 60 * 60 * 1000);

  data.forEach((row) => {
    const timestamp = new Date(row[0]);
    const price = parseFloat(row[1]);
    const description = row[2] || "Unknown";
    const category = row[3] || "Other";
    const paymentMethod = row[4] || "Unknown";
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
      const weekday = timestamp.getDay();
      const hour = timestamp.getHours();

      analytics.total += price;
      analytics.totalTransactions++;
      analytics.daysActive.add(dateKey);

      // Category analysis
      analytics.byCategory[category] =
        (analytics.byCategory[category] || 0) + price;

      // Monthly analysis
      analytics.byMonth[monthKey] = (analytics.byMonth[monthKey] || 0) + price;

      // Daily analysis
      analytics.byDate[dateKey] = (analytics.byDate[dateKey] || 0) + price;

      // Weekday analysis
      analytics.byWeekday[weekday] += price;

      // Payment method analysis
      analytics.byPaymentMethod[paymentMethod] =
        (analytics.byPaymentMethod[paymentMethod] || 0) + price;

      // Vendor analysis
      analytics.byVendor[vendor] = (analytics.byVendor[vendor] || 0) + price;

      // Location analysis
      analytics.byLocation[location] =
        (analytics.byLocation[location] || 0) + price;

      // Hour analysis
      analytics.byHour[hour] = (analytics.byHour[hour] || 0) + price;

      // Current vs last month
      if (
        timestamp.getMonth() === currentMonth &&
        timestamp.getFullYear() === currentYear
      ) {
        analytics.thisMonth += price;
      } else if (
        timestamp.getMonth() === lastMonth &&
        timestamp.getFullYear() === lastMonthYear
      ) {
        analytics.lastMonth += price;
      }

      // Weekly analysis
      if (timestamp >= oneWeekAgo) {
        analytics.thisWeek += price;
      } else if (timestamp >= twoWeeksAgo) {
        analytics.lastWeek += price;
      }

      // Top spending day
      if (analytics.byDate[dateKey] > analytics.topSpendingDay.amount) {
        analytics.topSpendingDay = {
          date: dateKey,
          amount: analytics.byDate[dateKey],
        };
      }

      // Largest and smallest transactions
      if (price > analytics.largestTransaction.amount) {
        analytics.largestTransaction = {
          amount: price,
          description,
          date: timestamp,
        };
      }
      if (price < analytics.smallestTransaction.amount) {
        analytics.smallestTransaction = {
          amount: price,
          description,
          date: timestamp,
        };
      }

      // Store transaction
      analytics.transactions.push({
        date: timestamp,
        price,
        description,
        category,
        paymentMethod,
        vendor,
        location,
        dateKey,
        hour,
      });
    }
  });

  // Calculate derived metrics
  analytics.avgDaily = analytics.total / Math.max(analytics.daysActive.size, 1);
  analytics.avgWeekly = analytics.avgDaily * 7;
  analytics.avgTransaction =
    analytics.total / Math.max(analytics.totalTransactions, 1);
  analytics.monthlyGrowth =
    analytics.lastMonth > 0
      ? ((analytics.thisMonth - analytics.lastMonth) / analytics.lastMonth) *
        100
      : 0;
  analytics.spendingVelocity =
    analytics.thisWeek > 0
      ? (analytics.thisWeek / analytics.lastWeek) * 100 - 100
      : 0;

  // Sort transactions by date
  analytics.transactions.sort((a, b) => b.date - a.date);

  // Generate weekly trend data
  const last8Weeks = [];
  for (let i = 7; i >= 0; i--) {
    const weekStart = new Date(now.getTime() - i * 7 * 24 * 60 * 60 * 1000);
    const weekEnd = new Date(weekStart.getTime() + 7 * 24 * 60 * 60 * 1000);
    const weekSpending = analytics.transactions
      .filter((t) => t.date >= weekStart && t.date < weekEnd)
      .reduce((sum, t) => sum + t.price, 0);
    last8Weeks.push({
      week: `Week ${8 - i}`,
      amount: weekSpending,
      startDate: weekStart,
    });
  }
  analytics.weeklyTrend = last8Weeks;

  return analytics;
}

function createHeader(dashboard) {
  // Main title
  dashboard
    .getRange("A1:H1")
    .merge()
    .setValue("💰 Advanced Expense Analytics Dashboard")
    .setFontSize(22)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Current date
  dashboard
    .getRange("A2:H2")
    .merge()
    .setValue(`Updated: ${new Date().toLocaleString()}`)
    .setFontSize(11)
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center")
    .setFontStyle("italic");
}

function createQuickStats(dashboard, analytics, budget) {
  const row = 4;

  // Quick stats header
  dashboard
    .getRange(row, 1, 1, 8)
    .merge()
    .setValue("📊 Key Performance Indicators")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#34a853")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Stats boxes
  const stats = [
    {
      label: "Total Spent",
      value: `$${analytics.total.toFixed(2)}`,
      color: "#fce8e6",
    },
    {
      label: "This Month",
      value: `$${analytics.thisMonth.toFixed(2)}`,
      color: "#e8f5e8",
    },
    {
      label: "Daily Average",
      value: `$${analytics.avgDaily.toFixed(2)}`,
      color: "#fff3e0",
    },
    {
      label: "Avg Transaction",
      value: `$${analytics.avgTransaction.toFixed(2)}`,
      color: "#f3e5f5",
    },
  ];

  stats.forEach((stat, index) => {
    const col = index * 2 + 1;
    dashboard
      .getRange(row + 1, col, 1, 2)
      .merge()
      .setValue(stat.label)
      .setFontWeight("bold")
      .setBackground("#f8f9fa")
      .setHorizontalAlignment("center");

    dashboard
      .getRange(row + 2, col, 1, 2)
      .merge()
      .setValue(stat.value)
      .setFontSize(16)
      .setFontWeight("bold")
      .setBackground(stat.color)
      .setHorizontalAlignment("center");
  });
}

function createBudgetOverview(dashboard, analytics, budget) {
  const row = 8;

  // Budget header
  dashboard
    .getRange(row, 1, 1, 8)
    .merge()
    .setValue("🎯 Budget Analysis & Projections")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#ff9800")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const remaining = budget - analytics.thisMonth;
  const percentage = (analytics.thisMonth / budget) * 100;
  const status =
    percentage > 100
      ? "Over Budget!"
      : percentage > 80
      ? "Getting Close"
      : "On Track";
  const statusColor =
    percentage > 100 ? "#ffcdd2" : percentage > 80 ? "#fff3e0" : "#e8f5e8";

  // Calculate projected spending
  const daysInMonth = new Date(
    new Date().getFullYear(),
    new Date().getMonth() + 1,
    0
  ).getDate();
  const daysElapsed = new Date().getDate();
  const projectedSpending = analytics.thisMonth * (daysInMonth / daysElapsed);
  const projectedOverage = Math.max(0, projectedSpending - budget);

  const budgetData = [
    ["Monthly Budget", `$${budget.toFixed(2)}`, ""],
    ["Spent This Month", `$${analytics.thisMonth.toFixed(2)}`, ""],
    ["Remaining", `$${remaining.toFixed(2)}`, ""],
    [
      "Projected Total",
      `$${projectedSpending.toFixed(2)}`,
      projectedOverage > 0
        ? `⚠️ $${projectedOverage.toFixed(2)} over`
        : "✅ Within budget",
    ],
    ["Status", status, `${percentage.toFixed(1)}%`],
  ];

  budgetData.forEach((item, index) => {
    dashboard
      .getRange(row + 1 + index, 1, 1, 2)
      .merge()
      .setValue(item[0])
      .setFontWeight("bold")
      .setBackground("#f8f9fa");

    dashboard
      .getRange(row + 1 + index, 3, 1, 2)
      .merge()
      .setValue(item[1])
      .setBackground(index === 4 ? statusColor : "#ffffff");

    dashboard
      .getRange(row + 1 + index, 5, 1, 2)
      .merge()
      .setValue(item[2])
      .setBackground("#f8f9fa");
  });

  // Enhanced budget progress bar
  dashboard.getRange(row + 6, 1).setValue("Progress:");
  const progressBars = Math.min(Math.floor(percentage / 5), 20);
  const progressText = `${"█".repeat(progressBars)}${"░".repeat(
    20 - progressBars
  )} ${percentage.toFixed(1)}%`;
  dashboard
    .getRange(row + 6, 2, 1, 6)
    .merge()
    .setValue(progressText)
    .setFontFamily("Courier New")
    .setBackground(statusColor);
}

function createTopCategories(dashboard, analytics) {
  const row = 16;

  // Categories header
  dashboard
    .getRange(row, 1, 1, 4)
    .merge()
    .setValue("🏷️ Category Breakdown")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#9c27b0")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const categories = Object.entries(analytics.byCategory)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  dashboard
    .getRange(row + 1, 1)
    .setValue("Category")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 2)
    .setValue("Amount")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 3)
    .setValue("Percentage")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 4)
    .setValue("Transactions")
    .setFontWeight("bold");

  categories.forEach(([category, amount], index) => {
    const percentage = ((amount / analytics.total) * 100).toFixed(1);
    const transactionCount = analytics.transactions.filter(
      (t) => t.category === category
    ).length;

    dashboard.getRange(row + 2 + index, 1).setValue(category);
    dashboard.getRange(row + 2 + index, 2).setValue(`$${amount.toFixed(2)}`);
    dashboard.getRange(row + 2 + index, 3).setValue(`${percentage}%`);
    dashboard.getRange(row + 2 + index, 4).setValue(transactionCount);

    // Add color coding
    const colors = [
      "#ffebee",
      "#f3e5f5",
      "#e8f5e8",
      "#fff3e0",
      "#e3f2fd",
      "#fce4ec",
      "#f9fbe7",
      "#fff8e1",
    ];
    dashboard.getRange(row + 2 + index, 1, 1, 4).setBackground(colors[index]);
  });
}

function createRecentActivity(dashboard, analytics) {
  const row = 16;
  const col = 5;

  // Recent activity header
  dashboard
    .getRange(row, col, 1, 4)
    .merge()
    .setValue("⏰ Recent Activity")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#607d8b")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const recentTransactions = analytics.transactions.slice(0, 8);

  dashboard
    .getRange(row + 1, col)
    .setValue("Date")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, col + 1)
    .setValue("Description")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, col + 2)
    .setValue("Amount")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, col + 3)
    .setValue("Category")
    .setFontWeight("bold");

  recentTransactions.forEach((transaction, index) => {
    const displayDate = transaction.date.toLocaleDateString();
    const shortDesc =
      transaction.description.length > 12
        ? transaction.description.substring(0, 12) + "..."
        : transaction.description;

    dashboard.getRange(row + 2 + index, col).setValue(displayDate);
    dashboard.getRange(row + 2 + index, col + 1).setValue(shortDesc);
    dashboard
      .getRange(row + 2 + index, col + 2)
      .setValue(`$${transaction.price.toFixed(2)}`);
    dashboard.getRange(row + 2 + index, col + 3).setValue(transaction.category);

    // Alternate row colors
    if (index % 2 === 0) {
      dashboard.getRange(row + 2 + index, col, 1, 4).setBackground("#f8f9fa");
    }
  });
}

function createAdvancedCharts(dashboard, analytics) {
  dashboard.getCharts().forEach((chart) => dashboard.removeChart(chart));

  // Category pie chart
  createCategoryChart(dashboard, analytics);

  // Monthly trend chart
  createMonthlyChart(dashboard, analytics);

  // Weekly trend chart
  createWeeklyTrendChart(dashboard, analytics);

  // Spending by day of week chart
  createWeekdayChart(dashboard, analytics);

  // Spending by hour chart
  createHourlyChart(dashboard, analytics);
}

function createCategoryChart(dashboard, analytics) {
  const chartData = Object.entries(analytics.byCategory)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const range = dashboard.getRange(26, 1, chartData.length + 1, 2);
  range.clearContent();
  range.setValues([["Category", "Amount"], ...chartData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(1, 10, 0, 0)
    .setOption("title", "💰 Spending by Category")
    .setOption("pieHole", 0.4)
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
    .setOption("titleTextStyle", { fontSize: 16, bold: true })
    .setOption("legend", { position: "right", textStyle: { fontSize: 12 } })
    .setOption("pieSliceTextStyle", { fontSize: 10 })
    .build();

  dashboard.insertChart(chart);
}

function createMonthlyChart(dashboard, analytics) {
  const monthlyData = Object.entries(analytics.byMonth).sort().slice(-12); // Last 12 months

  const range = dashboard.getRange(26, 4, monthlyData.length + 1, 2);
  range.clearContent();
  range.setValues([["Month", "Amount"], ...monthlyData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(20, 10, 0, 0)
    .setOption("title", "📈 Monthly Spending Trend")
    .setOption("colors", ["#4285f4"])
    .setOption("titleTextStyle", { fontSize: 16, bold: true })
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { textStyle: { fontSize: 10 } })
    .setOption("vAxis", { format: "$#,###" })
    .build();

  dashboard.insertChart(chart);
}

function createWeeklyTrendChart(dashboard, analytics) {
  const weeklyData = analytics.weeklyTrend.map((week) => [
    week.week,
    week.amount,
  ]);

  const range = dashboard.getRange(45, 1, weeklyData.length + 1, 2);
  range.clearContent();
  range.setValues([["Week", "Amount"], ...weeklyData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(40, 10, 0, 0)
    .setOption("title", "📊 8-Week Spending Trend")
    .setOption("colors", ["#34a853"])
    .setOption("titleTextStyle", { fontSize: 16, bold: true })
    .setOption("legend", { position: "none" })
    .setOption("curveType", "function")
    .setOption("pointSize", 5)
    .setOption("vAxis", { format: "$#,###" })
    .build();

  dashboard.insertChart(chart);
}

function createWeekdayChart(dashboard, analytics) {
  const weekdays = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];
  const weekdayData = weekdays.map((day, index) => [
    day,
    analytics.byWeekday[index],
  ]);

  const range = dashboard.getRange(45, 4, weekdayData.length + 1, 2);
  range.clearContent();
  range.setValues([["Day", "Amount"], ...weekdayData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(60, 10, 0, 0)
    .setOption("title", "📅 Spending by Day of Week")
    .setOption("colors", ["#ff9800"])
    .setOption("titleTextStyle", { fontSize: 16, bold: true })
    .setOption("legend", { position: "none" })
    .setOption("vAxis", { format: "$#,###" })
    .build();

  dashboard.insertChart(chart);
}

function createHourlyChart(dashboard, analytics) {
  const hourlyData = [];
  for (let hour = 0; hour < 24; hour++) {
    hourlyData.push([`${hour}:00`, analytics.byHour[hour] || 0]);
  }

  const range = dashboard.getRange(65, 1, hourlyData.length + 1, 2);
  range.clearContent();
  range.setValues([["Hour", "Amount"], ...hourlyData]);

  const chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(range)
    .setPosition(80, 10, 0, 0)
    .setOption("title", "🕐 Spending by Hour of Day")
    .setOption("colors", ["#9c27b0"])
    .setOption("titleTextStyle", { fontSize: 16, bold: true })
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { textStyle: { fontSize: 8 } })
    .setOption("vAxis", { format: "$#,###" })
    .setOption("isStacked", false)
    .build();

  dashboard.insertChart(chart);
}

function createSpendingPatterns(dashboard, analytics) {
  const row = 85;

  // Spending patterns header
  dashboard
    .getRange(row, 1, 1, 8)
    .merge()
    .setValue("🔍 Spending Pattern Analysis")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#795548")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const patterns = [
    [
      "Highest Spending Day",
      analytics.topSpendingDay.date,
      `$${analytics.topSpendingDay.amount.toFixed(2)}`,
    ],
    [
      "Most Active Day",
      Object.entries(analytics.byWeekday).sort((a, b) => b[1] - a[1])[0][0],
      "",
    ],
    [
      "Peak Spending Hour",
      Object.entries(analytics.byHour).sort((a, b) => b[1] - a[1])[0][0] +
        ":00",
      "",
    ],
    [
      "Largest Transaction",
      analytics.largestTransaction.description,
      `$${analytics.largestTransaction.amount.toFixed(2)}`,
    ],
    [
      "Smallest Transaction",
      analytics.smallestTransaction.description,
      `$${analytics.smallestTransaction.amount.toFixed(2)}`,
    ],
    [
      "Weekly Growth",
      analytics.spendingVelocity > 0 ? "↗️ Increasing" : "↘️ Decreasing",
      `${Math.abs(analytics.spendingVelocity).toFixed(1)}%`,
    ],
  ];

  patterns.forEach((pattern, index) => {
    dashboard
      .getRange(row + 1 + index, 1, 1, 2)
      .merge()
      .setValue(pattern[0])
      .setFontWeight("bold")
      .setBackground("#f8f9fa");

    dashboard
      .getRange(row + 1 + index, 3, 1, 3)
      .merge()
      .setValue(pattern[1])
      .setBackground("#ffffff");

    dashboard
      .getRange(row + 1 + index, 6, 1, 2)
      .merge()
      .setValue(pattern[2])
      .setBackground(
        index === 5
          ? analytics.spendingVelocity > 0
            ? "#ffcdd2"
            : "#e8f5e8"
          : "#ffffff"
      );
  });
}

function createMonthlySnapshot(dashboard, analytics) {
  const row = 93;

  // Monthly comparison
  dashboard
    .getRange(row, 1, 1, 8)
    .merge()
    .setValue("📈 Monthly & Weekly Comparison")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#34a853")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const monthChange = analytics.thisMonth - analytics.lastMonth;
  const monthChangePercent =
    analytics.lastMonth > 0 ? (monthChange / analytics.lastMonth) * 100 : 0;
  const weekChange = analytics.thisWeek - analytics.lastWeek;
  const weekChangePercent =
    analytics.lastWeek > 0 ? (weekChange / analytics.lastWeek) * 100 : 0;

  const comparisonData = [
    [
      "This Month",
      `$${analytics.thisMonth.toFixed(2)}`,
      "Last Month",
      `$${analytics.lastMonth.toFixed(2)}`,
    ],
    [
      "Monthly Change",
      `${monthChange > 0 ? "↗️" : "↘️"} $${Math.abs(monthChange).toFixed(2)}`,
      "Change %",
      `${Math.abs(monthChangePercent).toFixed(1)}%`,
    ],
    [
      "This Week",
      `$${analytics.thisWeek.toFixed(2)}`,
      "Last Week",
      `$${analytics.lastWeek.toFixed(2)}`,
    ],
    [
      "Weekly Change",
      `${weekChange > 0 ? "↗️" : "↘️"} $${Math.abs(weekChange).toFixed(2)}`,
      "Change %",
      `${Math.abs(weekChangePercent).toFixed(1)}%`,
    ],
  ];

  comparisonData.forEach((item, index) => {
    dashboard
      .getRange(row + 1 + index, 1, 1, 2)
      .merge()
      .setValue(item[0])
      .setFontWeight("bold")
      .setBackground("#f8f9fa");

    dashboard
      .getRange(row + 1 + index, 3, 1, 2)
      .merge()
      .setValue(item[1])
      .setBackground("#ffffff");

    dashboard
      .getRange(row + 1 + index, 5, 1, 2)
      .merge()
      .setValue(item[2])
      .setFontWeight("bold")
      .setBackground("#f8f9fa");

    dashboard
      .getRange(row + 1 + index, 7, 1, 2)
      .merge()
      .setValue(item[3])
      .setBackground(
        index % 2 === 1
          ? index === 1
            ? monthChange > 0
              ? "#ffcdd2"
              : "#e8f5e8"
            : weekChange > 0
            ? "#ffcdd2"
            : "#e8f5e8"
          : "#ffffff"
      );
  });
}

function createPredictiveAnalytics(dashboard, analytics, budget) {
  const row = 99;

  // Predictive analytics header
  dashboard
    .getRange(row, 1, 1, 8)
    .merge()
    .setValue("🔮 Predictive Analytics & Insights")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#673ab7")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const daysInMonth = new Date(
    new Date().getFullYear(),
    new Date().getMonth() + 1,
    0
  ).getDate();
  const daysElapsed = new Date().getDate();
  const daysRemaining = daysInMonth - daysElapsed;

  const projectedMonthly = analytics.thisMonth * (daysInMonth / daysElapsed);
  const budgetBurnRate = analytics.thisMonth / daysElapsed;
  const daysToDepleteBudget =
    budget > analytics.thisMonth
      ? (budget - analytics.thisMonth) / budgetBurnRate
      : 0;

  const insights = [
    [
      "Projected Monthly Total",
      `${projectedMonthly.toFixed(2)}`,
      projectedMonthly > budget ? "⚠️ Over Budget" : "✅ Within Budget",
    ],
    ["Daily Burn Rate", `${budgetBurnRate.toFixed(2)}`, "per day"],
    [
      "Days to Budget Depletion",
      daysToDepleteBudget > 0
        ? `${Math.ceil(daysToDepleteBudget)} days`
        : "Budget Exceeded",
      "",
    ],
    [
      "Recommended Daily Limit",
      `${((budget - analytics.thisMonth) / Math.max(daysRemaining, 1)).toFixed(
        2
      )}`,
      "for remaining days",
    ],
    [
      "Spending Velocity",
      `${Math.abs(analytics.spendingVelocity).toFixed(1)}%`,
      analytics.spendingVelocity > 0 ? "↗️ Increasing" : "↘️ Decreasing",
    ],
    [
      "Category Concentration",
      Object.keys(analytics.byCategory).length,
      "categories tracked",
    ],
  ];

  insights.forEach((insight, index) => {
    dashboard
      .getRange(row + 1 + index, 1, 1, 2)
      .merge()
      .setValue(insight[0])
      .setFontWeight("bold")
      .setBackground("#f8f9fa");

    dashboard
      .getRange(row + 1 + index, 3, 1, 3)
      .merge()
      .setValue(insight[1])
      .setBackground("#ffffff");

    dashboard
      .getRange(row + 1 + index, 6, 1, 2)
      .merge()
      .setValue(insight[2])
      .setBackground(
        index === 0
          ? projectedMonthly > budget
            ? "#ffcdd2"
            : "#e8f5e8"
          : index === 4
          ? analytics.spendingVelocity > 0
            ? "#ffcdd2"
            : "#e8f5e8"
          : "#f8f9fa"
      );
  });
}

function createVendorAnalysis(dashboard, analytics) {
  const row = 107;

  // Vendor analysis header
  dashboard
    .getRange(row, 1, 1, 4)
    .merge()
    .setValue("🏪 Top Vendors & Payment Methods")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#e91e63")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const topVendors = Object.entries(analytics.byVendor)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  dashboard
    .getRange(row + 1, 1)
    .setValue("Vendor")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 2)
    .setValue("Amount")
    .setFontWeight("bold");

  topVendors.forEach(([vendor, amount], index) => {
    dashboard.getRange(row + 2 + index, 1).setValue(vendor);
    dashboard.getRange(row + 2 + index, 2).setValue(`${amount.toFixed(2)}`);
    dashboard
      .getRange(row + 2 + index, 1, 1, 2)
      .setBackground(index % 2 === 0 ? "#f8f9fa" : "#ffffff");
  });

  // Payment methods analysis
  dashboard
    .getRange(row, 5, 1, 4)
    .merge()
    .setValue("💳 Payment Method Breakdown")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#009688")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const paymentMethods = Object.entries(analytics.byPaymentMethod)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  dashboard
    .getRange(row + 1, 5)
    .setValue("Payment Method")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 6)
    .setValue("Amount")
    .setFontWeight("bold");
  dashboard
    .getRange(row + 1, 7)
    .setValue("Percentage")
    .setFontWeight("bold");

  paymentMethods.forEach(([method, amount], index) => {
    const percentage = ((amount / analytics.total) * 100).toFixed(1);
    dashboard.getRange(row + 2 + index, 5).setValue(method);
    dashboard.getRange(row + 2 + index, 6).setValue(`${amount.toFixed(2)}`);
    dashboard.getRange(row + 2 + index, 7).setValue(`${percentage}%`);
    dashboard
      .getRange(row + 2 + index, 5, 1, 3)
      .setBackground(index % 2 === 0 ? "#f8f9fa" : "#ffffff");
  });

  // Advanced insights
  dashboard
    .getRange(row + 8, 1, 1, 8)
    .merge()
    .setValue("💡 Advanced Insights")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#ff5722")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const advancedInsights = [
    `You've made ${analytics.totalTransactions} transactions across ${analytics.daysActive.size} active days`,
    `Your most expensive transaction was ${
      analytics.largestTransaction.description
    } for ${analytics.largestTransaction.amount.toFixed(2)}`,
    `You spend most on ${
      Object.entries(analytics.byWeekday).sort((a, b) => b[1] - a[1])[0][0] ===
      "0"
        ? "Sundays"
        : Object.entries(analytics.byWeekday).sort(
            (a, b) => b[1] - a[1]
          )[0][0] === "1"
        ? "Mondays"
        : Object.entries(analytics.byWeekday).sort(
            (a, b) => b[1] - a[1]
          )[0][0] === "2"
        ? "Tuesdays"
        : Object.entries(analytics.byWeekday).sort(
            (a, b) => b[1] - a[1]
          )[0][0] === "3"
        ? "Wednesdays"
        : Object.entries(analytics.byWeekday).sort(
            (a, b) => b[1] - a[1]
          )[0][0] === "4"
        ? "Thursdays"
        : Object.entries(analytics.byWeekday).sort(
            (a, b) => b[1] - a[1]
          )[0][0] === "5"
        ? "Fridays"
        : "Saturdays"
    } and between ${
      Object.entries(analytics.byHour).sort((a, b) => b[1] - a[1])[0][0]
    }:00-${
      parseInt(
        Object.entries(analytics.byHour).sort((a, b) => b[1] - a[1])[0][0]
      ) + 1
    }:00`,
    `Your spending has ${
      analytics.monthlyGrowth > 0 ? "increased" : "decreased"
    } by ${Math.abs(analytics.monthlyGrowth).toFixed(
      1
    )}% compared to last month`,
    `Top spending category "${
      Object.entries(analytics.byCategory).sort((a, b) => b[1] - a[1])[0][0]
    }" accounts for ${(
      (Object.entries(analytics.byCategory).sort((a, b) => b[1] - a[1])[0][1] /
        analytics.total) *
      100
    ).toFixed(1)}% of total expenses`,
    `You average ${(
      analytics.totalTransactions / analytics.daysActive.size
    ).toFixed(1)} transactions per active day`,
  ];

  advancedInsights.forEach((insight, index) => {
    dashboard
      .getRange(row + 9 + index, 1, 1, 8)
      .merge()
      .setValue(`• ${insight}`)
      .setBackground(index % 2 === 0 ? "#f8f9fa" : "#ffffff")
      .setWrap(true);
  });
}

function formatDashboard(dashboard) {
  // Auto-resize columns
  dashboard.autoResizeColumns(1, 8);

  // Set column widths for better readability
  dashboard.setColumnWidth(1, 140);
  dashboard.setColumnWidth(2, 120);
  dashboard.setColumnWidth(3, 120);
  dashboard.setColumnWidth(4, 120);
  dashboard.setColumnWidth(5, 140);
  dashboard.setColumnWidth(6, 120);
  dashboard.setColumnWidth(7, 120);
  dashboard.setColumnWidth(8, 120);

  // Add borders to main sections
  const ranges = [
    "A1:H2",
    "A4:H7",
    "A8:H15",
    "A16:H25",
    "A85:H92",
    "A93:H98",
    "A99:H106",
    "A107:H120",
  ];

  ranges.forEach((range) => {
    dashboard.getRange(range).setBorder(true, true, true, true, true, true);
  });

  // Freeze header
  dashboard.setFrozenRows(3);

  // Add conditional formatting for budget status
  const budgetRange = dashboard.getRange("A8:H15");
  const budgetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Over Budget")
    .setBackground("#ffcdd2")
    .setRanges([budgetRange])
    .build();

  const rules = dashboard.getConditionalFormatRules();
  rules.push(budgetRule);
  dashboard.setConditionalFormatRules(rules);

  // Add last updated timestamp
  dashboard
    .getRange("A120")
    .setValue(`Last Updated: ${new Date().toLocaleString()}`);
  dashboard.getRange("A120").setFontStyle("italic").setFontSize(10);
}

// Enhanced setup function
function setupExpenseTracker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // Create Config sheet if it doesn't exist
  let config = sheet.getSheetByName("Config");
  if (!config) {
    config = sheet.insertSheet("Config");
    config.getRange("A1").setValue("📊 Expense Tracker Configuration");
    config.getRange("A1").setFontSize(16).setFontWeight("bold");

    // Settings
    config.getRange("A3").setValue("Monthly Budget:");
    config.getRange("B3").setValue("$");
    config.getRange("B4").setValue(1000);
    config.getRange("A5").setValue("Currency:");
    config.getRange("B5").setValue("USD");
    config.getRange("A6").setValue("Alert Threshold:");
    config.getRange("B6").setValue("80%");
    config.getRange("A7").setValue("Tracking Started:");
    config.getRange("B7").setValue(new Date().toLocaleDateString());

    // Categories
    config.getRange("D1").setValue("💡 Suggested Categories");
    config.getRange("D1").setFontSize(14).setFontWeight("bold");
    const categories = [
      "Food & Dining",
      "Transportation",
      "Shopping",
      "Entertainment",
      "Bills & Utilities",
      "Health & Fitness",
      "Travel",
      "Education",
      "Insurance",
      "Other",
    ];
    categories.forEach((cat, index) => {
      config.getRange(3 + index, 4).setValue(cat);
    });

    // Payment Methods
    config.getRange("F1").setValue("💳 Payment Methods");
    config.getRange("F1").setFontSize(14).setFontWeight("bold");
    const paymentMethods = [
      "Credit Card",
      "Debit Card",
      "Cash",
      "Bank Transfer",
      "Digital Wallet",
      "Check",
    ];
    paymentMethods.forEach((method, index) => {
      config.getRange(3 + index, 6).setValue(method);
    });

    // Format config sheet
    config.getRange("A3:A7").setFontWeight("bold");
    config.getRange("B4").setNumberFormat("$#,##0.00");
    config.getRange("D1:F1").setBackground("#e8f5e8");
    config.getRange("D3:D12").setBackground("#f8f9fa");
    config.getRange("F3:F8").setBackground("#f8f9fa");
  }

  // Create Dashboard sheet if it doesn't exist
  let dashboard = sheet.getSheetByName("Dashboard");
  if (!dashboard) {
    dashboard = sheet.insertSheet("Dashboard");
  }

  // Create Master sheet if it doesn't exist
  let master = sheet.getSheetByName("Master");
  if (!master) {
    master = sheet.insertSheet("Master");
    const headers = [
      "Timestamp",
      "Amount",
      "Description",
      "Category",
      "Payment Method",
      "Vendor",
      "Location",
    ];
    master.getRange(1, 1, 1, headers.length).setValues([headers]);
    master
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#e8f5e8");

    // Add sample data for testing
    const sampleData = [
      [
        new Date(),
        25.5,
        "Grocery Shopping",
        "Food & Dining",
        "Credit Card",
        "SuperMart",
        "Downtown",
      ],
      [
        new Date(Date.now() - 86400000),
        45.0,
        "Gas Station",
        "Transportation",
        "Debit Card",
        "Shell",
        "Highway",
      ],
      [
        new Date(Date.now() - 172800000),
        12.99,
        "Netflix Subscription",
        "Entertainment",
        "Credit Card",
        "Netflix",
        "Online",
      ],
      [
        new Date(Date.now() - 259200000),
        89.99,
        "Sneakers",
        "Shopping",
        "Credit Card",
        "Nike Store",
        "Mall",
      ],
      [
        new Date(Date.now() - 345600000),
        150.0,
        "Electric Bill",
        "Bills & Utilities",
        "Bank Transfer",
        "Power Company",
        "Online",
      ],
    ];

    master
      .getRange(2, 1, sampleData.length, sampleData[0].length)
      .setValues(sampleData);
  }

  // Run initial dashboard update
  updatePersonalDashboard();

  // Show completion message
  SpreadsheetApp.getUi().alert(
    "✅ Enhanced Expense Tracker Setup Complete!",
    "Your advanced expense tracker is now ready with:\n\n" +
      "• Comprehensive analytics dashboard\n" +
      "• Multiple chart visualizations\n" +
      "• Predictive insights\n" +
      "• Spending pattern analysis\n" +
      "• Budget tracking & projections\n" +
      "• Vendor & payment method analysis\n\n" +
      "Start adding expenses to see the analytics in action!",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Utility function to create form trigger
function createFormTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create("Expense Tracker Form");

  // Add form questions
  form.addDateTimeItem().setTitle("Date & Time").setRequired(true);
  form.addTextItem().setTitle("Amount ($)").setRequired(true);
  form.addTextItem().setTitle("Description").setRequired(true);
  form
    .addListItem()
    .setTitle("Category")
    .setChoiceValues([
      "Food & Dining",
      "Transportation",
      "Shopping",
      "Entertainment",
      "Bills & Utilities",
      "Health & Fitness",
      "Travel",
      "Education",
      "Insurance",
      "Other",
    ])
    .setRequired(true);
  form
    .addListItem()
    .setTitle("Payment Method")
    .setChoiceValues([
      "Credit Card",
      "Debit Card",
      "Cash",
      "Bank Transfer",
      "Digital Wallet",
      "Check",
    ])
    .setRequired(true);
  form.addTextItem().setTitle("Vendor");
  form.addTextItem().setTitle("Location");

  // Link form to spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  // Create trigger
  ScriptApp.newTrigger("onFormSubmit").create();

  return form.getEditUrl();
}
