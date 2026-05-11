const STORAGE_KEY = "order-payment-tracker-v2";
const OLD_STORAGE_KEY = "order-payment-tracker-v1";
const USERS_KEY = "order-payment-users-v1";
const SESSION_KEY = "order-payment-session-v1";

const pageLabels = {
  dashboard: "總覽",
  orders: "訂單管理",
  payments: "付款紀錄",
  accounting: "會計檢視",
  import: "資料匯入",
  admin: "員工權限",
};

const allPages = Object.keys(pageLabels);
const departmentPermissions = {
  客服: ["dashboard", "orders", "payments"],
  OP: ["dashboard", "orders", "import"],
  會計: ["dashboard", "payments", "accounting"],
  管理: allPages,
  其他: ["dashboard"],
};

const today = () => new Date().toISOString().slice(0, 10);

const sampleData = {
  orders: [
    {
      id: crypto.randomUUID(),
      businessType: "公司辦理",
      groupNo: "VN-2026-001",
      customer: "晨星旅遊",
      serviceType: "企業簽證代辦",
      orderDate: "2026-05-01",
      processTime: "2026-05-01T10:30",
      firstSeenTime: "2026-05-03T09:00",
      customerAmount: 68000,
      incomeCurrency: "TWD",
      receivingAccount: "台新 812",
      customerPaymentStatus: "已付款",
      companyPaidAmount: 42000,
      supplierName: "",
      supplierAmount: 0,
      costCurrency: "TWD",
      supplierPaymentStatus: "不適用",
      note: "公司辦理案件，已完成收款。",
    },
    {
      id: crypto.randomUUID(),
      businessType: "VISA",
      groupNo: "VN-2026-002",
      customer: "越南電商 A",
      serviceType: "90天一次",
      orderDate: "2026-05-05",
      processTime: "",
      firstSeenTime: "",
      customerAmount: 125000,
      incomeCurrency: "TWD",
      receivingAccount: "國泰 013",
      customerPaymentStatus: "未付款",
      companyPaidAmount: 0,
      supplierName: "JNB",
      supplierAmount: 2400,
      costCurrency: "USD",
      supplierPaymentStatus: "已付款",
      note: "已提醒客戶補款。",
    },
    {
      id: crypto.randomUUID(),
      businessType: "FAST TRACK",
      groupNo: "FT-2026-003",
      customer: "好物選品",
      serviceType: "送機 S - Fast Track - 出境",
      orderDate: "2026-05-08",
      processTime: "",
      firstSeenTime: "",
      customerAmount: 43000,
      incomeCurrency: "TWD",
      receivingAccount: "",
      customerPaymentStatus: "部分付款",
      companyPaidAmount: 0,
      supplierName: "Airport NCC",
      supplierAmount: 600,
      costCurrency: "USD",
      supplierPaymentStatus: "未付款",
      note: "新客戶，收尾款後排單。",
    },
  ],
  payments: [],
};

sampleData.payments = [
  {
    id: crypto.randomUUID(),
    orderId: sampleData.orders[0].id,
    direction: "income",
    date: "2026-05-03",
    amount: 68000,
    method: "銀行轉帳",
    accountCode: "公司辦理收入",
    note: "末五碼 39218",
  },
  {
    id: crypto.randomUUID(),
    orderId: sampleData.orders[2].id,
    direction: "income",
    date: "2026-05-09",
    amount: 20000,
    method: "信用卡",
    accountCode: "Fast Track收入",
    note: "訂金",
  },
  {
    id: crypto.randomUUID(),
    orderId: sampleData.orders[1].id,
    direction: "expense",
    date: "2026-05-06",
    amount: 2400,
    method: "銀行轉帳",
    accountCode: "NCC成本",
    note: "JNB 已付款",
  },
];

let state = loadState();
let users = loadUsers();
let currentUser = getSessionUser();

const views = {
  dashboard: document.querySelector("#dashboardView"),
  orders: document.querySelector("#ordersView"),
  payments: document.querySelector("#paymentsView"),
  accounting: document.querySelector("#accountingView"),
  import: document.querySelector("#importView"),
  admin: document.querySelector("#adminView"),
};

const viewTitles = {
  ...pageLabels,
};

const moneyFormatter = new Intl.NumberFormat("zh-Hant-TW", {
  maximumFractionDigits: 0,
});

function loadState() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) return parseStoredState(stored);

  const oldStored = localStorage.getItem(OLD_STORAGE_KEY);
  if (oldStored) {
    const oldState = parseStoredState(oldStored);
    return {
      orders: oldState.orders.map((order) => normalizeOrder(order)),
      payments: oldState.payments.map((payment) => ({
        ...payment,
        direction: payment.direction || "income",
      })),
    };
  }

  return structuredClone(sampleData);
}

function parseStoredState(stored) {
  try {
    const parsed = JSON.parse(stored);
    return {
      orders: Array.isArray(parsed.orders) ? parsed.orders.map((order) => normalizeOrder(order)) : [],
      payments: Array.isArray(parsed.payments) ? parsed.payments : [],
    };
  } catch {
    return structuredClone(sampleData);
  }
}

function normalizeOrder(order) {
  const groupNo = order.groupNo || order.orderNo || "";
  const customerAmount = Number(order.customerAmount ?? order.totalAmount ?? 0);
  return {
    id: order.id || crypto.randomUUID(),
    businessType: order.businessType || "VISA",
    groupNo,
    customer: order.customer || "",
    serviceType: order.serviceType || order.service || "",
    orderDate: order.orderDate || today(),
    processTime: order.processTime || "",
    firstSeenTime: order.firstSeenTime || "",
    customerAmount,
    incomeCurrency: order.incomeCurrency || order.currency || "TWD",
    receivingAccount: order.receivingAccount || "",
    customerPaymentStatus: order.customerPaymentStatus || deriveCustomerPaymentStatus(order.id, customerAmount),
    companyPaidAmount: Number(order.companyPaidAmount || 0),
    supplierName: order.supplierName || "",
    supplierAmount: Number(order.supplierAmount || 0),
    costCurrency: order.costCurrency || "USD",
    supplierPaymentStatus: order.supplierPaymentStatus || (order.supplierAmount ? "未付款" : "不適用"),
    note: order.note || "",
  };
}

function deriveCustomerPaymentStatus(orderId, customerAmount) {
  return "未付款";
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadUsers() {
  const stored = localStorage.getItem(USERS_KEY);
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(normalizeUser);
    } catch {
      return [defaultAdminUser()];
    }
  }

  const initialUsers = [defaultAdminUser()];
  localStorage.setItem(USERS_KEY, JSON.stringify(initialUsers));
  return initialUsers;
}

function defaultAdminUser() {
  return {
    id: "admin-root",
    name: "主管理員",
    username: "admin",
    password: "admin123",
    department: "管理",
    role: "admin",
    permissions: allPages,
  };
}

function normalizeUser(user) {
  const role = user.role === "admin" ? "admin" : "staff";
  return {
    id: user.id || crypto.randomUUID(),
    name: user.name || user.username || "未命名",
    username: user.username || "",
    password: user.password || "",
    department: user.department || "其他",
    role,
    permissions: role === "admin" ? allPages : sanitizePermissions(user.permissions || departmentPermissions[user.department] || ["dashboard"]),
  };
}

function saveUsers() {
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
}

function getSessionUser() {
  const userId = localStorage.getItem(SESSION_KEY);
  if (!userId) return null;
  return users.find((user) => user.id === userId) || null;
}

function login(username, password) {
  const matchedUser = users.find((user) => user.username === username && user.password === password);
  if (!matchedUser) return false;
  currentUser = matchedUser;
  localStorage.setItem(SESSION_KEY, matchedUser.id);
  applyAuthState();
  return true;
}

function logout() {
  currentUser = null;
  localStorage.removeItem(SESSION_KEY);
  applyAuthState();
}

function hasPermission(viewName) {
  if (!currentUser) return false;
  return currentUser.role === "admin" || currentUser.permissions.includes(viewName);
}

function getFirstAllowedView() {
  return allPages.find((page) => hasPermission(page)) || "dashboard";
}

function sanitizePermissions(permissions) {
  const unique = [...new Set(permissions.filter((permission) => allPages.includes(permission)))];
  return unique.length ? unique : ["dashboard"];
}

function applyAuthState() {
  const loginScreen = document.querySelector("#loginScreen");
  const appShell = document.querySelector("#appShell");

  if (!currentUser) {
    loginScreen.classList.remove("is-hidden");
    appShell.classList.add("is-hidden");
    return;
  }

  loginScreen.classList.add("is-hidden");
  appShell.classList.remove("is-hidden");
  document.querySelector("#currentUserName").textContent = currentUser.name;
  document.querySelector("#currentUserMeta").textContent = `${currentUser.department}｜${currentUser.role === "admin" ? "主管理員" : "員工"}`;
  renderNavigation();
  render();
  switchView(getFirstAllowedView());
}

function renderNavigation() {
  document.querySelectorAll(".nav-tab").forEach((button) => {
    button.classList.toggle("is-hidden", !hasPermission(button.dataset.view));
  });
  document.querySelector("#exportCsvButton").classList.toggle("is-hidden", !hasPermission("accounting"));
  document.querySelector("#openOrderModalButton").classList.toggle("is-hidden", !hasPermission("orders"));
}

function getPaymentAmount(orderId, direction) {
  return state.payments
    .filter((payment) => payment.orderId === orderId && (payment.direction || "income") === direction)
    .reduce((sum, payment) => sum + Number(payment.amount || 0), 0);
}

function getOrderStatus(order) {
  const customerPaid = getPaymentAmount(order.id, "income");
  const supplierPaid = getPaymentAmount(order.id, "expense");
  const customerDone = order.customerPaymentStatus === "已付款" || customerPaid >= Number(order.customerAmount || 0);
  const supplierDone =
    order.supplierPaymentStatus === "已付款" ||
    order.supplierPaymentStatus === "不適用" ||
    supplierPaid >= Number(order.supplierAmount || 0);

  if (customerDone && supplierDone) return "paid";
  if (customerPaid > 0 || order.customerPaymentStatus === "部分付款") return "partial";
  return "unpaid";
}

function statusLabel(status) {
  return {
    paid: "已結清",
    partial: "部分處理",
    unpaid: "待處理",
    overdue: "逾期",
  }[status];
}

function formatMoney(amount, currency = "TWD") {
  return `${currency} ${moneyFormatter.format(Number(amount || 0))}`;
}

function render() {
  if (!currentUser) return;
  renderDashboard();
  renderOrders();
  renderPaymentSelect();
  renderPayments();
  renderAccounting();
  renderUsers();
}

function renderDashboard() {
  const totals = state.orders.reduce(
    (acc, order) => {
      acc.revenue += Number(order.customerAmount || 0);
      acc.cost += Number(order.supplierAmount || 0);
      acc.incomePaid += getPaymentAmount(order.id, "income");
      acc.expensePaid += getPaymentAmount(order.id, "expense");
      acc[getOrderStatus(order)] += 1;
      return acc;
    },
    { revenue: 0, cost: 0, incomePaid: 0, expensePaid: 0, paid: 0, partial: 0, unpaid: 0, overdue: 0 },
  );

  document.querySelector("#metricsGrid").innerHTML = [
    metricTemplate("客戶應收", formatMoney(totals.revenue), `${state.orders.length} 筆案件`),
    metricTemplate("客戶已收", formatMoney(totals.incomePaid), "實際登記收款"),
    metricTemplate("應付供應商", formatMoney(totals.cost, "USD"), "依各案成本幣別查看明細"),
    metricTemplate("未完成", `${totals.unpaid + totals.partial} 筆`, "客戶或供應商狀態待處理"),
  ].join("");

  const followups = state.orders
    .filter((order) => getOrderStatus(order) !== "paid")
    .sort((a, b) => String(a.orderDate || "").localeCompare(String(b.orderDate || "")))
    .slice(0, 6);

  document.querySelector("#followupList").innerHTML = followups.length
    ? followups.map(followupTemplate).join("")
    : emptyTemplate("目前沒有待追蹤案件");

  const recentPayments = [...state.payments]
    .sort((a, b) => String(b.date).localeCompare(String(a.date)))
    .slice(0, 6);

  document.querySelector("#recentPaymentsList").innerHTML = recentPayments.length
    ? recentPayments.map(paymentItemTemplate).join("")
    : emptyTemplate("目前沒有付款紀錄");
}

function metricTemplate(label, value, detail) {
  return `<article class="metric-card"><span>${label}</span><strong>${value}</strong><span>${detail}</span></article>`;
}

function followupTemplate(order) {
  const customerBalance = Math.max(Number(order.customerAmount || 0) - getPaymentAmount(order.id, "income"), 0);
  const supplierBalance = Math.max(Number(order.supplierAmount || 0) - getPaymentAmount(order.id, "expense"), 0);
  return `
    <article class="list-item">
      <header>
        <span>${escapeHtml(order.businessType)} · ${escapeHtml(order.groupNo)}</span>
        ${statusPill(getOrderStatus(order))}
      </header>
      <p>${escapeHtml(order.customer)}｜${escapeHtml(order.serviceType || "-")}</p>
      <p>客戶未收 ${formatMoney(customerBalance, order.incomeCurrency)}；NCC未付 ${formatMoney(supplierBalance, order.costCurrency)}</p>
    </article>
  `;
}

function renderOrders() {
  const keyword = document.querySelector("#orderSearchInput").value.trim().toLowerCase();
  const statusFilter = document.querySelector("#statusFilter").value;
  const monthFilter = document.querySelector("#monthFilter").value;

  const rows = state.orders
    .filter((order) => {
      const status = getOrderStatus(order);
      const searchable = [
        order.businessType,
        order.groupNo,
        order.customer,
        order.serviceType,
        order.supplierName,
        order.note,
      ].join(" ").toLowerCase();
      const matchKeyword = !keyword || searchable.includes(keyword);
      const matchStatus = statusFilter === "all" || status === statusFilter;
      const matchMonth = !monthFilter || String(order.orderDate || "").startsWith(monthFilter);
      return matchKeyword && matchStatus && matchMonth;
    })
    .sort((a, b) => String(b.orderDate || "").localeCompare(String(a.orderDate || "")));

  document.querySelector("#ordersTableBody").innerHTML = rows.length
    ? rows.map(orderRowTemplate).join("")
    : `<tr><td colspan="11">${emptyTemplate("沒有符合條件的案件")}</td></tr>`;
}

function orderRowTemplate(order) {
  const status = getOrderStatus(order);
  const paidIncome = getPaymentAmount(order.id, "income");
  const paidExpense = getPaymentAmount(order.id, "expense");

  return `
    <tr>
      <td><strong>${escapeHtml(order.businessType)}</strong><br><span class="muted">${order.orderDate || ""}</span></td>
      <td>${escapeHtml(order.groupNo || "-")}</td>
      <td>${escapeHtml(order.customer)}</td>
      <td>${escapeHtml(order.serviceType || "-")}</td>
      <td>${formatMoney(order.customerAmount, order.incomeCurrency)}<br><span class="muted">已收 ${formatMoney(paidIncome, order.incomeCurrency)}</span></td>
      <td>${paymentStatusPill(order.customerPaymentStatus)}</td>
      <td>${escapeHtml(order.supplierName || "-")}</td>
      <td>${formatMoney(order.supplierAmount, order.costCurrency)}<br><span class="muted">已付 ${formatMoney(paidExpense, order.costCurrency)}</span></td>
      <td>${paymentStatusPill(order.supplierPaymentStatus)}</td>
      <td>${statusPill(status)}</td>
      <td>
        <div class="row-actions">
          <button type="button" data-action="edit" data-id="${order.id}">編輯</button>
          <button type="button" data-action="quickpay" data-id="${order.id}">收款</button>
        </div>
      </td>
    </tr>
  `;
}

function statusPill(status) {
  return `<span class="status-pill status-${status}">${statusLabel(status)}</span>`;
}

function paymentStatusPill(status) {
  const normalized = status || "未付款";
  const className = normalized === "已付款" || normalized === "不適用" ? "paid" : normalized === "部分付款" ? "partial" : "unpaid";
  return `<span class="status-pill status-${className}">${escapeHtml(normalized)}</span>`;
}

function renderPaymentSelect() {
  const select = document.querySelector("#paymentOrderSelect");
  select.innerHTML = state.orders
    .map((order) => `<option value="${order.id}">${escapeHtml(order.businessType)} · ${escapeHtml(order.groupNo)} · ${escapeHtml(order.customer)}</option>`)
    .join("");
}

function renderPayments() {
  const payments = [...state.payments].sort((a, b) => String(b.date).localeCompare(String(a.date)));
  document.querySelector("#paymentsList").innerHTML = payments.length
    ? payments.map(paymentItemTemplate).join("")
    : emptyTemplate("尚未新增付款紀錄");
}

function paymentItemTemplate(payment) {
  const order = state.orders.find((item) => item.id === payment.orderId);
  const direction = (payment.direction || "income") === "expense" ? "供應商付款" : "客戶收款";
  const currency = (payment.direction || "income") === "expense" ? order?.costCurrency : order?.incomeCurrency;
  return `
    <article class="list-item">
      <header>
        <span>${payment.date} · ${escapeHtml(order?.groupNo || "未知案件")}</span>
        <span>${formatMoney(payment.amount, currency || "TWD")}</span>
      </header>
      <p>${escapeHtml(direction)}｜${escapeHtml(order?.customer || "-")}｜${escapeHtml(payment.method || "-")}｜${escapeHtml(payment.accountCode || "未填科目")}</p>
      ${payment.note ? `<p>${escapeHtml(payment.note)}</p>` : ""}
    </article>
  `;
}

function renderAccounting() {
  const start = document.querySelector("#accountingStartInput").value;
  const end = document.querySelector("#accountingEndInput").value;
  const method = document.querySelector("#accountingMethodFilter").value;

  const rows = state.payments
    .filter((payment) => {
      const matchStart = !start || payment.date >= start;
      const matchEnd = !end || payment.date <= end;
      const matchMethod = method === "all" || payment.method === method;
      return matchStart && matchEnd && matchMethod;
    })
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));

  const income = rows
    .filter((payment) => (payment.direction || "income") === "income")
    .reduce((sum, payment) => sum + Number(payment.amount || 0), 0);
  const expense = rows
    .filter((payment) => (payment.direction || "income") === "expense")
    .reduce((sum, payment) => sum + Number(payment.amount || 0), 0);

  document.querySelector("#accountingSummary").innerHTML = [
    summaryTemplate("收入合計", formatMoney(income)),
    summaryTemplate("成本合計", formatMoney(expense)),
    summaryTemplate("付款筆數", `${rows.length} 筆`),
  ].join("");

  document.querySelector("#accountingTableBody").innerHTML = rows.length
    ? rows.map(accountingRowTemplate).join("")
    : `<tr><td colspan="9">${emptyTemplate("沒有符合條件的付款紀錄")}</td></tr>`;
}

function summaryTemplate(label, value) {
  return `<article class="summary-chip"><span class="muted">${label}</span><strong>${value}</strong></article>`;
}

function accountingRowTemplate(payment) {
  const order = state.orders.find((item) => item.id === payment.orderId);
  const direction = (payment.direction || "income") === "expense" ? "供應商付款" : "客戶收款";
  const currency = (payment.direction || "income") === "expense" ? order?.costCurrency : order?.incomeCurrency;
  return `
    <tr>
      <td>${payment.date}</td>
      <td>${escapeHtml(order?.businessType || "-")}</td>
      <td>${escapeHtml(order?.groupNo || "-")}</td>
      <td>${escapeHtml(order?.customer || "-")}</td>
      <td>${direction}</td>
      <td>${escapeHtml(payment.method || "-")}</td>
      <td>${escapeHtml(payment.accountCode || "-")}</td>
      <td>${formatMoney(payment.amount, currency || "TWD")}</td>
      <td>${escapeHtml(payment.note || "")}</td>
    </tr>
  `;
}

function renderUsers() {
  const tableBody = document.querySelector("#usersTableBody");
  if (!tableBody) return;

  tableBody.innerHTML = users
    .map((user) => {
      const canDelete = currentUser?.id !== user.id && user.id !== "admin-root";
      return `
        <tr>
          <td><strong>${escapeHtml(user.name)}</strong></td>
          <td>${escapeHtml(user.username)}</td>
          <td>${escapeHtml(user.department)}</td>
          <td>${user.role === "admin" ? "主管理員" : "員工"}</td>
          <td>${user.permissions.map((permission) => pageLabels[permission]).join("、")}</td>
          <td>
            <div class="row-actions">
              <button type="button" data-user-action="edit" data-id="${user.id}">編輯</button>
              <button type="button" data-user-action="delete" data-id="${user.id}" ${canDelete ? "" : "disabled"}>刪除</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");
}

function emptyTemplate(text) {
  return `<div class="empty-state">${text}</div>`;
}

function openOrderDialog(order = null) {
  const dialog = document.querySelector("#orderDialog");
  document.querySelector("#orderModalTitle").textContent = order ? "編輯案件" : "新增案件";
  document.querySelector("#editingOrderId").value = order?.id || "";
  document.querySelector("#businessTypeInput").value = order?.businessType || "VISA";
  document.querySelector("#groupNoInput").value = order?.groupNo || "";
  document.querySelector("#customerInput").value = order?.customer || "";
  document.querySelector("#serviceTypeInput").value = order?.serviceType || "";
  document.querySelector("#orderDateInput").value = order?.orderDate || today();
  document.querySelector("#processTimeInput").value = order?.processTime || "";
  document.querySelector("#firstSeenTimeInput").value = order?.firstSeenTime || "";
  document.querySelector("#customerAmountInput").value = order?.customerAmount || "";
  document.querySelector("#incomeCurrencyInput").value = order?.incomeCurrency || "TWD";
  document.querySelector("#receivingAccountInput").value = order?.receivingAccount || "";
  document.querySelector("#customerPaymentStatusInput").value = order?.customerPaymentStatus || "未付款";
  document.querySelector("#companyPaidAmountInput").value = order?.companyPaidAmount || "";
  document.querySelector("#supplierNameInput").value = order?.supplierName || "";
  document.querySelector("#supplierAmountInput").value = order?.supplierAmount || "";
  document.querySelector("#costCurrencyInput").value = order?.costCurrency || "USD";
  document.querySelector("#supplierPaymentStatusInput").value = order?.supplierPaymentStatus || "未付款";
  document.querySelector("#orderNoteInput").value = order?.note || "";
  dialog.showModal();
}

function openUserDialog(user = null) {
  document.querySelector("#userModalTitle").textContent = user ? "編輯員工" : "新增員工";
  document.querySelector("#editingUserId").value = user?.id || "";
  document.querySelector("#employeeNameInput").value = user?.name || "";
  document.querySelector("#employeeUsernameInput").value = user?.username || "";
  document.querySelector("#employeeUsernameInput").disabled = Boolean(user);
  document.querySelector("#employeePasswordInput").value = "";
  document.querySelector("#employeePasswordInput").required = !user;
  document.querySelector("#employeeDepartmentInput").value = user?.department || "客服";
  document.querySelector("#employeeRoleInput").value = user?.role || "staff";
  document.querySelector("#userFormError").textContent = "";

  const permissions = user?.permissions || departmentPermissions[user?.department || "客服"];
  document.querySelectorAll('input[name="permissions"]').forEach((checkbox) => {
    checkbox.checked = permissions.includes(checkbox.value);
    checkbox.disabled = user?.role === "admin";
  });

  document.querySelector("#userDialog").showModal();
}

function handleUserSubmit(event) {
  event.preventDefault();
  if (!hasPermission("admin")) return;

  const id = document.querySelector("#editingUserId").value || crypto.randomUUID();
  const existing = users.find((user) => user.id === id);
  const username = document.querySelector("#employeeUsernameInput").value.trim();
  const password = document.querySelector("#employeePasswordInput").value;
  const role = document.querySelector("#employeeRoleInput").value;
  const selectedPermissions = [...document.querySelectorAll('input[name="permissions"]:checked')].map((checkbox) => checkbox.value);
  const formError = document.querySelector("#userFormError");

  if (!existing && users.some((user) => user.username === username)) {
    formError.textContent = "這個登入帳號已經存在。";
    return;
  }

  if (!existing && !password) {
    formError.textContent = "新增員工時必須設定密碼。";
    return;
  }

  const user = {
    id,
    name: document.querySelector("#employeeNameInput").value.trim(),
    username: existing?.username || username,
    password: password || existing?.password || "",
    department: document.querySelector("#employeeDepartmentInput").value,
    role,
    permissions: role === "admin" ? allPages : sanitizePermissions(selectedPermissions),
  };

  const existingIndex = users.findIndex((item) => item.id === id);
  if (existingIndex >= 0) {
    users[existingIndex] = user;
    if (currentUser?.id === id) currentUser = user;
  } else {
    users.push(user);
  }

  saveUsers();
  document.querySelector("#userDialog").close();
  applyAuthState();
}

function applyDepartmentTemplate(department) {
  const permissions = departmentPermissions[department] || departmentPermissions.其他;
  document.querySelectorAll('input[name="permissions"]').forEach((checkbox) => {
    checkbox.checked = permissions.includes(checkbox.value);
  });
}

function handleOrderSubmit(event) {
  event.preventDefault();
  if (!hasPermission("orders")) return;
  const id = document.querySelector("#editingOrderId").value || crypto.randomUUID();
  const order = {
    id,
    businessType: document.querySelector("#businessTypeInput").value,
    groupNo: document.querySelector("#groupNoInput").value.trim(),
    customer: document.querySelector("#customerInput").value.trim(),
    serviceType: document.querySelector("#serviceTypeInput").value.trim(),
    orderDate: document.querySelector("#orderDateInput").value,
    processTime: document.querySelector("#processTimeInput").value,
    firstSeenTime: document.querySelector("#firstSeenTimeInput").value,
    customerAmount: Number(document.querySelector("#customerAmountInput").value || 0),
    incomeCurrency: document.querySelector("#incomeCurrencyInput").value,
    receivingAccount: document.querySelector("#receivingAccountInput").value.trim(),
    customerPaymentStatus: document.querySelector("#customerPaymentStatusInput").value,
    companyPaidAmount: Number(document.querySelector("#companyPaidAmountInput").value || 0),
    supplierName: document.querySelector("#supplierNameInput").value.trim(),
    supplierAmount: Number(document.querySelector("#supplierAmountInput").value || 0),
    costCurrency: document.querySelector("#costCurrencyInput").value,
    supplierPaymentStatus: document.querySelector("#supplierPaymentStatusInput").value,
    note: document.querySelector("#orderNoteInput").value.trim(),
  };

  const existingIndex = state.orders.findIndex((item) => item.id === id);
  if (existingIndex >= 0) {
    state.orders[existingIndex] = order;
  } else {
    state.orders.push(order);
  }

  saveState();
  document.querySelector("#orderDialog").close();
  render();
}

function handlePaymentSubmit(event) {
  event.preventDefault();
  if (!hasPermission("payments")) return;
  const orderId = document.querySelector("#paymentOrderSelect").value;
  const direction = document.querySelector("#paymentDirectionInput").value;
  const payment = {
    id: crypto.randomUUID(),
    orderId,
    direction,
    date: document.querySelector("#paymentDateInput").value,
    amount: Number(document.querySelector("#paymentAmountInput").value || 0),
    method: document.querySelector("#paymentMethodInput").value,
    accountCode: document.querySelector("#accountCodeInput").value.trim(),
    note: document.querySelector("#paymentNoteInput").value.trim(),
  };

  state.payments.push(payment);
  syncPaymentStatus(orderId, direction);
  saveState();
  event.currentTarget.reset();
  document.querySelector("#paymentDateInput").value = today();
  render();
}

function syncPaymentStatus(orderId, direction) {
  const order = state.orders.find((item) => item.id === orderId);
  if (!order) return;

  if (direction === "income") {
    const paid = getPaymentAmount(orderId, "income");
    order.customerPaymentStatus = paid >= Number(order.customerAmount || 0) ? "已付款" : paid > 0 ? "部分付款" : "未付款";
  } else {
    const paid = getPaymentAmount(orderId, "expense");
    order.supplierPaymentStatus = paid >= Number(order.supplierAmount || 0) ? "已付款" : "未付款";
  }
}

function switchView(viewName) {
  if (!hasPermission(viewName)) {
    viewName = getFirstAllowedView();
  }

  Object.values(views).forEach((view) => view.classList.remove("is-active"));
  views[viewName].classList.add("is-active");
  document.querySelector("#viewTitle").textContent = viewTitles[viewName];

  document.querySelectorAll(".nav-tab").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.view === viewName);
  });
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && next === '"') {
      current += '"';
      i += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }

  row.push(current.trim());
  if (row.some(Boolean)) rows.push(row);
  if (rows.length < 2) return [];

  const headers = rows[0].map(normalizeHeader);
  return rows.slice(1).map((cells) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = cells[index] || "";
    });
    return record;
  });
}

function normalizeHeader(header) {
  const key = header.toLowerCase().replace(/\s+/g, "");
  const map = {
    日期ngàythánglàm: "orderDate",
    日期: "orderDate",
    團體編號mãđoàn: "groupNo",
    團體編號: "groupNo",
    客戶名稱tênkháchhàng: "customer",
    客戶名稱: "customer",
    簽證類型loạivisa: "serviceType",
    簽證類型: "serviceType",
    快速通關類型loạithôngquan: "serviceType",
    快速通關類型: "serviceType",
    收客人金額sốtiềnthukhách: "customerAmount",
    客戶收款金額sốtiềnthukhách: "customerAmount",
    客戶收款金額: "customerAmount",
    收客人金額: "customerAmount",
    幣別loạitiền: "incomeCurrency",
    收入幣別loạitiềnthu: "incomeCurrency",
    收入幣別: "incomeCurrency",
    幣別: "incomeCurrency",
    公司付款金額sốtiềncôngtythanhtoán: "companyPaidAmount",
    公司付款金額: "companyPaidAmount",
    收款帳戶tàikhoảnnhậntiền: "receivingAccount",
    收款帳戶: "receivingAccount",
    客戶付款狀態trạngtháithanhtoánkh: "customerPaymentStatus",
    客戶付款狀態: "customerPaymentStatus",
    供應商名稱tênnhàcungcấp: "supplierName",
    供應商名稱: "supplierName",
    應付供應商金額sốtiềntrảncc: "supplierAmount",
    應付供應商金額: "supplierAmount",
    成本幣別loạitiềnchi: "costCurrency",
    成本幣別: "costCurrency",
    供應商付款狀態trạngtháithanhtoánncc: "supplierPaymentStatus",
    供應商付款狀態: "supplierPaymentStatus",
    辦理時間thờigianlàm: "processTime",
    辦理時間: "processTime",
    初見時間thờigianra: "firstSeenTime",
    初見時間: "firstSeenTime",
    備註ghichú: "note",
    備註: "note",
  };
  return map[key] || key;
}

function importCsvData() {
  if (!hasPermission("import")) return;
  const groups = [
    { type: "公司辦理", records: parseCsv(document.querySelector("#companyCsvInput").value) },
    { type: "VISA", records: parseCsv(document.querySelector("#visaCsvInput").value) },
    { type: "FAST TRACK", records: parseCsv(document.querySelector("#fastTrackCsvInput").value) },
  ];

  const ordersByKey = new Map(state.orders.map((order) => [`${order.businessType}::${order.groupNo}`, order]));
  let importedCount = 0;

  groups.forEach((group) => {
    group.records.forEach((record) => {
      if (!record.groupNo && !record.customer) return;
      const businessType = group.type;
      const groupNo = record.groupNo || `${businessType}-${crypto.randomUUID().slice(0, 8)}`;
      const key = `${businessType}::${groupNo}`;
      const existing = ordersByKey.get(key);
      const order = {
        id: existing?.id || crypto.randomUUID(),
        businessType,
        groupNo,
        customer: record.customer || existing?.customer || "",
        serviceType: record.serviceType || existing?.serviceType || "",
        orderDate: normalizeDate(record.orderDate) || existing?.orderDate || today(),
        processTime: normalizeDateTime(record.processTime) || existing?.processTime || "",
        firstSeenTime: normalizeDateTime(record.firstSeenTime) || existing?.firstSeenTime || "",
        customerAmount: toNumber(record.customerAmount || existing?.customerAmount),
        incomeCurrency: record.incomeCurrency || existing?.incomeCurrency || "TWD",
        receivingAccount: record.receivingAccount || existing?.receivingAccount || "",
        customerPaymentStatus: record.customerPaymentStatus || existing?.customerPaymentStatus || "未付款",
        companyPaidAmount: toNumber(record.companyPaidAmount || existing?.companyPaidAmount),
        supplierName: record.supplierName || existing?.supplierName || "",
        supplierAmount: toNumber(record.supplierAmount || existing?.supplierAmount),
        costCurrency: record.costCurrency || existing?.costCurrency || "USD",
        supplierPaymentStatus: record.supplierPaymentStatus || existing?.supplierPaymentStatus || (businessType === "公司辦理" ? "不適用" : "未付款"),
        note: record.note || existing?.note || "",
      };

      if (existing) {
        Object.assign(existing, order);
      } else {
        state.orders.push(order);
        ordersByKey.set(key, order);
      }
      importedCount += 1;
    });
  });

  saveState();
  render();
  alert(`已匯入或更新 ${importedCount} 筆案件。`);
}

function normalizeDate(value) {
  if (!value) return "";
  const trimmed = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(trimmed)) {
    const [year, month, day] = trimmed.split("/");
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return trimmed;
}

function normalizeDateTime(value) {
  if (!value) return "";
  const trimmed = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(trimmed)) return trimmed;
  return trimmed.replace(" ", "T");
}

function toNumber(value) {
  return Number(String(value || "0").replace(/[^\d.-]/g, "")) || 0;
}

function exportCsv() {
  if (!hasPermission("accounting")) return;
  const headers = [
    "日期",
    "類別",
    "團體編號",
    "客戶名稱",
    "簽證/通關類型",
    "客戶收款金額",
    "收入幣別",
    "客戶付款狀態",
    "供應商名稱",
    "應付供應商金額",
    "成本幣別",
    "供應商付款狀態",
    "備註",
  ];
  const rows = state.orders.map((order) => [
    order.orderDate,
    order.businessType,
    order.groupNo,
    order.customer,
    order.serviceType,
    order.customerAmount,
    order.incomeCurrency,
    order.customerPaymentStatus,
    order.supplierName,
    order.supplierAmount,
    order.costCurrency,
    order.supplierPaymentStatus,
    order.note,
  ]);

  const csv = [headers, ...rows]
    .map((row) => row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `訂單付款追蹤-${today()}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

document.querySelectorAll(".nav-tab").forEach((button) => {
  button.addEventListener("click", () => switchView(button.dataset.view));
});

document.querySelector("#loginForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const username = document.querySelector("#loginUsernameInput").value.trim();
  const password = document.querySelector("#loginPasswordInput").value;
  const success = login(username, password);
  document.querySelector("#loginError").textContent = success ? "" : "帳號或密碼不正確。";
});

document.querySelector("#logoutButton").addEventListener("click", logout);
document.querySelector("#openOrderModalButton").addEventListener("click", () => {
  if (hasPermission("orders")) openOrderDialog();
});
document.querySelector("#closeOrderModalButton").addEventListener("click", () => document.querySelector("#orderDialog").close());
document.querySelector("#cancelOrderButton").addEventListener("click", () => document.querySelector("#orderDialog").close());
document.querySelector("#orderForm").addEventListener("submit", handleOrderSubmit);
document.querySelector("#paymentForm").addEventListener("submit", handlePaymentSubmit);
document.querySelector("#importCsvButton").addEventListener("click", importCsvData);
document.querySelector("#exportCsvButton").addEventListener("click", exportCsv);
document.querySelector("#openUserModalButton").addEventListener("click", () => openUserDialog());
document.querySelector("#closeUserModalButton").addEventListener("click", () => document.querySelector("#userDialog").close());
document.querySelector("#cancelUserButton").addEventListener("click", () => document.querySelector("#userDialog").close());
document.querySelector("#userForm").addEventListener("submit", handleUserSubmit);
document.querySelector("#employeeDepartmentInput").addEventListener("change", (event) => {
  if (document.querySelector("#employeeRoleInput").value !== "admin") applyDepartmentTemplate(event.target.value);
});
document.querySelector("#employeeRoleInput").addEventListener("change", (event) => {
  const isAdmin = event.target.value === "admin";
  document.querySelectorAll('input[name="permissions"]').forEach((checkbox) => {
    checkbox.checked = isAdmin || (departmentPermissions[document.querySelector("#employeeDepartmentInput").value] || []).includes(checkbox.value);
    checkbox.disabled = isAdmin;
  });
});
document.querySelector("#loadSampleButton").addEventListener("click", () => {
  state = structuredClone(sampleData);
  saveState();
  render();
});

document.querySelector("#ordersTableBody").addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button) return;

  const order = state.orders.find((item) => item.id === button.dataset.id);
  if (!order) return;

  if (button.dataset.action === "edit") {
    if (!hasPermission("orders")) return;
    openOrderDialog(order);
  }

  if (button.dataset.action === "quickpay") {
    if (!hasPermission("payments")) return;
    switchView("payments");
    document.querySelector("#paymentOrderSelect").value = order.id;
    document.querySelector("#paymentDirectionInput").value = "income";
    document.querySelector("#paymentAmountInput").value = Math.max(Number(order.customerAmount || 0) - getPaymentAmount(order.id, "income"), 0);
    document.querySelector("#paymentDateInput").value = today();
  }
});

document.querySelector("#usersTableBody").addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button || !hasPermission("admin")) return;

  const user = users.find((item) => item.id === button.dataset.id);
  if (!user) return;

  if (button.dataset.userAction === "edit") {
    openUserDialog(user);
  }

  if (button.dataset.userAction === "delete") {
    if (user.id === currentUser.id || user.id === "admin-root") return;
    users = users.filter((item) => item.id !== user.id);
    saveUsers();
    renderUsers();
  }
});

[
  "#orderSearchInput",
  "#statusFilter",
  "#monthFilter",
  "#accountingStartInput",
  "#accountingEndInput",
  "#accountingMethodFilter",
].forEach((selector) => {
  document.querySelector(selector).addEventListener("input", render);
});

document.querySelector("#paymentDateInput").value = today();
applyAuthState();
