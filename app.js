// Configuration
const SPREADSHEET_ID = "1vbbeEhfw2BdVyUsv3bUewfh6SyCnWpyjIIODOxIOCEc";

// Global Application State
let projectsData = [];
let filteredProjects = [];
let deptChart = null;
let statusChart = null;

// DOM Elements
const valTotalProjects = document.getElementById("valTotalProjects");
const valTotalBudget = document.getElementById("valTotalBudget");
const valSpent = document.getElementById("valSpent");
const valRemaining = document.getElementById("valRemaining");
const valProgress = document.getElementById("valProgress");

const footerTotalProjects = document.getElementById("footerTotalProjects");
const footerTotalBudget = document.getElementById("footerTotalBudget");
const footerSpent = document.getElementById("footerSpent");
const footerRemaining = document.getElementById("footerRemaining");
const footerProgress = document.getElementById("footerProgress");

const tableBody = document.getElementById("tableBody");
const tableSearch = document.getElementById("tableSearch");
const filterDept = document.getElementById("filterDept");
const filterStatus = document.getElementById("filterStatus");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const refreshBtn = document.getElementById("refreshBtn");
const themeToggleBtn = document.getElementById("themeToggleBtn");
const syncStatus = document.getElementById("syncStatus");

// Modal Elements
const projectModal = document.getElementById("projectModal");
const modalCloseBtn = document.getElementById("modalCloseBtn");
const modalProjId = document.getElementById("modalProjId");
const modalProjName = document.getElementById("modalProjName");
const modalProgressText = document.getElementById("modalProgressText");
const modalProgressFill = document.getElementById("modalProgressFill");
const modalOwner = document.getElementById("modalOwner");
const modalDept = document.getElementById("modalDept");
const modalBudget = document.getElementById("modalBudget");
const modalSpent = document.getElementById("modalSpent");
const modalRemaining = document.getElementById("modalRemaining");
const modalStatusBadge = document.getElementById("modalStatusBadge");

// Initialize the Application
document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    loadDashboardData();
    
    // Set up event listeners
    refreshBtn.addEventListener("click", refreshData);
    themeToggleBtn.addEventListener("click", toggleTheme);
    tableSearch.addEventListener("input", applyFilters);
    filterDept.addEventListener("change", applyFilters);
    filterStatus.addEventListener("change", applyFilters);
    resetFiltersBtn.addEventListener("click", resetFilters);
    
    // Modal events
    modalCloseBtn.addEventListener("click", closeModal);
    projectModal.addEventListener("click", (e) => {
        if (e.target === projectModal) closeModal();
    });
});

// Theme Management
function initTheme() {
    const savedTheme = localStorage.getItem("theme") || "dark";
    document.documentElement.setAttribute("data-theme", savedTheme);
    updateThemeIcon(savedTheme);
}

function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute("data-theme");
    const newTheme = currentTheme === "light" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", newTheme);
    localStorage.setItem("theme", newTheme);
    updateThemeIcon(newTheme);
    
    // Re-render charts to pick up new theme variables/colors if initialized
    if (projectsData.length > 0) {
        renderCharts();
    }
}

function updateThemeIcon(theme) {
    const icon = themeToggleBtn.querySelector("i");
    if (theme === "light") {
        icon.className = "fa-solid fa-sun";
    } else {
        icon.className = "fa-solid fa-moon";
    }
}

// Fetch and Parse Data via JSONP to bypass CORS (especially from file:// origins)
function loadDashboardData() {
    showLoadingState();
    
    return new Promise((resolve, reject) => {
        // Clear any previous script tag
        const oldScript = document.getElementById("gviz-script");
        if (oldScript) oldScript.remove();
        
        // Define global callback function
        window.handleGoogleSheetResponse = function(data) {
            clearTimeout(timeoutId);
            
            // Clean up script tag
            const scriptTag = document.getElementById("gviz-script");
            if (scriptTag) scriptTag.remove();
            
            if (data && data.status === "ok" && data.table) {
                processAndRenderData(data.table);
                updateSyncStatus(true);
                resolve();
            } else {
                console.error("Google Sheet responded with error status:", data);
                showErrorState("ข้อมูล Google Sheet มีโครงสร้างไม่ถูกต้อง");
                updateSyncStatus(false);
                reject(new Error("Invalid structure"));
            }
        };
        
        // Set timeout to handle network failure
        const timeoutId = setTimeout(() => {
            const scriptTag = document.getElementById("gviz-script");
            if (scriptTag) {
                scriptTag.remove();
                showErrorState("การเชื่อมต่อหมดเวลา (Timeout) กรุณาตรวจสอบลิงก์หรืออินเทอร์เน็ต");
                updateSyncStatus(false);
                reject(new Error("Timeout"));
            }
        }, 10000); // 10 seconds timeout
        
        // Inject dynamic script tag
        const script = document.createElement("script");
        script.id = "gviz-script";
        script.src = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=responseHandler:handleGoogleSheetResponse&t=${new Date().getTime()}`;
        script.onerror = () => {
            clearTimeout(timeoutId);
            const scriptTag = document.getElementById("gviz-script");
            if (scriptTag) scriptTag.remove();
            showErrorState("ไม่สามารถเชื่อมต่อเพื่อดึงข้อมูลจาก Google Sheets ได้ (CORS/Network Error)");
            updateSyncStatus(false);
            reject(new Error("CORS/Network Error"));
        };
        
        document.body.appendChild(script);
    });
}

function updateSyncStatus(success) {
    const dateStr = new Date().toLocaleTimeString("th-TH", { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    if (success) {
        syncStatus.innerHTML = `<span></span> อัปเดตล่าสุด: ${dateStr} น.`;
        syncStatus.style.color = "var(--success)";
        syncStatus.querySelector("span").style.backgroundColor = "var(--success)";
        syncStatus.querySelector("span").style.boxShadow = "0 0 8px var(--success)";
    } else {
        syncStatus.innerHTML = `<span style="background-color: var(--danger); box-shadow: 0 0 8px var(--danger);"></span> ดึงข้อมูลผิดพลาด`;
        syncStatus.style.color = "var(--danger)";
    }
}

// Manual Refresh Function
function refreshData() {
    const icon = refreshBtn.querySelector("i");
    icon.classList.add("spin");
    refreshBtn.disabled = true;
    
    loadDashboardData().finally(() => {
        setTimeout(() => {
            icon.classList.remove("spin");
            refreshBtn.disabled = false;
        }, 800);
    });
}

// Data Processing
function processAndRenderData(table) {
    const cols = table.cols;
    
    // Find column indices by checking column labels dynamically
    const colIndices = {
        id: cols.findIndex(c => c.label === "รหัสโครงการ"),
        name: cols.findIndex(c => c.label === "ชื่อโครงการ"),
        owner: cols.findIndex(c => c.label === "ผู้รับผิดชอบ"),
        department: cols.findIndex(c => c.label === "กลุ่มงาน"),
        budget: cols.findIndex(c => c.label === "งบประมาณ"),
        spent: cols.findIndex(c => c.label === "ใช้ไปแล้ว"),
        remaining: cols.findIndex(c => c.label === "คงเหลือ"),
        progress: cols.findIndex(c => c.label === "ความคืบหน้า"),
        status: cols.findIndex(c => c.label === "สถานะ")
    };
    
    // Helper function to get cell value safely
    const getVal = (row, index, defaultValue = "") => {
        if (index === -1 || !row.c || !row.c[index]) return defaultValue;
        const val = row.c[index].v;
        return (val !== undefined && val !== null) ? val : defaultValue;
    };

    projectsData = table.rows.map(row => {
        const budget = Number(getVal(row, colIndices.budget !== -1 ? colIndices.budget : 4, 0));
        const spent = Number(getVal(row, colIndices.spent !== -1 ? colIndices.spent : 5, 0));
        const remainingVal = getVal(row, colIndices.remaining !== -1 ? colIndices.remaining : 6, null);
        const remaining = remainingVal !== null ? Number(remainingVal) : (budget - spent);
        const progress = Number(getVal(row, colIndices.progress !== -1 ? colIndices.progress : 7, 0));

        return {
            id: String(getVal(row, colIndices.id !== -1 ? colIndices.id : 0, "N/A")),
            name: String(getVal(row, colIndices.name !== -1 ? colIndices.name : 1, "โครงการไม่มีชื่อ")),
            owner: String(getVal(row, colIndices.owner !== -1 ? colIndices.owner : 2, "ไม่ระบุ")),
            department: String(getVal(row, colIndices.department !== -1 ? colIndices.department : 3, "ทั่วไป")),
            budget: budget,
            spent: spent,
            remaining: remaining,
            progress: progress,
            status: String(getVal(row, colIndices.status !== -1 ? colIndices.status : 8, "ยังไม่ดำเนินการ")).trim()
        };
    });

    // Populate Department Filter dropdown options dynamically
    populateDepartmentDropdown();

    // Initial Filter Apply (which triggers KPI calculation and rendering)
    applyFilters();
}

function populateDepartmentDropdown() {
    // Extract unique departments
    const departments = [...new Set(projectsData.map(p => p.department))].filter(Boolean);
    
    // Save current selection
    const currentSelection = filterDept.value;
    
    // Clear dynamic options (keep first one "ทุกกลุ่มงาน")
    filterDept.innerHTML = '<option value="">ทุกกลุ่มงาน</option>';
    
    departments.sort().forEach(dept => {
        const option = document.createElement("option");
        option.value = dept;
        option.textContent = dept;
        filterDept.appendChild(option);
    });
    
    // Restore selection if still exists
    if (departments.includes(currentSelection)) {
        filterDept.value = currentSelection;
    }
}

// KPI and Statistics Calculations
function calculateMetrics() {
    const totalProjects = filteredProjects.length;
    let totalBudget = 0;
    let totalSpent = 0;
    let totalRemaining = 0;
    let progressSum = 0;
    
    // Count projects by status in filtered set
    let completedCount = 0;
    let inProgressCount = 0;
    let unstartedCount = 0;

    filteredProjects.forEach(proj => {
        totalBudget += proj.budget;
        totalSpent += proj.spent;
        totalRemaining += proj.remaining;
        progressSum += proj.progress;

        if (proj.status === "ดำเนินการแล้ว") completedCount++;
        else if (proj.status === "อยู่ระหว่างดำเนินการ") inProgressCount++;
        else unstartedCount++;
    });

    const avgProgress = totalProjects > 0 ? (progressSum / totalProjects) : 0;
    const spentPercent = totalBudget > 0 ? ((totalSpent / totalBudget) * 100) : 0;
    const remainingPercent = totalBudget > 0 ? ((totalRemaining / totalBudget) * 100) : 0;

    // Remove skeletons and update values
    updateKpiCard("valTotalProjects", totalProjects.toLocaleString("th-TH"), "โครงการ");
    updateKpiCard("valTotalBudget", "฿" + totalBudget.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    updateKpiCard("valSpent", "฿" + totalSpent.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    updateKpiCard("valRemaining", "฿" + totalRemaining.toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    updateKpiCard("valProgress", avgProgress.toFixed(2) + "%");

    // Update Card Footers with extra insights
    footerTotalProjects.innerHTML = `<span class="trend-up"><i class="fa-solid fa-circle-check"></i> ${completedCount} เสร็จสิ้น</span> | <span class="trend-down" style="color: var(--warning)"><i class="fa-solid fa-spinner"></i> ${inProgressCount} ทำอยู่</span>`;
    footerTotalBudget.innerHTML = `งบประมาณของโครงการที่คัดกรอง`;
    footerSpent.innerHTML = `คิดเป็น <span class="trend-down font-semibold">${spentPercent.toFixed(1)}%</span> ของงบทั้งหมด`;
    footerRemaining.innerHTML = `คิดเป็น <span class="trend-up font-semibold">${remainingPercent.toFixed(1)}%</span> ของงบทั้งหมด`;
    footerProgress.innerHTML = `เป้าหมายภาพรวมทุกโครงการ`;
}

function updateKpiCard(elementId, valueText) {
    const el = document.getElementById(elementId);
    el.classList.remove("skeleton");
    el.textContent = valueText;
}

// Rendering Table
function renderTable() {
    // Clear existing body
    tableBody.innerHTML = "";

    if (filteredProjects.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="9">
                    <div class="no-data">
                        <i class="fa-regular fa-folder-open"></i>
                        <p>ไม่พบข้อมูลโครงการที่ตรงกับเงื่อนไขการค้นหา</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    filteredProjects.forEach(proj => {
        const tr = document.createElement("tr");
        tr.addEventListener("click", () => openModal(proj));

        // Format numbers
        const budgetFormatted = proj.budget.toLocaleString("th-TH", { minimumFractionDigits: 2 });
        const spentFormatted = proj.spent > 0 ? proj.spent.toLocaleString("th-TH", { minimumFractionDigits: 2 }) : "-";
        const remainingFormatted = proj.remaining.toLocaleString("th-TH", { minimumFractionDigits: 2 });

        // Status Badge class selector
        let badgeClass = "badge-unstarted";
        if (proj.status === "ดำเนินการแล้ว") badgeClass = "badge-completed";
        else if (proj.status === "อยู่ระหว่างดำเนินการ") badgeClass = "badge-pending";

        tr.innerHTML = `
            <td><span class="project-id-badge">${proj.id}</span></td>
            <td><div class="project-title-cell" title="${proj.name}">${proj.name}</div></td>
            <td>${proj.owner}</td>
            <td>${proj.department}</td>
            <td style="text-align: right;" class="money-text">${budgetFormatted}</td>
            <td style="text-align: right;" class="money-text">${spentFormatted}</td>
            <td style="text-align: right;" class="money-text">${remainingFormatted}</td>
            <td>
                <div class="progress-cell-wrapper">
                    <div class="progress-track">
                        <div class="progress-fill" style="width: ${proj.progress}%"></div>
                    </div>
                    <span class="progress-percent-label">${proj.progress}%</span>
                </div>
            </td>
            <td><span class="badge ${badgeClass}">${proj.status}</span></td>
        `;

        tableBody.appendChild(tr);
    });
}

// Rendering Charts
function renderCharts() {
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const textThemeColor = isDark ? "#9ca3af" : "#64748b";
    const gridColor = isDark ? "rgba(255, 255, 255, 0.08)" : "rgba(15, 23, 42, 0.08)";

    // 1. Prepare Department Data
    // Sum budget and spent per department
    const deptStats = {};
    filteredProjects.forEach(p => {
        if (!deptStats[p.department]) {
            deptStats[p.department] = { budget: 0, spent: 0 };
        }
        deptStats[p.department].budget += p.budget;
        deptStats[p.department].spent += p.spent;
    });

    const deptLabels = Object.keys(deptStats).sort();
    const deptBudgets = deptLabels.map(dept => deptStats[dept].budget);
    const deptSpents = deptLabels.map(dept => deptStats[dept].spent);

    // Render/Update Column Chart
    const deptChartOptions = {
        series: [{
            name: 'งบประมาณรวม (บาท)',
            data: deptBudgets
        }, {
            name: 'งบประมาณที่ใช้ไป (บาท)',
            data: deptSpents
        }],
        chart: {
            type: 'bar',
            height: '100%',
            background: 'transparent',
            toolbar: { show: false },
            fontFamily: 'Prompt, sans-serif'
        },
        plotOptions: {
            bar: {
                horizontal: false,
                columnWidth: '55%',
                endingShape: 'rounded',
                borderRadius: 4
            },
        },
        dataLabels: { enabled: false },
        stroke: {
            show: true,
            width: 2,
            colors: ['transparent']
        },
        xaxis: {
            categories: deptLabels,
            labels: {
                style: {
                    colors: textThemeColor,
                    fontSize: '12px'
                }
            },
            axisBorder: { show: false },
            axisTicks: { show: false }
        },
        yaxis: {
            title: {
                text: 'จำนวนเงิน (บาท)',
                style: { color: textThemeColor, fontWeight: 500 }
            },
            labels: {
                style: { colors: textThemeColor },
                formatter: function (val) {
                    return val >= 1000000 ? (val / 1000000).toFixed(1) + 'M' : val.toLocaleString("th-TH");
                }
            }
        },
        grid: {
            borderColor: gridColor,
            strokeDashArray: 4
        },
        fill: {
            opacity: 1,
            colors: ['#6366f1', '#10b981'] // Indigo and Emerald
        },
        colors: ['#6366f1', '#10b981'],
        legend: {
            position: 'top',
            horizontalAlign: 'right',
            labels: { colors: isDark ? '#f3f4f6' : '#0f172a' }
        },
        tooltip: {
            theme: isDark ? 'dark' : 'light',
            y: {
                formatter: function (val) {
                    return val.toLocaleString("th-TH") + " บาท";
                }
            }
        }
    };

    if (deptChart) {
        deptChart.destroy();
    }
    deptChart = new ApexCharts(document.querySelector("#departmentBudgetChart"), deptChartOptions);
    deptChart.render();

    // 2. Prepare Status Data
    let completed = 0;
    let inProgress = 0;
    let unstarted = 0;

    filteredProjects.forEach(p => {
        if (p.status === "ดำเนินการแล้ว") completed++;
        else if (p.status === "อยู่ระหว่างดำเนินการ") inProgress++;
        else unstarted++;
    });

    const statusChartOptions = {
        series: [completed, inProgress, unstarted],
        chart: {
            type: 'donut',
            height: '100%',
            background: 'transparent',
            fontFamily: 'Prompt, sans-serif'
        },
        labels: ['ดำเนินการแล้ว', 'อยู่ระหว่างดำเนินการ', 'ยังไม่ดำเนินการ'],
        colors: ['#10b981', '#f59e0b', '#64748b'], // Emerald, Amber, Slate
        legend: {
            position: 'bottom',
            labels: { colors: isDark ? '#f3f4f6' : '#0f172a' }
        },
        stroke: { show: false },
        plotOptions: {
            pie: {
                donut: {
                    size: '70%',
                    background: 'transparent',
                    labels: {
                        show: true,
                        name: {
                            show: true,
                            fontSize: '14px',
                            color: textThemeColor,
                            offsetY: -5
                        },
                        value: {
                            show: true,
                            fontSize: '24px',
                            fontWeight: 'bold',
                            color: isDark ? '#f3f4f6' : '#0f172a',
                            offsetY: 5,
                            formatter: function (val) {
                                return val;
                            }
                        },
                        total: {
                            show: true,
                            label: 'โครงการทั้งหมด',
                            color: textThemeColor,
                            formatter: function (w) {
                                return w.globals.seriesTotals.reduce((a, b) => a + b, 0);
                            }
                        }
                    }
                }
            }
        },
        tooltip: {
            theme: isDark ? 'dark' : 'light',
            y: {
                formatter: function (val) {
                    return val + " โครงการ";
                }
            }
        }
    };

    if (statusChart) {
        statusChart.destroy();
    }
    statusChart = new ApexCharts(document.querySelector("#statusDonutChart"), statusChartOptions);
    statusChart.render();
}

// Search and Filter Functions
function applyFilters() {
    const searchText = tableSearch.value.toLowerCase().trim();
    const selectedDept = filterDept.value;
    const selectedStatus = filterStatus.value;

    filteredProjects = projectsData.filter(proj => {
        const matchesSearch = 
            proj.name.toLowerCase().includes(searchText) || 
            proj.owner.toLowerCase().includes(searchText) ||
            proj.id.toLowerCase().includes(searchText);
            
        const matchesDept = selectedDept === "" || proj.department === selectedDept;
        const matchesStatus = selectedStatus === "" || proj.status === selectedStatus;

        return matchesSearch && matchesDept && matchesStatus;
    });

    calculateMetrics();
    renderTable();
    renderCharts();
}

function resetFilters() {
    tableSearch.value = "";
    filterDept.value = "";
    filterStatus.value = "";
    applyFilters();
}

// Modal View Functions
function openModal(project) {
    modalProjId.textContent = `รหัสโครงการ: ${project.id}`;
    modalProjName.textContent = project.name;
    modalProgressText.textContent = `${project.progress}%`;
    modalProgressFill.style.width = `${project.progress}%`;
    modalOwner.textContent = project.owner;
    modalDept.textContent = project.department;
    
    modalBudget.textContent = project.budget.toLocaleString("th-TH", { minimumFractionDigits: 2 }) + " บาท";
    modalSpent.textContent = project.spent.toLocaleString("th-TH", { minimumFractionDigits: 2 }) + " บาท";
    modalRemaining.textContent = project.remaining.toLocaleString("th-TH", { minimumFractionDigits: 2 }) + " บาท";
    
    // Status Badge inside Modal
    let badgeClass = "badge-unstarted";
    if (project.status === "ดำเนินการแล้ว") badgeClass = "badge-completed";
    else if (project.status === "อยู่ระหว่างดำเนินการ") badgeClass = "badge-pending";
    
    modalStatusBadge.className = `badge ${badgeClass}`;
    modalStatusBadge.textContent = project.status;

    projectModal.classList.add("active");
}

function closeModal() {
    projectModal.classList.remove("active");
}

// UI State Handling
function showLoadingState() {
    // Add skeleton classes to KPI values
    const kpiValues = ["valTotalProjects", "valTotalBudget", "valSpent", "valRemaining", "valProgress"];
    kpiValues.forEach(id => {
        const el = document.getElementById(id);
        el.className = "kpi-value skeleton";
        el.textContent = "--";
    });

    // Populate table with skeleton rows
    tableBody.innerHTML = `
        <tr class="skeleton-row">
            <td class="skeleton" style="height: 53px;"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
        </tr>
        <tr class="skeleton-row">
            <td class="skeleton" style="height: 53px;"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
            <td class="skeleton"></td>
        </tr>
    `;
}

function showErrorState(message) {
    const kpiValues = ["valTotalProjects", "valTotalBudget", "valSpent", "valRemaining", "valProgress"];
    kpiValues.forEach(id => {
        const el = document.getElementById(id);
        el.classList.remove("skeleton");
        el.textContent = "error";
    });

    tableBody.innerHTML = `
        <tr>
            <td colspan="9">
                <div class="no-data" style="color: var(--danger)">
                    <i class="fa-solid fa-triangle-exclamation"></i>
                    <p>${message}</p>
                    <button class="btn btn-primary" onclick="loadDashboardData()" style="margin-top: 8px;">ลองใหม่อีกครั้ง</button>
                </div>
            </td>
        </tr>
    `;
}
