// 财务数据看板主控

// ========== 数据加载与解析 ==========
class DataLoader {
    // Excel 列映射（A=0, B=1, ...）
    static COLUMN_MAP = {
        totalAssets: 0,      // A 总资产
        currentDeposit: 1,   // B 活期
        emergencyFund: 3,    // D 应急资金
        fixedDeposit: 5,     // F 定期
        foreignExchange: 7,  // H 外汇
        reserveFund: 9,      // J 投资储备金
        stableFund: 11,      // L 稳健性基金
        riskFund: 13,        // N 风险性基金
        stock: 15,           // P 股票
        providentFund: 17,   // R 公积金
        debt: 19,            // T 负债
        date: 21             // V 时间列
    };

    // 读取 Excel 文件
    static async loadFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // 解析工作簿为记录数组
    static parseWorkbook(workbook) {
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        const records = [];
        jsonData.forEach((row, index) => {
            if (index === 0) return; // 跳过表头

            const dateValue = row[this.COLUMN_MAP.date];
            if (!dateValue) return;

            const date = this.parseDate(dateValue);
            if (!date) return;

            const getValue = (key) => {
                const val = row[this.COLUMN_MAP[key]];
                if (val === null || val === undefined || val === '') return 0;
                const num = parseFloat(val);
                return isNaN(num) ? 0 : num;
            };

            records.push({
                date,
                dateStr: this.formatDate(date),
                totalAssets: getValue('totalAssets'),
                currentDeposit: getValue('currentDeposit'),
                emergencyFund: getValue('emergencyFund'),
                fixedDeposit: getValue('fixedDeposit'),
                foreignExchange: getValue('foreignExchange'),
                reserveFund: getValue('reserveFund'),
                stableFund: getValue('stableFund'),
                riskFund: getValue('riskFund'),
                stock: getValue('stock'),
                providentFund: getValue('providentFund'),
                debt: getValue('debt')
            });
        });

        // 按日期升序排序
        records.sort((a, b) => a.date - b.date);
        return records;
    }

    // 计算衍生指标
    static calculateDerivedIndicators(records) {
        return records.map(r => {
            const liquidAssets = r.currentDeposit + r.emergencyFund;
            const investment = r.reserveFund + r.stableFund + r.riskFund + r.stock;
            const netAssets = r.totalAssets - r.debt;
            const fixedAssets = r.totalAssets - liquidAssets - investment;

            return {
                ...r,
                liquidAssets,
                investment,
                netAssets,
                fixedAssets,
                riskPosition: r.riskFund + r.stock
            };
        });
    }

    static parseDate(value) {
        if (value instanceof Date) return value;
        if (typeof value === 'number') {
            // Excel 序列号日期
            return new Date((value - 25569) * 86400 * 1000);
        }
        if (typeof value === 'string') {
            const date = new Date(value.replace(/\//g, '-'));
            if (!isNaN(date.getTime())) return date;
        }
        return null;
    }

    static formatDate(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
    }

    static formatMoney(value) {
        const wan = value / 10000;
        if (wan >= 1 || wan <= -1) {
            return `¥ ${wan.toFixed(1)}万`;
        }
        return `¥ ${value.toFixed(0)}`;
    }

    static formatMoneyTable(value) {
        return `¥ ${value.toFixed(0)}`;
    }
}

// ========== 图表渲染 ==========
class ChartRenderer {
    constructor(data) {
        this.data = data;
        this.charts = new Map(); // 存储所有Chart.js实例
    }

    // 渲染所有图表
    renderAll() {
        if (!this.data || this.data.length === 0) return;

        const latest = this.data[this.data.length - 1];
        this.updateBadges(latest);
        this.renderChart1();
        this.renderChart2();
        this.renderChart3();
        this.renderChart4();
        this.renderChart5(latest);
        this.renderChart6(latest);
    }

    // 更新右上角金额徽章
    updateBadges(latest) {
        const badges = [
            { selector: '.card:nth-child(1) .net-badge .value', value: latest.netAssets },
            { selector: '.card:nth-child(2) .net-badge .value', value: latest.liquidAssets },
            { selector: '.card:nth-child(3) .net-badge .value', value: latest.investment },
            { selector: '.card:nth-child(4) .net-badge .value', value: latest.providentFund }
        ];

        badges.forEach(item => {
            const el = document.querySelector(item.selector);
            if (el) el.textContent = DataLoader.formatMoney(item.value);
        });

        // 卡片2 的 mini-stat
        const miniStats = document.querySelectorAll('.card:nth-child(2) .mini-stat b');
        if (miniStats.length >= 2) {
            miniStats[0].textContent = DataLoader.formatMoney(latest.currentDeposit);
            miniStats[1].textContent = DataLoader.formatMoney(latest.emergencyFund);
        }

        // 卡片5 饼图图例
        this.updatePieLegend('.card:nth-child(5)', [
            { color: 'var(--s1)', label: '流动资金', value: latest.liquidAssets, total: latest.totalAssets },
            { color: 'var(--s2)', label: '固定资本', value: latest.fixedAssets, total: latest.totalAssets },
            { color: 'var(--s3)', label: '投资', value: latest.investment, total: latest.totalAssets }
        ]);

        // 卡片6 饼图图例
        this.updatePieLegend('.card:nth-child(6)', [
            { color: 'var(--s6)', label: '储备金', value: latest.reserveFund, total: latest.investment },
            { color: 'var(--s4)', label: '风险仓位', value: latest.riskPosition, total: latest.investment },
            { color: 'var(--s3)', label: '稳健仓位', value: latest.stableFund, total: latest.investment }
        ]);
    }

    updatePieLegend(cardSelector, items) {
        const card = document.querySelector(cardSelector);
        if (!card) return;
        const legendItems = card.querySelectorAll('.legend-item');
        items.forEach((item, index) => {
            if (legendItems[index]) {
                const ratio = item.total > 0 ? (item.value / item.total * 100).toFixed(0) : 0;
                legendItems[index].innerHTML = `
                    <span class="dot" style="background:${item.color}"></span>
                    ${item.label} ${ratio}%
                `;
            }
        });
    }

    // 渲染折线图使用Chart.js
    renderLineChart(container, chartTitle, labels, datasets) {
        // 清除之前创建的图表
        if (this.charts.has(container)) {
            this.charts.get(container).destroy();
        }

        // 创建新的Chart.js图表
        const ctx = document.createElement('canvas');
        ctx.width = container.offsetWidth;
        ctx.height = container.offsetHeight;
        container.innerHTML = '';
        container.appendChild(ctx);

        // 检查Chart是否存在
        if (typeof Chart !== 'undefined' && Chart && Chart.defaults) {
            const chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: datasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: false
                            // 隐藏图表标题
                        },
                        legend: {
                            display: false
                            // 隐藏图例，因为图例已在卡片左上角显示
                        },
                        tooltip: {
                            mode: 'index',
                            intersect: false
                        }
                    },
                    hover: {
                        mode: 'nearest',
                        intersect: true
                    },
                    scales: {
                        x: {
                            display: true,
                            title: {
                                display: false
                                // 隐藏标题文字，只显示数字
                            },
                            ticks: {
                                // 格式化横坐标显示，去掉年份前缀
                                callback: function(value, index, values) {
                                    // Chart.js v3+ 传入的是索引，需要通过getLabelForValue获取真实标签
                                    const label = this.getLabelForValue(value);
                                    // 如果是日期格式，格式化为YY/MM/DD
                                    if (label && typeof label === 'string' && label.includes('/')) {
                                        const parts = label.split('/');
                                        if (parts.length >= 3) {
                                            const year = parts[0];
                                            const month = parts[1];
                                            const day = parts[2];
                                            // 只保留年份后两位
                                            const shortYear = year.length === 4 ? year.slice(2) : year;
                                            return shortYear + '/' + month + '/' + day;
                                        }
                                    }
                                    return label;
                                }
                            }
                        },
                        y: {
                            display: true,
                            title: {
                                display: false
                                // 隐藏标题文字，只显示数字
                            },
                            ticks: {
                                // 格式化纵坐标显示，以w为单位
                                callback: function(value) {
                                    if (value >= 10000) {
                                        return (value / 10000).toFixed(1) + 'w';
                                    }
                                    return value;
                                }
                            }
                        }
                    },
                    elements: {
                        line: {
                            tension: 0.4, // 平滑曲线
                            borderWidth: 2
                        },
                        point: {
                            display: false // 不在节点处显示图形
                        }
                    }
                }
            });

            this.charts.set(container, chart);
        }
    }

    // 渲染柱状图使用Chart.js
    renderBarChart(container, chartTitle, labels, datasets) {
        // 清除之前创建的图表
        if (this.charts.has(container)) {
            this.charts.get(container).destroy();
        }

        // 创建新的Chart.js图表
        const ctx = document.createElement('canvas');
        ctx.width = container.offsetWidth;
        ctx.height = container.offsetHeight;
        container.innerHTML = '';
        container.appendChild(ctx);

        // 检查Chart是否存在
        if (typeof Chart !== 'undefined' && Chart && Chart.defaults) {
            const chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: datasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: chartTitle
                        },
                        legend: {
                            display: true,
                            position: 'top',
                        }
                    },
                    scales: {
                        x: {
                            display: true,
                            title: {
                                display: false
                                // 隐藏标题文字，只显示数字
                            }
                        },
                        y: {
                            display: true,
                            title: {
                                display: false
                                // 隐藏标题文字，只显示数字
                            }
                        }
                    }
                }
            });

            this.charts.set(container, chart);
        }
    }

    // 渲染饼图使用Chart.js
    renderPieChart(container, chartTitle, labels, data, backgroundColors) {
        // 清除之前创建的图表
        if (this.charts.has(container)) {
            this.charts.get(container).destroy();
        }

        // 创建新的Chart.js图表
        const ctx = document.createElement('canvas');
        ctx.width = container.offsetWidth;
        ctx.height = container.offsetHeight;
        container.innerHTML = '';
        container.appendChild(ctx);

        // 检查Chart是否存在
        if (typeof Chart !== 'undefined' && Chart && Chart.defaults) {
            const chart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        data: data,
                        backgroundColor: backgroundColors
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: chartTitle
                        },
                        legend: {
                            display: true,
                            position: 'right',
                        }
                    }
                }
            });

            this.charts.set(container, chart);
        }
    }

    renderChart1() {
        const container = document.querySelector('.card:nth-child(1) .chart');
        const labels = this.data.map(d => d.dateStr);
        const datasets = [
            {
                label: '总资产',
                data: this.data.map(d => d.totalAssets),
                borderColor: 'var(--s1)',
                backgroundColor: 'rgba(42, 120, 214, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '负债',
                data: this.data.map(d => d.debt),
                borderColor: 'var(--s2)',
                backgroundColor: 'rgba(227, 73, 72, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '净资产',
                data: this.data.map(d => d.netAssets),
                borderColor: 'var(--s3)',
                backgroundColor: 'rgba(27, 175, 122, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            }
        ];

        this.renderLineChart(container, '总资产 / 负债 / 净资产', labels, datasets);
    }

    renderChart2() {
        const container = document.querySelector('.card:nth-child(2) .chart');
        const labels = this.data.map(d => d.dateStr);
        const datasets = [
            {
                label: '活期',
                data: this.data.map(d => d.currentDeposit),
                borderColor: 'var(--s1)',
                backgroundColor: 'rgba(42, 120, 214, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '应急资金',
                data: this.data.map(d => d.emergencyFund),
                borderColor: 'var(--s6)',
                backgroundColor: 'rgba(237, 161, 0, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '流动资金',
                data: this.data.map(d => d.liquidAssets),
                borderColor: 'var(--s3)',
                backgroundColor: 'rgba(27, 175, 122, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            }
        ];

        this.renderLineChart(container, '流动资金', labels, datasets);
    }

    renderChart3() {
        const container = document.querySelector('.card:nth-child(3) .chart');
        const labels = this.data.map(d => d.dateStr);
        const datasets = [
            {
                label: '储备金',
                data: this.data.map(d => d.reserveFund),
                borderColor: 'var(--s7)',
                backgroundColor: 'rgba(0, 131, 0, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '稳健基金',
                data: this.data.map(d => d.stableFund),
                borderColor: 'var(--s3)',
                backgroundColor: 'rgba(27, 175, 122, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '风险基金',
                data: this.data.map(d => d.riskFund),
                borderColor: 'var(--s4)',
                backgroundColor: 'rgba(235, 104, 52, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '股票',
                data: this.data.map(d => d.stock),
                borderColor: 'var(--s5)',
                backgroundColor: 'rgba(232, 123, 164, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            }
        ];

        this.renderLineChart(container, '投资资金', labels, datasets);
    }

    renderChart4() {
        const container = document.querySelector('.card:nth-child(4) .chart');
        const labels = this.data.map(d => d.dateStr);
        const datasets = [
            {
                label: '公积金',
                data: this.data.map(d => d.providentFund),
                borderColor: 'var(--s4)',
                backgroundColor: 'rgba(235, 104, 52, 0.2)',
                borderWidth: 2,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            }
        ];

        this.renderLineChart(container, '公积金', labels, datasets);
    }

    renderChart5(latest) {
        const container = document.querySelector('.card:nth-child(5) .chart');
        const labels = ['流动资金', '固定资本', '投资'];
        const data = [latest.liquidAssets, latest.fixedAssets, latest.investment];
        const backgroundColors = ['var(--s1)', 'var(--s2)', 'var(--s3)'];

        this.renderPieChart(container, '总资产构成占比', labels, data, backgroundColors);
    }

    renderChart6(latest) {
        const container = document.querySelector('.card:nth-child(6) .chart');
        const labels = ['储备金', '风险仓位', '稳健仓位'];
        const data = [latest.reserveFund, latest.riskPosition, latest.stableFund];
        const backgroundColors = ['var(--s6)', 'var(--s4)', 'var(--s3)'];

        this.renderPieChart(container, '投资资金构成占比', labels, data, backgroundColors);
    }
}

// ========== 交互处理 ==========
class FinancialChartHandler {
    static totalYears = 1;
    static defaultYears = 1;
    static minYears = 0.5;

    constructor() {
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.setupTimeFilter();
        this.setupRefresh();
    }

    setupEventListeners() {
        const refreshBtn = document.getElementById('refreshBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => window.financialDashboard.handleRefresh());
        }

        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.addEventListener('change', (e) => {
                if (e.target.files && e.target.files[0]) {
                    window.financialDashboard.handleFileSelect(e.target.files[0]);
                }
            });
        }
    }

    setupTimeFilter() {
        const filterPills = document.querySelectorAll('.filter-pill');
        filterPills.forEach(pill => {
            pill.addEventListener('click', (e) => this.handleTimeFilter(e));
        });
    }

    handleTimeFilter(e) {
        const filterPills = document.querySelectorAll('.filter-pill');
        filterPills.forEach(pill => pill.classList.remove('active'));
        e.target.classList.add('active');

        const text = e.target.textContent.trim();
        const { totalYears } = FinancialChartHandler;
        let targetYears = totalYears;
        switch (text) {
            case '最近1年': targetYears = 1; break;
            case '最近3年': targetYears = 3; break;
            case '最近5年': targetYears = 5; break;
            case '全部': targetYears = totalYears; break;
        }

        // 重新渲染图表
        window.financialDashboard.renderChartsWithTimeRange(targetYears);

        console.log('时间范围切换到: ' + text);
    }

    // 新增方法：根据时间范围重新渲染图表
    renderChartsWithTimeRange(targetYears) {
        if (window.financialDashboard && window.financialDashboard.data) {
            // 获取当前数据并调整显示范围
            const data = window.financialDashboard.data;
            const startIndex = Math.max(0, data.length - Math.floor(targetYears * 12)); // 假设每月一个数据点
            const filteredData = data.slice(startIndex);

            // 更新图表显示（这里只是重新初始化图表，实际可用性需要根据需要完善）
            if (filteredData.length > 0) {
                const renderer = new ChartRenderer(filteredData);
                renderer.renderAll();
                // 重新初始化交互部分
                window.financialDashboard.chartHandler.setupTimeFilter();
            }
        }
    }

    setupRefresh() {
        // 刷新按钮事件在 setupEventListeners 中统一处理
    }
}

// ========== 数据表格 ==========
class DataTable {
    static render(data) {
        const section = document.getElementById('dataTableSection');
        const tbody = document.querySelector('#dataTable tbody');
        if (!section || !tbody) return;

        tbody.innerHTML = '';
        // 显示最近 20 条记录
        const recentData = data.slice(-20).reverse();

        recentData.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${row.dateStr}</td>
                <td>${DataLoader.formatMoneyTable(row.totalAssets)}</td>
                <td>${DataLoader.formatMoneyTable(row.debt)}</td>
                <td>${DataLoader.formatMoneyTable(row.netAssets)}</td>
                <td>${DataLoader.formatMoneyTable(row.liquidAssets)}</td>
                <td>${DataLoader.formatMoneyTable(row.investment)}</td>
                <td>${DataLoader.formatMoneyTable(row.providentFund)}</td>
            `;
            tbody.appendChild(tr);
        });

        section.style.display = 'block';
    }
}

// ========== 主控制器 ==========
class FinancialDashboard {
    constructor() {
        this.data = [];
        this.chartHandler = new FinancialChartHandler();
    }

    async handleFileSelect(file) {
        try {
            this.updateDataSourceInfo(`正在读取 ${file.name}...`);
            const workbook = await DataLoader.loadFile(file);
            const records = DataLoader.parseWorkbook(workbook);
            if (records.length === 0) {
                alert('未解析到有效数据，请检查 Excel 列是否与文档约定一致。');
                return;
            }

            // 允许再次选择同一文件
            const fileInput = document.getElementById('fileInput');
            if (fileInput) fileInput.value = '';

            this.data = DataLoader.calculateDerivedIndicators(records);

            // 根据数据时间跨度更新配置
            const years = this.calculateDataSpanYears();
            FinancialChartHandler.totalYears = years;

            // 默认只显示最近一年的数据（如果数据量足够）
            let displayData = this.data;
            if (this.data.length > 12) { // 至少有12个月的数据
                const oneYearAgo = new Date();
                oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

                // 找到最近一年的数据开始点
                let startIndex = 0;
                for (let i = 0; i < this.data.length; i++) {
                    if (this.data[i].date >= oneYearAgo) {
                        startIndex = i;
                        break;
                    }
                }

                displayData = this.data.slice(startIndex);
            }

            // 渲染图表
            const renderer = new ChartRenderer(displayData);
            renderer.renderAll();

            // 重新初始化交互（包含提取固定标签）
            // 注意：此处不调用reinit，因为我们已移除了鼠标交互功能

            // 填充表格
            DataTable.render(displayData);

            // 更新数据源信息
            const latest = displayData[displayData.length - 1];
            this.updateDataSourceInfo(`数据源：${file.name} · 最新数据 ${latest.dateStr}`);

        } catch (err) {
            console.error(err);
            alert('读取 Excel 失败：' + err.message);
        }
    }

    handleRefresh() {
        const fileInput = document.getElementById('fileInput');
        if (fileInput && fileInput.files && fileInput.files[0]) {
            this.handleFileSelect(fileInput.files[0]);
        } else {
            alert('请先选择 Excel 文件');
        }
    }

    calculateDataSpanYears() {
        if (this.data.length < 2) return 1;
        const start = this.data[0].date;
        const end = this.data[this.data.length - 1].date;
        const days = (end - start) / (1000 * 60 * 60 * 24);
        return Math.max(1, days / 365);
    }

    updateDataSourceInfo(text) {
        const el = document.getElementById('dataSourceInfo');
        if (el) el.textContent = text;
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    window.financialDashboard = new FinancialDashboard();
});