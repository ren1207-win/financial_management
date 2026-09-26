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
        date: 21,            // V 时间列
        totalInvestment: 28  // AC 总投资额
    };

    // 解析投资收益 CSV（格式：日期,收益,总收益）
    static parseInvestmentIncomeCsv(text) {
        const lines = text.trim().replace(/^﻿/, '').split(/\r?\n/);
        if (lines.length < 2) return [];

        const parseNumber = (v) => {
            if (!v) return 0;
            const n = parseFloat(v.replace(/,/g, ''));
            return isNaN(n) ? 0 : n;
        };

        // 总收益只在 CSV 最后一行的第三列记录，统一读取该值作为累计总收益
        const lastRow = lines[lines.length - 1].split(',').map(v => v.trim());
        const lastTotalIncome = lastRow.length >= 3 ? parseNumber(lastRow[2]) : 0;

        const records = [];
        for (let i = 1; i < lines.length; i++) {
            const row = lines[i].split(',').map(v => v.trim());
            if (row.length < 2) continue;

            const date = this.parseMonthlyDate(row[0]);
            if (!date) continue;

            const income = parseNumber(row[1]);
            const totalIncome = lastTotalIncome;

            // 直接使用原始日期字符串，不进行额外格式化
            const originalDateStr = row[0];

            records.push({
                date,
                dateStr: originalDateStr,  // 直接使用原始日期字符串
                monthLabel: originalDateStr,  // 直接使用原始日期字符串
                income,
                totalIncome  // 添加总收益字段
            });
        }

        // 确保记录按年月正确排序，解决跨年排序问题
        records.sort((a, b) => {
            // 先按年份排序
            const yearA = a.date.getFullYear();
            const yearB = b.date.getFullYear();
            if (yearA !== yearB) {
                return yearA - yearB;
            }
            // 年份相同时，按月份排序
            const monthA = a.date.getMonth();
            const monthB = b.date.getMonth();
            return monthA - monthB;
        });
        return records;
    }

    // 读取 Excel 文件（支持 File 对象和 ArrayBuffer）
    static async loadFile(file) {
        if (file instanceof File) {
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
        } else if (file instanceof ArrayBuffer) {
            return new Promise((resolve, reject) => {
                try {
                    const data = new Uint8Array(file);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (err) {
                    reject(err);
                }
            });
        } else if (file instanceof Uint8Array) {
            return new Promise((resolve, reject) => {
                try {
                    const workbook = XLSX.read(file, { type: 'array' });
                    resolve(workbook);
                } catch (err) {
                    reject(err);
                }
            });
        } else {
            return Promise.reject(new Error('不受支持的文件类型'));
        }
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
                debt: getValue('debt'),
                totalInvestment: getValue('totalInvestment')
            });
        });

        // 按日期升序排序
        // 确保记录按年月正确排序，解决跨年排序问题
        records.sort((a, b) => {
            // 先按年份排序
            const yearA = a.date.getFullYear();
            const yearB = b.date.getFullYear();
            if (yearA !== yearB) {
                return yearA - yearB;
            }
            // 年份相同时，按月份排序
            const monthA = a.date.getMonth();
            const monthB = b.date.getMonth();
            return monthA - monthB;
        });
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

    // 读取月度收支 CSV 文件
    static async loadMonthlyCsv(path) {
        const response = await fetch(path);
        if (!response.ok) {
            throw new Error(`无法读取月度数据：${path}，状态码：${response.status}`);
        }
        const text = await response.text();
        return this.parseMonthlyCsv(text);
    }

    // 解析月度收支 CSV（格式：日期,收入,支出,结余）
    static parseMonthlyCsv(text) {
        const lines = text.trim().replace(/^﻿/, '').split(/\r?\n/);
        if (lines.length < 2) return [];

        const records = [];
        for (let i = 1; i < lines.length; i++) {
            const row = lines[i].split(',').map(v => v.trim());
            if (row.length < 4) continue;

            const date = this.parseMonthlyDate(row[0]);
            if (!date) continue;

            const parseNumber = (v) => {
                if (!v) return 0;
                const n = parseFloat(v.replace(/,/g, ''));
                return isNaN(n) ? 0 : n;
            };

            const income = parseNumber(row[1]);
            const expense = parseNumber(row[2]);
            const balance = parseNumber(row[3]);

            // 直接使用原始日期格式作为monthLabel
            const monthLabel = row[0]; // 原始格式如 "25/6"

            records.push({
                date,
                dateStr: this.formatDate(date),
                monthLabel: monthLabel,
                income,
                expense,
                balance
            });
        }

        // 确保记录按年月正确排序，解决跨年排序问题
        records.sort((a, b) => {
            // 先按年份排序
            const yearA = a.date.getFullYear();
            const yearB = b.date.getFullYear();
            if (yearA !== yearB) {
                return yearA - yearB;
            }
            // 年份相同时，按月份排序
            const monthA = a.date.getMonth();
            const monthB = b.date.getMonth();
            return monthA - monthB;
        });
        return records;
    }

    // 解析月度数据中的日期（处理 "25/5" 这种 "YY/M" 年份/月份格式）
    static parseMonthlyDate(value) {
        if (!value) return null;

        const parts = value.split('/');
        if (parts.length === 2) {
            const first = parseInt(parts[0], 10);
            const second = parseInt(parts[1], 10);

            // 当第一部分大于 12 时，无法是合法的 DD/MM 中的 day，
            // 应视为年份（两位或四位），第二部分为月份
            if (!isNaN(first) && !isNaN(second) && first > 12) {
                const year = first < 100 ? 2000 + first : first;
                const month = second - 1; // 月份从0开始
                const date = new Date(year, month, 1);
                if (!isNaN(date.getTime())) {
                    return date;
                }
            }

            // 兼容旧的 "DD/M" 格式（day <= 12）
            if (!isNaN(first) && !isNaN(second) && first >= 1 && first <= 12 && second >= 1 && second <= 12) {
                const year = new Date().getFullYear();
                const month = second - 1;
                const date = new Date(year, month, first);
                if (!isNaN(date.getTime())) {
                    return date;
                }
            }
        }

        // 如果上面的方法失败，使用原来的方法
        return this.parseDate(value);
    }
}

// ========== 图表渲染 ==========
class ChartRenderer {
    constructor(data, monthlyData = null) {
        this.data = data;
        this.monthlyData = monthlyData;
        this.charts = new Map(); // 存储所有Chart.js实例
    }

    // 获取 CSS 变量对应的真实颜色值
    static getCssColor(variableName) {
        try {
            const root = document.documentElement;
            const value = getComputedStyle(root).getPropertyValue(variableName).trim();
            if (value) return value;
        } catch (err) {
            console.warn('读取 CSS 变量失败:', variableName, err);
        }
        // 回退颜色表
        const fallback = {
            '--s1': '#dc2626',
            '--s2': '#2563eb',
            '--s3': '#16a34a',
            '--s4': '#f97316',
            '--s5': '#9333ea',
            '--s6': '#ca8a04',
            '--s7': '#475569',
            '--s8': '#dc2626'
        };
        return fallback[variableName] || '#000000';
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
        this.renderMonthlyIncomeExpense();
        this.renderInvestmentIncomeChart();
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
            { color: 'var(--s2)', label: '储备金', value: latest.reserveFund, total: latest.investment },
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
                            position: 'right',
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
                            borderWidth: 1
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

    // 渲染投资收益折线图
    renderInvestmentIncomeChart() {
        const container = document.querySelector('.card:nth-child(8) .chart');
        if (!container) {
            console.log('未找到投资收益图表容器');
            return;
        }

        // 检查是否已有数据，如果没有则退出
        if (!this.monthlyData || this.monthlyData.length === 0) {
            console.log('没有月度数据可供渲染投资收益图表');
            return;
        }

        // 计算最近12个月的数据
        const data = this.monthlyData.slice(-12);
        const labels = data.map(d => d.monthLabel);
        const income = data.map(d => d.income);

        // 清理旧图表
        const existingChart = this.charts.get(container);
        if (existingChart) {
            existingChart.destroy();
            this.charts.delete(container);
        }

        if (typeof Chart === 'undefined' || !Chart.defaults) {
            console.log('Chart.js未加载');
            return;
        }

        // 创建新的canvas元素
        const canvas = document.createElement('canvas');
        container.innerHTML = '';
        container.appendChild(canvas);
        const ctx = canvas.getContext('2d');

        const secondaryColor = ChartRenderer.getCssColor('--ink-secondary');
        const gridColor = 'rgba(11,11,11,0.06)';

        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: '月度收益',
                        data: income,
                        borderColor: ChartRenderer.getCssColor('--s1'),
                        backgroundColor: 'rgba(220, 38, 38, 0.1)',
                        borderWidth: 2,
                        tension: 0.4, // 平滑曲线
                        pointRadius: 3,
                        pointHoverRadius: 5,
                        pointBackgroundColor: ChartRenderer.getCssColor('--s1'),
                        fill: false,
                        yAxisID: 'y'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: { display: false },
                    legend: { display: false },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                const value = context.parsed.y;
                                if (Math.abs(value) >= 10000) {
                                    label += (value / 10000).toFixed(1) + '万';
                                } else {
                                    label += value.toLocaleString();
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        display: true,
                        grid: { display: false },
                        ticks: { color: secondaryColor }
                    },
                    y: {
                        display: true,
                        position: 'left',
                        beginAtZero: true,
                        grid: { color: gridColor },
                        ticks: {
                            color: secondaryColor,
                            callback: function(value) {
                                if (Math.abs(value) >= 10000) {
                                    return (value / 10000).toFixed(1) + 'w';
                                }
                                return value;
                            }
                        }
                    }
                }
            }
        });

        this.charts.set(container, chart);

        // 更新总收益金额显示（卡片右上角）
        this.updateInvestmentTotalIncome(data);
    }

    // 更新投资总收益显示
    updateInvestmentTotalIncome(monthlyData) {
        if (!monthlyData || monthlyData.length === 0) return;

        // 获取最新记录中的总收益
        const latest = monthlyData[monthlyData.length - 1];
        const totalIncome = latest.totalIncome;

        // 更新页面上的显示
        const totalIncomeElement = document.querySelector('.card:nth-child(8) .net-badge .value');
        if (totalIncomeElement && totalIncome !== undefined) {
            const formatted = DataLoader.formatMoney(totalIncome);
            totalIncomeElement.textContent = formatted;
        } else {
            console.log('Could not find total income element or invalid total income');
        }
    }

    renderChart1() {
        const container = document.querySelector('.card:nth-child(1) .chart');
        const labels = this.data.map(d => d.dateStr);
        const datasets = [
            {
                label: '总资产',
                data: this.data.map(d => d.totalAssets),
                borderColor: ChartRenderer.getCssColor('--s2'),
                backgroundColor: 'rgba(37, 99, 235, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '负债',
                data: this.data.map(d => d.debt),
                borderColor: ChartRenderer.getCssColor('--s7'),
                backgroundColor: 'rgba(71, 85, 105, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '净资产',
                data: this.data.map(d => d.netAssets),
                borderColor: ChartRenderer.getCssColor('--s1'),
                backgroundColor: 'rgba(220, 38, 38, 0.2)',
                borderWidth: 1,
                order: 1,
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
                borderColor: ChartRenderer.getCssColor('--s2'),
                backgroundColor: 'rgba(37, 99, 235, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '应急资金',
                data: this.data.map(d => d.emergencyFund),
                borderColor: ChartRenderer.getCssColor('--s6'),
                backgroundColor: 'rgba(202, 138, 4, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '流动资金',
                data: this.data.map(d => d.liquidAssets),
                borderColor: ChartRenderer.getCssColor('--s1'),
                backgroundColor: 'rgba(220, 38, 38, 0.2)',
                borderWidth: 1,
                order: 1,
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
                borderColor: ChartRenderer.getCssColor('--s2'),
                backgroundColor: 'rgba(37, 99, 235, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '稳健基金',
                data: this.data.map(d => d.stableFund),
                borderColor: ChartRenderer.getCssColor('--s3'),
                backgroundColor: 'rgba(22, 163, 74, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '风险基金',
                data: this.data.map(d => d.riskFund),
                borderColor: ChartRenderer.getCssColor('--s4'),
                backgroundColor: 'rgba(249, 115, 22, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '股票',
                data: this.data.map(d => d.stock),
                borderColor: ChartRenderer.getCssColor('--s5'),
                backgroundColor: 'rgba(147, 51, 234, 0.2)',
                borderWidth: 1,
                tension: 0.4,
                pointStyle: 'circle',
                pointRadius: 0,
                pointHoverRadius: 0
            },
            {
                label: '总投资额',
                data: this.data.map(d => d.totalInvestment),
                borderColor: ChartRenderer.getCssColor('--s8'),
                backgroundColor: 'rgba(220, 38, 38, 0.2)',
                borderWidth: 1,
                order: 1,
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
                borderColor: ChartRenderer.getCssColor('--s1'),
                backgroundColor: 'rgba(220, 38, 38, 0.2)',
                borderWidth: 1,
                order: 1,
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
        const backgroundColors = [ChartRenderer.getCssColor('--s1'), ChartRenderer.getCssColor('--s2'), ChartRenderer.getCssColor('--s3')];

        this.renderPieChart(container, '总资产构成占比', labels, data, backgroundColors);
    }

    renderChart6(latest) {
        const container = document.querySelector('.card:nth-child(6) .chart');
        const labels = ['储备金', '风险仓位', '稳健仓位'];
        const data = [latest.reserveFund, latest.riskPosition, latest.stableFund];
        const backgroundColors = [ChartRenderer.getCssColor('--s2'), ChartRenderer.getCssColor('--s4'), ChartRenderer.getCssColor('--s3')];

        this.renderPieChart(container, '投资资金构成占比', labels, data, backgroundColors);
    }

    // 渲染 p3 卡片：月度收支条形图（最近 12 个月）
    renderMonthlyIncomeExpense() {
        if (!this.monthlyData || this.monthlyData.length === 0) {
            console.log('没有月度数据可供渲染');
            return;
        }

        const container = document.querySelector('.card:nth-child(7) .chart');
        if (!container) {
            console.log('未找到图表容器');
            return;
        }

        // 只取最近 12 个月，不足则取全部
        const data = this.monthlyData.slice(-12);
        const labels = data.map(d => d.monthLabel);
        const income = data.map(d => d.income);
        const expense = data.map(d => d.expense);
        // 结余数据已解析并保留，便于后续直接叠加折线图
        const balance = data.map(d => d.balance);

        // 清理旧图表
        const existingChart = this.charts.get(container);
        if (existingChart) {
            existingChart.destroy();
            this.charts.delete(container);
        }

        if (typeof Chart === 'undefined' || !Chart.defaults) {
            console.log('Chart.js未加载');
            return;
        }

        // 创建新的canvas元素
        const canvas = document.createElement('canvas');
        container.innerHTML = '';
        container.appendChild(canvas);
        const ctx = canvas.getContext('2d');

        const secondaryColor = ChartRenderer.getCssColor('--ink-secondary');
        const gridColor = 'rgba(11,11,11,0.06)';

        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: '收入',
                        data: income,
                        backgroundColor: ChartRenderer.getCssColor('--s1'),
                        borderColor: ChartRenderer.getCssColor('--s1'),
                        borderWidth: 0,
                        borderRadius: 4,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8,
                        order: 2,
                        yAxisID: 'y'
                    },
                    {
                        label: '支出',
                        data: expense,
                        backgroundColor: ChartRenderer.getCssColor('--s3'),
                        borderColor: ChartRenderer.getCssColor('--s3'),
                        borderWidth: 0,
                        borderRadius: 4,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8,
                        order: 2,
                        yAxisID: 'y'
                    },
                    {
                        label: '结余',
                        data: balance,
                        type: 'line',
                        borderColor: ChartRenderer.getCssColor('--s2'),
                        backgroundColor: 'rgba(37, 99, 235, 0.1)',
                        borderWidth: 2,
                        tension: 0,
                        pointRadius: 3,
                        pointHoverRadius: 5,
                        pointBackgroundColor: ChartRenderer.getCssColor('--s2'),
                        fill: false,
                        yAxisID: 'y1',
                        order: 1
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: { display: false },
                    legend: { display: false },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                const value = context.parsed.y;
                                if (Math.abs(value) >= 10000) {
                                    label += (value / 10000).toFixed(1) + '万';
                                } else {
                                    label += value.toLocaleString();
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        display: true,
                        grid: { display: false },
                        ticks: { color: secondaryColor }
                    },
                    y: {
                        display: true,
                        position: 'left',
                        beginAtZero: true,
                        grid: { color: gridColor },
                        ticks: {
                            color: secondaryColor,
                            callback: function(value) {
                                if (Math.abs(value) >= 10000) {
                                    return (value / 10000).toFixed(1) + 'w';
                                }
                                return value;
                            }
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        beginAtZero: true,
                        grid: { drawOnChartArea: false },
                        ticks: {
                            color: secondaryColor,
                            callback: function(value) {
                                if (Math.abs(value) >= 10000) {
                                    return (value / 10000).toFixed(1) + 'w';
                                }
                                return value;
                            }
                        }
                    }
                }
            }
        });

        this.charts.set(container, chart);

        // 更新结余总金额显示
        this.updateMonthlyBalance(data);
    }

    // 渲染投资收益折线图
    renderInvestmentIncomeChart() {
        const container = document.querySelector('.card:nth-child(8) .chart');
        if (!container) {
            console.log('未找到投资收益图表容器');
            return;
        }

        // 检查是否已有数据，如果没有则退出
        if (!this.monthlyData || this.monthlyData.length === 0) {
            console.log('没有月度数据可供渲染投资收益图表');
            return;
        }

        // 只取最近 12 个月的数据
        const data = this.monthlyData.slice(-12);
        const labels = data.map(d => d.monthLabel);
        const income = data.map(d => d.income);
        // 获取最新的总收益用于显示

        // 清理旧图表
        const existingChart = this.charts.get(container);
        if (existingChart) {  
            existingChart.destroy();
            this.charts.delete(container);
        }

        if (typeof Chart === 'undefined' || !Chart.defaults) {
            console.log('Chart.js未加载');
            return;
        }

        // 创建新的canvas元素
        const canvas = document.createElement('canvas');
        container.innerHTML = '';
        container.appendChild(canvas);
        const ctx = canvas.getContext('2d');

        const secondaryColor = ChartRenderer.getCssColor('--ink-secondary');
        const gridColor = 'rgba(11,11,11,0.06)';

        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: '月度收益',
                        data: income,
                        borderColor: ChartRenderer.getCssColor('--s1'),
                        backgroundColor: 'rgba(220, 38, 38, 0.1)',
                        borderWidth: 2,
                        tension: 0.4, // 平滑曲线
                        pointRadius: 3,
                        pointHoverRadius: 5,
                        pointBackgroundColor: ChartRenderer.getCssColor('--s1'),
                        fill: false,
                        yAxisID: 'y'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: { display: false },
                    legend: { display: false },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                const value = context.parsed.y;
                                if (Math.abs(value) >= 10000) {
                                    label += (value / 10000).toFixed(1) + '万';
                                } else {
                                    label += value.toLocaleString();
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        display: true,
                        border: { display: false },
                        grid: { display: false, drawBorder: false },
                        ticks: { color: secondaryColor }
                    },
                    y: {
                        display: true,
                        position: 'left',
                        beginAtZero: false,
                        grid: {
                            // 让 y=0 的网格线充当 x 轴
                            color: function(context) {
                                if (context.tick.value === 0) {
                                    return 'rgba(11, 11, 11, 0.5)';
                                }
                                return gridColor;
                            },
                            lineWidth: function(context) {
                                if (context.tick.value === 0) {
                                    return 1.5;
                                }
                                return 1;
                            },
                            drawOnChartArea: true
                        },
                        ticks: {
                            color: secondaryColor,
                            callback: function(value) {
                                if (Math.abs(value) >= 10000) {
                                    return (value / 10000).toFixed(1) + 'w';
                                }
                                return value;
                            }
                        }
                    }
                },
                // 添加点配置确保图表效果正确
                elements: {
                    line: {
                        borderWidth: 2
                    },
                    point: {
                        radius: 3,
                        hoverRadius: 5
                    }
                }
            }
        });

        this.charts.set(container, chart);

        // 更新总收益金额显示（卡片右上角）
        this.updateInvestmentTotalIncome(data);
    }

    // 更新投资总收益显示
    updateInvestmentTotalIncome(monthlyData) {
        if (!monthlyData || monthlyData.length === 0) return;

        // 获取最新记录中的总收益
        const latest = monthlyData[monthlyData.length - 1];
        const totalIncome = latest.totalIncome;

        // 添加调试输出
        console.log('Updating investment total income:', { latest, totalIncome });

        // 更新页面上的显示
        const totalIncomeElement = document.querySelector('.card:nth-child(8) .net-badge .value');
        if (totalIncomeElement && totalIncome !== undefined) {
            const formatted = DataLoader.formatMoney(totalIncome);
            console.log('Setting element text to:', formatted);
            totalIncomeElement.textContent = formatted;
        } else {
            console.log('Could not find total income element or invalid total income');
        }
    }

    // 更新月度收支卡的结余总金额显示
    updateMonthlyBalance(monthlyData) {
        if (!monthlyData || monthlyData.length === 0) return;

        // 计算最近12个月的总结余
        let totalBalance = 0;
        monthlyData.forEach(record => {
            totalBalance += record.balance;
        });

        // 更新页面上的显示
        const balanceElement = document.getElementById('monthlyBalance');
        if (balanceElement) {
            balanceElement.textContent = DataLoader.formatMoney(totalBalance);
        }
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
        this.setupCompoundCalculator();
    }

    setupEventListeners() {
        const refreshBtn = document.getElementById('refreshBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => window.financialDashboard.handleRefresh());
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
        let targetYears = null;
        switch (text) {
            case '最近1年': targetYears = 1; break;
            case '最近3年': targetYears = 3; break;
            case '最近5年': targetYears = 5; break;
            case '全部': targetYears = null; break;
        }

        // 重新渲染图表（null 表示全部）
        this.renderChartsWithTimeRange(targetYears, text);
    }

    // 根据时间范围重新渲染图表
    renderChartsWithTimeRange(targetYears, label) {
        if (!window.financialDashboard || !window.financialDashboard.data) return;

        const data = window.financialDashboard.data;
        let filteredData = data;

        if (targetYears) {
            const latestDate = data[data.length - 1].date;
            const cutoff = new Date(latestDate);
            cutoff.setFullYear(cutoff.getFullYear() - targetYears);

            let startIndex = 0;
            for (let i = 0; i < data.length; i++) {
                if (data[i].date >= cutoff) {
                    startIndex = i;
                    break;
                }
            }
            filteredData = data.slice(startIndex);
        }

        if (filteredData.length > 0) {
            // 确保月度数据也按相同时间范围筛选
            let filteredMonthlyData = window.financialDashboard.monthlyData;
            if (targetYears) {
                const latestDate = data[data.length - 1].date;
                const cutoff = new Date(latestDate);
                cutoff.setFullYear(cutoff.getFullYear() - targetYears);

                // 过滤月度数据以匹配时间范围
                filteredMonthlyData = window.financialDashboard.monthlyData.filter(d => d.date >= cutoff);
            }

            const renderer = new ChartRenderer(filteredData, filteredMonthlyData || null);
            // 重新渲染除月度收支和投资收益外的图表
            renderer.renderChart1();
            renderer.renderChart2();
            renderer.renderChart3();
            renderer.renderChart4();
            renderer.renderChart5(filteredData[filteredData.length - 1]);
            renderer.renderChart6(filteredData[filteredData.length - 1]);
            // 不调用 renderer.renderMonthlyIncomeExpense() 和 renderer.renderInvestmentIncomeChart()
            // 以避免刷新特定卡片

            DataTable.render(filteredData);

            const latest = filteredData[filteredData.length - 1];
            const displayLabel = label || '全部';
            window.financialDashboard.updateDataSourceInfo(
                `数据源：${CONFIG.DATA_SOURCE_NAME} · ${displayLabel} · 最新 ${latest.dateStr}`
            );
        }
    }

    setupRefresh() {
        // 刷新按钮事件在 setupEventListeners 中统一处理
    }

    setupCompoundCalculator() {
        const startAmountInput = document.getElementById('compoundStartAmount');
        const amountInput = document.getElementById('compoundAmount');
        const rateInput = document.getElementById('compoundRate');
        const btn = document.getElementById('compoundCalcBtn');
        const resultEl = document.getElementById('compoundResult');
        const chartContainer = document.querySelector('.compound-chart');
        if (!amountInput || !rateInput || !btn || !resultEl) return;

        // Set the default value of startAmount to latest total assets if we have data
        if (window.financialDashboard && window.financialDashboard.data.length > 0) {
            const latest = window.financialDashboard.data[window.financialDashboard.data.length - 1];
            startAmountInput.value = (latest.totalAssets / 10000).toFixed(1);
        }

        const formatWan = (value) => {
            const wan = value / 10000;
            if (Math.abs(wan) >= 1) return `¥${wan.toFixed(1)}万`;
            return `¥${value.toFixed(0)}`;
        };

        const calculateFV = (startAmount, amount, ratePercent, years) => {
            const r = ratePercent / 100;
            // Future value of compound interest with initial amount + annual contributions
            if (r === 0) {
                return startAmount + amount * years;
            }
            const futureValue = startAmount * Math.pow(1 + r, years) + amount * (Math.pow(1 + r, years) - 1) / r;
            return futureValue;
        };

        let compoundChart = null;

        const renderChart = (startAmount, amount, rate) => {
            if (!chartContainer || typeof Chart === 'undefined') return;
            if (compoundChart) {
                compoundChart.destroy();
            }

            const labels = Array.from({ length: 30 }, (_, i) => i + 1);
            const data = labels.map(year => calculateFV(startAmount, amount, rate, year));

            const ctx = document.createElement('canvas');
            ctx.width = chartContainer.offsetWidth;
            ctx.height = chartContainer.offsetHeight;
            chartContainer.innerHTML = '';
            chartContainer.appendChild(ctx);

            compoundChart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [{
                        label: '复利金额',
                        data: data,
                        borderColor: '#dc2626',
                        backgroundColor: 'rgba(220, 38, 38, 0.08)',
                        borderWidth: 1,
                        tension: 0.4,
                        pointRadius: 0,
                        pointHoverRadius: 0,
                        order: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: { display: false },
                        legend: { display: false },
                        tooltip: {
                            mode: 'index',
                            intersect: false
                        }
                    },
                    scales: {
                        x: {
                            display: true,
                            title: { display: false }
                        },
                        y: {
                            display: true,
                            position: 'right',
                            title: { display: false },
                            ticks: {
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
                        line: { borderWidth: 1 },
                        point: { display: false }
                    }
                }
            });
        };

        const renderResult = () => {
            const startAmountWan = parseFloat(startAmountInput.value) || 0;
            const startAmount = startAmountWan * 10000;
            const amountWan = parseFloat(amountInput.value) || 0;
            const amount = amountWan * 10000;
            const rate = parseFloat(rateInput.value) || 0;
            if (amountWan <= 0 || rate < 0) {
                resultEl.innerHTML = '<span style="color:var(--s2)">请输入有效的年储蓄金额和年化利率</span>';
                return;
            }

            const fv10 = calculateFV(startAmount, amount, rate, 10);
            const fv20 = calculateFV(startAmount, amount, rate, 20);
            const principal10 = startAmount + amount * 10;
            const principal20 = startAmount + amount * 20;

            resultEl.innerHTML = `
                <div>10 年后：<span class="highlight">${formatWan(fv10)}</span>
                    <span style="color:var(--ink-secondary)">（本金 ${formatWan(principal10)}，收益 ${formatWan(fv10 - principal10)}）</span>
                </div>
                <div>20 年后：<span class="highlight">${formatWan(fv20)}</span>
                    <span style="color:var(--ink-secondary)">（本金 ${formatWan(principal20)}，收益 ${formatWan(fv20 - principal20)}）</span>
                </div>
            `;

            renderChart(startAmount, amount, rate);
        };

        // Add event listeners for all inputs
        startAmountInput.addEventListener('input', renderResult);
        amountInput.addEventListener('input', renderResult);
        rateInput.addEventListener('input', renderResult);
        btn.addEventListener('click', renderResult);
        // Default calculation
        renderResult();
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
        this.monthlyData = [];
        this.chartHandler = new FinancialChartHandler();
    }

    async handleFileSelect(file) {
        try {
            let fileName = CONFIG.DATA_SOURCE_NAME;
            if (file instanceof File) {
                fileName = file.name;
            }

            this.updateDataSourceInfo(`正在读取 ${fileName}...`);
            const workbook = await DataLoader.loadFile(file);
            const records = DataLoader.parseWorkbook(workbook);
            if (records.length === 0) {
                alert('未解析到有效数据，请检查 Excel 列是否与文档约定一致。');
                return;
            }

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
            this.updateDataSourceInfo(`数据源：${fileName} · 最新数据 ${latest.dateStr}`);

        } catch (err) {
            console.error(err);
            alert('读取 Excel 失败：' + err.message);
        }
    }

    handleRefresh() {
        this.loadFixedDataSource();
    }

    // 加载固定数据源
    async loadFixedDataSource() {
        try {
            this.updateDataSourceInfo(`正在读取 ${CONFIG.DATA_SOURCE_NAME}...`);

            // 使用 fetch 获取固定路径的 Excel 文件
            const response = await fetch(CONFIG.DATA_SOURCE_PATH);
            if (!response.ok) {
                throw new Error(`无法读取文件：${CONFIG.DATA_SOURCE_PATH}，状态码：${response.status}`);
            }

            const arrayBuffer = await response.arrayBuffer();
            await this.handleFileSelect(arrayBuffer);

        } catch (err) {
            console.error(err);
            alert('读取固定数据源失败：' + err.message);
            this.updateDataSourceInfo('数据源：读取失败');
        }
    }

    // 加载月度收支 CSV 并渲染 p3 卡片条形图
    async loadMonthlyData() {
        try {
            const records = await DataLoader.loadMonthlyCsv(CONFIG.MONTHLY_DATA_PATH);
            this.monthlyData = records;
            this.renderMonthlyChart();
        } catch (err) {
            console.error('读取月度数据失败：', err);
        }
    }

    // 重新设计：为投资收益图表专门加载并渲染
    async loadAndRenderInvestmentIncomeData() {
        try {
            const response = await fetch(CONFIG.INVEST_DATA_PATH);
            if (!response.ok) {
                throw new Error(`无法读取投资收益数据：${CONFIG.INVEST_DATA_PATH}，状态码：${response.status}`);
            }
            const text = await response.text();
            const records = DataLoader.parseInvestmentIncomeCsv(text);

            // 正确更新 monthlyData 以供后续使用
            this.monthlyData = records;

            // 用专门的 ChartRenderer 对象渲染投资收益图表（独立数据源）
            const renderer = new ChartRenderer(this.data, records);
            renderer.renderInvestmentIncomeChart();
        } catch (err) {
            console.error('读取投资收益数据失败：', err);
        }
    }

    // 加载投资收益数据
    async loadInvestmentIncomeData() {
        try {
            // 使用DataLoader解析invest_income_data.csv
            const response = await fetch(CONFIG.INVEST_DATA_PATH);
            if (!response.ok) {
                throw new Error(`无法读取投资收益数据：${CONFIG.INVEST_DATA_PATH}，状态码：${response.status}`);
            }
            const text = await response.text();
            const records = DataLoader.parseInvestmentIncomeCsv(text);
            this.monthlyData = records; // 使用investment data 替代原来的monthly data
            // 立即渲染投资收益图表
            this.renderInvestmentIncomeChart();
        } catch (err) {
            console.error('读取投资收益数据失败：', err);
        }
    }

    renderMonthlyChart() {
        if (this.monthlyData.length === 0) return;
        const renderer = new ChartRenderer(this.data, this.monthlyData);
        renderer.renderMonthlyIncomeExpense();
    }

    // 渲染投资收益图表（从FinancialDashboard调用）
    renderInvestmentIncomeChart() {
        if (this.monthlyData.length === 0) return;
        const renderer = new ChartRenderer(this.data, this.monthlyData);
        renderer.renderInvestmentIncomeChart();
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
document.addEventListener('DOMContentLoaded', async function() {
    window.financialDashboard = new FinancialDashboard();

    // 如果使用固定数据源，则自动加载
    if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.USE_FIXED_DATA_SOURCE) {
        await window.financialDashboard.loadFixedDataSource();
    }

    // 先加载月度收支数据
    await window.financialDashboard.loadMonthlyData();

    // 再加载投资收益数据用于投资收益卡片
    await window.financialDashboard.loadAndRenderInvestmentIncomeData();
});