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

    // 通用折线图渲染
    renderLineChart(container, series, viewBox, plotArea, colors) {
        const svg = container.querySelector('svg');
        if (!svg) return;

        // 保存原始 viewBox 用于交互
        const [vbX, vbY, vbW, vbH] = viewBox;
        svg.setAttribute('viewBox', `${vbX} ${vbY} ${vbW} ${vbH}`);
        svg.setAttribute('preserveAspectRatio', 'none');

        // 清空 SVG（保留 defs 如果有的话）
        const defs = svg.querySelector('defs');
        svg.innerHTML = '';
        if (defs) svg.appendChild(defs);

        const { left, right, top, bottom } = plotArea;
        const plotWidth = right - left;
        const plotHeight = bottom - top;

        // 获取所有数值范围
        const allValues = series.flatMap(s => s.data.map(d => d.value));
        const minValue = Math.min(...allValues, 0);
        const maxValue = Math.max(...allValues);
        const valueRange = maxValue - minValue || 1;

        // 时间范围
        const minTime = this.data[0].date.getTime();
        const maxTime = this.data[this.data.length - 1].date.getTime();
        const timeRange = maxTime - minTime || 1;

        // 映射函数
        const mapX = (date) => left + ((date.getTime() - minTime) / timeRange) * plotWidth;
        const mapY = (value) => bottom - ((value - minValue) / valueRange) * plotHeight;

        // 绘制网格线
        for (let i = 0; i <= 4; i++) {
            const y = top + (plotHeight / 4) * i;
            const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
            line.setAttribute('class', 'gridline');
            line.setAttribute('x1', left);
            line.setAttribute('y1', y);
            line.setAttribute('x2', right);
            line.setAttribute('y2', y);
            svg.appendChild(line);
        }

        // 绘制 x 轴线
        const axisLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        axisLine.setAttribute('class', 'axis-line');
        axisLine.setAttribute('x1', left);
        axisLine.setAttribute('y1', bottom);
        axisLine.setAttribute('x2', right);
        axisLine.setAttribute('y2', bottom);
        svg.appendChild(axisLine);

        // 绘制每个系列
        series.forEach((s, index) => {
            const points = s.data.map(d => ({
                x: mapX(d.date),
                y: mapY(d.value)
            }));

            const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            path.setAttribute('class', 'ln');
            path.setAttribute('style', `stroke:${colors[index]}`);
            path.setAttribute('d', this.generateSmoothPath(points));
            svg.appendChild(path);
        });

        // 生成 x 轴标签（固定位置，避免缩放变形）
        this.generateAxisLabels(container, left, right, minTime, maxTime);

        // 存储参数，供缩放时动态更新标签
        container.dataset.minTime = minTime;
        container.dataset.maxTime = maxTime;
        container.dataset.plotLeft = left;
        container.dataset.plotRight = right;
        container.dataset.origViewBoxW = vbW;
    }

    // 生成平滑曲线路径（三次贝塞尔）
    generateSmoothPath(points) {
        if (points.length === 0) return '';
        if (points.length === 1) return `M${points[0].x},${points[0].y}`;
        if (points.length === 2) return `M${points[0].x},${points[0].y} L${points[1].x},${points[1].y}`;

        let d = `M${points[0].x},${points[0].y}`;
        for (let i = 0; i < points.length - 1; i++) {
            const p0 = points[i === 0 ? 0 : i - 1];
            const p1 = points[i];
            const p2 = points[i + 1];
            const p3 = points[i + 2] || p2;

            const cp1x = p1.x + (p2.x - p0.x) / 6;
            const cp1y = p1.y + (p2.y - p0.y) / 6;
            const cp2x = p2.x - (p3.x - p1.x) / 6;
            const cp2y = p2.y - (p3.y - p1.y) / 6;

            d += ` C${cp1x},${cp1y} ${cp2x},${cp2y} ${p2.x},${p2.y}`;
        }
        return d;
    }

    // 在 SVG 内生成时间标签（后续会被提取为 HTML）
    generateAxisLabels(container, left, right, minTime, maxTime) {
        const svg = container.querySelector('svg');
        const labelCount = 5;
        for (let i = 0; i < labelCount; i++) {
            const t = minTime + (maxTime - minTime) * (i / (labelCount - 1));
            const date = new Date(t);
            const x = left + (right - left) * (i / (labelCount - 1));
            const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            text.setAttribute('class', 'tick');
            text.setAttribute('x', x);
            text.setAttribute('y', parseFloat(svg.getAttribute('viewBox').split(/\s+/)[3]) - 8);
            text.textContent = `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}`;
            svg.appendChild(text);
        }
    }

    renderChart1() {
        const container = document.querySelector('.card:nth-child(1) .chart');
        this.renderLineChart(container, [
            { name: '总资产', data: this.data.map(d => ({ date: d.date, value: d.totalAssets })) },
            { name: '负债', data: this.data.map(d => ({ date: d.date, value: d.debt })) },
            { name: '净资产', data: this.data.map(d => ({ date: d.date, value: d.netAssets })) }
        ], [0, 0, 640, 220], { left: 50, right: 640, top: 20, bottom: 200 }, ['var(--s1)', 'var(--s2)', 'var(--s3)']);
    }

    renderChart2() {
        const container = document.querySelector('.card:nth-child(2) .chart');
        this.renderLineChart(container, [
            { name: '活期', data: this.data.map(d => ({ date: d.date, value: d.currentDeposit })) },
            { name: '应急资金', data: this.data.map(d => ({ date: d.date, value: d.emergencyFund })) },
            { name: '流动资金', data: this.data.map(d => ({ date: d.date, value: d.liquidAssets })) }
        ], [0, 0, 300, 200], { left: 40, right: 300, top: 20, bottom: 180 }, ['var(--s1)', 'var(--s6)', 'var(--s3)']);
    }

    renderChart3() {
        const container = document.querySelector('.card:nth-child(3) .chart');
        this.renderLineChart(container, [
            { name: '储备金', data: this.data.map(d => ({ date: d.date, value: d.reserveFund })) },
            { name: '稳健基金', data: this.data.map(d => ({ date: d.date, value: d.stableFund })) },
            { name: '风险基金', data: this.data.map(d => ({ date: d.date, value: d.riskFund })) },
            { name: '股票', data: this.data.map(d => ({ date: d.date, value: d.stock })) }
        ], [0, 0, 300, 200], { left: 40, right: 300, top: 20, bottom: 180 }, ['var(--s7)', 'var(--s3)', 'var(--s4)', 'var(--s5)']);
    }

    renderChart4() {
        const container = document.querySelector('.card:nth-child(4) .chart');
        this.renderLineChart(container, [
            { name: '公积金', data: this.data.map(d => ({ date: d.date, value: d.providentFund })) }
        ], [0, 0, 300, 200], { left: 40, right: 300, top: 20, bottom: 180 }, ['var(--s4)']);
    }

    renderPieChart(svg, items, centerX, centerY, radius) {
        svg.innerHTML = '';
        const total = items.reduce((sum, item) => sum + item.value, 0);
        const circumference = 2 * Math.PI * radius;
        let offset = 0;

        items.forEach(item => {
            const ratio = total > 0 ? item.value / total : 0;
            const dashArray = `${circumference * ratio} ${circumference}`;
            const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
            circle.setAttribute('cx', centerX);
            circle.setAttribute('cy', centerY);
            circle.setAttribute('r', radius);
            circle.setAttribute('fill', 'none');
            circle.setAttribute('stroke', item.color);
            circle.setAttribute('stroke-width', '24');
            circle.setAttribute('stroke-dasharray', dashArray);
            circle.setAttribute('stroke-dashoffset', -offset);
            circle.setAttribute('transform', `rotate(-90 ${centerX} ${centerY})`);
            svg.appendChild(circle);
            offset += circumference * ratio;
        });

        // 中心文字
        const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        text.setAttribute('class', 'tick');
        text.setAttribute('x', centerX);
        text.setAttribute('y', centerY - 5);
        text.setAttribute('text-anchor', 'middle');
        text.setAttribute('fill', 'var(--ink-primary)');
        text.textContent = items[0]?.label === '流动资金' ? '总资产' : '投资资金';
        svg.appendChild(text);
    }

    renderChart5(latest) {
        const svg = document.querySelector('.card:nth-child(5) .chart svg');
        if (!svg) return;
        this.renderPieChart(svg, [
            { color: 'var(--s1)', value: latest.liquidAssets },
            { color: 'var(--s2)', value: latest.fixedAssets },
            { color: 'var(--s3)', value: latest.investment }
        ], 130, 85, 52);
    }

    renderChart6(latest) {
        const svg = document.querySelector('.card:nth-child(6) .chart svg');
        if (!svg) return;
        this.renderPieChart(svg, [
            { color: 'var(--s6)', value: latest.reserveFund },
            { color: 'var(--s4)', value: latest.riskPosition },
            { color: 'var(--s3)', value: latest.stableFund }
        ], 130, 85, 52);
    }
}

// ========== 交互处理 ==========
class FinancialChartHandler {
    static totalYears = 2;
    static defaultYears = 1;
    static minYears = 0.5;

    constructor() {
        this.chartStates = new Map();
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.setupZoomAndPan();
        this.setupHoverEffects();
        this.setupTimeFilter();
        this.setupRefresh();
    }

    reinit() {
        this.chartStates.clear();
        this.setupFixedAxisLabels();
        this.setupZoomAndPan();
        this.setupHoverEffects();
    }

    setupFixedAxisLabels() {
        const charts = document.querySelectorAll('.chart');
        charts.forEach(chart => {
            // 移除已有的外部标签
            const existing = chart.querySelector('.x-axis-labels');
            if (existing) existing.remove();

            const svg = chart.querySelector('svg');
            if (!svg) return;

            const texts = svg.querySelectorAll('text.tick, text.tick-label');
            if (texts.length === 0) return;

            const [ox, , ow] = (svg.getAttribute('viewBox') || '0 0 600 220').split(/\s+/).map(Number);

            let startX = Infinity;
            texts.forEach(text => {
                const x = parseFloat(text.getAttribute('x')) || 0;
                startX = Math.min(startX, x);
            });
            const endX = ox + ow;
            const plotWidth = endX - startX;
            const leftMarginRatio = (startX - ox) / ow;

            const xAxis = document.createElement('div');
            xAxis.className = 'x-axis-labels';
            xAxis.style.left = `${leftMarginRatio * 100}%`;
            xAxis.style.right = `${leftMarginRatio * 100}%`;

            texts.forEach(text => {
                const x = parseFloat(text.getAttribute('x')) || 0;
                const label = text.textContent.trim();
                const span = document.createElement('span');
                span.textContent = label;
                span.style.left = `${((x - startX) / plotWidth) * 100}%`;
                xAxis.appendChild(span);
                text.remove();
            });

            chart.appendChild(xAxis);
        });
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

    setupZoomAndPan() {
        const charts = document.querySelectorAll('.chart');
        charts.forEach(chart => {
            if (this.isPieChart(chart)) return;
            this.initChartInteraction(chart);
        });
    }

    isPieChart(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        return svg ? svg.querySelectorAll('circle').length > 1 : false;
    }

    initChartInteraction(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        if (!svg) return;

        const originalViewBox = svg.getAttribute('viewBox') || '0 0 600 220';
        const [ox, oy, ow, oh] = originalViewBox.split(/\s+/).map(Number);

        const { totalYears, defaultYears, minYears } = FinancialChartHandler;

        const state = {
            ox, oy, ow, oh,
            timeRatio: defaultYears / totalYears,
            panX: 0,
            isDragging: false,
            startX: 0,
            startPanX: 0,
            minTimeRatio: minYears / totalYears,
            maxTimeRatio: 1
        };

        this.chartStates.set(chartContainer, state);

        chartContainer.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            const s = this.chartStates.get(chartContainer);
            s.isDragging = true;
            s.startX = e.clientX;
            s.startPanX = s.panX;
            chartContainer.style.cursor = 'grabbing';
        });

        chartContainer.addEventListener('mousemove', (e) => {
            const s = this.chartStates.get(chartContainer);
            if (!s || !s.isDragging) return;
            const pixelDeltaX = e.clientX - s.startX;
            const svgDeltaX = pixelDeltaX * (s.ow / chartContainer.clientWidth) * s.timeRatio;
            s.panX = s.startPanX + svgDeltaX;
            this.applyTimeWindow(chartContainer);
        });

        document.addEventListener('mouseup', () => {
            const s = this.chartStates.get(chartContainer);
            if (s && s.isDragging) {
                s.isDragging = false;
                chartContainer.style.cursor = 'grab';
            }
        });

        chartContainer.addEventListener('wheel', (e) => {
            e.preventDefault();
            const s = this.chartStates.get(chartContainer);
            const delta = e.deltaY > 0 ? -0.1 : 0.1;
            s.timeRatio = Math.max(s.minTimeRatio, Math.min(s.maxTimeRatio, s.timeRatio + delta));
            this.applyTimeWindow(chartContainer);
        }, { passive: false });

        chartContainer.addEventListener('mouseleave', () => {
            const s = this.chartStates.get(chartContainer);
            if (s && s.isDragging) {
                s.isDragging = false;
                chartContainer.style.cursor = 'grab';
            }
        });

        this.applyTimeWindow(chartContainer);
    }

    applyTimeWindow(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        if (!svg) return;
        const s = this.chartStates.get(chartContainer);
        if (!s) return;

        const windowWidth = s.ow * s.timeRatio;
        let windowX = s.ox + s.ow - windowWidth + s.panX;

        const minWindowX = s.ox;
        const maxWindowX = s.ox + s.ow - windowWidth;
        windowX = Math.min(maxWindowX, Math.max(minWindowX, windowX));

        svg.setAttribute('viewBox', `${windowX} ${s.oy} ${windowWidth} ${s.oh}`);

        // 更新 x 轴标签以反映当前可见的时间范围
        this.updateXLabels(chartContainer);
    }

    // 动态更新 x 轴标签，使其随缩放/平移而变化
    updateXLabels(chartContainer) {
        const container = chartContainer;
        const svg = container.querySelector('svg');
        const s = this.chartStates.get(container);
        if (!svg || !s) return;

        // 读取渲染时存储的参数
        const minTime = parseInt(container.dataset.minTime);
        const maxTime = parseInt(container.dataset.maxTime);
        const plotLeft = parseFloat(container.dataset.plotLeft);
        const plotRight = parseFloat(container.dataset.plotRight);
        const origOw = parseFloat(container.dataset.origViewBoxW);

        if (isNaN(minTime) || isNaN(maxTime) || isNaN(plotLeft) || isNaN(plotRight) || isNaN(origOw)) return;

        const timeRange = maxTime - minTime || 1;
        const plotWidth = plotRight - plotLeft;

        // 当前 viewBox
        const vb = svg.getAttribute('viewBox').split(/\s+/).map(Number);
        const windowX = vb[0];
        const windowWidth = vb[2];

        // viewBox x 坐标 ⟼ 时间
        const mapXToTime = (x) => minTime + ((x - plotLeft) / plotWidth) * timeRange;

        // 可见时间范围（钳位到数据范围）
        const visibleStartX = Math.max(windowX, plotLeft);
        const visibleEndX = Math.min(windowX + windowWidth, plotRight);
        if (visibleStartX >= visibleEndX) {
            const xAxisDiv = container.querySelector('.x-axis-labels');
            if (xAxisDiv) xAxisDiv.innerHTML = '';
            return;
        }

        const visibleStartTime = Math.max(minTime, mapXToTime(visibleStartX));
        const visibleEndTime = Math.min(maxTime, mapXToTime(visibleEndX));
        const visibleTimeRange = visibleEndTime - visibleStartTime || 1;

        const labelCount = 5;
        const leftMarginRatio = plotLeft / origOw;
        const xAxisDiv = container.querySelector('.x-axis-labels');
        if (!xAxisDiv) return;

        xAxisDiv.innerHTML = '';

        for (let i = 0; i < labelCount; i++) {
            const t = visibleStartTime + visibleTimeRange * (i / (labelCount - 1));
            const date = new Date(t);
            // 该时间对应的 viewBox x 坐标
            const vx = plotLeft + ((t - minTime) / timeRange) * plotWidth;
            // 在图表中的位置比例
            const chartFrac = (vx - windowX) / windowWidth;
            // 在 x 轴标签容器内的位置百分比
            const posInContainer = (chartFrac - leftMarginRatio) / (1 - 2 * leftMarginRatio) * 100;

            const span = document.createElement('span');
            span.textContent = `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}`;
            span.style.left = `${posInContainer}%`;
            xAxisDiv.appendChild(span);
        }
    }

    setupHoverEffects() {
        const chartPaths = document.querySelectorAll('.ln, .line');
        chartPaths.forEach(path => {
            path.addEventListener('mouseenter', (e) => this.highlightElement(e.target));
            path.addEventListener('mouseleave', (e) => this.resetHighlight(e.target));
        });
    }

    highlightElement(element) {
        element.style.strokeWidth = '3px';
        element.classList.add('highlight');
    }

    resetHighlight(element) {
        element.style.strokeWidth = '2px';
        element.classList.remove('highlight');
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

        this.chartStates.forEach((state, chart) => {
            const targetRatio = Math.min(state.maxTimeRatio, Math.max(state.minTimeRatio, targetYears / totalYears));
            state.timeRatio = targetRatio;
            state.panX = 0;
            this.applyTimeWindow(chart);
        });

        console.log('时间范围切换到: ' + text);
    }

    setupRefresh() {
        // 刷新按钮事件在 setupEventListeners 中统一处理
    }

    resetAllCharts() {
        const { defaultYears, totalYears } = FinancialChartHandler;
        this.chartStates.forEach((state, chart) => {
            state.timeRatio = defaultYears / totalYears;
            state.panX = 0;
            this.applyTimeWindow(chart);
        });
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

            // 渲染图表
            const renderer = new ChartRenderer(this.data);
            renderer.renderAll();

            // 重新初始化交互（包含提取固定标签）
            this.chartHandler.reinit();

            // 填充表格
            DataTable.render(this.data);

            // 更新数据源信息
            const latest = this.data[this.data.length - 1];
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
        if (this.data.length < 2) return 2;
        const start = this.data[0].date;
        const end = this.data[this.data.length - 1].date;
        const days = (end - start) / (1000 * 60 * 60 * 24);
        return Math.max(2, days / 365);
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