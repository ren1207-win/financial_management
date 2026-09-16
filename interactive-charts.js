// 交互功能实现
class FinancialChartHandler {
    // 时间跨度配置
    // 假设当前静态 SVG 的完整宽度代表 2 年时间跨度
    static totalYears = 2;
    static defaultYears = 1;
    static minYears = 0.5;

    constructor() {
        // 记录每个图表的交互状态
        this.chartStates = new Map();
        this.init();
    }

    init() {
        this.setupFixedAxisLabels();
        this.setupEventListeners();
        this.setupZoomAndPan();
        this.setupHoverEffects();
        this.setupTimeFilter();
        this.setupRefresh();
    }

    // 将 SVG 内部的时间轴文字提取为 HTML 元素，避免缩放时变形或大小变化
    setupFixedAxisLabels() {
        const charts = document.querySelectorAll('.chart');
        charts.forEach(chart => {
            const svg = chart.querySelector('svg');
            if (!svg) return;

            // 只提取时间轴刻度文字（tick / tick-label 类）
            const texts = svg.querySelectorAll('text.tick, text.tick-label');
            if (texts.length === 0) return;

            const [ox, , ow] = (svg.getAttribute('viewBox') || '0 0 600 220').split(/\s+/).map(Number);

            // 计算 SVG 绘图区域边界，使标签左右边距对称
            // 左侧边距由最小刻度 x 决定，右侧对齐 viewBox 右边界
            let startX = Infinity;
            texts.forEach(text => {
                const x = parseFloat(text.getAttribute('x')) || 0;
                startX = Math.min(startX, x);
            });
            const endX = ox + ow;
            const plotWidth = endX - startX;
            const leftMarginRatio = (startX - ox) / ow;

            // 创建外部时间轴标签容器，左右边距与左侧边距相同以保持对称
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

    // 设置基础事件监听器
    setupEventListeners() {
        const refreshBtn = document.querySelector('.refresh-btn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => this.handleRefresh());
        }

        const filterPills = document.querySelectorAll('.filter-pill');
        filterPills.forEach(pill => {
            pill.addEventListener('click', (e) => this.handleTimeFilter(e));
        });
    }

    // 设置时间轴缩放和平移（仅针对曲线图，不针对饼图）
    setupZoomAndPan() {
        const charts = document.querySelectorAll('.chart');
        charts.forEach(chart => {
            // 饼图不参与时间轴缩放
            if (this.isPieChart(chart)) return;
            this.initChartInteraction(chart);
        });
    }

    // 判断是否为饼图（通过是否存在 circle 元素简单判断）
    isPieChart(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        return svg ? svg.querySelectorAll('circle').length > 1 : false;
    }

    // 初始化单个曲线图的交互
    initChartInteraction(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        if (!svg) return;

        // 保存原始 viewBox，代表完整时间序列
        const originalViewBox = svg.getAttribute('viewBox') || '0 0 600 220';
        const [ox, oy, ow, oh] = originalViewBox.split(/\s+/).map(Number);

        // timeRatio：当前窗口占全部时间跨度的比例
        // 1.0 = 显示全部时间；0.5 = 显示一半时间；0.25 = 显示半年
        const { totalYears, defaultYears, minYears } = FinancialChartHandler;

        const state = {
            ox, oy, ow, oh,
            timeRatio: defaultYears / totalYears, // 默认显示一年
            panX: 0,
            isDragging: false,
            startX: 0,
            startPanX: 0,
            minTimeRatio: minYears / totalYears,  // 最小半年
            maxTimeRatio: 1                       // 最大全部时间
        };

        this.chartStates.set(chartContainer, state);

        // 鼠标按下开始拖拽
        chartContainer.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            const s = this.chartStates.get(chartContainer);
            s.isDragging = true;
            s.startX = e.clientX;
            s.startPanX = s.panX;
            chartContainer.style.cursor = 'grabbing';
        });

        // 鼠标移动：平移时间窗口
        chartContainer.addEventListener('mousemove', (e) => {
            const s = this.chartStates.get(chartContainer);
            if (!s || !s.isDragging) return;

            // 把屏幕像素位移转换为 SVG 坐标系位移
            const pixelDeltaX = e.clientX - s.startX;
            const svgDeltaX = pixelDeltaX * (s.ow / chartContainer.clientWidth) * s.timeRatio;
            s.panX = s.startPanX + svgDeltaX;

            this.applyTimeWindow(chartContainer);
        });

        // 鼠标释放
        document.addEventListener('mouseup', () => {
            const s = this.chartStates.get(chartContainer);
            if (s && s.isDragging) {
                s.isDragging = false;
                chartContainer.style.cursor = 'grab';
            }
        });

        // 鼠标滚轮：调整时间窗口大小，右侧固定
        chartContainer.addEventListener('wheel', (e) => {
            e.preventDefault();
            const s = this.chartStates.get(chartContainer);

            // 滚轮向上（deltaY < 0）放大时间窗口（看到更久）
            // 滚轮向下（deltaY > 0）缩小时间窗口（聚焦近期）
            const delta = e.deltaY > 0 ? -0.1 : 0.1;
            s.timeRatio = Math.max(s.minTimeRatio, Math.min(s.maxTimeRatio, s.timeRatio + delta));

            this.applyTimeWindow(chartContainer);
        }, { passive: false });

        // 鼠标移出时取消拖拽
        chartContainer.addEventListener('mouseleave', () => {
            const s = this.chartStates.get(chartContainer);
            if (s && s.isDragging) {
                s.isDragging = false;
                chartContainer.style.cursor = 'grab';
            }
        });

        // 初始应用一次默认窗口
        this.applyTimeWindow(chartContainer);
    }

    // 应用时间窗口到 viewBox
    // 策略：右侧固定，左侧移动；timeRatio 表示当前窗口占完整时间跨度的比例
    applyTimeWindow(chartContainer) {
        const svg = chartContainer.querySelector('svg');
        if (!svg) return;

        const s = this.chartStates.get(chartContainer);
        if (!s) return;

        // 窗口宽度 = 原始宽度 × 时间跨度倍数
        const windowWidth = s.ow * s.timeRatio;

        // 右侧固定：窗口的右端对齐原始 viewBox 右端，再加上平移偏移
        // 平移为正值时，窗口向右（更新）移动；为负值时，窗口向左（更旧）移动
        let windowX = s.ox + s.ow - windowWidth + s.panX;

        // 限制窗口不超出完整时间序列范围
        const minWindowX = s.ox;
        const maxWindowX = s.ox + s.ow - windowWidth;
        windowX = Math.min(maxWindowX, Math.max(minWindowX, windowX));

        svg.setAttribute('viewBox', `${windowX} ${s.oy} ${windowWidth} ${s.oh}`);
    }

    // 设置悬停效果
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

    // 时间筛选处理
    handleTimeFilter(e) {
        const filterPills = document.querySelectorAll('.filter-pill');
        filterPills.forEach(pill => pill.classList.remove('active'));
        e.target.classList.add('active');

        const text = e.target.textContent.trim();
        const { totalYears } = FinancialChartHandler;
        let targetYears = totalYears; // 默认全部
        switch (text) {
            case '最近1年': targetYears = 1; break;
            case '最近3年': targetYears = 3; break;
            case '最近5年': targetYears = 5; break;
            case '全部': targetYears = totalYears; break;
        }

        // 对所有曲线图应用对应时间窗口
        this.chartStates.forEach((state, chart) => {
            const targetRatio = Math.min(state.maxTimeRatio, Math.max(state.minTimeRatio, targetYears / totalYears));
            state.timeRatio = targetRatio;
            state.panX = 0;
            this.applyTimeWindow(chart);
        });

        console.log('时间范围切换到: ' + text);
    }

    // 刷新处理
    handleRefresh() {
        console.log('数据刷新按钮被点击，开始加载数据...');
        const refreshBtn = document.querySelector('.refresh-btn');
        refreshBtn.innerHTML = '↻ 刷新中...';
        refreshBtn.disabled = true;

        setTimeout(() => {
            refreshBtn.innerHTML = '↻ 刷新数据';
            refreshBtn.disabled = false;
            console.log('数据刷新完成');
            this.resetAllCharts();
        }, 1000);
    }

    // 重置所有曲线图到默认窗口
    resetAllCharts() {
        const { defaultYears, totalYears } = FinancialChartHandler;
        this.chartStates.forEach((state, chart) => {
            state.timeRatio = defaultYears / totalYears;
            state.panX = 0;
            this.applyTimeWindow(chart);
        });
    }

    // Tooltip 工具方法（供后续扩展）
    addTooltip(chartContainer, x, y, text) {
        let tooltip = chartContainer.querySelector('.chart-tooltip');
        if (!tooltip) {
            tooltip = document.createElement('div');
            tooltip.className = 'chart-tooltip';
            tooltip.style.cssText = `
                position: absolute;
                background: rgba(0,0,0,0.8);
                color: white;
                padding: 5px 10px;
                border-radius: 4px;
                font-size: 12px;
                pointer-events: none;
                z-index: 1000;
                display: none;
            `;
            chartContainer.appendChild(tooltip);
        }

        tooltip.textContent = text;
        tooltip.style.display = 'block';
        tooltip.style.left = (x + 10) + 'px';
        tooltip.style.top = (y - 10) + 'px';

        setTimeout(() => {
            tooltip.style.display = 'none';
        }, 3000);
    }
}

// 页面加载完成后初始化交互
document.addEventListener('DOMContentLoaded', function() {
    new FinancialChartHandler();
});