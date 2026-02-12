/**
 * 图表渲染模块
 * 负责所有图表的创建、更新和交互
 */

const ChartRenderer = (function() {
    let loadCurveChart = null;
    let dailyBarChart = null;
    let heatmapCanvas = null;

    const CHART_COLORS = [
        'rgb(79, 70, 229)',
        'rgb(99, 102, 241)',
        'rgb(139, 92, 246)',
        'rgb(168, 85, 247)',
        'rgb(217, 70, 239)',
        'rgb(236, 72, 153)',
        'rgb(244, 63, 94)',
        'rgb(239, 68, 68)',
        'rgb(249, 115, 22)',
        'rgb(245, 158, 11)',
        'rgb(234, 179, 8)',
        'rgb(132, 204, 22)'
    ];

    function getColor(index) {
        return CHART_COLORS[index % CHART_COLORS.length];
    }

    function initLoadCurveChart(canvasId) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return null;

        const ctx = canvas.getContext('2d');
        
        loadCurveChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: Array.from({ length: 24 }, (_, i) => `${i}:00`),
                datasets: []
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: {
                        display: false,
                        position: 'top'
                    },
                    tooltip: {
                        enabled: true,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        titleColor: '#fff',
                        bodyColor: '#cbd5e1',
                        borderColor: 'rgba(79, 70, 229, 0.3)',
                        borderWidth: 1,
                        cornerRadius: 12,
                        padding: 12,
                        displayColors: true,
                        callbacks: {
                            label: function(context) {
                                return `${context.dataset.label}: ${context.parsed.y.toFixed(2)} kWh`;
                            }
                        }
                    },
                    title: {
                        display: true,
                        text: '24小时负荷曲线',
                        font: { size: 16, weight: 'bold' },
                        color: '#1e293b'
                    }
                },
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: '时间 (小时)',
                            font: { size: 12, weight: 'bold' },
                            color: '#64748b'
                        },
                        grid: {
                            color: 'rgba(148, 163, 184, 0.1)'
                        }
                    },
                    y: {
                        display: true,
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '用电量 (kWh)',
                            font: { size: 12, weight: 'bold' },
                            color: '#64748b'
                        },
                        grid: {
                            color: 'rgba(148, 163, 184, 0.1)'
                        }
                    }
                },
                animation: {
                    duration: 750,
                    easing: 'easeOutQuart'
                }
            }
        });

        return loadCurveChart;
    }

    function updateLoadCurveChart(data, options = {}) {
        if (!loadCurveChart) {
            initLoadCurveChart('loadCurveChart');
        }

        if (!data || data.length === 0) {
            loadCurveChart.data.datasets = [];
            loadCurveChart.update('none');
            return;
        }

        const {
            lineStyle = 'solid',
            showPoints = false,
            smoothCurve = true,
            curveTension = 0.4,
            focusStart = 0,
            focusEnd = 23
        } = options;

        const datasets = data.map((day, index) => {
            const color = getColor(index);
            const filteredData = day.hourlyData.slice(focusStart, focusEnd + 1);
            const labels = Array.from({ length: focusEnd - focusStart + 1 }, (_, i) => `${focusStart + i}:00`);
            
            loadCurveChart.data.labels = labels;

            return {
                label: day.date,
                data: filteredData,
                borderColor: color,
                backgroundColor: color.replace('rgb', 'rgba').replace(')', ', 0.1)'),
                borderWidth: 2,
                borderDash: lineStyle === 'dashed' ? [10, 5] : lineStyle === 'dotted' ? [2, 2] : [],
                pointRadius: showPoints ? 3 : 0,
                pointHoverRadius: showPoints ? 5 : 0,
                tension: smoothCurve ? curveTension : 0,
                fill: false,
                spanGaps: true
            };
        });

        loadCurveChart.data.datasets = datasets;
        loadCurveChart.update('active');
    }

    function initDailyBarChart(canvasId) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return null;

        const ctx = canvas.getContext('2d');
        
        dailyBarChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [],
                datasets: [{
                    label: '日用电量 (kWh)',
                    data: [],
                    backgroundColor: 'rgba(79, 70, 229, 0.7)',
                    borderColor: 'rgb(79, 70, 229)',
                    borderWidth: 1,
                    borderRadius: 6,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        enabled: true,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        titleColor: '#fff',
                        bodyColor: '#cbd5e1',
                        cornerRadius: 12,
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                return `用电量: ${context.parsed.y.toFixed(2)} kWh`;
                            }
                        }
                    },
                    title: {
                        display: true,
                        text: '日期总用电量对比',
                        font: { size: 16, weight: 'bold' },
                        color: '#1e293b'
                    }
                },
                scales: {
                    x: {
                        display: true,
                        grid: {
                            display: false
                        }
                    },
                    y: {
                        display: true,
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '用电量 (kWh)',
                            font: { size: 12, weight: 'bold' },
                            color: '#64748b'
                        },
                        grid: {
                            color: 'rgba(148, 163, 184, 0.1)'
                        }
                    }
                },
                animation: {
                    duration: 750,
                    easing: 'easeOutQuart'
                }
            }
        });

        return dailyBarChart;
    }

    function updateDailyBarChart(data, aggregation = 'daily') {
        if (!dailyBarChart) {
            initDailyBarChart('dailyTotalBarChart');
        }

        if (!data || data.length === 0) {
            dailyBarChart.data.labels = [];
            dailyBarChart.data.datasets[0].data = [];
            dailyBarChart.update('none');
            return;
        }

        let labels = [];
        let values = [];

        if (aggregation === 'monthly') {
            const monthlyData = new Map();
            
            data.forEach(day => {
                const date = day.dateObj || new Date(day.date);
                const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                
                if (!monthlyData.has(monthKey)) {
                    monthlyData.set(monthKey, 0);
                }
                
                const dayTotal = day.hourlyData.reduce((sum, val) => sum + (val || 0), 0);
                monthlyData.set(monthKey, monthlyData.get(monthKey) + dayTotal);
            });

            monthlyData.forEach((value, key) => {
                labels.push(key);
                values.push(value);
            });
        } else {
            data.forEach(day => {
                labels.push(day.date);
                values.push(day.hourlyData.reduce((sum, val) => sum + (val || 0), 0));
            });
        }

        dailyBarChart.data.labels = labels;
        dailyBarChart.data.datasets[0].data = values;
        dailyBarChart.update('active');
    }

    function updateHeatmap(data, canvasId) {
        const canvas = document.getElementById(canvasId);
        if (!canvas || !data || data.length === 0) return;

        const ctx = canvas.getContext('2d');
        const container = canvas.parentElement;
        const width = container.clientWidth;
        const height = container.clientHeight;

        canvas.width = width;
        canvas.height = height;

        ctx.clearRect(0, 0, width, height);

        const cellWidth = width / 24;
        const cellHeight = Math.min(30, height / data.length);

        let maxVal = -Infinity;
        let minVal = Infinity;

        data.forEach(day => {
            day.hourlyData.forEach(val => {
                if (val !== null) {
                    maxVal = Math.max(maxVal, val);
                    minVal = Math.min(minVal, val);
                }
            });
        });

        const range = maxVal - minVal || 1;

        data.forEach((day, dayIndex) => {
            day.hourlyData.forEach((val, hourIndex) => {
                const x = hourIndex * cellWidth;
                const y = dayIndex * cellHeight;
                
                let color;
                if (val === null) {
                    color = '#f1f5f9';
                } else {
                    const ratio = (val - minVal) / range;
                    color = getHeatmapColor(ratio);
                }
                
                ctx.fillStyle = color;
                ctx.fillRect(x, y, cellWidth - 1, cellHeight - 1);
            });
        });
    }

    function getHeatmapColor(ratio) {
        if (ratio < 0.5) {
            const r = Math.round(219 + (253 - 219) * (ratio * 2));
            const g = Math.round(234 + (224 - 234) * (ratio * 2));
            const b = Math.round(254 + (97 - 254) * (ratio * 2));
            return `rgb(${r}, ${g}, ${b})`;
        } else {
            const adjustedRatio = (ratio - 0.5) * 2;
            const r = Math.round(253 + (248 - 253) * adjustedRatio);
            const g = Math.round(224 + (113 - 224) * adjustedRatio);
            const b = Math.round(97 + (113 - 97) * adjustedRatio);
            return `rgb(${r}, ${g}, ${b})`;
        }
    }

    function exportChartAsImage(chartId, format = 'png', quality = 0.9) {
        const canvas = document.getElementById(chartId);
        if (!canvas) return null;

        if (format === 'svg') {
            return canvas.toDataURL('image/svg+xml');
        }

        const mimeType = format === 'jpg' ? 'image/jpeg' : 'image/png';
        return canvas.toDataURL(mimeType, quality);
    }

    function destroyCharts() {
        if (loadCurveChart) {
            loadCurveChart.destroy();
            loadCurveChart = null;
        }
        if (dailyBarChart) {
            dailyBarChart.destroy();
            dailyBarChart = null;
        }
    }

    function showChartLoading(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const loading = container.querySelector('.chart-loading');
        if (loading) {
            loading.classList.remove('hidden');
        }
    }

    function hideChartLoading(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const loading = container.querySelector('.chart-loading');
        if (loading) {
            loading.classList.add('hidden');
        }
    }

    function showNoDataMessage(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const noData = container.querySelector('.no-data-message');
        if (noData) {
            noData.classList.remove('hidden');
        }
    }

    function hideNoDataMessage(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const noData = container.querySelector('.no-data-message');
        if (noData) {
            noData.classList.add('hidden');
        }
    }

    return {
        initLoadCurveChart,
        updateLoadCurveChart,
        initDailyBarChart,
        updateDailyBarChart,
        updateHeatmap,
        exportChartAsImage,
        destroyCharts,
        showChartLoading,
        hideChartLoading,
        showNoDataMessage,
        hideNoDataMessage,
        getColor,
        CHART_COLORS
    };
})();

window.ChartRenderer = ChartRenderer;

export default ChartRenderer;
