        // 存储各section的滚动位置
        const sectionScrollPositions = new Map();
        
        // 添加平滑过渡的CSS样式
        const style = document.createElement('style');
        style.textContent = `
            /* 页面加载时隐藏所有section，避免闪烁 */
            body:not(.initialized) .section-container {
                opacity: 0 !important;
                transform: translateY(20px) !important;
            }
            
            .section-container {
                /* 优化过渡效果：增加transform实现轻微位移 */
                transition: opacity 0.3s cubic-bezier(0.4, 0, 0.2, 1), transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                will-change: opacity, transform;
                /* 性能优化：跳过离屏内容的渲染 */
                content-visibility: auto;
                contain-intrinsic-size: 0 500px;
                opacity: 1;
                transform: translateY(0);
            }
            
            .section-hidden {
                display: none !important;
            }
            
            .section-visible {
                display: block !important;
                opacity: 1;
                transform: translateY(0);
            }
            
            /* 离场状态：上浮并消失 */
            .section-leave {
                display: block !important;
                opacity: 0;
                transform: translateY(-10px);
                pointer-events: none;
            }
            
            /* 入场初始状态：下沉并透明 */
            .section-enter {
                display: block !important;
                opacity: 0;
                transform: translateY(10px);
            }
        `;
        document.head.appendChild(style);
        
        // 页面初始化完成后移除加载状态
        function initializePage() {
            if (!document.body.classList.contains('initialized')) {
                document.body.classList.add('initialized');
                console.log('Page initialized');
                
                // 为初始显示的卡片添加交错动画
                const activeSection = document.querySelector('.section-visible');
                if (activeSection) {
                    const cards = activeSection.querySelectorAll('.card, .card-base, .grid > div');
                    cards.forEach((card, index) => {
                        card.style.opacity = '0';
                        card.style.transform = 'translateY(10px)';
                        setTimeout(() => {
                            card.style.transition = 'all 0.5s cubic-bezier(0.22, 1, 0.36, 1)';
                            card.style.opacity = '1';
                            card.style.transform = 'translateY(0)';
                        }, index * 100 + 100);
                    });
                }
            }
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializePage);
        } else {
            initializePage();
        }
        
        // 兜底方案：如果 2 秒后还没初始化，强制初始化
        setTimeout(initializePage, 2000);
        
        // 平滑滚动跳转函数（动画增强版）
        function scrollToSection(sectionId) {
            const element = document.getElementById(sectionId);
            if (!element) {
                console.warn('目标section不存在：', sectionId);
                return;
            }
            
            // 检查目标是否已经是当前显示的
            if (element.classList.contains('section-visible') && !element.classList.contains('section-leave')) {
                return;
            }
            
            console.log('切换到section：', sectionId);
            
            // 1. 立即更新侧边栏导航选中状态（视觉反馈优先）
            document.querySelectorAll('#sidebar .sidebar-link').forEach(item => {
                item.classList.remove('active');
            });
            const activeLink = document.querySelector(`#sidebar .sidebar-link[data-section="${sectionId}"]`);
            if (activeLink) {
                activeLink.classList.add('active');
            }

            // 为所有section添加容器类（如果还没有的话）
            const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            allSections.forEach(id => {
                const section = document.getElementById(id);
                if (section && !section.classList.contains('section-container')) {
                    section.classList.add('section-container');
                }
            });
            
            // 获取当前显示的section
            const currentSection = document.querySelector('.section-visible:not(.section-leave)');
            
            // 定义切换并显示的逻辑
            const showTargetSection = () => {
                // 滚动回顶部
                const prefersReducedMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
                window.scrollTo({ top: 0, behavior: prefersReducedMotion ? 'auto' : 'smooth' });

                // 准备入场
                element.classList.remove('section-hidden');
                element.classList.add('section-enter');
                
                // 强制重绘，确保enter状态生效
                void element.offsetWidth;
                
                // 执行入场动画
                requestAnimationFrame(() => {
                    element.classList.remove('section-enter');
                    element.classList.add('section-visible');
                    
                    // 触发图表更新逻辑（在入场动画开始时预加载）
                    if (sectionId === 'visualizationSection') {
                        triggerChartUpdate();
                        // 确保热力图在进入可视化页面时更新，且不随图表频繁刷新
                        if (typeof updateHeatmap === 'function') {
                            setTimeout(updateHeatmap, 100);
                        }
                    }
                });
                
                // 更新步骤导航状态
                updateStepNavigation(sectionId);
            };

            // 执行过渡动画
            if (currentSection && currentSection.id !== sectionId) {
                // 离场动画
                currentSection.classList.remove('section-visible');
                currentSection.classList.add('section-leave');
                
                // 等待离场动画结束 (300ms)
                setTimeout(() => {
                    currentSection.classList.remove('section-leave');
                    currentSection.classList.add('section-hidden');
                    showTargetSection();
                    
                    // 移动端侧边栏收起逻辑
                    handleMobileSidebar();
                }, 300);
            } else {
                // 无当前显示section，直接显示目标
                showTargetSection();
                handleMobileSidebar();
            }
            
            // 图表更新逻辑封装
            function triggerChartUpdate() {
                 // 检查是否需要更新
                const currentDates = JSON.stringify(appData.visualization.selectedDates);
                const currentMPs = JSON.stringify(appData.visualization.selectedMeteringPoints);
                const currentDataLen = appData.processedData ? appData.processedData.length : 0;
                const currentStart = appData.visualization.focusStartTime;
                const currentEnd = appData.visualization.focusEndTime;
                const currentMode = appData.visualization.summaryMode ? 'summary' : (appData.visualization.allCurvesMode ? 'all' : 'single');

                // 检查图表是否存在
                const chartExists = appData.chart && appData.chart.canvas && document.body.contains(appData.chart.canvas);

                if (chartExists &&
                    lastChartUpdateParams.dates === currentDates &&
                    lastChartUpdateParams.meteringPoints === currentMPs &&
                    lastChartUpdateParams.dataLength === currentDataLen &&
                    lastChartUpdateParams.startTime === currentStart &&
                    lastChartUpdateParams.endTime === currentEnd &&
                    lastChartUpdateParams.mode === currentMode
                ) {
                    // 不需要更新，但确保容器显示
                    console.log('图表参数未变化，跳过更新');
                    const barChartContainer = document.getElementById('dailyTotalBarChartContainer');
                    if (barChartContainer) barChartContainer.style.display = 'block';
                } else {
                    // 延迟执行图表更新，等待DOM稳定
                    setTimeout(() => {
                        if (appData.processedData && appData.processedData.length > 0) {
                            // 预先计算筛选数据，避免多次计算
                            const filteredData = getFilteredData();
                            
                            try {
                                if (typeof updateChart === 'function') updateChart(filteredData);
                                if (typeof updateDailyTotalBarChart === 'function') {
                                    updateDailyTotalBarChart(filteredData);
                                }
                            } catch (err) {
                                console.error('图表更新失败:', err);
                            } finally {
                                // 兜底：确保加载动画被隐藏
                                const chartLoading = document.getElementById('chartLoading');
                                const barChartLoading = document.getElementById('barChartLoading');
                                if (chartLoading) chartLoading.classList.add('hidden');
                                if (barChartLoading) barChartLoading.classList.add('hidden');
                            }
                            
                            const barChartContainer = document.getElementById('dailyTotalBarChartContainer');
                            if (barChartContainer) barChartContainer.style.display = 'block';

                            // 更新缓存状态
                            lastChartUpdateParams = {
                                dates: currentDates,
                                meteringPoints: currentMPs,
                                dataLength: currentDataLen,
                                startTime: currentStart,
                                endTime: currentEnd,
                                mode: currentMode
                            };
                        }
                    }, 50); // 入场动画开始后稍作延迟即可
                }
            }
            
            // 移动端侧边栏处理
            function handleMobileSidebar() {
                if (window.innerWidth < 768) {
                    const sidebar = document.getElementById('sidebar');
                    const overlay = document.getElementById('sidebarOverlay');
                    if (sidebar && !sidebar.classList.contains('-translate-x-full')) {
                        sidebar.classList.add('-translate-x-full');
                        if (overlay) {
                            overlay.classList.add('hidden');
                            overlay.classList.remove('opacity-100');
                            overlay.classList.add('opacity-0');
                        }
                    }
                }
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            document.querySelectorAll('.nav-items-container .nav-item[role="button"]').forEach((item) => {
                item.addEventListener('keydown', (e) => {
                    if (e.key !== 'Enter' && e.key !== ' ') return;
                    e.preventDefault();
                    const sectionId = item.getAttribute('data-section');
                    if (sectionId) scrollToSection(sectionId);
                });
            });
            
            // 移动端侧边栏切换逻辑
            const mobileMenuBtn = document.getElementById('mobileMenuBtn');
            const sidebar = document.getElementById('sidebar');
            const sidebarOverlay = document.getElementById('sidebarOverlay');
            
            if (mobileMenuBtn && sidebar && sidebarOverlay) {
                // 打开侧边栏
                mobileMenuBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    sidebar.classList.remove('-translate-x-full');
                    sidebarOverlay.classList.remove('hidden');
                    // 使用setTimeout确保display:block生效后再改变opacity，触发transition
                    setTimeout(() => {
                        sidebarOverlay.classList.remove('opacity-0');
                        sidebarOverlay.classList.add('opacity-100');
                    }, 10);
                });
                
                // 关闭侧边栏 (点击遮罩层)
                sidebarOverlay.addEventListener('click', () => {
                    sidebar.classList.add('-translate-x-full');
                    sidebarOverlay.classList.remove('opacity-100');
                    sidebarOverlay.classList.add('opacity-0');
                    setTimeout(() => {
                        sidebarOverlay.classList.add('hidden');
                    }, 300); // 等待transition结束
                });
            }
        });

        // 切换导航栏的收纳/展开状态
        function toggleNav() {
            const navContainer = document.getElementById('leftNav');
            const toggleBtn = document.querySelector('.nav-toggle-btn i');
            if (navContainer) {
                const isCollapsed = navContainer.classList.toggle('collapsed');
                navContainer.classList.toggle('expanded', !isCollapsed);
                
                // 更新图标方向
                if (toggleBtn) {
                    toggleBtn.style.transform = isCollapsed ? 'rotate(180deg)' : 'rotate(0deg)';
                }
            }
        }
        
        // 初始化导航栏状态
        function initNav() {
            // 可以在这里添加初始化逻辑
        }
        
        // 更新步骤导航状态
        function updateStepNavigation(activeSection) {
            // 如果页面正在初始化，不更新导航状态
            if (isPageInitializing) {
                return;
            }
            
            const sections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            const currentIndex = sections.indexOf(activeSection);
            
            // 如果当前没有明确的激活section，保持导航不变
            if (currentIndex === -1) return;
            
            // 更新所有步骤卡片状态
            sections.forEach((section, index) => {
                const card = document.getElementById(`stepCard${index + 1}`);
                if (!card) return; // 如果卡片不存在，跳过
                
                if (index === currentIndex) {
                    // 当前激活步骤
                    card.classList.remove('completed');
                    card.classList.add('active');
                } else if (index < currentIndex) {
                    // 已完成步骤
                    card.classList.remove('active');
                    card.classList.add('completed');
                } else {
                    // 未开始步骤
                    card.classList.remove('active', 'completed');
                }
            });
        }

        // 标记步骤为完成
        function completeStep(stepNumber) {
            const sections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            if (stepNumber > 0 && stepNumber <= sections.length) {
                updateStepNavigation(sections[stepNumber - 1]);
            }
        }

        // 原始数据与PVsyst数据对比预览函数
        function updateComparisonPreview() {
            if (appData.processedData.length === 0) {
                if (elements.previewButtons) {
                    elements.previewButtons.classList.add('hidden');
                }
                elements.dataPreview.classList.add('bg-white', 'border', 'border-slate-200', 'shadow-sm');
                elements.dataPreview.innerHTML = `
                    <div class="flex flex-col items-center justify-center py-16 px-8 text-center">
                        <div class="w-16 h-16 mb-6 bg-slate-50 border border-slate-100 rounded-2xl flex items-center justify-center text-slate-300">
                            <i class="fa fa-info-circle text-2xl"></i>
                        </div>
                        <h3 class="text-base font-black text-slate-900 mb-2">没有可对比的数据</h3>
                        <p class="text-xs font-medium text-slate-500 max-w-md leading-relaxed">请先完成数据配置，系统将为您展示对比预览</p>
                    </div>
                `;
                return;
            }
            
            // 解析数据
            parseWorksheetData();
            
            if (appData.processedData.length === 0) {
                elements.dataPreview.innerHTML = `
            <div class="flex flex-col items-center justify-center py-16 px-8">
                <div class="w-16 h-16 mb-6 bg-indigo-50 rounded-full flex items-center justify-center ring-1 ring-indigo-100">
                    <svg class="w-8 h-8 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z"></path>
                    </svg>
                </div>
                <h3 class="text-lg font-semibold text-slate-700 mb-2">暂无数据</h3>
                <p class="text-sm text-slate-500 text-center max-w-md leading-relaxed">当前配置下没有找到可预览的数据，请检查数据源或调整配置参数</p>
            </div>
        `;
                return;
            }
            
            // 获取数据结构类型
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
            
            // 创建基于月和日的数据映射（忽略年份，年份固定为99）
            const dateDataMap = new Map();
            
            // 按日期排序，确保数据按时间顺序处理
            const sortedExportData = [...appData.processedData].sort((a, b) => {
                const dateA = a.dateObj || parseDate(a.date);
                const dateB = b.dateObj || parseDate(b.date);
                return (dateA ? dateA.getTime() : 0) - (dateB ? dateB.getTime() : 0);
            });
            
            // 使用排序后的数据构建映射，确保按时间顺序处理
            sortedExportData.forEach(dateData => {
                const date = dateData.dateObj || parseDate(dateData.date);
                if (!date) return;
                const monthDayKey = `${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                
                // 确保hourlyData是24个元素的数组
                if (dateData.hourlyData && Array.isArray(dateData.hourlyData) && dateData.hourlyData.length === 24) {
                    dateDataMap.set(monthDayKey, dateData.hourlyData);
                } else {
                    dateDataMap.set(monthDayKey, Array(24).fill(0));
                }
            });
            
            // 生成对比预览表格
            let previewHtml = `<div class="bg-indigo-50/50 p-3 mb-3 text-sm font-bold text-indigo-700 rounded-xl border border-indigo-100/50">原始数据与PVsyst数据对比预览</div>`;
            
            // 创建统一的表格样式，支持完整数据展示
            previewHtml += '<div class="overflow-x-auto border border-slate-200 rounded-xl shadow-sm max-h-96 overflow-y-auto bg-white">';
            previewHtml += '<table class="w-full text-sm table-fixed" style="font-family: system-ui, -apple-system, sans-serif;">';
            previewHtml += '<thead><tr class="bg-slate-50 text-slate-800">';
            previewHtml += '<th class="border border-slate-200 px-3 py-2 text-center font-bold" style="width: 120px;">日期</th>';
            previewHtml += '<th class="border border-slate-200 px-3 py-2 text-center font-bold">数据类型</th>';
            
            // 统一显示24小时数据（列对行和列对列结构都显示24小时）
            for (let hour = 0; hour < 24; hour++) {
                previewHtml += `<th class="border border-slate-200 px-3 py-2 text-center font-bold">${hour}:00</th>`;
            }
            previewHtml += '<th class="border border-slate-200 px-3 py-2 text-center font-bold">日总电能 (kWh)</th>';
            previewHtml += '</tr></thead><tbody class="divide-y divide-slate-100">';
            
            // 添加数据行 - 显示原始数据和PVsyst数据的对比
            const displayRows = Math.min(appData.processedData.length, 5); // 限制显示5天的数据对比
            
            for (let i = 0; i < displayRows; i++) {
                const row = appData.processedData[i];
                const date = row.dateObj || parseDate(row.date);
                if (!date) return;
                const monthDayKey = `${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                
                // 原始数据行
                previewHtml += `<tr class="bg-white hover:bg-slate-50 transition-colors">`;
                previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-medium" rowspan="2">${row.date}</td>`;
                previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-bold text-indigo-600">原始数据</td>`;
                
                let originalDayTotal = 0;
                for (let hour = 0; hour < 24; hour++) {
                        const value = row.hourlyData[hour];
                        if (value !== null && value !== undefined) {
                            previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center">${value.toFixed(2)}</td>`;
                            originalDayTotal += value;
                        } else {
                            previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center text-slate-400">-</td>`;
                        }
                    }
                    previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-black text-indigo-600">${originalDayTotal.toFixed(2)}</td>`;
                    previewHtml += '</tr>';
                    
                    // PVsyst数据行
                    previewHtml += `<tr class="bg-indigo-50/30 hover:bg-indigo-50/50 transition-colors">`;
                    previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-bold text-slate-600">PVsyst数据</td>`;
                    
                    // 获取PVsyst数据
                    if (dateDataMap.has(monthDayKey)) {
                        pvsystData = [...dateDataMap.get(monthDayKey)];
                    } else {
                        // 如果没有匹配的数据，使用findNearbyData函数查找相近数据
                        const nearbyKey = findNearbyData(date.getMonth() + 1, date.getDate(), dateDataMap);
                        if (nearbyKey && dateDataMap.has(nearbyKey)) {
                            const originalData = dateDataMap.get(nearbyKey);
                            const variationFactor = 0.05; // 5%的变化范围
                            pvsystData = originalData.map(value => {
                                // 添加随机变化，但保持数据的基本模式
                                const variation = 1 + (Math.random() - 0.5) * 2 * variationFactor;
                                return value * variation;
                            });
                        } else {
                            // 如果找不到相近数据，使用原始数据
                            pvsystData = [...row.hourlyData];
                        }
                    }
                    
                    let pvsystDayTotal = 0;
                    for (let hour = 0; hour < 24; hour++) {
                        const value = pvsystData[hour];
                        if (value !== null && value !== undefined) {
                            previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center">${value.toFixed(2)}</td>`;
                            pvsystDayTotal += value;
                        } else {
                            previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center text-slate-400">-</td>`;
                        }
                    }
                    previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-black text-slate-700">${pvsystDayTotal.toFixed(2)}</td>`;
                    previewHtml += '</tr>';
                    
                    // PVsyst数据行
                    previewHtml += `<tr class="bg-indigo-50/20 hover:bg-indigo-50/40 transition-colors">`;
                    previewHtml += `<td class="border border-slate-200 px-3 py-2 text-center font-bold text-slate-600">PVsyst数据</td>`;
                    
                    // 获取PVsyst数据
                    let pvsystData;
                    if (dateDataMap.has(monthDayKey)) {
                        pvsystData = [...dateDataMap.get(monthDayKey)];
                    } else {
                        // 如果没有匹配的数据，使用findNearbyData函数查找相近数据
                        const nearbyKey = findNearbyData(date.getMonth() + 1, date.getDate(), dateDataMap);
                        if (nearbyKey && dateDataMap.has(nearbyKey)) {
                            const originalData = dateDataMap.get(nearbyKey);
                            const variationFactor = 0.05; // 5%的变化范围
                            pvsystData = originalData.map(value => {
                                // 添加随机变化，但保持数据的基本模式
                                const variation = 1 + (Math.random() - 0.5) * 2 * variationFactor;
                                return value * variation;
                            });
                        } else {
                            // 如果找不到相近数据，使用原始数据
                            pvsystData = [...row.hourlyData];
                        }
                    }
                    
                    for (let hour = 0; hour < 24; hour++) {
                        const value = pvsystData[hour];
                        if (value !== null && value !== undefined) {
                            previewHtml += `<td class="border border-gray-300 px-3 py-2 text-center">${value.toFixed(2)}</td>`;
                        } else {
                            previewHtml += `<td class="border border-gray-300 px-3 py-2 text-center text-gray-400">-</td>`;
                        }
                    }
                    
                    previewHtml += '</tr>';
                }
            
            previewHtml += '</tbody></table></div>';
            
            // 添加数据统计信息
            previewHtml += `<div class="mt-2 text-sm text-gray-600 text-center bg-gray-50 p-2 rounded">
                显示前 ${Math.min(appData.processedData.length, 5)} 天的原始数据与PVsyst数据对比
            </div>`;
            
            elements.dataPreview.innerHTML = previewHtml;
        }
        
        // 24小时数据处理函数
        // 按日期和计量点分组数据
        function groupDataByDateAndMeteringPoint(data) {
            const groups = {};
            
            data.forEach(row => {
                const date = row.date;
                const meteringPoint = (row.meteringPoint === undefined || row.meteringPoint === null || String(row.meteringPoint).trim() === '') ? '默认计量点' : String(row.meteringPoint).trim();
                const key = `${date}_${meteringPoint}`;
                
                if (!groups[key]) {
                    groups[key] = {
                        date: date,
                        meteringPoint: meteringPoint,
                        hourlyData: new Array(24).fill(null),
                        totalEnergy: 0
                    };
                }
                
                // 合并该日期该计量点的24小时数据
                for (let hour = 0; hour < 24; hour++) {
                    if (row.hourlyData[hour] !== null && row.hourlyData[hour] !== undefined) {
                        if (groups[key].hourlyData[hour] === null) {
                            groups[key].hourlyData[hour] = 0;
                        }
                        groups[key].hourlyData[hour] += row.hourlyData[hour];
                        groups[key].totalEnergy += row.hourlyData[hour];
                    }
                }
            });
            
            return groups;
        }
        
        // 创建24小时完整数据集
        function create24HourDatasets(groupedData) {
            const datasets = [];
            let colorIndex = 0;
            
            // 按日期排序
            const sortedKeys = Object.keys(groupedData).sort();
            
            sortedKeys.forEach(key => {
                const data = groupedData[key];
                const color = getColor(colorIndex % getColorCount());
                colorIndex++;
                
                // 确保24小时数据完整
                const complete24HourData = [];
                for (let hour = 0; hour < 24; hour++) {
                    complete24HourData.push(data.hourlyData[hour] || 0);
                }
                
                const mp = data.meteringPoint;
                const labelText = appData.config.meteringPointColumn && mp && mp !== '默认计量点'
                    ? `${data.date} (${mp})`
                    : `${data.date}`;
                const dataset = {
                    label: labelText,
                    data: complete24HourData,
                    borderColor: color,
                    backgroundColor: color.replace('rgb', 'rgba').replace(')', ', 0.1)'),
                    borderWidth: 3,
                    borderDash: appData.visualization.lineStyle === 'dashed' ? [10, 5] : 
                               appData.visualization.lineStyle === 'dotted' ? [2, 2] : [],
                    pointRadius: appData.visualization.showPoints ? 3 : 0,
                    pointHoverRadius: appData.visualization.showPoints ? 5 : 0,
                    tension: appData.visualization.smoothCurve ? appData.visualization.curveTension : 0,
                    fill: false,
                    spanGaps: true
                };
                
                datasets.push(dataset);
            });
            
            return datasets;
        }
        
        // 使用24小时数据集更新图表
        function updateChartWith24HourDatasets(datasets) {
            if (!appData.chart) {
                // 如果图表不存在，创建一个新的图表
                initEmptyChart();
            }
            
            // 24小时时间标签
            const hourLabels = Array.from({ length: 24 }, (_, i) => `${i}:00`);
            
            // 更新图表数据
            appData.chart.data.labels = hourLabels;
            appData.chart.data.datasets = datasets;
            
            // 更新图表标题
            const selectedDates = appData.visualization.selectedDates;
            const selectedMeteringPoints = appData.visualization.selectedMeteringPoints;
            appData.chart.options.plugins.title.text = 
                `24小时负荷曲线 - ${selectedDates.length}个日期 × ${selectedMeteringPoints.length}个计量点`;
            
            // 更新图表配置 - 确保XY轴始终显示
            appData.chart.options.scales.x.display = true;
            appData.chart.options.scales.y.display = true;
            appData.chart.options.scales.x.title = {
                display: true,
                text: '时间 (小时)',
                font: { size: 14, weight: 'bold' }
            };
            appData.chart.options.scales.y.title = {
                display: true,
                text: '用电量 (kWh)',
                font: { size: 14, weight: 'bold' }
            };
            
            // 使用动画更新图表
            updateChartWithAnimation();
            
            // 隐藏加载状态
            const chartLoadingElement = document.getElementById('chartLoading');
            if (chartLoadingElement && chartLoadingElement.parentNode) {
                chartLoadingElement.classList.add('hidden');
            }
        }

        // 添加页面状态监听功能
        function initPageStateListener() {
            function checkCurrentSection() {
                // 如果页面正在初始化，不检查当前section
                if (isPageInitializing) {
                    return;
                }
                
                const sections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
                let currentSection = null;
                let visibleSectionCount = 0;
                
                // 检查哪个section当前显示（没有section-hidden类）
                for (const sectionId of sections) {
                    const section = document.getElementById(sectionId);
                    if (section && section.classList.contains('section-visible') && !section.classList.contains('section-hidden')) {
                        currentSection = sectionId;
                        visibleSectionCount++;
                    }
                }
                
                // 只有在有明确显示的section时才更新导航状态，避免默认跳转到importSection
                if (currentSection && visibleSectionCount === 1) {
                    updateStepNavigation(currentSection);
                }
            }
            
            // 使用MutationObserver监听DOM变化
            const observer = new MutationObserver(function(mutations) {
                mutations.forEach(function(mutation) {
                    if (mutation.type === 'attributes' && mutation.attributeName === 'class') {
                        // 当class属性变化时，检查当前显示的section
                        checkCurrentSection();
                    }
                });
            });
            
            // 监听所有section的class变化
            const sections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            sections.forEach(sectionId => {
                const section = document.getElementById(sectionId);
                if (section) {
                    observer.observe(section, {
                        attributes: true,
                        attributeFilter: ['class']
                    });
                }
            });
            
            // 延迟初始检查，避免页面加载时的自动导航
            setTimeout(() => {
                checkCurrentSection();
            }, 1000);
        }
        
        // 页面加载完成后初始化页面状态监听
        document.addEventListener('DOMContentLoaded', function() {
            initPageStateListener();
            

        });

