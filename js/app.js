        // 全局数据存储
        // 缓存上次图表更新时的参数，用于避免不必要的重绘
        let lastChartUpdateParams = {
            dates: null,
            meteringPoints: null,
            dataLength: 0,
            startTime: null,
            endTime: null,
            mode: null
        };

        const appData = {
            files: [],
            worksheets: [],
            parsedData: [],
            processedData: [],
            originalData: null,
            config: {
                dataType: 'instantPower',
                multiplier: 1.0,
                useMultiplierColumn: false,
                multiplierColumn: '',
                multiplierMode: 'single', // 'single' | 'ptct' - 倍率模式：单列或PT/CT分开
                ptColumn: '', // PT列
                ctColumn: '', // CT列
                dateColumn: '',
                timeColumn: '',
                useDateRange: false,
                dataStartColumn: '',
                dataEndColumn: '',
                dataStartIndex: 0,
                dataEndIndex: -1,
                meteringPointColumn: '',
                meteringPointFilter: '',
                dateFormat: 'YYYY-MM-DD',
                invalidDataHandling: 'ignore',
                timeInterval: 15,
                additionalFilters: []
            },
            visualization: {
        selectedDates: [],
        selectedMeteringPoints: [],
        focusStartTime: 0,
        focusEndTime: 23,
        lineStyle: 'solid',
        showPoints: false,
        secondaryAxis: false,
        smoothCurve: true,
        curveTension: 0.4,
        showDailyTotalBarChart: true,
        barChartAggregation: 'daily', // 'daily' or 'monthly'
        allCurvesMode: false, // 默认显示单日日期曲线
        summaryMode: false, // 是否处于汇总曲线模式
            summarySubMode: 'all' // 'all' | 'extreme' | 'quarter' | 'month' - 汇总子模式
        },
            chart: null,
            dailyTotalBarChart: null, // 存储日期总用电量图表引用
            meteringPoints: [], // 存储所有计量点编号
            // 内存管理配置
            memoryManagement: {
                maxDataRows: 100000, // 单个文件最大行数限制
                maxTotalMemoryMB: 200, // 总内存使用限制（MB）
                currentMemoryUsageMB: 0, // 当前内存使用量估算
                enableGarbageCollection: true // 是否启用垃圾回收优化
            }
        };
        
        // 智能数字格式化函数（处理大数值显示）
        function smartFormatNumber(n, fraction = 2) {
            if (typeof n !== 'number' || !isFinite(n)) return '—';
            const abs = Math.abs(n);
            // 超过10亿 -> G
            if (abs >= 1e9) return (n / 1e9).toFixed(fraction) + ' G';
            // 超过100万 -> M
            if (abs >= 1e6) return (n / 1e6).toFixed(fraction) + ' M';
            // 其他情况使用千分位
            return n.toLocaleString(undefined, { minimumFractionDigits: fraction, maximumFractionDigits: fraction });
        }

        // DOM 元素
        const elements = {
            // 步骤指示器
            
            // 导入部分
            dropArea: document.getElementById('dropArea'),
            fileInput: document.getElementById('fileInput'),
            fileList: document.getElementById('fileList'),
            importedFiles: document.getElementById('importedFiles'),
            removeAllFiles: document.getElementById('removeAllFiles'),

            uploadProgress: document.getElementById('uploadProgress'),
            uploadProgressBar: document.getElementById('progressBar'),
        uploadProgressText: document.getElementById('progressText'),
            importStatus: document.getElementById('importStatus'),
            fileListContent: document.getElementById('fileListContent'),
            
            // 配置部分
            dataTypeRadios: document.querySelectorAll('input[name="dataType"]'),
            multiplierInput: document.getElementById('multiplier'),
            useMultiplierColumnCheckbox: document.getElementById('useMultiplierColumn'),
            multiplierOptions: document.getElementById('multiplierOptions'),
            multiplierContent: document.getElementById('multiplierContent'),
            multiplierColumnOption: document.getElementById('multiplierColumnOption'),
            multiplierColumnSelect: document.getElementById('multiplierColumn'),
            ptColumnSelect: document.getElementById('ptColumn'),
            ctColumnSelect: document.getElementById('ctColumn'),
            singleMultiplierSection: document.getElementById('singleMultiplierSection'),
            ptCtMultiplierSection: document.getElementById('ptCtMultiplierSection'),
            ptCtPreview: document.getElementById('ptCtPreview'),
            multiplierModeSelector: document.getElementById('multiplierModeSelector'),
            notificationIconContainer: document.getElementById('notificationIconContainer'),

            dateColumnSelect: document.getElementById('dateColumn'),
                dataStartColumnSelect: document.getElementById('dataStartColumn'),
            dataEndColumnSelect: document.getElementById('dataEndColumn'),
            meteringPointColumnSelect: document.getElementById('meteringPointColumn'),
            meteringPointFilterSelect: document.getElementById('meteringPointFilter'),
            meteringPointFilterSelect_v2: document.getElementById('meteringPointFilter_v2'),
            meteringPointColumnSelect_v2: document.getElementById('meteringPointColumn_v2'),
            clearMeteringPointColumnBtn_v2: document.getElementById('clearMeteringPointColumn_v2'),
            clearMeteringPointFilterBtn_v2: document.getElementById('clearMeteringPointFilter_v2'),

            invalidDataHandlingSelect: document.getElementById('invalidDataHandling'),
            timeIntervalSelect: document.getElementById('timeInterval'),
            configStatus: document.getElementById('configStatus'),
            dataPreview: document.getElementById('dataPreview'),
            previewButtons: document.getElementById('previewButtons'),
            previewFirstHourCalculationBtn: document.getElementById('previewFirstHourCalculation'),
            firstHourCalculationModal: document.getElementById('firstHourCalculationModal'),
            closeFirstHourCalculationModalBtn: document.getElementById('closeFirstHourCalculationModal'),
            previewFirstDayCalculationBtn: document.getElementById('previewFirstDayCalculation'),
            previewDate: document.getElementById('previewDate'),
            previewInterval: document.getElementById('previewInterval'),
            previewDataType: document.getElementById('previewDataType'),
            previewMultiplier: document.getElementById('previewMultiplier'),
            previewRawData: document.getElementById('previewRawData'),
            previewCleanedData: document.getElementById('previewCleanedData'),
            previewCalculationProcess: document.getElementById('previewCalculationProcess'),
            previewResult: document.getElementById('previewResult'),
            previewFormula: document.getElementById('previewFormula'),
            
            // 第一天计算过程相关
            firstDayCalculationModal: document.getElementById('firstDayCalculationModal'),
            closeFirstDayCalculationModalBtn: document.getElementById('closeFirstDayCalculationModal'),
            previewFirstDayDate: document.getElementById('previewFirstDayDate'),
            previewFirstDayInterval: document.getElementById('previewFirstDayInterval'),
            previewFirstDayDataType: document.getElementById('previewFirstDayDataType'),
            previewFirstDayMultiplier: document.getElementById('previewFirstDayMultiplier'),
            previewFirstDayRawData: document.getElementById('previewFirstDayRawData'),
            previewFirstDayCleanedData: document.getElementById('previewFirstDayCleanedData'),
            previewFirstDayCalculationProcess: document.getElementById('previewFirstDayCalculationProcess'),
            previewFirstDayResult: document.getElementById('previewFirstDayResult'),
            previewFirstDayFormula: document.getElementById('previewFirstDayFormula'),

            
            focusStartTimeSelect: document.getElementById('focusStartTime'),
            focusEndTimeSelect: document.getElementById('focusEndTime'),
            startDateInput: document.getElementById('startDateInput'),
            endDateInput: document.getElementById('endDateInput'),
            prevDayBtn: document.getElementById('prevDayBtn'),
            nextDayBtn: document.getElementById('nextDayBtn'),
            resetTimeFocusBtn: document.getElementById('resetTimeFocus'),
            openAdvancedFiltersBtn: document.getElementById('openAdvancedFilters'),
            openAdvancedFiltersBtn_v2: document.getElementById('openAdvancedFilters_v2'),
            openCurveStyleBtn: document.getElementById('openCurveStyleBtn'),
            curveStyleModal: document.getElementById('curveStyleModal'),
            closeCurveStyleModalBtn: document.getElementById('closeCurveStyleModal'),
            applyCurveStyleBtn: document.getElementById('applyCurveStyle'),
            lineStyleSelect: document.getElementById('lineStyle'),
            showPointsSelect: document.getElementById('showPoints'),
            smoothCurveSelect: document.getElementById('smoothCurve'),
            curveTensionSelect: document.getElementById('curveTension'),
            secondaryAxisCheckbox: document.getElementById('secondaryAxis'),
            loadCurveChart: document.getElementById('loadCurveChart'),
            chartLoading: document.getElementById('chartLoading'),
            noDataMessage: document.getElementById('noDataMessage'),
            dailyStats: document.getElementById('dailyStats'),
            periodStats: document.getElementById('periodStats'),

            // 热力图部分
            loadHeatmap: document.getElementById('loadHeatmap'),
            heatmapLoading: document.getElementById('heatmapLoading'),
            heatmapNoDataMessage: document.getElementById('heatmapNoDataMessage'),

            // 聚合切换按钮
            barChartAggregationDailyBtn: document.getElementById('barChartAggregationDailyBtn'),
            barChartAggregationMonthlyBtn: document.getElementById('barChartAggregationMonthlyBtn'),
            
            // 导出部分
            exportPNGBtn: document.getElementById('exportPNG'),
            pngQualitySelect: document.getElementById('pngQuality'),
            exportJPGBtn: document.getElementById('exportJPG'),
            exportSVGBtn: document.getElementById('exportSVG'),
            exportPDFBtn: document.getElementById('exportPDF'),
            export15MinDataBtn: document.getElementById('export15MinData'),
            exportHourlyDataBtn: document.getElementById('exportHourlyData'),
            exportDailyStatsBtn: document.getElementById('exportDailyStats'),
            exportSummaryCurveDataBtn: document.getElementById('exportSummaryCurveData'),
            exportFullPackageBtn: document.getElementById('exportFullPackage'),
            exportFormatSelect: document.getElementById('exportFormat'),
            exportFileNameInput: document.getElementById('exportFileName'),
            exportNameTemplateInput: document.getElementById('exportNameTemplate'),
            includeTimestampCheckbox: document.getElementById('includeTimestamp'),
            includeFiltersInNameCheckbox: document.getElementById('includeFiltersInName'),
            exportWeekdaysOnlyCheckbox: document.getElementById('exportWeekdaysOnly'),
            exportAnomalyMarksCheckbox: document.getElementById('exportAnomalyMarks'),
            exportSamplingIntervalSelect: document.getElementById('exportSamplingInterval'),
            exportSamplingAggSelect: document.getElementById('exportSamplingAgg'),
            exportZipBundleBtn: document.getElementById('exportZipBundle'),

            exportPreviewModal: document.getElementById('exportPreviewModal'),
            exportPreviewCloseBtn: document.getElementById('exportPreviewClose'),
            exportPreviewCancelBtn: document.getElementById('exportPreviewCancel'),
            exportPreviewConfirmBtn: document.getElementById('exportPreviewConfirm'),
            exportPreviewFileName: document.getElementById('exportPreviewFileName'),
            exportPreviewSummary: document.getElementById('exportPreviewSummary'),
            exportPreviewRange: document.getElementById('exportPreviewRange'),
            exportPreviewChartWrap: document.getElementById('exportPreviewChartWrap'),
            exportPreviewChartImg: document.getElementById('exportPreviewChartImg'),
            exportPreviewBarChartWrap: document.getElementById('exportPreviewBarChartWrap'),
            exportPreviewBarChartImg: document.getElementById('exportPreviewBarChartImg'),
            exportChartSelect: document.getElementById('exportChartSelect'),

            startOverBtn: document.getElementById('startOverBtn'),
            
            // 通知和模态框
            notification: document.getElementById('notification'),
            notificationIcon: document.getElementById('notificationIcon'),
            notificationTitle: document.getElementById('notificationTitle'),
            notificationMessage: document.getElementById('notificationMessage'),
            closeNotificationBtn: document.getElementById('closeNotification'),
            modalOverlay: document.getElementById('modalOverlay'),
            modal: document.getElementById('modal'),
            modalTitle: document.getElementById('modalTitle'),
            modalContent: document.getElementById('modalContent'),
            closeModalBtn: document.getElementById('closeModal'),
            helpBtn: document.getElementById('helpBtn'),
            aboutBtn: document.getElementById('aboutBtn'),
            
            //  sections
            importSection: document.getElementById('importSection'),
            configSection: document.getElementById('configSection'),
            visualizationSection: document.getElementById('visualizationSection'),
            exportSection: document.getElementById('exportSection'),
            loadExampleDataBtn: document.getElementById('loadExampleData')
        };

        // 初始化事件监听
        function initEventListeners() {
            const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

            // 导入部分
            elements.dropArea.addEventListener('click', () => elements.fileInput.click());
            elements.fileInput.addEventListener('change', handleFileSelection);
            elements.dropArea.addEventListener('dragover', handleDragOver);
            elements.dropArea.addEventListener('dragleave', handleDragLeave);
            elements.dropArea.addEventListener('drop', handleDrop);
            elements.removeAllFiles.addEventListener('click', (e) => {
                e.stopPropagation();
                removeAllFiles();
            });

            if (elements.loadExampleDataBtn) {
                elements.loadExampleDataBtn.addEventListener('click', loadExampleData);
            }

            
            // 配置部分
            elements.dataTypeRadios.forEach(radio => {
                radio.addEventListener('change', (e) => {
                    appData.config.dataType = e.target.value;
                    if (appData.originalData) {
                        reprocessDataWithConfig();
                    } else {
                        updateDataPreview();
                    }
                });
            });
            
            elements.multiplierInput.addEventListener('change', (e) => {
                appData.config.multiplier = parseFloat(e.target.value) || 1.0;
                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
            });
            
            elements.useMultiplierColumnCheckbox.addEventListener('change', (e) => {
                appData.config.useMultiplierColumn = e.target.checked;
                if (e.target.checked) {
                    if (elements.multiplierOptions) elements.multiplierOptions.classList.add('hidden');
                    if (elements.multiplierColumnOption) elements.multiplierColumnOption.classList.remove('hidden');
                    if (elements.multiplierModeSelector) {
                        elements.multiplierModeSelector.classList.remove('hidden');
                    }
                } else {
                    if (elements.multiplierOptions) elements.multiplierOptions.classList.remove('hidden');
                    if (elements.multiplierColumnOption) elements.multiplierColumnOption.classList.add('hidden');
                    if (elements.multiplierModeSelector) {
                        elements.multiplierModeSelector.classList.add('hidden');
                    }
                }
                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
            });
            
            // 柱状图聚合切换
            if (elements.barChartAggregationDailyBtn && elements.barChartAggregationMonthlyBtn) {
                const applyAggregationActive = () => {
                    const isMonthly = appData?.visualization?.barChartAggregation === 'monthly';
                    
                    // 日按钮样式
                    elements.barChartAggregationDailyBtn.classList.toggle('bg-white', !isMonthly);
                    elements.barChartAggregationDailyBtn.classList.toggle('text-indigo-600', !isMonthly);
                    elements.barChartAggregationDailyBtn.classList.toggle('shadow-sm', !isMonthly);
                    elements.barChartAggregationDailyBtn.classList.toggle('text-slate-500', isMonthly);
                    elements.barChartAggregationDailyBtn.classList.toggle('hover:text-slate-700', isMonthly);
                    
                    // 月按钮样式
                    elements.barChartAggregationMonthlyBtn.classList.toggle('bg-white', isMonthly);
                    elements.barChartAggregationMonthlyBtn.classList.toggle('text-indigo-600', isMonthly);
                    elements.barChartAggregationMonthlyBtn.classList.toggle('shadow-sm', isMonthly);
                    elements.barChartAggregationMonthlyBtn.classList.toggle('text-slate-500', !isMonthly);
                    elements.barChartAggregationMonthlyBtn.classList.toggle('hover:text-slate-700', !isMonthly);
                };
                applyAggregationActive();
                
                elements.barChartAggregationDailyBtn.addEventListener('click', () => {
                    if (appData.visualization.barChartAggregation !== 'daily') {
                        appData.visualization.barChartAggregation = 'daily';
                        applyAggregationActive();
                        updateDailyTotalBarChart(getFilteredData());
                        showNotification('显示模式', '柱状图已切换为按日聚合', 'success');
                    }
                });
                elements.barChartAggregationMonthlyBtn.addEventListener('click', () => {
                    if (appData.visualization.barChartAggregation !== 'monthly') {
                        appData.visualization.barChartAggregation = 'monthly';
                        applyAggregationActive();
                        updateDailyTotalBarChart(getFilteredData());
                        showNotification('显示模式', '柱状图已切换为按月聚合', 'success');
                    }
                });
            }

            elements.multiplierColumnSelect.addEventListener('change', (e) => {
                appData.config.multiplierColumn = e.target.value;
                updatePtCtPreview();
                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
            });
            
            // PT/CT 列选择事件
            if (elements.ptColumnSelect) {
                elements.ptColumnSelect.addEventListener('change', (e) => {
                    appData.config.ptColumn = e.target.value;
                    updatePtCtPreview();
                    if (appData.originalData) {
                        reprocessDataWithConfig();
                    } else {
                        updateDataPreview();
                    }
                });
            }
            
            if (elements.ctColumnSelect) {
                elements.ctColumnSelect.addEventListener('change', (e) => {
                    appData.config.ctColumn = e.target.value;
                    updatePtCtPreview();
                    if (appData.originalData) {
                        reprocessDataWithConfig();
                    } else {
                        updateDataPreview();
                    }
                });
            }
            
            elements.dateColumnSelect.addEventListener('change', (e) => {
                appData.config.dateColumn = e.target.value;
                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
                checkConfigValidity();
            });
            
            elements.dataStartColumnSelect.addEventListener('change', (e) => {
                const val = e.target.value;
                appData.config.dataStartColumn = val;
                
                // 如果是列对列 (长表) 模式，自动同步结束列
                const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value;
                if (dataStructure === 'columnToColumn') {
                    if (elements.dataEndColumnSelect) {
                        elements.dataEndColumnSelect.value = val;
                        appData.config.dataEndColumn = val;
                    }
                }

                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
                checkConfigValidity();
            });
            
            elements.dataEndColumnSelect.addEventListener('change', (e) => {
                appData.config.dataEndColumn = e.target.value;
                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
                checkConfigValidity();
            });
            
            elements.meteringPointColumnSelect.addEventListener('change', (e) => {
                appData.config.meteringPointColumn = e.target.value;
                if (elements.meteringPointColumnSelect_v2) elements.meteringPointColumnSelect_v2.value = e.target.value;
                updateMeteringPointFilterOptions();
                
                // 如果清空了计量点列，则同时也重置筛选值
                if (!e.target.value) {
                    appData.config.meteringPointFilter = '';
                    if (elements.meteringPointFilterSelect) elements.meteringPointFilterSelect.value = '';
                    if (elements.meteringPointFilterSelect_v2) elements.meteringPointFilterSelect_v2.value = '';
                    appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                }

                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }
            });

            if (elements.meteringPointColumnSelect_v2) {
                elements.meteringPointColumnSelect_v2.addEventListener('change', (e) => {
                    appData.config.meteringPointColumn = e.target.value;
                    elements.meteringPointColumnSelect.value = e.target.value;
                    updateMeteringPointFilterOptions();
                    
                    if (!e.target.value) {
                        appData.config.meteringPointFilter = '';
                        if (elements.meteringPointFilterSelect) elements.meteringPointFilterSelect.value = '';
                        if (elements.meteringPointFilterSelect_v2) elements.meteringPointFilterSelect_v2.value = '';
                        appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                    }

                    if (appData.originalData) {
                        reprocessDataWithConfig();
                    } else {
                        updateDataPreview();
                    }
                });
            }
            
            elements.meteringPointFilterSelect.addEventListener('change', (e) => {
                const val = e.target.value;
                appData.config.meteringPointFilter = val;
                
                // 同步 v2
                if (elements.meteringPointFilterSelect_v2) {
                    elements.meteringPointFilterSelect_v2.value = val;
                }

                if (appData.config.meteringPointFilter) {
                    appData.visualization.selectedMeteringPoints = [appData.config.meteringPointFilter];
                } else {
                    appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                }

                if (appData.originalData) {
                    reprocessDataWithConfig();
                } else {
                    updateDataPreview();
                }

                setTimeout(() => {
                    if (appData.chart && appData.processedData.length > 0) {
                        updateChart();
                        // 计量点变更时，热力图需要同步刷新
                        updateHeatmap();
                    }
                }, 0);
            });

                // 计量点筛选 v2 同步逻辑
                if (elements.meteringPointFilterSelect_v2) {
                    elements.meteringPointFilterSelect_v2.addEventListener('change', (e) => {
                        const val = e.target.value;
                        
                        // 同步 v1
                        if (elements.meteringPointFilterSelect) {
                            elements.meteringPointFilterSelect.value = val;
                            // 触发 v1 的 change 事件以更新数据
                            elements.meteringPointFilterSelect.dispatchEvent(new Event('change'));
                        }
                    });

                    // 绑定 v2 的清除按钮事件
                    if (elements.clearMeteringPointFilterBtn_v2) {
                        elements.clearMeteringPointFilterBtn_v2.addEventListener('click', () => {
                            elements.meteringPointFilterSelect_v2.value = '';
                            elements.meteringPointFilterSelect_v2.dispatchEvent(new Event('change'));
                            showNotification('已清除', '计量点编号筛选已清除', 'success');
                        });
                    }
                    
                    if (elements.clearMeteringPointColumnBtn_v2) {
                        elements.clearMeteringPointColumnBtn_v2.addEventListener('click', () => {
                            elements.meteringPointColumnSelect_v2.value = '';
                            elements.meteringPointColumnSelect_v2.dispatchEvent(new Event('change'));
                            showNotification('已清除', '计量点编号所在列选择已清除', 'success');
                        });
                    }
                }

                // 绑定 v1 的清除按钮事件
                const clearMeteringPointColumn = document.getElementById('clearMeteringPointColumn');
                if (clearMeteringPointColumn) {
                    clearMeteringPointColumn.addEventListener('click', () => {
                        elements.meteringPointColumnSelect.value = '';
                        elements.meteringPointColumnSelect.dispatchEvent(new Event('change'));
                        showNotification('已清除', '计量点编号所在列选择已清除', 'success');
                    });
                }

                const clearMeteringPointFilter = document.getElementById('clearMeteringPointFilter');
                if (clearMeteringPointFilter) {
                    clearMeteringPointFilter.addEventListener('click', () => {
                        elements.meteringPointFilterSelect.value = '';
                        elements.meteringPointFilterSelect.dispatchEvent(new Event('change'));
                        showNotification('已清除', '计量点编号筛选已清除', 'success');
                    });
                }
            

            
            elements.invalidDataHandlingSelect.addEventListener('change', (e) => {
                appData.config.invalidDataHandling = e.target.value;
                if (appData.originalData) {
                    reprocessDataWithConfig();
                }
            });
            
            elements.timeIntervalSelect.addEventListener('change', (e) => {
        appData.config.timeInterval = parseInt(e.target.value);
        if (appData.originalData) {
            reprocessDataWithConfig();
        } else {
            updateDataPreview();
        }
        checkConfigValidity();
    });
    
    // 可视化部分
            elements.resetTimeFocusBtn.addEventListener('click', resetTimeFocus);
            if (elements.openAdvancedFiltersBtn) {
                elements.openAdvancedFiltersBtn.addEventListener('click', openAdvancedFiltersModal);
            }
            if (elements.openAdvancedFiltersBtn_v2) {
                elements.openAdvancedFiltersBtn_v2.addEventListener('click', openAdvancedFiltersModal);
            }

            // 时段聚焦自动应用
            elements.focusStartTimeSelect.addEventListener('change', applyTimeFocus);
            elements.focusEndTimeSelect.addEventListener('change', applyTimeFocus);
            if (elements.startDateInput && elements.endDateInput) {
                elements.startDateInput.addEventListener('change', (e) => {
                    if (elements.endDateInput.value && elements.endDateInput.value < e.target.value) {
                        elements.endDateInput.value = e.target.value;
                    }
                    applyVisualizationDateRange();
                    updateChart();
                    updateDailyTotalBarChart(getFilteredData());
                });
                elements.endDateInput.addEventListener('change', (e) => {
                    if (elements.startDateInput.value && elements.startDateInput.value > e.target.value) {
                        elements.startDateInput.value = e.target.value;
                    }
                    applyVisualizationDateRange();
                    updateChart();
                    updateDailyTotalBarChart(getFilteredData());
                });
            }
            elements.lineStyleSelect.addEventListener('change', (e) => {
                const allowed = new Set(['solid', 'dashed', 'dotted']);
                appData.visualization.lineStyle = allowed.has(e.target.value) ? e.target.value : 'solid';
                if (appData.visualization.lineStyle !== e.target.value) {
                    e.target.value = appData.visualization.lineStyle;
                }
                // 添加视觉反馈
                e.target.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.style.transform = 'scale(1)';
                }, 200);
                updateChart();
                showNotification('样式更新', `线条样式已更改为${e.target.options[e.target.selectedIndex].text}`, 'success');
            });
            
            elements.showPointsSelect.addEventListener('change', (e) => {
                appData.visualization.showPoints = e.target.value === 'true';
                // 添加视觉反馈
                e.target.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.style.transform = 'scale(1)';
                }, 200);
                updateChart();
                showNotification('样式更新', `数据点显示已${e.target.value === 'true' ? '开启' : '关闭'}`, 'success');
            });
            
            elements.smoothCurveSelect.addEventListener('change', (e) => {
                appData.visualization.smoothCurve = e.target.value === 'true';
                if (elements.curveTensionSelect) {
                    elements.curveTensionSelect.disabled = !appData.visualization.smoothCurve;
                    elements.curveTensionSelect.classList.toggle('opacity-50', !appData.visualization.smoothCurve);
                }
                // 添加视觉反馈
                e.target.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.style.transform = 'scale(1)';
                }, 200);
                updateChart();
                showNotification('样式更新', `曲线平滑已${e.target.value === 'true' ? '开启' : '关闭'}`, 'success');
            });
            
            elements.curveTensionSelect.addEventListener('change', (e) => {
                appData.visualization.curveTension = parseFloat(e.target.value);
                // 添加视觉反馈
                e.target.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.style.transform = 'scale(1)';
                }, 200);
                updateChart();
                showNotification('样式更新', `曲线平滑度已设置为${e.target.options[e.target.selectedIndex].text}`, 'success');
            });
            
            elements.secondaryAxisCheckbox.addEventListener('change', (e) => {
                appData.visualization.secondaryAxis = e.target.checked;
                // 添加视觉反馈
                e.target.parentElement.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.parentElement.style.transform = 'scale(1)';
                }, 200);
                updateChart();
                showNotification('样式更新', `辅助Y轴已${e.target.checked ? '开启' : '关闭'}`, 'success');
            });
            
            // 曲线图图例显示开关事件
            document.getElementById('toggleCurveLegend').addEventListener('change', (e) => {
                const isVisible = e.target.checked;
                const legendContainer = document.getElementById('curveLegendContainer');
                
                // 更新图例容器显示状态
                if (legendContainer) {
                    legendContainer.style.display = isVisible ? 'block' : 'none';
                }
                
                // 更新图表图例显示状态
                if (appData.chart) {
                    appData.chart.options.plugins.legend.display = isVisible;
                    appData.chart.update();
                }
                
                // 添加视觉反馈
                e.target.parentElement.style.transform = 'scale(1.05)';
                setTimeout(() => {
                    e.target.parentElement.style.transform = 'scale(1)';
                }, 200);
                
                showNotification('样式更新', `曲线图例已${isVisible ? '开启' : '关闭'}`, 'success');
            });
            
            // 新增：汇总模式切换按钮
            const summaryBtn = document.getElementById('toggleSummaryMode');
            const summarySubModeGroup = document.getElementById('summarySubModeGroup');
            
            if (summaryBtn) {
                summaryBtn.addEventListener('click', () => {
                    appData.visualization.summaryMode = !appData.visualization.summaryMode;
                    // 如果开启汇总模式，关闭全部曲线模式
                    if (appData.visualization.summaryMode) {
                        appData.visualization.allCurvesMode = false;
                        const allCurvesBtn = document.getElementById('toggleAllCurves');
                        if (allCurvesBtn) {
                            allCurvesBtn.querySelector('span').textContent = '显示全部日期';
                            allCurvesBtn.classList.remove('bg-indigo-600', 'text-white', 'border-indigo-600');
                            allCurvesBtn.classList.add('bg-white', 'text-slate-600', 'border-slate-200');
                        }
                        // 显示子模式切换组并同步当前值
                        if (summarySubModeGroup) {
                            summarySubModeGroup.classList.remove('hidden');
                            summarySubModeGroup.value = appData.visualization.summarySubMode || 'extreme';
                        }
                    } else {
                        // 隐藏子模式切换组
                        if (summarySubModeGroup) summarySubModeGroup.classList.add('hidden');
                    }
                    // 更新按钮文案/样式
                    {
                        const span = summaryBtn.querySelector('span');
                        if (span) span.textContent = appData.visualization.summaryMode ? '返回明细' : '汇总曲线';
                    }
                    summaryBtn.classList.toggle('bg-indigo-600', !appData.visualization.summaryMode);
                    summaryBtn.classList.toggle('bg-gray-600', appData.visualization.summaryMode);
                    // 重新应用日期筛选逻辑
                    const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                    if (appData.visualization.allCurvesMode || appData.visualization.summaryMode) {
                        elements.startDateInput.value = '';
                        elements.endDateInput.value = '';
                        if (appData.visualization.summaryMode) {
                            appData.visualization.selectedDates = allDates;
                        }
                    } else {
                        if (allDates.length > 0) {
                            elements.startDateInput.value = allDates[0];
                            elements.endDateInput.value = allDates[0];
                        }
                    }
                    
                    applyVisualizationDateRange();
                    // 刷新图表
                    updateChart();
                    // 汇总模式下隐藏分页控件
                    if (appData.visualization.summaryMode) {
                        hidePaginationControls();
                    }
                });
            }

            // 修改：汇总子模式切换下拉菜单事件
            const summarySubModeSelect = document.getElementById('summarySubModeGroup');
            if (summarySubModeSelect) {
                summarySubModeSelect.addEventListener('change', (e) => {
                    const mode = e.target.value;
                    appData.visualization.summarySubMode = mode;
                    
                    // 刷新图表
                    updateChart();
                });
            }

            // 新增：显示全部日期切换按钮
            const allCurvesBtn = document.getElementById('toggleAllCurves');
            if (allCurvesBtn) {
                allCurvesBtn.addEventListener('click', () => {
                    appData.visualization.allCurvesMode = !appData.visualization.allCurvesMode;
                    // 如果开启全部曲线模式，关闭汇总模式
                    if (appData.visualization.allCurvesMode) {
                        appData.visualization.summaryMode = false;
                        const summaryBtn = document.getElementById('toggleSummaryMode');
                        const summarySubModeGroup = document.getElementById('summarySubModeGroup');
                        if (summarySubModeGroup) summarySubModeGroup.classList.add('hidden');
                        if (summaryBtn) {
                            {
                                const span = summaryBtn.querySelector('span');
                                if (span) span.textContent = '汇总曲线';
                            }
                            summaryBtn.classList.remove('bg-gray-600');
                            summaryBtn.classList.add('bg-indigo-600');
                        }
                    }
                    // 更新按钮文案/样式
                    allCurvesBtn.querySelector('span').textContent = appData.visualization.allCurvesMode ? '显示单个日期' : '显示全部日期';
                    if (appData.visualization.allCurvesMode) {
                        allCurvesBtn.classList.add('bg-indigo-600', 'text-white', 'border-indigo-600');
                        allCurvesBtn.classList.remove('bg-white', 'text-slate-600', 'border-slate-200');
                    } else {
                        allCurvesBtn.classList.remove('bg-indigo-600', 'text-white', 'border-indigo-600');
                        allCurvesBtn.classList.add('bg-white', 'text-slate-600', 'border-slate-200');
                    }
                    
                    // 重新应用日期筛选逻辑
                    if (appData.visualization.allCurvesMode) {
                        // 如果切换到“全部日期”模式，清空日期选择器的值，强制显示全部
                        elements.startDateInput.value = '';
                        elements.endDateInput.value = '';
                    } else {
                        // 如果切换到“单日期”模式，默认恢复到第一个日期
                        const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                        if (allDates.length > 0) {
                            elements.startDateInput.value = allDates[0];
                            elements.endDateInput.value = allDates[0];
                        }
                    }

                    applyVisualizationDateRange();
                    // 刷新图表
                    updateChart();
                });
            }

            // 新增：前一天/后一天按钮
            if (elements.prevDayBtn) {
                elements.prevDayBtn.addEventListener('click', () => {
                    const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                    if (allDates.length === 0) return;
                    
                    let currentIndex = allDates.indexOf(elements.startDateInput.value);
                    if (currentIndex === -1) currentIndex = 0;
                    
                    const nextIndex = Math.max(0, currentIndex - 1);
                    elements.startDateInput.value = allDates[nextIndex];
                    elements.endDateInput.value = allDates[nextIndex];
                    
                    // 切换回单日期模式
                    if (appData.visualization.allCurvesMode) {
                        appData.visualization.allCurvesMode = false;
                        if (allCurvesBtn) {
                            allCurvesBtn.querySelector('span').textContent = '显示全部日期';
                            allCurvesBtn.classList.remove('bg-indigo-600', 'text-white', 'border-indigo-600');
                            allCurvesBtn.classList.add('bg-white', 'text-slate-600', 'border-slate-200');
                        }
                    }

                    applyVisualizationDateRange();
                    updateChart();
                });
            }

            if (elements.nextDayBtn) {
                elements.nextDayBtn.addEventListener('click', () => {
                    const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                    if (allDates.length === 0) return;

                    let currentIndex = allDates.indexOf(elements.startDateInput.value);
                    if (currentIndex === -1) currentIndex = 0;

                    const nextIndex = Math.min(allDates.length - 1, currentIndex + 1);
                    elements.startDateInput.value = allDates[nextIndex];
                    elements.endDateInput.value = allDates[nextIndex];

                    // 切换回单日期模式
                    if (appData.visualization.allCurvesMode) {
                        appData.visualization.allCurvesMode = false;
                        if (allCurvesBtn) {
                            allCurvesBtn.querySelector('span').textContent = '显示全部日期';
                            allCurvesBtn.classList.remove('bg-indigo-600', 'text-white', 'border-indigo-600');
                            allCurvesBtn.classList.add('bg-white', 'text-slate-600', 'border-slate-200');
                        }
                    }
                    
                    applyVisualizationDateRange();
                    updateChart();
                });
            }

            // 柱状图始终显示，无需事件处理

            

            
            // 导出部分（统一走“导出预览”弹窗确认）
            if (elements.exportPNGBtn) elements.exportPNGBtn.addEventListener('click', () => requestExportAction('chart_png'));
            if (elements.exportJPGBtn) elements.exportJPGBtn.addEventListener('click', () => requestExportAction('chart_jpg'));
            if (elements.exportSVGBtn) elements.exportSVGBtn.addEventListener('click', () => requestExportAction('chart_svg'));
            // PDF报告直接生成，不经过预览模态框
            if (elements.exportPDFBtn) elements.exportPDFBtn.addEventListener('click', () => {
                const fileName = getExportFileName('专业报告', 'pdf');
                exportPDFReport({ fileName });
            });

            // 图表选择下拉菜单事件
            if (elements.exportChartSelect) {
                elements.exportChartSelect.addEventListener('change', updateChartExportPreview);
            }

            // 报告预览模态框事件
            const closeReportPreviewBtn = document.getElementById('closeReportPreviewBtn');
            const reportPrevBtn = document.getElementById('reportPrevBtn');
            const reportNextBtn = document.getElementById('reportNextBtn');
            const reportDownloadBtn = document.getElementById('reportDownloadBtn');

            if (closeReportPreviewBtn) {
                closeReportPreviewBtn.addEventListener('click', closeReportPreview);
            }
            if (reportPrevBtn) {
                reportPrevBtn.addEventListener('click', () => navigateReportPage(-1));
            }
            if (reportNextBtn) {
                reportNextBtn.addEventListener('click', () => navigateReportPage(1));
            }
            if (reportDownloadBtn) {
                reportDownloadBtn.addEventListener('click', () => {
                    closeReportPreview();
                    requestExportAction('report_pdf');
                });
            }

            // 报告点导航
            document.querySelectorAll('.report-dot').forEach(dot => {
                dot.addEventListener('click', () => {
                    const page = parseInt(dot.dataset.page);
                    goToReportPage(page);
                });
            });

            // 点击模态框背景关闭
            const reportPreviewPanel = document.getElementById('reportPreviewPanel');
            if (reportPreviewPanel) {
                reportPreviewPanel.addEventListener('click', (e) => {
                    if (e.target === reportPreviewPanel) {
                        closeReportPreview();
                    }
                });
            }

            if (elements.export15MinDataBtn) elements.export15MinDataBtn.addEventListener('click', () => requestExportAction('data_15min'));
            if (elements.exportHourlyDataBtn) elements.exportHourlyDataBtn.addEventListener('click', () => requestExportAction('data_24h'));
            if (elements.exportDailyStatsBtn) elements.exportDailyStatsBtn.addEventListener('click', () => requestExportAction('data_daily_stats'));
            if (elements.exportSummaryCurveDataBtn) elements.exportSummaryCurveDataBtn.addEventListener('click', () => requestExportAction('data_summary_curve'));
            const pvBtn = document.getElementById('exportPVsystData');
            if (pvBtn) pvBtn.addEventListener('click', () => requestExportAction('data_pvsyst'));
            if (elements.exportFullPackageBtn) elements.exportFullPackageBtn.addEventListener('click', () => requestExportAction('package_full'));
            if (elements.exportZipBundleBtn) elements.exportZipBundleBtn.addEventListener('click', () => requestExportAction('bundle_zip'));

            if (elements.exportPreviewCloseBtn) elements.exportPreviewCloseBtn.addEventListener('click', closeExportPreview);
            if (elements.exportPreviewCancelBtn) elements.exportPreviewCancelBtn.addEventListener('click', closeExportPreview);
            if (elements.exportPreviewConfirmBtn) elements.exportPreviewConfirmBtn.addEventListener('click', confirmExportPreview);
            if (elements.exportPreviewModal) {
                elements.exportPreviewModal.addEventListener('click', (e) => {
                    if (e.target === elements.exportPreviewModal) closeExportPreview();
                });
            }
            
            // 导出日期范围选择器事件处理
            document.getElementById('exportAllDates').addEventListener('change', function() {
                const startDate = document.getElementById('exportStartDate');
                const endDate = document.getElementById('exportEndDate');
                
                if (this.checked) {
                    startDate.disabled = true;
                    endDate.disabled = true;
                } else {
                    startDate.disabled = false;
                    endDate.disabled = false;
                }
            });
            
            // 初始化日期范围选择器状态
            const exportAllDatesBtn = document.getElementById('exportAllDates');
            if (exportAllDatesBtn) {
                exportAllDatesBtn.dispatchEvent(new Event('change'));
            }
            
            // 联动：开始日期变更时，确保结束日期不早于开始日期
            document.getElementById('exportStartDate').addEventListener('change', function() {
                const startDate = this.value;
                const endDateSelect = document.getElementById('exportEndDate');
                if (startDate && endDateSelect.value && endDateSelect.value < startDate) {
                    endDateSelect.value = startDate;
                }
            });
            
            // 联动：结束日期变更时，确保不早于开始日期
            document.getElementById('exportEndDate').addEventListener('change', function() {
                const endDate = this.value;
                const startDateSelect = document.getElementById('exportStartDate');
                if (endDate && startDateSelect.value && startDateSelect.value > endDate) {
                    startDateSelect.value = endDate;
                }
            });

            if (elements.startOverBtn) {
                elements.startOverBtn.addEventListener('click', startOver);
            }
            
            // 通知和模态框
            if (elements.closeNotificationBtn) {
                elements.closeNotificationBtn.addEventListener('click', hideNotification);
            }
            if (elements.closeModalBtn) {
                elements.closeModalBtn.addEventListener('click', closeModal);
            }
            if (elements.modalOverlay) {
                elements.modalOverlay.addEventListener('click', (e) => {
                    if (e.target === elements.modalOverlay) closeModal();
                });
            }
            

            if (elements.aboutBtn) {
                elements.aboutBtn.addEventListener('click', showAboutModal);
            }
            
            // 第一个小时计算过程预览
            if (elements.previewFirstHourCalculationBtn) {
            elements.previewFirstHourCalculationBtn.addEventListener('click', previewFirstHourCalculation);
        }
        if (elements.closeFirstHourCalculationModalBtn) {
            elements.closeFirstHourCalculationModalBtn.addEventListener('click', closeFirstHourCalculationModal);
        }
        
        // 第一天计算过程预览
        if (elements.previewFirstDayCalculationBtn) {
            elements.previewFirstDayCalculationBtn.addEventListener('click', previewFirstDayCalculation);
        }
        if (elements.closeFirstDayCalculationModalBtn) {
            elements.closeFirstDayCalculationModalBtn.addEventListener('click', closeFirstDayCalculationModal);
        }

        // 曲线样式弹窗
        if (elements.openCurveStyleBtn) {
            elements.openCurveStyleBtn.addEventListener('click', openCurveStyleModal);
        }
        if (elements.closeCurveStyleModalBtn) {
            elements.closeCurveStyleModalBtn.addEventListener('click', closeCurveStyleModal);
        }
        if (elements.applyCurveStyleBtn) {
            elements.applyCurveStyleBtn.addEventListener('click', closeCurveStyleModal);
        }
        if (elements.curveStyleModal) {
            elements.curveStyleModal.addEventListener('click', (e) => {
                if (e.target === elements.curveStyleModal) closeCurveStyleModal();
            });
        }
        
        // 统一更新数据结构相关的UI显示状态
        function updateDataStructureUI() {
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
            const endColSelect = document.getElementById('dataEndColumn');
            const timeIntervalLabel = document.querySelector('label[for="timeInterval"]');
            const crossDateContainer = document.getElementById('crossDateStructureContainer');
            
            // 根据数据结构类型显示/隐藏相关选项
            if (dataStructure === 'columnToRow') {
                // --- 列对行 (宽表) ---
                // 1. 显示数据结束列选项
                if (endColSelect && endColSelect.parentElement) {
                    endColSelect.parentElement.classList.remove('hidden');
                    // 同时显示箭头
                    const arrow = endColSelect.parentElement.previousElementSibling;
                    if (arrow) arrow.classList.remove('hidden');
                }
                
                // 2. 更新时间间隔提示
                if (timeIntervalLabel) {
                    timeIntervalLabel.textContent = '时间间隔 (分钟，宽表可自动检测)';
                }
                
                // 3. 隐藏跨日期结构选项
                if (crossDateContainer) {
                    crossDateContainer.classList.add('hidden');
                }
            } else {
                // --- 列对列 (长表) ---
                // 1. 隐藏数据结束列选项 (长表起始列即结束列)
                if (endColSelect && endColSelect.parentElement) {
                    endColSelect.parentElement.classList.add('hidden');
                    // 同时隐藏箭头
                    const arrow = endColSelect.parentElement.previousElementSibling;
                    if (arrow) arrow.classList.add('hidden');
                    
                    // 强制设置结束列等于起始列
                    const startColSelect = document.getElementById('dataStartColumn');
                    if (startColSelect && endColSelect) {
                        endColSelect.value = startColSelect.value;
                        appData.config.dataEndColumn = startColSelect.value;
                    }
                }
                
                // 2. 更新时间间隔提示
                if (timeIntervalLabel) {
                    timeIntervalLabel.textContent = '时间间隔 (分钟)';
                }
                
                // 3. 显示跨日期结构选项
                if (crossDateContainer) {
                    crossDateContainer.classList.remove('hidden');
                }
            }
        }

        // 添加数据结构类型变更事件处理
        document.querySelectorAll('input[name="dataStructure"]').forEach(radio => {
            radio.addEventListener('change', function() {
                // 更新 UI
                updateDataStructureUI();
                
                // 更新数据预览
                updateDataPreview();
                
                // 更新按钮状态
                checkConfigValidity();
            });
        });
        
        // 初始化数据结构类型UI
        updateDataStructureUI();


        }

        // 文件处理辅助函数
        function estimateMemoryUsage(jsonData) {
            // 估算数据的内存使用量（MB）
            const rowCount = jsonData.length;
            const avgColumnsPerRow = jsonData[0] ? jsonData[0].length : 0;
            const avgBytesPerCell = 20; // 估算每个单元格平均字节数
            const estimatedBytes = rowCount * avgColumnsPerRow * avgBytesPerCell;
            return estimatedBytes / (1024 * 1024); // 转换为MB
        }
        
        function checkMemoryLimits(file, jsonData) {
            const rowCount = jsonData.length;
            const estimatedMB = estimateMemoryUsage(jsonData);
            
            // 检查单个文件行数限制
            if (rowCount > appData.memoryManagement.maxDataRows) {
                const proceed = confirm(
                    `文件 "${file.name}" 包含 ${rowCount} 行数据，超过建议的 ${appData.memoryManagement.maxDataRows} 行限制。\n\n` +
                    `大量数据可能导致浏览器响应缓慢或崩溃。\n\n是否仍要继续处理？`
                );
                if (!proceed) {
                    return false;
                }
            }
            
            // 检查总内存使用限制
            const newTotalMemory = appData.memoryManagement.currentMemoryUsageMB + estimatedMB;
            if (newTotalMemory > appData.memoryManagement.maxTotalMemoryMB) {
                const proceed = confirm(
                    `添加此文件将使总内存使用量达到约 ${newTotalMemory.toFixed(1)}MB，超过 ${appData.memoryManagement.maxTotalMemoryMB}MB 限制。\n\n` +
                    `建议先清理现有数据或减少数据量。\n\n是否仍要继续？`
                );
                if (!proceed) {
                    return false;
                }
            }
            
            // 更新内存使用量
            appData.memoryManagement.currentMemoryUsageMB += estimatedMB;
            return true;
        }
        
        function optimizeMemoryUsage() {
            if (appData.memoryManagement.enableGarbageCollection) {
                // 强制垃圾回收（如果浏览器支持）
                if (window.gc && typeof window.gc === 'function') {
                    try {
                        window.gc();
                    } catch (e) {
                        // 忽略错误，某些浏览器不支持手动GC
                    }
                }
                
                // 清理不必要的引用
                setTimeout(() => {
                    // 延迟执行，让浏览器有机会进行垃圾回收
                    if (appData.memoryManagement.currentMemoryUsageMB > appData.memoryManagement.maxTotalMemoryMB * 0.8) {
                        showNotification('提示', '内存使用量较高，建议清理部分数据以提升性能', 'warning');
                    }
                }, 1000);
            }
        }
        
        function completeFileProcessing(file, jsonData) {
            showNotification('成功', `文件 "${file.name}" 导入成功`, 'success');

            setImportStatus('done', '文件导入完成');

            // 导入完成后自动刷新数据预览（无需手动下拉选择）
            setTimeout(() => {
                try {
                    updateDataPreview();
                } catch (err) {
                    console.error('自动刷新数据预览失败:', err);
                }
            }, 0);
        }
        
        // 倍率模式切换函数
        function setMultiplierMode(mode) {
            appData.config.multiplierMode = mode;
            
            const singleBtn = document.getElementById('multiplierModeSingle');
            const ptCtBtn = document.getElementById('multiplierModePtCt');
            
            if (mode === 'single') {
                // 切换到单列模式
                singleBtn.classList.add('bg-white', 'text-indigo-600', 'shadow-sm');
                singleBtn.classList.remove('text-slate-500');
                ptCtBtn.classList.remove('bg-white', 'text-indigo-600', 'shadow-sm');
                ptCtBtn.classList.add('text-slate-500');
                
                if (elements.singleMultiplierSection) elements.singleMultiplierSection.classList.remove('hidden');
                if (elements.ptCtMultiplierSection) elements.ptCtMultiplierSection.classList.add('hidden');
            } else {
                // 切换到 PT/CT 模式
                ptCtBtn.classList.add('bg-white', 'text-indigo-600', 'shadow-sm');
                ptCtBtn.classList.remove('text-slate-500');
                singleBtn.classList.remove('bg-white', 'text-indigo-600', 'shadow-sm');
                singleBtn.classList.add('text-slate-500');
                
                if (elements.singleMultiplierSection) elements.singleMultiplierSection.classList.add('hidden');
                if (elements.ptCtMultiplierSection) elements.ptCtMultiplierSection.classList.remove('hidden');
            }
            
            updatePtCtPreview();
        }
        
        // 更新 PT/CT 预览
        function updatePtCtPreview() {
            if (!elements.ptCtPreview) return;
            
            const ptCol = parseInt(appData.config.ptColumn);
            const ctCol = parseInt(appData.config.ctColumn);
            
            if (!isNaN(ptCol) && !isNaN(ctCol) && appData.worksheets && appData.worksheets.length > 0) {
                // 获取第一行数据计算示例
                const data = appData.worksheets[0].data;
                if (data && data.length > 1) {
                    const firstRow = data[1];
                    const ptVal = parseFloat(firstRow[ptCol]) || 1;
                    const ctVal = parseFloat(firstRow[ctCol]) || 1;
                    const multiplier = ptVal * ctVal;
                    elements.ptCtPreview.textContent = `= ${multiplier}`;
                    elements.ptCtPreview.classList.remove('hidden');
                }
            } else {
                elements.ptCtPreview.textContent = '';
                elements.ptCtPreview.classList.add('hidden');
            }
        }
        
        function handleFileProcessingError(file, error) {
            console.error('Error processing file:', error);
            showNotification('错误', `处理文件 "${file.name}" 时出错: ${error.message}`, 'error');
            setImportStatus('error', '导入失败');
        }

        // 文件处理函数
        function handleFileSelection(e) {
            const files = Array.from(e.target.files);
            processFiles(files);
            e.target.value = '';
        }

        function handleDragOver(e) {
            e.preventDefault();
            if (elements.dropArea && elements.dropArea.parentNode) {
                elements.dropArea.classList.add('border-indigo-400');
                elements.dropArea.classList.add('bg-indigo-50/40');
                elements.dropArea.classList.add('shadow-inner-soft');
            }
        }

        function handleDragLeave() {
            if (elements.dropArea && elements.dropArea.parentNode) {
                elements.dropArea.classList.remove('border-indigo-400');
                elements.dropArea.classList.remove('bg-indigo-50/40');
                elements.dropArea.classList.remove('shadow-inner-soft');
            }
        }

        function handleDrop(e) {
            e.preventDefault();
            if (elements.dropArea && elements.dropArea.parentNode) {
                elements.dropArea.classList.remove('border-indigo-400');
                elements.dropArea.classList.remove('bg-indigo-50/40');
                elements.dropArea.classList.remove('shadow-inner-soft');
            }
            
            const files = Array.from(e.dataTransfer.files);
            processFiles(files);
        }

        if (elements.dropArea) {
            elements.dropArea.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter' && e.key !== ' ') return;
                e.preventDefault();
                if (elements.fileInput) elements.fileInput.click();
            });
        }

        function processFiles(files) {
            if (files.length === 0) return;

            setImportStatus('processing');

            let importedNewCount = 0;
            let failureCount = 0;
            let skippedCount = 0;
            
            // 在非追加模式下，导入新文件前先清空现有数据与配置
            const appendModeAtImport = document.getElementById('appendMode')?.checked || false;
            if (!appendModeAtImport) {
                removeAllFiles({ silent: true, keepStatus: true, force: true });
            }
            
            // 显示上传进度
            if (elements.uploadProgress && elements.uploadProgress.parentNode) {
                if (elements.uploadProgress && elements.uploadProgress.parentNode) {
                elements.uploadProgress.classList.remove('hidden');
            }
            }
            
            let processedCount = 0;
            const totalFiles = files.length;
            
            files.forEach((file, index) => {
                // 检查文件是否已导入
                const isDuplicate = appData.files.some(f => 
                    f.name === file.name && f.size === file.size && f.lastModified === file.lastModified
                );
                
                if (isDuplicate) {
                    showNotification('警告', `文件 "${file.name}" 已导入，跳过重复文件`, 'warning');
                    skippedCount++;
                    processedCount++;
                    updateUploadProgress(processedCount, totalFiles);
                    if (processedCount >= totalFiles) {
                        if (importedNewCount > 0) {
                            setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                        } else if (failureCount > 0) {
                            setImportStatus('error', '导入失败');
                        } else {
                            setImportStatus('idle', '未导入新文件');
                        }
                    }
                    return;
                }
                
                // 检查文件大小限制（50MB）
                const maxFileSize = 50 * 1024 * 1024; // 50MB in bytes
                if (file.size > maxFileSize) {
                    const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
                    showNotification('警告', `文件 "${file.name}" 过大 (${fileSizeMB}MB)，建议文件大小不超过50MB。大文件可能导致浏览器响应缓慢或崩溃。`, 'warning');
                    
                    // 更新计数器并继续
                    if (!confirm(`文件 "${file.name}" 大小为 ${fileSizeMB}MB，超过建议的50MB限制。\n\n继续处理可能导致浏览器响应缓慢或崩溃。\n\n是否仍要继续导入此文件？`)) {
                        skippedCount++;
                        processedCount++;
                        updateUploadProgress(processedCount, totalFiles);
                        if (processedCount >= totalFiles) {
                            if (importedNewCount > 0) {
                                setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                            } else if (failureCount > 0) {
                                setImportStatus('error', '导入失败');
                            } else {
                                setImportStatus('idle', '未导入新文件');
                            }
                        }
                        return;
                    }
                }
                
                if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
                    showNotification('错误', `文件 "${file.name}" 格式不支持，请上传Excel或CSV文件`, 'error');
                    
                    failureCount++;
                    processedCount++;
                    updateUploadProgress(processedCount, totalFiles);
                    if (processedCount >= totalFiles) {
                        if (importedNewCount > 0) {
                            setImportStatus('done', '导入完成（部分失败）');
                        } else {
                            setImportStatus('error', '导入失败');
                        }
                    }
                    return;
                }
                
                appData.files.push(file);
                addFileToUI(file);
                
                const reader = new FileReader();
                
                // 添加文件读取错误处理
                reader.onerror = function(e) {
                    console.error('FileReader error:', e);
                    const errorMessage = e.target.error ? e.target.error.message : '文件读取失败';
                    showNotification('错误', `读取文件 "${file.name}" 时出错: ${errorMessage}`, 'error');
                    
                    failureCount++;
                    processedCount++;
                    updateUploadProgress(processedCount, totalFiles);
                    if (processedCount >= totalFiles) {
                        if (importedNewCount > 0) {
                            setImportStatus('done', '导入完成（部分失败）');
                        } else {
                            setImportStatus('error', '导入失败');
                        }
                    }
                };
                
                reader.onload = function(e) {
                    try {
                        // 显示处理进度（对于大文件）
                        const fileSizeMB = file.size / (1024 * 1024);
                        if (fileSizeMB > 10) { // 对于超过10MB的文件显示处理提示
                            showNotification('信息', `正在处理大文件 "${file.name}" (${fileSizeMB.toFixed(2)}MB)，请稍候...`, 'info');
                        }
                        
                        let jsonData;
                        let firstSheetName;
                        
                        if (file.name.endsWith('.csv')) {
                            const buffer = e.target.result;
                            const bytes = new Uint8Array(buffer);
                            const isZipLike = bytes.length > 3 && bytes[0] === 0x50 && bytes[1] === 0x4B;
                            if (isZipLike) {
                                const readOptions = fileSizeMB > 20 ? 
                                    { type: 'array', cellDates: false, cellNF: false, cellStyles: false } : 
                                    { type: 'array' };
                                const workbook = XLSX.read(bytes, readOptions);
                                firstSheetName = workbook.SheetNames[0];
                                const worksheet = workbook.Sheets[firstSheetName];
                                repairWorksheetRef(worksheet);
                                jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                            } else {
                                const csvText = decodeCsvArrayBuffer(buffer);
                                const normalizedCsv = normalizeCsvContent(csvText);
                                let workbook = null;
                                try {
                                    workbook = XLSX.read(normalizedCsv, { type: 'string' });
                                } catch (err) {
                                    console.error('XLSX 解析 CSV 失败，尝试降级为手动解析:', err);
                                }
                                if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
                                    firstSheetName = workbook.SheetNames[0];
                                    const worksheet = workbook.Sheets[firstSheetName];
                                    repairWorksheetRef(worksheet);
                                    jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                                }
                                if (!jsonData || jsonData.length === 0) {
                                    jsonData = parseCsvText(normalizedCsv);
                                    firstSheetName = firstSheetName || 'CSV';
                                }
                            }
                        } else {
                            const data = new Uint8Array(e.target.result);
                            const readOptions = fileSizeMB > 20 ? 
                                { type: 'array', cellDates: false, cellNF: false, cellStyles: false } : 
                                { type: 'array' };
                            const workbook = XLSX.read(data, readOptions);
                            firstSheetName = workbook.SheetNames[0];
                            const worksheet = workbook.Sheets[firstSheetName];
                            repairWorksheetRef(worksheet);
                            jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        }

                        // 追加模式下的格式校验
                        const isAppendMode = document.getElementById('appendMode')?.checked || false;
                        console.log('Append Mode Status:', isAppendMode);
                        console.log('Existing Worksheets:', appData.worksheets.length);

                        if (isAppendMode && appData.worksheets.length > 0) {
                            const existingHeader = appData.worksheets[0].data[0];
                            const newHeader = jsonData[0];
                            console.log('Existing Header:', existingHeader);
                            console.log('New Header:', newHeader);
                            
                            let formatMatch = true;
                            let diffReason = '';

                            if (!existingHeader || !newHeader) {
                                formatMatch = false;
                                diffReason = '无法识别表头';
                            } else {
                                // 过滤掉空列，只比较有内容的列
                                const cleanExistingHeader = existingHeader.filter(h => h !== undefined && h !== null && String(h).trim() !== '');
                                const cleanNewHeader = newHeader.filter(h => h !== undefined && h !== null && String(h).trim() !== '');

                                if (cleanExistingHeader.length !== cleanNewHeader.length) {
                                    formatMatch = false;
                                    diffReason = `列数不一致 (当前有效列: ${cleanNewHeader.length}, 已有有效列: ${cleanExistingHeader.length})`;
                                } else {
                                    for (let i = 0; i < cleanExistingHeader.length; i++) {
                                        if (String(cleanExistingHeader[i]).trim() !== String(cleanNewHeader[i]).trim()) {
                                            formatMatch = false;
                                            diffReason = `第 ${i+1} 列标题不一致 ("${cleanNewHeader[i]}" vs "${cleanExistingHeader[i]}")`;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (!formatMatch) {
                                console.warn('Format mismatch detected:', diffReason);
                                if (!confirm(`【追加模式提醒】\n\n当前文件 "${file.name}" 的格式与已导入文件不匹配：\n${diffReason}\n\n追加模式建议使用相同格式的文件以保证数据处理正确。\n\n是否仍要继续导入？`)) {
                                    skippedCount++;
                                    processedCount++;
                                    updateUploadProgress(processedCount, totalFiles);
                                    if (processedCount >= totalFiles) {
                                        setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                                        updateStepStatus(1, true);
                                    }
                                    return;
                                }
                            } else {
                                console.log('Format validation passed.');
                            }
                        }
                        
                        // 检查内存限制
                        if (!checkMemoryLimits(file, jsonData)) {
                            // 用户取消处理大数据文件
                            skippedCount++;
                            processedCount++;
                            updateUploadProgress(processedCount, totalFiles);
                            if (processedCount >= totalFiles) {
                                if (importedNewCount > 0) {
                                    setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                                } else if (failureCount > 0) {
                                    setImportStatus('error', '导入失败');
                                } else {
                                    setImportStatus('idle', '未导入新文件');
                                }
                            }
                            return;
                        }
                        
                        // 对于大数据集，分批处理以避免UI阻塞
                if (jsonData.length > 10000) {
                    showNotification('信息', `文件包含 ${jsonData.length} 行数据，正在优化处理...`, 'info');
                            
                            // 使用setTimeout分批处理，避免阻塞UI
                            setTimeout(() => {
                                try {
                                    // 填充空列，确保列索引从0开始连续
                                    const filledJsonData = fillEmptyColumns(jsonData);
                                    appData.worksheets.push({
                                        file: file,
                                        sheetName: firstSheetName,
                                        data: filledJsonData
                                    });
                                    
                                    // 更新统计信息
                                    updateDataStats();
                                    
                                    // 非追加模式或首个文件时，自动填充列选择并识别数据结构
                                    const appendModeDetect1 = document.getElementById('appendMode')?.checked || false;
                                    if (!appendModeDetect1 || appData.worksheets.length === 1) {
                                        populateColumnSelects(jsonData[0] || [], jsonData);
                                        autoDetectDataStructure(jsonData);
                                        // 检测日期重复情况并提醒用户
                                        detectDuplicateDatesAndNotify(jsonData);
                                        // 自动识别完成后，尝试刷新预览（若配置已充分）
                                        setTimeout(() => { try { updateDataPreview(); } catch(e){} }, 0);
                                    }
                                    
                                    // 构建原始数据预览
                                    buildRawDataPreview();
                                    
                                    completeFileProcessing(file, jsonData);
                                    importedNewCount++;

                                    processedCount++;
                                    updateUploadProgress(processedCount, totalFiles);
                                    if (processedCount >= totalFiles) {
                                        setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                                        updateStepStatus(1, true);
                                    }
                                    
                                    // 优化内存使用
                                    optimizeMemoryUsage();
                                } catch (error) {
                                    handleFileProcessingError(file, error);
                                    failureCount++;
                                    processedCount++;
                                    updateUploadProgress(processedCount, totalFiles);
                                    if (processedCount >= totalFiles) {
                                        if (importedNewCount > 0) {
                                            setImportStatus('done', '导入完成（部分失败）');
                                        } else {
                                            setImportStatus('error', '导入失败');
                                        }
                                    }
                                }
                            }, 100);
                        } else {
                            // 填充空列，确保列索引从0开始连续
                            const filledJsonData = fillEmptyColumns(jsonData);
                            appData.worksheets.push({
                                file: file,
                                sheetName: firstSheetName,
                                data: filledJsonData
                            });
                            
                            // 更新统计信息
                            updateDataStats();
                            
                            // 非追加模式或首个文件时，自动填充列选择并识别数据结构（与大数据分支一致）
                            const appendModeDetect2 = document.getElementById('appendMode')?.checked || false;
                            if (!appendModeDetect2 || appData.worksheets.length === 1) {
                                populateColumnSelects(jsonData[0] || [], jsonData);
                                autoDetectDataStructure(jsonData);
                                // 检测日期重复情况并提醒用户
                                detectDuplicateDatesAndNotify(jsonData);
                                // 自动识别完成后，尝试刷新预览（若配置已充分）
                                setTimeout(() => { try { updateDataPreview(); } catch(e){} }, 0);
                            }
                            
                            // 构建原始数据预览
                            buildRawDataPreview();
                            
                            completeFileProcessing(file, jsonData);
                            importedNewCount++;

                            processedCount++;
                            updateUploadProgress(processedCount, totalFiles);
                            if (processedCount >= totalFiles) {
                                setImportStatus('done', failureCount > 0 ? '导入完成（部分失败）' : '文件导入完成');
                                updateStepStatus(1, true);
                            }
                            
                            // 优化内存使用
                            optimizeMemoryUsage();
                        }
                        
                    } catch (error) {
                        handleFileProcessingError(file, error);
                        failureCount++;
                        processedCount++;
                        updateUploadProgress(processedCount, totalFiles);
                        if (processedCount >= totalFiles) {
                            if (importedNewCount > 0) {
                                setImportStatus('done', '导入完成（部分失败）');
                            } else {
                                setImportStatus('error', '导入失败');
                            }
                        }
                    }
                };
                
                if (file.name.endsWith('.csv')) {
                    reader.readAsArrayBuffer(file);
                } else {
                    reader.readAsArrayBuffer(file);
                }
            });

            // 文件列表现在固定展开，无需控制显示/隐藏
            if (elements.removeAllFiles && elements.removeAllFiles.parentNode) {
                elements.removeAllFiles.classList.remove('hidden');
            }
            
        }

        function addFileToUI(file) {
            const fileSize = (file.size / (1024 * 1024)).toFixed(2);
            const emptyEl = document.getElementById('importedFilesEmpty');
            if (emptyEl) emptyEl.classList.add('hidden');
            const li = document.createElement('li');
            li.dataset.fileName = file.name;
            li.className = 'flex items-center justify-between gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200';
            li.innerHTML = `
                <div class="flex items-center">
                    <i class="fa fa-file-excel-o text-success mr-3"></i>
                    <div>
                        <p class="font-bold text-sm text-slate-900 truncate max-w-xs">${file.name}</p>
                        <p class="text-xs text-slate-500">${fileSize} MB</p>
                    </div>
                </div>
                <button class="inline-flex h-8 w-8 items-center justify-center rounded-full border border-slate-200 bg-white text-slate-500 hover:text-indigo-600 hover:border-indigo-200 hover:bg-indigo-50 transition-all-300 remove-file" data-name="${file.name}" aria-label="删除文件 ${file.name}" title="删除文件 ${file.name}">
                    <i class="fa fa-times"></i>
                </button>
            `;
            
            elements.importedFiles.appendChild(li);
            
            // 添加删除单个文件的事件
            li.querySelector('.remove-file').addEventListener('click', function(e) {
                const fileName = e.target.closest('.remove-file').dataset.name;
                removeFile(fileName);
            });
        }

        function removeFile(fileName) {
            // 从数据中移除
            appData.files = appData.files.filter(file => file.name !== fileName);
            appData.worksheets = appData.worksheets.filter(ws => ws.file.name !== fileName);
            
            // 从UI中移除
            const fileElement = Array.from(elements.importedFiles.children).find(li => li.dataset.fileName === fileName);
            
            if (fileElement) {
                fileElement.remove();
            }
            
            // 如果没有文件了，保持列表展开但隐藏清空按钮
            if (appData.files.length === 0) {
                // 文件列表现在固定展开，无需隐藏
                if (elements.removeAllFiles && elements.removeAllFiles.parentNode) {
                    elements.removeAllFiles.classList.add('hidden');
                }

                const appendMode = document.getElementById('appendMode')?.checked || false;
                if (!appendMode) {
                    // 非追加模式：彻底重置数据与配置
                    appData.worksheets = [];
                    appData.parsedData = [];
                    appData.processedData = [];

                    appData.config = appData.config || {};
                    appData.config.dateColumn = '';
                    appData.config.dataStartColumn = '';
                    appData.config.dataEndColumn = '';
                    appData.config.meteringPointColumn = '';
                    appData.config.multiplierColumn = '';
                    appData.config.dataStructure = '';

                    // 清空列选择（含计量点、多倍率列）
                    elements.dateColumnSelect.innerHTML = '<option value="">时间索引列</option>';
                    elements.dataStartColumnSelect.innerHTML = '<option value="">起始列</option>';
                    elements.dataEndColumnSelect.innerHTML = '<option value="">结束列</option>';
                    elements.meteringPointColumnSelect.innerHTML = '<option value="">计量点列</option>';
                    if (elements.meteringPointColumnSelect_v2) elements.meteringPointColumnSelect_v2.innerHTML = '<option value="">计量点列</option>';
                    elements.multiplierColumnSelect.innerHTML = '<option value="">倍率列</option>';
                } else {
                    // 追加模式：保留数据与配置，不清空列选择
                }

                const emptyEl = document.getElementById('importedFilesEmpty');
                if (emptyEl) emptyEl.classList.remove('hidden');
            } else if (appData.worksheets.length > 0) {
                // 重新填充列选择
                populateColumnSelects(appData.worksheets[0].data[0] || [], appData.worksheets[0].data);
            }

            updateDataStats();
            try { updateDataPreview(); } catch (e) {}
            if (typeof checkConfigValidity === 'function') {
                try { checkConfigValidity(); } catch (e) {}
            }
            
            // 更新原始数据预览
            buildRawDataPreview();
            
            showNotification('信息', `文件 "${fileName}" 已移除`, 'info');
        }

        // 加载示例数据
        function loadExampleData() {
            // 如果已有数据，提示用户
            if (appData.files.length > 0) {
                if (!confirm('加载示例数据将清除当前已上传的所有文件，是否继续？')) {
                    return;
                }
            }

            // 清除当前数据
            removeAllFiles({ silent: true, force: true });
            
            // 构造模拟数据（标准24小时负荷数据）
            const fileName = '示例负荷数据.csv';
            const headers = ['日期', '计量点编号', '0:00', '1:00', '2:00', '3:00', '4:00', '5:00', '6:00', '7:00', '8:00', '9:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00', '19:00', '20:00', '21:00', '22:00', '23:00'];
            
            const generateDailyLoad = (base, noise) => {
                return Array.from({length: 24}, (_, i) => {
                    // 模拟典型的双峰负荷曲线
                    const hour = i;
                    let factor = 0.5;
                    if (hour >= 8 && hour <= 12) factor = 0.8 + Math.random() * 0.2; // 上午峰
                    else if (hour >= 13 && hour <= 17) factor = 0.7 + Math.random() * 0.1; // 下午平
                    else if (hour >= 18 && hour <= 21) factor = 0.9 + Math.random() * 0.1; // 晚高峰
                    else if (hour >= 22 || hour <= 5) factor = 0.3 + Math.random() * 0.1; // 深夜谷
                    return (base * factor + (Math.random() - 0.5) * noise).toFixed(2);
                });
            };

            const mockRows = [headers];
            const startDate = new Date();
            startDate.setDate(startDate.getDate() - 7); // 从7天前开始

            for (let i = 0; i < 7; i++) {
                const date = new Date(startDate);
                date.setDate(date.getDate() + i);
                const dateStr = date.toISOString().split('T')[0];
                
                // 模拟两个不同的计量点
                mockRows.push([dateStr, 'MP001', ...generateDailyLoad(1000, 50)]);
                mockRows.push([dateStr, 'MP002', ...generateDailyLoad(1500, 80)]);
            }

            // 模拟文件对象添加到 appData
            const fileInfo = {
                name: fileName,
                size: 1024,
                type: 'text/csv',
                status: 'success',
                progress: 100,
                id: 'example_' + Date.now()
            };

            appData.files.push(fileInfo);
            appData.parsedData.push({
                fileName: fileName,
                sheets: [{
                    name: 'Sheet1',
                    data: mockRows
                }]
            });

            // 渲染 UI
            renderFileList();
            setImportStatus('success');
            updateStepStatus(1, true);
            
            // 自动配置常用选项
            appData.config.dateColumn = '日期';
            appData.config.dataStartColumn = '0:00';
            appData.config.dataEndColumn = '23:00';
            appData.config.meteringPointColumn = '计量点编号';
            appData.config.timeInterval = 60;
            
            // 手动触发一次 UI 更新和数据处理
            setTimeout(() => {
                // 更新下拉框选项
                updateColumnSelects(mockRows);
                
                // 设置下拉框初始值
                elements.dateColumnSelect.value = '日期';
                elements.dataStartColumnSelect.value = '0:00';
                elements.dataEndColumnSelect.value = '23:00';
                elements.meteringPointColumnSelect.value = '计量点编号';
                elements.timeIntervalSelect.value = '60';
                
                // 处理数据
                reprocessDataWithConfig();
                
                // 提示用户
                showNotification('成功', '示例数据已加载，包含过去7天的模拟负荷。', 'success');
            }, 100);
        }

        // 重置所有文件和数据
        function removeAllFiles(options = {}) {
            const { silent = false, keepStatus = false, force = false } = options;
            if (appData.files.length === 0 && !force) return;
            
            // 重置顶部显示
            const headerDateRangeEl = document.getElementById('headerDateRange');
            const headerIntervalEl = document.getElementById('headerInterval');
            if (headerDateRangeEl) headerDateRangeEl.textContent = '--';
            if (headerIntervalEl) headerIntervalEl.textContent = '--';
            
            // 清空数据 - 追加模式下不清除数据
            const appendMode = document.getElementById('appendMode')?.checked || false;
            if (!appendMode) {
                appData.files = [];
                appData.worksheets = [];
                appData.parsedData = [];
                appData.processedData = [];
                
                // 同步重置配置
                appData.config = appData.config || {};
                appData.config.dateColumn = '';
                appData.config.timeColumn = '';
                appData.config.dataStartColumn = '';
                appData.config.dataEndColumn = '';
                appData.config.meteringPointColumn = '';
                appData.config.multiplierColumn = '';
                appData.config.dataStructure = '';

                // 重置步骤指示灯
                updateStepStatus(1, false);
                updateStepStatus(2, false);
                updateStepStatus(3, false);
            } else {
                // 追加模式下只清空文件列表，保留数据
                appData.files = [];
                // 不清除worksheets, parsedData, processedData, config
            }
            
            // 清空UI
            elements.importedFiles.innerHTML = '';

            const emptyEl = document.getElementById('importedFilesEmpty');
            if (emptyEl) {
                elements.importedFiles.appendChild(emptyEl);
                emptyEl.classList.remove('hidden');
            }

            if (elements.uploadProgress && elements.uploadProgress.parentNode) {
                elements.uploadProgress.classList.add('hidden');
            }
            // 文件列表现在固定展开，无需隐藏
            if (elements.removeAllFiles && elements.removeAllFiles.parentNode) {
                elements.removeAllFiles.classList.add('hidden');
            }

            // 清空列选择（含计量点、多倍率列）
            elements.dateColumnSelect.innerHTML = '<option value="">时间索引列</option>';
            elements.dataStartColumnSelect.innerHTML = '<option value="">起始列</option>';
            elements.dataEndColumnSelect.innerHTML = '<option value="">结束列</option>';
            elements.meteringPointColumnSelect.innerHTML = '<option value="">计量点列</option>';
            if (elements.meteringPointColumnSelect_v2) elements.meteringPointColumnSelect_v2.innerHTML = '<option value="">计量点列</option>';
            elements.multiplierColumnSelect.innerHTML = '<option value="">倍率列</option>';
            
            // 重置筛选、预览与概览
            if (!appendMode) {
                updateDataStats();
                updateDataOverview();
                updateDataPreview();
            }
            
            if (!silent) {
                showNotification('信息', '所有文件已移除', 'info');
            }
            if (!keepStatus) {
                setImportStatus('idle');
            }
        }

        function normalizeCsvContent(csvText) {
            let text = String(csvText || '');
            if (text.charCodeAt(0) === 0xFEFF) {
                text = text.slice(1);
            }
            const firstLineEnd = text.indexOf('\n');
            const firstLine = firstLineEnd === -1 ? text.trim().toLowerCase() : text.slice(0, firstLineEnd).trim().toLowerCase();
            if (firstLine.startsWith('sep=')) {
                text = firstLineEnd === -1 ? '' : text.slice(firstLineEnd + 1);
            }
            return text;
        }

        function decodeCsvArrayBuffer(buffer) {
            const bytes = new Uint8Array(buffer);
            let text = '';
            try {
                text = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
            } catch (e) {
                text = '';
            }
            if (text) {
                const badCount = (text.match(/\uFFFD/g) || []).length;
                if (badCount / Math.max(text.length, 1) > 0.01) {
                    try {
                        text = new TextDecoder('gb18030', { fatal: false }).decode(bytes);
                    } catch (e) {}
                }
            }
            return text;
        }

        function parseCsvText(csvText) {
            const text = String(csvText || '');
            if (!text.trim()) return [];
            const lines = text.split(/\r\n|\n|\r/);
            if (lines.length === 0) return [];

            // 自动检测分隔符
            const sampleLine = lines.find(l => l.trim().length > 0) || '';
            const commaCount = (sampleLine.match(/,/g) || []).length;
            const semicolonCount = (sampleLine.match(/;/g) || []).length;
            const tabCount = (sampleLine.match(/\t/g) || []).length;
            let delimiter = ',';
            if (semicolonCount > commaCount && semicolonCount >= tabCount) {
                delimiter = ';';
            } else if (tabCount > commaCount && tabCount > semicolonCount) {
                delimiter = '\t';
            }

            const rows = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                const row = [];
                let current = '';
                let inQuotes = false;

                for (let j = 0; j < line.length; j++) {
                    const ch = line[j];
                    if (ch === '"') {
                        if (inQuotes && j + 1 < line.length && line[j + 1] === '"') {
                            current += '"';
                            j++;
                        } else {
                            inQuotes = !inQuotes;
                        }
                    } else if (ch === delimiter && !inQuotes) {
                        row.push(current);
                        current = '';
                    } else {
                        current += ch;
                    }
                }
                row.push(current);

                // 跳过完全空行
                const hasValue = row.some(cell => String(cell).trim() !== '');
                if (hasValue) {
                    rows.push(row);
                }
            }

            return rows;
        }

        function repairWorksheetRef(worksheet) {
            if (!worksheet || typeof worksheet !== 'object') return;
            if (typeof XLSX === 'undefined' || !XLSX?.utils?.decode_cell || !XLSX?.utils?.encode_range) return;

            let minRow = Infinity;
            let minCol = Infinity;
            let maxRow = -1;
            let maxCol = -1;

            const keys = Object.keys(worksheet);
            for (let i = 0; i < keys.length; i++) {
                const addr = keys[i];
                if (!addr || addr[0] === '!') continue;
                if (!/^[A-Z]+[0-9]+$/.test(addr)) continue;
                const cell = XLSX.utils.decode_cell(addr);
                if (cell.r < minRow) minRow = cell.r;
                if (cell.c < minCol) minCol = cell.c;
                if (cell.r > maxRow) maxRow = cell.r;
                if (cell.c > maxCol) maxCol = cell.c;
            }

            if (!Number.isFinite(minRow) || !Number.isFinite(minCol) || maxRow < 0 || maxCol < 0) return;

            const computedRef = XLSX.utils.encode_range({ s: { r: minRow, c: minCol }, e: { r: maxRow, c: maxCol } });
            const existingRef = worksheet['!ref'];

            if (!existingRef || typeof existingRef !== 'string') {
                worksheet['!ref'] = computedRef;
                return;
            }

            let existingRange = null;
            try {
                existingRange = XLSX.utils.decode_range(existingRef);
            } catch (e) {
                existingRange = null;
            }

            if (!existingRange) {
                worksheet['!ref'] = computedRef;
                return;
            }

            const shouldExpand =
                minRow < existingRange.s.r ||
                minCol < existingRange.s.c ||
                maxRow > existingRange.e.r ||
                maxCol > existingRange.e.c;

            if (shouldExpand) {
                worksheet['!ref'] = computedRef;
            }
        }

        // 填充空列，确保列索引从0开始连续
        function fillEmptyColumns(jsonData) {
            if (!Array.isArray(jsonData) || jsonData.length === 0) return jsonData;

            // 找出所有行中最长的列数
            let maxColumns = 0;
            for (const row of jsonData) {
                if (Array.isArray(row) && row.length > maxColumns) {
                    maxColumns = row.length;
                }
            }

            // 对每一行进行填充，确保每行都有相同的列数
            const filledData = [];
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (Array.isArray(row)) {
                    const filledRow = [...row];
                    // 填充空列
                    while (filledRow.length < maxColumns) {
                        filledRow.push(''); // 使用空字符串填充
                    }
                    filledData.push(filledRow);
                } else {
                    // 如果不是数组，创建一个新的数组
                    const newRow = [];
                    for (let j = 0; j < maxColumns; j++) {
                        newRow.push('');
                    }
                    filledData.push(newRow);
                }
            }

            return filledData;
        }

        function normalizeTimeString(timeStr) {
            let text = String(timeStr || '').trim();
            if (!text) return '';
            text = text.replace(/：/g, ':');
            if (/^\d{1,2}$/.test(text)) {
                return `${text.padStart(2, '0')}:00`;
            }
            if (/^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(text)) {
                const parts = text.split(':').map(p => p.padStart(2, '0'));
                return parts.length === 2 ? `${parts[0]}:${parts[1]}` : `${parts[0]}:${parts[1]}:${parts[2]}`;
            }
            return text;
        }

        function buildDateTimeString(datePart, timePart) {
            const dateText = String(datePart || '').trim();
            if (!dateText) return '';
            if (/\d{1,2}:\d{2}/.test(dateText)) return dateText;
            const timeText = normalizeTimeString(timePart);
            return timeText ? `${dateText} ${timeText}` : dateText;
        }

        function detectTimeColumnIndex(headerRow, data, dateColumnIndex, startRow = 1) {
            if (!Array.isArray(headerRow) || headerRow.length === 0) return -1;
            const timeKeywords = ['时间', '时刻', 'time', 'timestamp', '采样时间', '时分', '分钟'];
            // 支持多种时间格式：00:00, 00:15, 0:00, 12:30 等
            const timePattern = /^([01]?\d|2[0-3])[:：][0-5]\d$/;
            const timePatternWithSeconds = /^([01]?\d|2[0-3])[:：][0-5]\d[:：][0-5]\d$/;
            const sampleRows = Array.isArray(data) ? data.slice(startRow, Math.min(startRow + 50, data.length)) : [];
            let bestIndex = -1;
            let bestScore = 0;

            headerRow.forEach((header, index) => {
                if (index === dateColumnIndex) return;
                const headerText = String(header || '').toLowerCase();
                let score = 0;
                
                // 表头关键字匹配（增强权重）
                if (timeKeywords.some(k => headerText.includes(k))) {
                    score += 25;
                    // 如果是精确匹配"时间"或"时刻"，额外加分
                    if (headerText === '时间' || headerText === '时刻' || headerText === 'time') {
                        score += 10;
                    }
                }
                
                // 数据内容分析
                if (sampleRows.length > 0) {
                    let timeCount = 0;
                    let totalCount = 0;
                    sampleRows.forEach(row => {
                        const val = row ? row[index] : null;
                        if (val !== undefined && val !== null && String(val).trim() !== '') {
                            totalCount++;
                            const valStr = String(val).trim();
                            if (timePattern.test(valStr) || timePatternWithSeconds.test(valStr)) {
                                timeCount++;
                            }
                        }
                    });
                    if (totalCount > 0) {
                        const timeRatio = timeCount / totalCount;
                        if (timeRatio > 0.8) score += 50; // 高置信度
                        else if (timeRatio > 0.5) score += 30; // 中等置信度
                        else if (timeRatio > 0.3) score += 10; // 低置信度
                    }
                }
                if (score > bestScore) {
                    bestScore = score;
                    bestIndex = index;
                }
            });

            return bestScore >= 20 ? bestIndex : -1;
        }

        function populateColumnSelects(headerRow, data) {
            // 清空现有选项
            elements.dateColumnSelect.innerHTML = '<option value="">时间索引列</option>';
            elements.dataStartColumnSelect.innerHTML = '<option value="">起始列</option>';
            elements.dataEndColumnSelect.innerHTML = '<option value="">结束列</option>';
            elements.meteringPointColumnSelect.innerHTML = '<option value="">计量点列</option>';
            if (elements.meteringPointColumnSelect_v2) elements.meteringPointColumnSelect_v2.innerHTML = '<option value="">计量点列</option>';
            elements.multiplierColumnSelect.innerHTML = '<option value="">倍率列</option>';
            
            // 添加新选项 (支持所有列，不再限制50列)
            for (let i = 0; i < headerRow.length; i++) {
                // 生成列字母，支持超过26列的情况（如AA, AB等）
                let columnLetter = "";
                let columnNum = i;
                while (columnNum >= 0) {
                    columnLetter = String.fromCharCode(65 + (columnNum % 26)) + columnLetter;
                    columnNum = Math.floor(columnNum / 26) - 1;
                    if (columnNum < 0) break;
                }
                
                const headerText = headerRow[i] ? String(headerRow[i]).substring(0, 20) : `列 ${columnLetter}`;
                
                const dateOption = document.createElement('option');
                dateOption.value = i;
                dateOption.textContent = `${columnLetter}: ${headerText}`;
                elements.dateColumnSelect.appendChild(dateOption);
                
                const startOption = document.createElement('option');
                startOption.value = i;
                startOption.textContent = `${columnLetter}: ${headerText}`;
                elements.dataStartColumnSelect.appendChild(startOption);
                
                const endOption = document.createElement('option');
                endOption.value = i;
                endOption.textContent = `${columnLetter}: ${headerText}`;
                elements.dataEndColumnSelect.appendChild(endOption);
                
                const meteringPointOption = document.createElement('option');
                meteringPointOption.value = i;
                meteringPointOption.textContent = `${columnLetter}: ${headerText}`;
                elements.meteringPointColumnSelect.appendChild(meteringPointOption);
                
                if (elements.meteringPointColumnSelect_v2) {
                    const meteringPointOption_v2 = document.createElement('option');
                    meteringPointOption_v2.value = i;
                    meteringPointOption_v2.textContent = `${columnLetter}: ${headerText}`;
                    elements.meteringPointColumnSelect_v2.appendChild(meteringPointOption_v2);
                }
                
                const multiplierColumnOption = document.createElement('option');
                multiplierColumnOption.value = i;
                multiplierColumnOption.textContent = `${columnLetter}: ${headerText}`;
                elements.multiplierColumnSelect.appendChild(multiplierColumnOption);
                
                // PT 列选项
                if (elements.ptColumnSelect) {
                    const ptOption = document.createElement('option');
                    ptOption.value = i;
                    ptOption.textContent = `${columnLetter}: ${headerText}`;
                    elements.ptColumnSelect.appendChild(ptOption);
                }
                
                // CT 列选项
                if (elements.ctColumnSelect) {
                    const ctOption = document.createElement('option');
                    ctOption.value = i;
                    ctOption.textContent = `${columnLetter}: ${headerText}`;
                    elements.ctColumnSelect.appendChild(ctOption);
                }
            }
            
            // 自动识别列
            if (data && data.length > 1) {
                autoDetectColumns(headerRow, data);
            }
        }

        // 自动识别数据结构类型函数（增强版 v2.0）
        function autoDetectDataStructure(jsonData) {
            if (!jsonData || jsonData.length < 2) return;
            
            const headerRow = jsonData[0];
            const dataRows = jsonData.slice(1, Math.min(51, jsonData.length));
            
            let timeSeriesScore = 0;  // 列对行（宽表）得分
            let columnSeriesScore = 0;  // 列对列（长表）得分
            
            // ========== 1. 检查表头中的时间模式（核心特征）==========
            let timeHeaderCount = 0;
            let timeHeaderIndices = [];
            const timePattern = /^([01]?\d|2[0-3])[:：][0-5]\d$/;  // 匹配 00:00, 23:59 等
            const hourPattern = /^\d{1,2}(时|点)$/;  // 匹配 1时, 12点 等
            const timePatternWithSeconds = /^([01]?\d|2[0-3])[:：][0-5]\d[:：][0-5]\d$/;  // 匹配 00:00:00
            
            headerRow.forEach((header, idx) => {
                const text = String(header).trim();
                if (timePattern.test(text) || hourPattern.test(text) || timePatternWithSeconds.test(text)) {
                    timeHeaderCount++;
                    timeHeaderIndices.push(idx);
                }
            });
            
            // 宽表核心特征：表头包含多个连续的时间点列
            let maxConsecutiveTimeHeaders = 0;
            let timeHeaderGroups = [];
            
            if (timeHeaderIndices.length >= 2) {
                let currentGroup = [timeHeaderIndices[0]];
                let consecutiveCount = 1;
                maxConsecutiveTimeHeaders = 1;
                
                for (let i = 1; i < timeHeaderIndices.length; i++) {
                    if (timeHeaderIndices[i] === timeHeaderIndices[i-1] + 1) {
                        currentGroup.push(timeHeaderIndices[i]);
                        consecutiveCount++;
                        maxConsecutiveTimeHeaders = Math.max(maxConsecutiveTimeHeaders, consecutiveCount);
                    } else {
                        if (currentGroup.length >= 4) {
                            timeHeaderGroups.push([...currentGroup]);
                        }
                        currentGroup = [timeHeaderIndices[i]];
                        consecutiveCount = 1;
                    }
                }
                if (currentGroup.length >= 4) {
                    timeHeaderGroups.push(currentGroup);
                }
            } else if (timeHeaderIndices.length === 1) {
                maxConsecutiveTimeHeaders = 1;
            }
            
            // 根据时间列数量评分（这是最重要的特征）
            if (timeHeaderCount >= 24) {
                // 强宽表特征：超过24个时间列（24小时数据）
                timeSeriesScore += 50;
            } else if (timeHeaderCount >= 12) {
                timeSeriesScore += 40;
            } else if (timeHeaderCount >= 4) {
                timeSeriesScore += 25;
            }
            
            // 连续时间列加分
            if (maxConsecutiveTimeHeaders >= 48) {
                timeSeriesScore += 25;  // 48个点（半小时间隔24小时）
            } else if (maxConsecutiveTimeHeaders >= 24) {
                timeSeriesScore += 20;  // 24个点（小时数据）
            } else if (maxConsecutiveTimeHeaders >= 12) {
                timeSeriesScore += 10;
            }
            
            // 多组时间列（可能是多天的数据）
            if (timeHeaderGroups.length >= 2) {
                timeSeriesScore += 15;
            }
            
            // ========== 2. 寻找关键列（提前到列数分析之前）==========
            let dateColIdx = -1;
            let timeColIdx = -1;
            
            // 找日期列
            const dateKeywords = ['日期', 'date', 'day', 'data date', '记录日期', '数据日期'];
            headerRow.forEach((header, idx) => {
                const text = String(header).toLowerCase().trim();
                if (dateKeywords.some(k => text.includes(k)) && dateColIdx === -1) {
                    dateColIdx = idx;
                }
            });
            
            // ========== 3. 列数分析 ==========
            const totalColumns = headerRow.length;
            if (totalColumns >= 50) {
                timeSeriesScore += 20;  // 大量列通常是宽表
            } else if (totalColumns >= 24) {
                timeSeriesScore += 15;
            } else if (totalColumns >= 12) {
                timeSeriesScore += 5;
            } else if (totalColumns <= 6) {
                columnSeriesScore += 15;  // 很少列可能是长表
            } else if (totalColumns <= 8) {
                columnSeriesScore += 8;
            }
            
            // 时间列占比
            const timeColumnRatio = timeHeaderCount / totalColumns;
            if (timeColumnRatio >= 0.5) {
                timeSeriesScore += 20;  // 超过一半是时间列，强宽表特征
            } else if (timeColumnRatio >= 0.3) {
                timeSeriesScore += 10;
            }
            
            // ========== 4. 检测"日期+指标"标准时间序列格式 ==========
            // 这种格式：第一列是日期时间，后面是多列指标（电流、电压、功率等）
            if (timeHeaderCount < 4 && dateColIdx !== -1) {
                // 检查是否有典型的电力指标列
                const powerKeywords = ['电流', '电压', '功率', '有功', '无功', '电量', '电能', '因数', 'A相', 'B相', 'C相', '零线'];
                let powerIndicatorCount = 0;
                
                headerRow.forEach((header, idx) => {
                    if (idx === dateColIdx) return;
                    const text = String(header).toLowerCase();
                    if (powerKeywords.some(k => text.includes(k))) {
                        powerIndicatorCount++;
                    }
                });
                
                // 如果有日期列 + 多个电力指标列，且没有时间格式的列，判定为列对列（长表）
                if (powerIndicatorCount >= 3) {
                    columnSeriesScore += 40;  // 强长表特征
                    debugLog('检测到标准时间序列格式（日期+指标）', 'info', {
                        powerIndicatorCount,
                        dateColIdx,
                        timeHeaderCount
                    });
                }
            }
            
            // 找时间列（用于长表检测）
            const timeKeywords = ['时间', 'time', '时刻', '小时'];
            headerRow.forEach((header, idx) => {
                const text = String(header).toLowerCase().trim();
                // 时间列不能是时间格式的表头（避免误判）
                if (timeKeywords.some(k => text.includes(k)) && 
                    !timeHeaderIndices.includes(idx) && 
                    timeColIdx === -1) {
                    timeColIdx = idx;
                }
            });
            
            // 长表特征：存在独立的时间列（非表头时间格式）
            if (timeColIdx !== -1 && timeHeaderCount < 4) {
                columnSeriesScore += 30;
                if (dateColIdx !== -1) columnSeriesScore += 10;
            }
            
            // ========== 4. 数据内容深度分析 ==========
            if (dateColIdx !== -1 && dataRows.length > 0) {
                const dateVals = dataRows.map(row => String(row[dateColIdx])).filter(d => d && d !== 'undefined');
                const uniqueDates = new Set(dateVals).size;
                
                if (dateVals.length > 5) {
                    const duplicationRate = (dateVals.length - uniqueDates) / dateVals.length;
                    
                    // 宽表情况下，日期重复是正常的（多个计量点同一天）
                    // 只有当没有大量时间列时，日期重复才是长表特征
                    if (timeHeaderCount < 4 && duplicationRate > 0.5) {
                        columnSeriesScore += 20;
                    }
                    
                    // 日期不重复是宽表特征
                    if (duplicationRate < 0.1 && timeHeaderCount >= 4) {
                        timeSeriesScore += 10;
                    }
                }
            }
            
            // ========== 5. 数值列分布分析 ==========
            let numericColumnCount = 0;
            let timeRangeNumericCount = 0;  // 时间列范围内的数值列
            const sampleRows = Math.min(10, dataRows.length);
            
            for (let col = 0; col < headerRow.length; col++) {
                let colNumericCount = 0;
                for (let row = 0; row < sampleRows; row++) {
                    const val = dataRows[row]?.[col];
                    if (val !== null && val !== undefined && !isNaN(parseFloat(val)) && isFinite(val)) {
                        colNumericCount++;
                    }
                }
                const numericRatio = colNumericCount / sampleRows;
                if (numericRatio >= 0.8) {
                    numericColumnCount++;
                    // 检查这个时间列是否也是数值列
                    if (timeHeaderIndices.includes(col)) {
                        timeRangeNumericCount++;
                    }
                }
            }
            
            // 大量数值列是宽表特征
            if (numericColumnCount >= 24) {
                timeSeriesScore += 20;
            } else if (numericColumnCount >= 12) {
                timeSeriesScore += 15;
            }
            
            // 时间列同时也是数值列（数据列）
            if (timeRangeNumericCount >= timeHeaderCount * 0.8 && timeHeaderCount >= 12) {
                timeSeriesScore += 15;  // 时间列包含数值数据，强宽表特征
            }
            
            // ========== 6. 计量点列检测 ==========
            const meteringPointKeywords = ['编号', 'id', '标识', '代码', '设备号', '计量点', '测点', '点位', '表号', '资产编号', '终端'];
            let hasMeteringPointCol = false;
            let meteringPointColIdx = -1;
            
            headerRow.forEach((header, idx) => {
                const text = String(header).toLowerCase();
                if (meteringPointKeywords.some(k => text.includes(k)) && meteringPointColIdx === -1) {
                    hasMeteringPointCol = true;
                    meteringPointColIdx = idx;
                }
            });
            
            // 有计量点列 + 大量时间列 = 典型的电力负荷宽表
            if (hasMeteringPointCol && timeHeaderCount >= 12) {
                timeSeriesScore += 15;
            }
            
            // ========== 7. 行数与列数比例分析 ==========
            const rowCount = dataRows.length;
            const colCount = headerRow.length;
            const rowColRatio = rowCount / colCount;
            
            // 宽表通常列数接近或超过行数，或者比例较小
            if (colCount > rowCount && timeHeaderCount >= 12) {
                timeSeriesScore += 10;
            }
            
            // ========== 8. 特殊模式检测 ==========
            // 检测是否像电力负荷数据（00:00, 00:15, 00:30... 这种15分钟间隔）
            if (timeHeaderCount >= 90 && timeHeaderCount <= 100) {
                // 96个点 = 15分钟间隔 × 24小时，典型的电力负荷数据
                const firstTimeIdx = timeHeaderIndices[0];
                const lastTimeIdx = timeHeaderIndices[timeHeaderIndices.length - 1];
                if (lastTimeIdx - firstTimeIdx + 1 === timeHeaderCount) {
                    // 连续的时间列
                    timeSeriesScore += 20;
                }
            }
            
            // ========== 综合决策 ==========
            let detectedStructure = 'columnToRow';
            let confidence = 'medium';
            const scoreDiff = Math.abs(columnSeriesScore - timeSeriesScore);
            
            // 决策逻辑
            if (columnSeriesScore > timeSeriesScore) {
                detectedStructure = 'columnToColumn';
            }
            
            // 置信度判断
            if (scoreDiff > 30) {
                confidence = 'high';
            } else if (scoreDiff > 15) {
                confidence = 'medium';
            } else {
                confidence = 'low';
            }
            
            // 特殊情况：如果有很多时间列，强制判定为宽表
            if (timeHeaderCount >= 24 && timeSeriesScore > columnSeriesScore * 0.5) {
                detectedStructure = 'columnToRow';
                if (timeHeaderCount >= 48) confidence = 'high';
            }
            
            // 调试日志
            debugLog('数据结构识别评分（增强版）', 'info', {
                timeSeriesScore,
                columnSeriesScore,
                timeHeaderCount,
                maxConsecutiveTimeHeaders,
                timeHeaderGroups: timeHeaderGroups.length,
                totalColumns,
                timeColumnRatio: timeColumnRatio.toFixed(2),
                dateColIdx,
                timeColIdx,
                numericColumnCount,
                timeRangeNumericCount,
                hasMeteringPointCol,
                rowColRatio: rowColRatio.toFixed(2),
                detectedStructure,
                confidence
            });
            
            // 更新 UI
            const radio = document.querySelector(`input[name="dataStructure"][value="${detectedStructure}"]`);
            if (radio) {
                radio.checked = true;
                appData.config.dataStructure = detectedStructure;
                
                const container = radio.closest('label');
                if (container) {
                    container.classList.add('ring-2', 'ring-indigo-500', 'ring-offset-2');
                    setTimeout(() => container.classList.remove('ring-2', 'ring-indigo-500', 'ring-offset-2'), 2000);
                }
                
                radio.dispatchEvent(new Event('change', { bubbles: true }));
                
                const structureName = detectedStructure === 'columnToRow' ? '列对行 (宽表)' : '列对列 (长表)';
                const confidenceName = confidence === 'high' ? '极高' : (confidence === 'medium' ? '较高' : '较低');
                showNotification('结构识别', `自动识别为：${structureName} (置信度: ${confidenceName})`, 'info');
            }
        }

        // 检测日期重复情况并提醒用户（仅在列对行结构时触发）
        function detectDuplicateDatesAndNotify(jsonData) {
            if (!jsonData || jsonData.length < 2) return;
            
            // 只在列对行（宽表）结构时进行检测
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
            if (dataStructure !== 'columnToRow') {
                console.log('[日期重复检测] 当前为列对列结构，跳过日期重复检测');
                return;
            }
            
            const headerRow = jsonData[0];
            const dataRows = jsonData.slice(1);
            
            // 查找日期列
            let dateColIdx = -1;
            const dateKeywords = ['日期', 'date', 'day', '数据日期', '记录日期'];
            headerRow.forEach((header, idx) => {
                const text = String(header).toLowerCase().trim();
                if (dateKeywords.some(k => text.includes(k)) && dateColIdx === -1) {
                    dateColIdx = idx;
                }
            });
            
            if (dateColIdx === -1) return;
            
            // 统计日期出现次数
            const dateCount = {};
            const dateSamples = {};
            
            dataRows.forEach((row, rowIdx) => {
                if (row && row[dateColIdx]) {
                    const dateVal = String(row[dateColIdx]).trim();
                    if (dateVal && dateVal !== 'undefined') {
                        dateCount[dateVal] = (dateCount[dateVal] || 0) + 1;
                        // 保存第一行样本用于显示
                        if (!dateSamples[dateVal]) {
                            dateSamples[dateVal] = { row: row, rowIndex: rowIdx + 1 };
                        }
                    }
                }
            });
            
            // 检测重复日期
            const duplicateDates = Object.entries(dateCount)
                .filter(([date, count]) => count > 1)
                .sort((a, b) => b[1] - a[1]); // 按重复次数降序
            
            if (duplicateDates.length === 0) return;
            
            // 获取计量点列（用于提示）
            let meteringPointColIdx = -1;
            const meteringPointKeywords = ['编号', 'id', '标识', '代码', '设备号', '计量点', '表号', '资产编号', '终端'];
            headerRow.forEach((header, idx) => {
                const text = String(header).toLowerCase();
                if (meteringPointKeywords.some(k => text.includes(k)) && meteringPointColIdx === -1) {
                    meteringPointColIdx = idx;
                }
            });
            
            // 构建提醒信息
            const totalUniqueDates = Object.keys(dateCount).length;
            const totalDuplicateCount = duplicateDates.length;
            const maxDuplicate = duplicateDates[0][1];
            const mostFrequentDate = duplicateDates[0][0];
            
            // 获取最频繁日期的不同计量点示例
            let meteringPointExamples = [];
            if (meteringPointColIdx !== -1) {
                dataRows.forEach(row => {
                    if (row && row[dateColIdx] && row[meteringPointColIdx]) {
                        const dateVal = String(row[dateColIdx]).trim();
                        if (dateVal === mostFrequentDate) {
                            const meteringPoint = String(row[meteringPointColIdx]).trim();
                            if (meteringPoint && !meteringPointExamples.includes(meteringPoint)) {
                                meteringPointExamples.push(meteringPoint);
                            }
                        }
                    }
                });
            }
            
            // 构建详细提示
            let detailMessage = `检测到 ${totalDuplicateCount} 个日期存在重复记录，最多重复 ${maxDuplicate} 次。\n\n`;
            detailMessage += `最频繁的日期：${mostFrequentDate}（出现 ${maxDuplicate} 次）\n`;
            
            if (meteringPointExamples.length > 0) {
                detailMessage += `\n该日期包含以下不同计量点：\n`;
                meteringPointExamples.slice(0, 5).forEach((mp, idx) => {
                    detailMessage += `  ${idx + 1}. ${mp}\n`;
                });
                if (meteringPointExamples.length > 5) {
                    detailMessage += `  ... 还有 ${meteringPointExamples.length - 5} 个计量点\n`;
                }
            }
            
            detailMessage += `\n💡 建议操作：\n`;
            if (meteringPointColIdx !== -1) {
                detailMessage += `1. 在「计量点列」选择「${headerRow[meteringPointColIdx]}」\n`;
                detailMessage += `2. 在「计量点筛选」中选择特定计量点以筛选数据\n`;
            }
            detailMessage += `3. 或使用「高级过滤器」添加更多筛选条件`;
            
            // 显示提醒（延迟显示，让结构识别通知先显示）
            setTimeout(() => {
                showNotification(
                    '⚠️ 检测到多计量点数据', 
                    `发现 ${totalDuplicateCount} 个日期有重复记录（如 ${mostFrequentDate} 出现 ${maxDuplicate} 次），建议配置计量点筛选`, 
                    'warning',
                    8000
                );
                
                // 在控制台输出详细信息
                console.log('%c[日期重复检测]', 'color: #f59e0b; font-weight: bold; font-size: 14px;', {
                    总唯一日期数: totalUniqueDates,
                    重复日期数: totalDuplicateCount,
                    最大重复次数: maxDuplicate,
                    最频繁日期: mostFrequentDate,
                    计量点列: meteringPointColIdx !== -1 ? headerRow[meteringPointColIdx] : '未检测到',
                    计量点示例: meteringPointExamples.slice(0, 5),
                    详细统计: duplicateDates.slice(0, 10)
                });
                console.log('%c[建议操作]', 'color: #3b82f6; font-weight: bold;', detailMessage);
                
                // 如果有计量点列，自动填充计量点筛选选项
                if (meteringPointColIdx !== -1 && meteringPointExamples.length > 0) {
                    updateMeteringPointFilterOptionsWithData(meteringPointExamples, headerRow[meteringPointColIdx]);
                }
            }, 1500);
        }
        
        // 更新计量点筛选选项（带数据）
        function updateMeteringPointFilterOptionsWithData(meteringPoints, columnName) {
            const selects = [
                elements.meteringPointFilterSelect,
                elements.meteringPointFilterSelect_v2
            ];
            
            selects.forEach(select => {
                if (!select) return;
                
                // 保存当前值
                const currentValue = select.value;
                
                // 清空并添加默认选项
                select.innerHTML = `<option value="">选择 ${columnName}</option>`;
                
                // 添加计量点选项（限制数量）
                meteringPoints.slice(0, 50).forEach(mp => {
                    const option = document.createElement('option');
                    option.value = mp;
                    option.textContent = mp;
                    select.appendChild(option);
                });
                
                // 恢复之前的值（如果存在）
                if (currentValue) {
                    select.value = currentValue;
                }
            });
            
            // 高亮计量点筛选区域
            const filterSection = document.getElementById('meteringPointFilter')?.closest('.rounded-2xl');
            if (filterSection) {
                filterSection.classList.add('ring-2', 'ring-indigo-400', 'ring-offset-2');
                setTimeout(() => {
                    filterSection.classList.remove('ring-2', 'ring-indigo-400', 'ring-offset-2');
                }, 5000);
            }
        }

        // 将数字转换为Excel列号 (0 -> A, 25 -> Z, 26 -> AA, 27 -> AB, ...)
        function getColumnLabel(index) {
            let label = '';
            let num = index;
            do {
                label = String.fromCharCode(65 + (num % 26)) + label;
                num = Math.floor(num / 26) - 1;
            } while (num >= 0);
            return label;
        }

        // 自动识别列函数
        function autoDetectColumns(headerRow, data) {
            let dateColumnIndex = -1;
            let dataStartColumnIndex = -1;
            let dataEndColumnIndex = -1;
            let meteringPointColumnIndex = -1;
            let multiplierColumnIndex = -1;
            
            const dateKeywords = ['日期', 'date', 'datetime', '记录日期', '数据日期', '抄表日期', '结算日期', '采集日期'];
            const datePatterns = [/\d{4}[\-\/]\d{1,2}[\-\/]\d{1,2}/, /\d{1,2}[\-\/]\d{1,2}[\-\/]\d{4}/, /\d{8}/];
            
            const powerKeywords = ['功率', 'power', 'p', 'kw', '有功', '无功', '视在', '瞬时', '实时', '负荷', '平均功率', 'kw)', '(kw'];
            const energyKeywords = ['电能', '电量', 'energy', 'kwh', '累积', '累计', '总量', '用电量', '耗电量', '读数', '正向有功', 'kwh)', '(kwh', '电度', '总电能', '总电量', '表码', '示数', '总加', '合计', '总'];
            const generalDataKeywords = ['数据', '数值', '值', 'data', 'value', '开始', '电流', '电压', 'a', 'v', 'ia', 'ua', '总有功', 'a相', 'b相', 'c相'];
            
            const meteringPointKeywords = ['编号', 'id', '标识', '代码', '设备号', '设备编号', '计量点', '测点', '点位', '序号', 'no', '名称', 'name', '户号', '用户', '表号', '资产编号'];
            // 倍率检测优先级：1. 包含"倍率"两字 2. PT×CT 模式 3. 其他变比/系数关键字
            const multiplierKeywords = ['倍率', '变比', '系数', 'multiplier', 'ratio', '综合倍率', '乘数'];
            const ptKeywords = ['pt', '电压变比', '电压互感器', 'potential', 'pt变比'];
            const ctKeywords = ['ct', '电流变比', '电流互感器', 'current', 'ct变比'];

            const sampleRows = data.slice(1, Math.min(31, data.length)); // 采样增加到30行
            
            // 检测表头中的时间模式（用于横向展开格式）
            const timeHeaderPattern = /^([01]?\d|2[0-3])[:：][0-5]\d$/;
            const timeHeaderIndices = [];
            headerRow.forEach((header, index) => {
                if (timeHeaderPattern.test(String(header).trim())) {
                    timeHeaderIndices.push(index);
                }
            });
            
            // 为每列计算各项分数
            const columnScores = headerRow.map((header, index) => {
                const headerText = String(header).toLowerCase().trim();
                const scoreObj = {
                    index,
                    dateScore: 0,
                    dataScore: 0,
                    meteringPointScore: 0,
                    multiplierScore: 0,
                    ptScore: 0,
                    ctScore: 0,
                    columnType: 'general'
                };
                
                // 1. 表头关键字识别
                if (dateKeywords.some(k => headerText.includes(k))) scoreObj.dateScore += 20;
                
                // 2. 表头模式识别
                if (datePatterns.some(p => p.test(headerText))) scoreObj.dateScore += 15;
                
                // 3. 数据关键字识别
                let pScore = 0, eScore = 0, gScore = 0;
                powerKeywords.forEach(k => { if (headerText.includes(k)) pScore += (k === 'kw' || k === '功率' ? 20 : 15); });
                energyKeywords.forEach(k => { if (headerText.includes(k)) eScore += (k === 'kwh' || k === '电量' ? 20 : 15); });
                generalDataKeywords.forEach(k => { if (headerText.includes(k)) gScore += 10; });
                
                // 3.5 表头是时间格式（如 00:00, 00:15）- 这是横向展开格式的数据列特征
                if (timeHeaderPattern.test(String(header).trim())) {
                    pScore += 50; // 高权重，优先识别为功率数据列
                    scoreObj.columnType = 'power';
                }
                
                // 4. 计量点和倍率关键字
                if (meteringPointKeywords.some(k => headerText.includes(k))) scoreObj.meteringPointScore += 20;
                
                // 倍率检测优先级：1. 包含"倍率"两字（最高优先级）
                if (headerText.includes('倍率')) {
                    scoreObj.multiplierScore += 100; // 最高优先级
                } else if (multiplierKeywords.some(k => headerText.includes(k))) {
                    scoreObj.multiplierScore += 25; // 其他变比/系数关键字
                }
                
                // 4.5 PT 和 CT 关键字识别（增强检测）
                // PT检测：优先检测独立的PT关键字（不是"电压"的一部分）
                if (headerText === 'pt' || headerText === 'PT' || headerText.includes('pt变比') || headerText.includes('PT变比') || 
                    headerText.includes('电压互感器') || headerText.includes('potential')) {
                    scoreObj.ptScore = 50; // 提高PT检测权重
                } else if (ptKeywords.some(k => headerText.includes(k))) {
                    scoreObj.ptScore = 30;
                }
                
                // CT检测：优先检测独立的CT关键字（不是"电流"的一部分）
                if (headerText === 'ct' || headerText === 'CT' || headerText.includes('ct变比') || headerText.includes('CT变比') || 
                    headerText.includes('电流互感器') || headerText.includes('current transformer')) {
                    scoreObj.ctScore = 50; // 提高CT检测权重
                } else if (ctKeywords.some(k => headerText.includes(k))) {
                    scoreObj.ctScore = 30;
                }
                
                // 5. 样本数据内容分析 (更深入)
                let dateCount = 0, numericCount = 0, totalCount = 0;
                let integerCount = 0; // 用于识别倍率或ID
                let stringLenSum = 0;
                
                sampleRows.forEach(row => {
                    const val = row[index];
                    if (val !== undefined && val !== null && String(val).trim() !== "") {
                        totalCount++;
                        const text = String(val).trim();
                        stringLenSum += text.length;
                        
                        if (isDateLike(text)) dateCount++;
                        
                        const num = parseFloat(text);
                        if (!isNaN(num) && isFinite(num)) {
                            numericCount++;
                            if (Number.isInteger(num)) integerCount++;
                        }
                    }
                });
                
                if (totalCount > 0) {
                    const numRate = numericCount / totalCount;
                    const dateRate = dateCount / totalCount;
                    const avgLen = stringLenSum / totalCount;
                    
                    // 日期内容加分
                    if (dateRate > 0.8) scoreObj.dateScore += 40;
                    
                    // 数值内容加分
                    if (numRate > 0.9) {
                        scoreObj.dataScore += 15;
                        // 如果平均长度较短且多为整数，可能是倍率
                        if (avgLen < 5 && integerCount / totalCount > 0.8) scoreObj.multiplierScore += 10;
                    }
                    
                    // 分析数值特征，判断是否为累积电能数据（独立于 numRate 条件）
                    const numericValues = sampleRows
                        .map(row => parseFloat(row[index]))
                        .filter(v => !isNaN(v) && isFinite(v));
                    
                    console.log(`列 ${index}: 数值率=${numRate.toFixed(2)}, 数值个数=${numericValues.length}`);
                    
                    if (numericValues.length >= 3) {
                        // 检查数值是否整体呈递增趋势（累积电能特征）
                        let increasingCount = 0;
                        let totalVariance = 0;
                        for (let i = 1; i < numericValues.length; i++) {
                            if (numericValues[i] >= numericValues[i-1]) increasingCount++;
                            totalVariance += Math.abs(numericValues[i] - numericValues[i-1]);
                        }
                        const increasingRate = increasingCount / (numericValues.length - 1);
                        const avgVariance = totalVariance / (numericValues.length - 1);
                        const avgValue = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
                        
                        // 调试日志
                        console.log(`列 ${index} 数据特征: 递增率=${increasingRate.toFixed(2)}, 平均值=${avgValue.toFixed(2)}, 变异系数=${avgValue > 0 ? (avgVariance / avgValue).toFixed(2) : 'N/A'}`);
                        
                        // 累积电能特征：递增趋势明显，数值较大，变化相对平稳
                        // 放宽条件：递增率 > 50%，平均值 > 50，变异系数 < 1.0
                        if (increasingRate > 0.5 && avgValue > 50 && avgVariance / avgValue < 1.0) {
                            eScore += 60; // 强烈倾向于累积电能，权重高于时间格式（50分）
                            scoreObj.columnType = 'energy';
                            scoreObj.isCumulativeEnergy = true; // 标记为累积电能
                            console.log(`列 ${index} 被识别为累积电能`);
                        }
                    }
                    
                    // 计量点ID识别：如果包含字母且长度适中，或者全是数字但长度很长
                    if (numRate < 0.5 && avgLen > 5 && avgLen < 20) scoreObj.meteringPointScore += 10;
                }
                
                // 确定最终数据得分和类型
                const maxDScore = Math.max(pScore, eScore, gScore);
                scoreObj.dataScore += maxDScore;
                // 如果已经标记为累积电能，保持该类型
                if (!scoreObj.isCumulativeEnergy) {
                    scoreObj.columnType = pScore >= eScore && pScore >= gScore ? 'power' : (eScore > pScore && eScore >= gScore ? 'energy' : 'general');
                }
                
                return scoreObj;
            });
            
            // 决策排序
            const getBest = (list, scoreKey) => list.sort((a, b) => b[scoreKey] - a[scoreKey])[0]?.index ?? -1;
            
            dateColumnIndex = getBest([...columnScores], 'dateScore');
            
            // 优先检查是否有时间格式的表头列（横向展开格式）
            if (timeHeaderIndices.length >= 4) {
                // 横向展开格式：数据列是时间格式的列
                dataStartColumnIndex = timeHeaderIndices[0];
                dataEndColumnIndex = timeHeaderIndices[timeHeaderIndices.length - 1];
                console.log(`检测到横向展开格式，时间列范围: ${dataStartColumnIndex} - ${dataEndColumnIndex}`);
                
                // 检查时间格式列的数据类型（可能是功率或累积电能）
                // 采样几列来判断整体数据类型
                let energyColCount = 0;
                const sampleCols = timeHeaderIndices.slice(0, Math.min(5, timeHeaderIndices.length));
                for (const colIdx of sampleCols) {
                    const colScore = columnScores.find(s => s.index === colIdx);
                    if (colScore && colScore.columnType === 'energy') {
                        energyColCount++;
                    }
                }
                // 如果超过一半的采样列被识别为累积电能，则整体设为累积电能
                if (energyColCount >= sampleCols.length / 2) {
                    console.log('时间格式列被识别为累积电能数据');
                    // 强制将所有时间格式列标记为energy类型
                    timeHeaderIndices.forEach(idx => {
                        const cs = columnScores.find(s => s.index === idx);
                        if (cs) cs.columnType = 'energy';
                    });
                }
            } else {
                // 纵向格式或其他格式：使用原来的逻辑
                const dataCandidates = columnScores.filter(s => s.index !== dateColumnIndex && s.dataScore > 10);
                if (dataCandidates.length > 0) {
                    dataCandidates.sort((a, b) => {
                        const pA = a.columnType === 'power' ? 3 : (a.columnType === 'energy' ? 2 : 1);
                        const pB = b.columnType === 'power' ? 3 : (b.columnType === 'energy' ? 2 : 1);
                        return pA !== pB ? pB - pA : b.dataScore - a.dataScore;
                    });
                    dataStartColumnIndex = dataCandidates[0].index;
                    
                    // 查找连续的数据列作为结束列
                    let lastIdx = dataStartColumnIndex;
                    for (let i = dataStartColumnIndex + 1; i < headerRow.length; i++) {
                        const s = columnScores[i];
                        if (s.dataScore > 8) lastIdx = i;
                        else if (i - lastIdx > 2) break; // 允许最多2列间隔
                    }
                    dataEndColumnIndex = lastIdx;
                }
            }
            
            meteringPointColumnIndex = getBest([...columnScores].filter(s => ![dateColumnIndex, dataStartColumnIndex].includes(s.index)), 'meteringPointScore');
            multiplierColumnIndex = getBest([...columnScores].filter(s => s.multiplierScore > 10), 'multiplierScore');
            
            // 自动识别 PT 和 CT 列
            // 注意：PT和CT列通常是数据列的一部分，所以只排除日期列
            const ptColumnIndex = getBest([...columnScores].filter(s => s.ptScore > 0 && s.index !== dateColumnIndex), 'ptScore');
            const ctColumnIndex = getBest([...columnScores].filter(s => s.ctScore > 0 && s.index !== dateColumnIndex), 'ctScore');
            
            // 调试日志：输出 PT/CT 识别结果
            console.log('[PT/CT识别] 结果:', {
                ptColumnIndex,
                ctColumnIndex,
                ptScores: columnScores.filter(s => s.ptScore > 0).map(s => ({index: s.index, score: s.ptScore, header: headerRow[s.index]})).slice(0, 3),
                ctScores: columnScores.filter(s => s.ctScore > 0).map(s => ({index: s.index, score: s.ctScore, header: headerRow[s.index]})).slice(0, 3),
                excludedIndices: [dateColumnIndex, dataStartColumnIndex, dataEndColumnIndex]
            });
            
            const timeColumnIndex = detectTimeColumnIndex(headerRow, data, dateColumnIndex);
            if (timeColumnIndex !== -1) {
                appData.config.timeColumn = timeColumnIndex;
            }

            // 如果是列对列模式，强制结束列等于起始列
            if (appData.config.dataStructure === 'columnToColumn') {
                dataEndColumnIndex = dataStartColumnIndex;
            }

            // 自动识别日期格式
            if (dateColumnIndex !== -1) {
                const dateValues = sampleRows.map(row => String(row[dateColumnIndex]).trim()).filter(v => v && v !== 'undefined');
                if (dateValues.length > 0) {
                    const formats = [
                        { name: 'YYYY-MM-DD', regex: /^\d{4}-\d{1,2}-\d{1,2}/ },
                        { name: 'YYYY/MM/DD', regex: /^\d{4}\/\d{1,2}\/\d{1,2}/ },
                        { name: 'YYYY年MM月DD日', regex: /^\d{4}年\d{1,2}月\d{1,2}日/ }
                    ];
                    
                    let detectedFormat = null;
                    for (const fmt of formats) {
                        if (dateValues.every(val => fmt.regex.test(val))) {
                            detectedFormat = fmt.name;
                            break;
                        }
                    }
                    
                    if (detectedFormat) {
                        appData.config.dateFormat = detectedFormat;
                        console.log('自动检测到日期格式:', detectedFormat);
                    }
                }
            }

            // 自动识别数据类型
            let detectedDataType = 'instantPower';
            if (dataStartColumnIndex !== -1) {
                const startCol = columnScores.find(s => s.index === dataStartColumnIndex);
                if (startCol && startCol.columnType === 'energy') detectedDataType = 'cumulativeEnergy';
            }

            // 应用配置并触发反馈
            const configs = [
                { key: 'dateColumn', val: dateColumnIndex, el: elements.dateColumnSelect },
                { key: 'dataStartColumn', val: dataStartColumnIndex, el: elements.dataStartColumnSelect },
                { key: 'dataEndColumn', val: dataEndColumnIndex, el: elements.dataEndColumnSelect },
                { key: 'multiplierColumn', val: multiplierColumnIndex, el: elements.multiplierColumnSelect }
            ];

            configs.forEach(c => {
                if (c.val !== -1 && c.el) {
                    c.el.value = c.val;
                    appData.config[c.key] = c.val;
                    c.el.classList.add('ring-2', 'ring-indigo-400', 'bg-indigo-50/50');
                    setTimeout(() => c.el.classList.remove('ring-2', 'ring-indigo-400', 'bg-indigo-50/50'), 2000);
                }
            });

            if (elements.meteringPointColumnSelect) {
                elements.meteringPointColumnSelect.value = '';
            }
            appData.config.meteringPointColumn = '';

            // 如果识别到倍率列，自动切换到倍率列选项
            if (multiplierColumnIndex !== -1 && elements.useMultiplierColumnCheckbox) {
                elements.useMultiplierColumnCheckbox.checked = true;
                appData.config.useMultiplierColumn = true;
                appData.config.multiplierMode = 'single';
                // 触发 UI 切换
                if (elements.multiplierOptions) elements.multiplierOptions.classList.add('hidden');
                if (elements.multiplierColumnOption) elements.multiplierColumnOption.classList.remove('hidden');
                if (elements.multiplierModeSelector) {
                    elements.multiplierModeSelector.classList.remove('hidden');
                }
                // 切换到单列模式
                setMultiplierMode('single');
            }

            // 如果识别到 PT 和 CT 列，自动切换到 PT/CT 模式
            if (ptColumnIndex !== -1 && ctColumnIndex !== -1 && elements.useMultiplierColumnCheckbox) {
                elements.useMultiplierColumnCheckbox.checked = true;
                appData.config.useMultiplierColumn = true;
                appData.config.multiplierMode = 'ptct';

                // 触发 UI 切换
                if (elements.multiplierOptions) elements.multiplierOptions.classList.add('hidden');
                if (elements.multiplierColumnOption) elements.multiplierColumnOption.classList.remove('hidden');
                if (elements.multiplierModeSelector) {
                    elements.multiplierModeSelector.classList.remove('hidden');
                }

                // 切换到 PT/CT 模式
                setMultiplierMode('ptct');

                // 设置 PT 和 CT 列
                if (elements.ptColumnSelect) {
                    elements.ptColumnSelect.value = ptColumnIndex;
                    appData.config.ptColumn = ptColumnIndex;
                }
                if (elements.ctColumnSelect) {
                    elements.ctColumnSelect.value = ctColumnIndex;
                    appData.config.ctColumn = ctColumnIndex;
                }

                showNotification('倍率识别', `自动识别到 PT 列和 CT 列，已切换为 PT×CT 模式`, 'info');

                // 更新预览
                updatePtCtPreview();
            }
            
            // 如果既没有检测到倍率列，也没有检测到PT/CT列，提示用户使用默认倍率
            if (multiplierColumnIndex === -1 && (ptColumnIndex === -1 || ctColumnIndex === -1)) {
                // 检查是否有任何列可能包含倍率信息（根据数据特征）
                const potentialMultiplierCols = columnScores.filter(s => {
                    // 排除已识别的日期列和数据列
                    if ([dateColumnIndex, dataStartColumnIndex, dataEndColumnIndex].includes(s.index)) return false;
                    // 检查是否为数值列且数值较小（可能是倍率）
                    const colData = sampleRows.map(row => parseFloat(row[s.index])).filter(v => !isNaN(v));
                    if (colData.length === 0) return false;
                    const avgVal = colData.reduce((a, b) => a + b, 0) / colData.length;
                    // 倍率通常在 0.1 - 10000 之间
                    return avgVal > 0.1 && avgVal < 10000 && colData.every(v => v > 0);
                });
                
                if (potentialMultiplierCols.length > 0) {
                    // 找到可能的倍率列，提示用户确认
                    const bestCol = potentialMultiplierCols[0];
                    showNotification('倍率提示', `检测到可能的倍率列（列 ${getColumnLabel(bestCol.index)}），请手动选择或输入默认倍率 1.0`, 'warning');
                } else {
                    // 完全没有检测到倍率信息，使用默认值
                    showNotification('倍率提示', '未检测到倍率列，将使用默认倍率 1.0，如需修改请手动配置', 'info');
                    // 确保固定倍率输入框显示默认值
                    const multiplierInput = document.getElementById('multiplier');
                    if (multiplierInput && !multiplierInput.value) {
                        multiplierInput.value = '1.0';
                    }
                }
            }

            if (detectedDataType) {
                const radio = document.querySelector(`input[name="dataType"][value="${detectedDataType}"]`);
                if (radio) {
                    radio.checked = true;
                    appData.config.dataType = detectedDataType;
                    radio.closest('label')?.classList.add('ring-2', 'ring-indigo-500', 'ring-offset-2');
                    setTimeout(() => radio.closest('label')?.classList.remove('ring-2', 'ring-indigo-500', 'ring-offset-2'), 2000);
                }
            }

            if (dateColumnIndex !== -1 && dataStartColumnIndex !== -1) {
                showNotification('自动识别', '已根据表头和数据样本自动匹配列', 'success');
                updateDataPreview();
                checkConfigValidity();
            }
        }

        // 配置部分函数
        // 重新设计的数据预览函数
        function updateDataPreview() {
            console.log('=== 开始数据预览更新 ===');
            
            // 防抖/防重入：避免大数据下多次触发导致阻塞
            if (appData.__isUpdatingPreview) {
                console.log('数据预览更新进行中，本次调用已忽略');
                return;
            }
            
            // 检查基础配置
            if (!validateDataConfiguration()) {
                console.log('配置验证未通过，进入等待状态');
                renderWaitingState();
                return;
            }
            
            // 先展示加载提示，避免长时间空白
            if (elements.dataPreview) {
                elements.dataPreview.classList.remove('bg-slate-50/50', 'border-2', 'border-dashed', 'border-slate-200');
                elements.dataPreview.innerHTML = `
                    <div class="flex flex-col items-center justify-center py-12 px-6">
                        <div class="w-12 h-12 mb-3 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin"></div>
                        <div class="text-sm text-slate-600 font-medium">正在解析与处理数据...</div>
                        <div class="text-xs text-slate-400 mt-2">大型数据集可能需要几秒钟时间</div>
                    </div>`;
            }
            
            // 异步调度重型解析与渲染，释放UI线程
            appData.__isUpdatingPreview = true;
            
            // 设置一个安全超时，防止某些情况下解析挂起导致一直转圈
            const watchdog = setTimeout(() => {
                if (appData.__isUpdatingPreview) {
                    console.error('数据预览更新超时（可能由于数据量过大或逻辑阻塞）');
                    appData.__isUpdatingPreview = false;
                    renderErrorState(new Error('数据处理超时，请尝试缩小数据范围或检查文件格式'));
                }
            }, 30000); // 30秒超时
            
            setTimeout(async () => {
                const startTime = performance.now();
                try {
                    // 解析工作表数据
                    console.time('parseWorksheetData');
                    await parseWorksheetData();
                    console.timeEnd('parseWorksheetData');

                    // 处理解析后的数据
                    console.time('processParsedData');
                    await processParsedData();
                    console.timeEnd('processParsedData');
                    
                    // 检查解析后的数据
                    if (!appData.processedData || appData.processedData.length === 0) {
                        console.warn('解析完成，但没有生成有效数据');
                        renderNoDataState();
                        return;
                    }
                    
                    // 渲染数据预览界面
                    console.time('renderDataPreview');
                    renderDataPreview();
                    console.timeEnd('renderDataPreview');
                    
                    // 触发数据预览入场动画
                    if (elements.dataPreview) {
                        elements.dataPreview.classList.remove('animate-graph-entry');
                        void elements.dataPreview.offsetWidth; // 触发重绘
                        elements.dataPreview.classList.add('animate-graph-entry');
                    }
                    
                    // 更新按钮状态
                    updatePreviewButtons();

                    try {
                        // 使用 getOverviewData 确保显示全局统计，而不是初始筛选的统计
                        updateGlobalSummary(getOverviewData(), appData.visualization.focusStartTime, appData.visualization.focusEndTime);
                    } catch (e) {
                        console.error('更新初始全局统计失败:', e);
                    }
                    
                    const endTime = performance.now();
                    console.log(`=== 数据预览更新完成，总耗时: ${(endTime - startTime).toFixed(2)}ms ===`);
                } catch (error) {
                    console.error('数据预览更新过程出错:', error);
                    renderErrorState(error);
                } finally {
                    clearTimeout(watchdog);
                    appData.__isUpdatingPreview = false;
                }
            }, 50); // 稍微延迟一下，确保UI先显示加载状态
        }

        // 渲染错误状态
        function renderErrorState(error) {
            if (!elements.dataPreview) return;
            
            const errorMessage = error.message || '未知错误';
            const errorStack = error.stack ? `<details class="mt-4 text-left w-full"><summary class="text-xs cursor-pointer opacity-50">错误详情</summary><pre class="text-xs mt-2 p-2 bg-indigo-50/50 rounded border border-indigo-100/30 overflow-auto max-h-32 text-indigo-900/70 font-mono">${error.stack}</pre></details>` : '';
            
            elements.dataPreview.innerHTML = `
                <div class="flex flex-col items-center justify-center py-16 px-8">
                    <div class="w-20 h-20 mb-6 bg-indigo-50 rounded-2xl flex items-center justify-center shadow-sm ring-1 ring-indigo-100">
                        <svg class="w-10 h-10 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                        </svg>
                    </div>
                    <h3 class="text-xl font-bold text-slate-800 mb-3">数据处理异常</h3>
                    <p class="text-sm text-slate-600 text-center max-w-lg leading-relaxed mb-6">
                        在解析或处理您的数据时发生了错误。这通常是由于文件格式不匹配或配置参数（如日期列、数据列）设置不当引起的。<br>
                        <span class="text-indigo-600 font-bold">错误原因：${errorMessage}</span>
                    </p>
                    <div class="flex space-x-3">
                        <button onclick="updateDataPreview()" class="px-6 py-2.5 bg-indigo-600 text-white text-sm font-bold rounded-xl hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200 active:scale-95">
                            重新尝试
                        </button>
                        <button onclick="location.reload()" class="px-6 py-2.5 bg-white text-slate-600 text-sm font-bold rounded-xl border border-slate-200 hover:bg-slate-50 transition-all active:scale-95">
                            刷新页面
                        </button>
                    </div>
                    ${errorStack}
                </div>
            `;
        }
        
        // 验证数据配置
        function validateDataConfiguration() {
            const hasWorksheets = Array.isArray(appData.worksheets) && appData.worksheets.length > 0;
            const hasDateColumn = appData.config.dateColumn !== undefined && appData.config.dateColumn !== null && appData.config.dateColumn !== '';
            const hasDataStartColumn = appData.config.dataStartColumn !== undefined && appData.config.dataStartColumn !== null && appData.config.dataStartColumn !== '';
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value;
            
            // 对于列对行结构，需要数据结束列；对于列对列结构，不需要
            let hasRequiredColumns = true;
            if (dataStructure === 'columnToRow') {
                hasRequiredColumns = (appData.config.dataEndColumn !== undefined && appData.config.dataEndColumn !== null && appData.config.dataEndColumn !== '');
            }
            
            console.log('验证配置状态:');
            console.log('- 工作表:', hasWorksheets);
            console.log('- 日期列:', hasDateColumn);
            console.log('- 数据开始列:', hasDataStartColumn);
            console.log('- 数据结构:', dataStructure);
            console.log('- 必需列配置:', hasRequiredColumns);
            
            return hasWorksheets && hasDateColumn && hasDataStartColumn && hasRequiredColumns;
        }
        
        // 渲染等待配置状态
        function renderWaitingState() {
            if (!elements.dataPreview) return;
            
            // 恢复等待状态样式
            elements.dataPreview.classList.add('bg-slate-50/50', 'border-2', 'border-dashed', 'border-slate-200');
            
            elements.dataPreview.innerHTML = `
                <div class="flex flex-col items-center justify-center py-16 px-8 text-center">
                    <div class="w-20 h-20 mb-6 bg-indigo-50 rounded-2xl flex items-center justify-center shadow-sm ring-1 ring-indigo-100">
                        <svg class="w-10 h-10 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                    </div>
                    <h3 class="text-xl font-bold text-slate-800 mb-3">等待配置完成</h3>
                    <p class="text-sm text-slate-600 text-center max-w-lg leading-relaxed mb-6">
                        请完成以下步骤以开始数据预览：<br>
                        1. 导入 Excel 文件<br>
                        2. 在“字段映射”中配置日期列和数据开始列<br>
                        3. 如选择“宽表”结构，还需配置数据结束列
                    </p>
                    <div class="flex space-x-2">
                        <div class="w-3 h-3 bg-indigo-400 rounded-full animate-bounce" style="animation-delay: 0ms"></div>
                        <div class="w-3 h-3 bg-indigo-500 rounded-full animate-bounce" style="animation-delay: 150ms"></div>
                        <div class="w-3 h-3 bg-indigo-600 rounded-full animate-bounce" style="animation-delay: 300ms"></div>
                    </div>
                </div>
            `;
        }
        
        // 渲染无数据状态
        function renderNoDataState() {
            elements.dataPreview.innerHTML = `
                <div class="flex flex-col items-center justify-center py-16 px-8">
                    <div class="w-20 h-20 mb-6 bg-indigo-50 rounded-2xl flex items-center justify-center shadow-sm ring-1 ring-indigo-100">
                        <svg class="w-10 h-10 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z"></path>
                        </svg>
                    </div>
                    <h3 class="text-xl font-bold text-slate-800 mb-3">暂无有效数据</h3>
                    <p class="text-sm text-slate-600 text-center max-w-lg leading-relaxed">
                        当前配置下没有找到可预览的数据。<br>
                        请检查数据源文件或调整配置参数。
                    </p>
                </div>
            `;
        }

        // 渲染数据预览界面
        function renderDataPreview() {
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
            
            // 移除等待状态样式并显示按钮
            if (elements.dataPreview) {
                elements.dataPreview.classList.remove('bg-white', 'border', 'border-slate-200', 'shadow-sm');
                if (elements.previewButtons) {
                    elements.previewButtons.classList.remove('hidden');
                }
            }
            
            // 构建新的预览界面 - 扁平化设计
            let previewHtml = `
                <div class="flex flex-col gap-6 p-6">
                    <!-- 头部摘要 -->
                    ${buildDataSummaryHeader()}
                    
                    <!-- 数据表格容器 -->
                    <div class="overflow-hidden rounded-2xl border border-slate-200/60 bg-white shadow-sm">
                        ${buildDataTable(dataStructure)}
                    </div>
                </div>
            `;
            
            elements.dataPreview.innerHTML = previewHtml;
            
            // 添加交互功能
            attachDataPreviewEvents();
        }
        
        // 构建数据汇总头部
        function buildDataSummaryHeader() {
            const totalRecords = appData.processedData.length;
            const dateRange = getDataDateRange();
            const interval = appData.config.timeInterval || '--';
            
            // 更新页面顶部的 headerDateRange 和 headerInterval
            const headerDateRangeEl = document.getElementById('headerDateRange');
            const headerIntervalEl = document.getElementById('headerInterval');
            
            if (headerDateRangeEl) headerDateRangeEl.textContent = dateRange;
            if (headerIntervalEl) {
                const intervalText = interval !== '--' ? `${interval} 分钟` : '--';
                headerIntervalEl.textContent = intervalText;
            }
            
            return `
                <div class="flex flex-col gap-6">
                    <!-- 顶部栏：标题 + 状态 -->
                    <div class="flex items-center justify-between">
                        <div class="flex items-center gap-4">
                            <div class="flex h-12 w-12 items-center justify-center rounded-2xl bg-indigo-600 text-white shadow-lg shadow-indigo-200">
                                <i class="fa fa-table-list text-xl"></i>
                            </div>
                            <div>
                                <h3 class="text-lg font-black text-slate-900 tracking-tight">数据预览</h3>
                                <p class="text-sm font-medium text-slate-500 mt-0.5">用于快速核对解析结果</p>
                            </div>
                        </div>
                        <div class="inline-flex items-center gap-2 px-3 py-1.5 bg-white/80 rounded-xl border border-white/60 shadow-sm backdrop-blur-sm">
                            <span class="flex h-2 w-2 rounded-full bg-green-500 animate-pulse"></span>
                            <span class="text-[11px] font-black text-slate-600 uppercase tracking-wider">数据解析就绪</span>
                        </div>
                    </div>

                    <!-- 统计卡片区 -->
                    <div class="grid grid-cols-1 gap-4 md:grid-cols-3">
                        <!-- 总记录数 -->
                        <div class="rounded-2xl bg-white border border-slate-200 p-5 shadow-sm transition-all hover:shadow-md">
                            <div class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">总记录数</div>
                            <div class="text-2xl font-black text-slate-900 tabular-nums">${totalRecords.toLocaleString()} <span class="text-sm font-bold text-slate-400 ml-1">条</span></div>
                        </div>
                        <!-- 时间范围 -->
                        <div class="rounded-2xl bg-white border border-slate-200 p-5 shadow-sm transition-all hover:shadow-md">
                            <div class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">时间范围</div>
                            <div class="text-lg font-extrabold text-slate-900 truncate" title="${dateRange}">${dateRange}</div>
                        </div>
                        <!-- 数据结构 -->
                        <div class="rounded-2xl bg-white border border-slate-200 p-5 shadow-sm transition-all hover:shadow-md">
                            <div class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">数据结构</div>
                            <div class="text-lg font-extrabold text-slate-900">${getDataStructureLabel()}</div>
                        </div>
                    </div>
                </div>`;
        }
        
        // 构建数据表格
         function buildDataTable(dataStructure) {
             let tableHtml = `<div class="max-h-[420px] overflow-auto bg-white/50 backdrop-blur-sm">`;
             
             if (dataStructure === 'columnToRow') {
                 tableHtml += buildHourlyDataTable();
             } else {
                 tableHtml += buildSummaryDataTable();
             }
             
             tableHtml += `</div>`;
             return tableHtml;
         }
         
         // 构建24小时数据表格
         function buildHourlyDataTable() {
             let tableHtml = `<div class="grid" style="grid-template-columns: 140px repeat(24, 84px) 140px; min-width: fit-content; content-visibility: auto;">`;

             // 表头
             tableHtml += `<div class=\"bg-slate-50/90 border-b border-slate-200 px-4 py-4 text-left text-xs font-extrabold text-slate-500 sticky top-0 left-0 z-30 backdrop-blur-sm\">日期</div>`;
             for (let hour = 0; hour < 24; hour++) {
                tableHtml += `<div class=\"bg-slate-50/90 border-b border-slate-200 px-2 py-4 text-center text-xs font-extrabold text-slate-500 sticky top-0 z-20 backdrop-blur-sm\">${hour}:00</div>`;
            }
             tableHtml += `<div class=\"bg-slate-50/90 border-b border-slate-200 px-3 py-4 text-center text-xs font-extrabold text-slate-500 sticky top-0 z-20 backdrop-blur-sm\">日总 (kWh)</div>`;

             const totalRows = appData.processedData.length;
             const headLimit = 25;
             const tailLimit = 25;
             const showEllipsis = totalRows > headLimit + tailLimit;

            // 渲染单行数据的辅助函数
            const renderRow = (row, index, rowIndex, isSummary = false) => {
                const isCumulative = isSummary && row.date === '累计汇总';
                const rowBgClass = isSummary 
                    ? (isCumulative ? 'bg-slate-100/95 font-bold' : 'bg-slate-50/95 font-bold') 
                    : (rowIndex % 2 === 0 ? 'bg-white' : 'bg-slate-50/50');
                
                const borderClass = isSummary ? 'border-t-2 border-slate-200' : 'border-b border-slate-100';
                const stickyClass = isSummary ? 'sticky bottom-0 z-40 backdrop-blur-md' : '';
                const stickyOffset = isSummary && row.date === '平均数值' ? 'bottom-0' : (isSummary ? 'bottom-[37px]' : '');
                let rowHtml = '';

                // 日期列 (冻结首列)
                rowHtml += `<div class="${rowBgClass} ${borderClass} px-4 py-3.5 text-left text-xs font-bold text-slate-700 sticky left-0 z-30 ${stickyClass} ${stickyOffset}" data-date="${row.date || ''}">${row.date}</div>`;

                // 24小时数据列
                let dayTotal = 0;
                for (let hour = 0; hour < 24; hour++) {
                    const value = row.hourlyData[hour];
                    const displayValue = (value !== null && value !== undefined) ? value.toFixed(2) : '-';
                    const textClass = isSummary ? 'text-slate-900' : ((value !== null && value !== undefined) ? 'text-slate-700' : 'text-slate-300');
                    
                    rowHtml += `<div class="${rowBgClass} ${borderClass} px-2 py-3.5 text-center text-xs font-medium ${textClass} ${stickyClass} ${stickyOffset}">${displayValue}</div>`;
                    if (value !== null && value !== undefined) {
                        dayTotal += value;
                    }
                }

                // 日总电能列
                const totalValue = row.total !== undefined ? row.total : dayTotal;
                const totalTextClass = isSummary ? 'text-indigo-700' : 'text-indigo-600';
                rowHtml += `<div class="${rowBgClass} ${borderClass} px-3 py-3.5 text-center text-xs font-bold ${totalTextClass} ${stickyClass} ${stickyOffset}">${totalValue.toFixed(2)}</div>`;
                return rowHtml;
            };

            // 计算汇总和平均
            const calculateSummaryRows = () => {
                const hourlySums = new Array(24).fill(0);
                const hourlyCounts = new Array(24).fill(0);
                let grandTotal = 0;

                appData.processedData.forEach(row => {
                    for (let hour = 0; hour < 24; hour++) {
                        const val = row.hourlyData[hour];
                        if (val !== null && val !== undefined) {
                            hourlySums[hour] += val;
                            hourlyCounts[hour]++;
                            grandTotal += val;
                        }
                    }
                });

                const hourlyAverages = hourlySums.map((sum, i) => hourlyCounts[i] > 0 ? sum / hourlyCounts[i] : null);
                const avgTotal = grandTotal / appData.processedData.length;

                return {
                    sumRow: {
                        date: '累计汇总',
                        hourlyData: hourlySums,
                        total: grandTotal
                    },
                    avgRow: {
                        date: '平均数值',
                        hourlyData: hourlyAverages,
                        total: avgTotal
                    }
                };
            };

            const { sumRow, avgRow } = calculateSummaryRows();

            if (showEllipsis) {
                // 显示前25行
                for (let i = 0; i < headLimit; i++) {
                    tableHtml += renderRow(appData.processedData[i], i, i);
                }

                // 显示省略行
                tableHtml += `<div class="bg-slate-100 border-b border-slate-200 px-4 py-3 text-left text-xs font-extrabold text-slate-500 sticky left-0 z-10" style="grid-column: 1 / -1;">... 省略 ${totalRows - headLimit - tailLimit} 行数据 ...</div>`;

                // 显示后25行
                for (let i = totalRows - tailLimit; i < totalRows; i++) {
                    tableHtml += renderRow(appData.processedData[i], i, i);
                }
            } else {
                // 数据少于50行，全部显示
                for (let i = 0; i < totalRows; i++) {
                    tableHtml += renderRow(appData.processedData[i], i, i);
                }
            }

            // 添加汇总和平均行
            tableHtml += renderRow(sumRow, -1, totalRows, true);
            tableHtml += renderRow(avgRow, -2, totalRows + 1, true);

            tableHtml += `</div>`;

             // 数据提示
             if (totalRows > headLimit + tailLimit) {
                tableHtml += `<div class=\"px-4 py-2 text-xs text-slate-500 bg-slate-50 border border-t-0 border-slate-200\">预览显示前 ${headLimit} 行和后 ${tailLimit} 行（共 ${totalRows} 条记录）。</div>`;
             } else {
                tableHtml += `<div class=\"px-4 py-2 text-xs text-slate-500 bg-slate-50 border border-t-0 border-slate-200\">共 ${totalRows} 条记录。</div>`;
             }
             return tableHtml;
         }
         
         // 构建汇总数据表格
         function buildSummaryDataTable() {
             // 统一使用24小时详细数据表格，而不是汇总表格
             return buildHourlyDataTable();
         }

         // 构建原始数据结构预览
         function buildRawDataPreview() {
             const container = document.getElementById('rawDataPreviewContainer');
             if (!container) return;

             // 获取原始工作表数据
             if (!appData.worksheets || appData.worksheets.length === 0) {
                 container.innerHTML = '';
                 return;
             }

             const worksheet = appData.worksheets[0];
            const rawData = worksheet.data || [];
            const totalRows = rawData.length;
            const originalColumnCount = rawData[0] ? rawData[0].length : 0;
            
            // 性能优化：限制预览列数，避免 DOM 节点过多导致卡顿
            const MAX_PREVIEW_COLUMNS = 40;
            const needsColumnTruncation = originalColumnCount > MAX_PREVIEW_COLUMNS;
            
            const getPreviewIndices = () => {
                if (!needsColumnTruncation) return Array.from({length: originalColumnCount}, (_, i) => i);
                // 显示前 20 列和最后 20 列
                const indices = [];
                for (let i = 0; i < 20; i++) indices.push(i);
                indices.push(-1); // 标记省略号
                for (let i = originalColumnCount - 20; i < originalColumnCount; i++) indices.push(i);
                return indices;
            };
            
            const previewIndices = getPreviewIndices();
            const headLimit = 25;
            const tailLimit = 25;
            const showEllipsis = totalRows > headLimit + tailLimit;

             let previewHtml = `
                 <div class="p-4 bg-white border border-green-200/60 rounded-2xl shadow-sm">
                    <div class="flex items-center gap-2 mb-3">
                        <div class="flex h-7 w-7 items-center justify-center rounded-lg bg-green-50 text-green-600">
                            <i class="fa fa-table-list text-xs"></i>
                        </div>
                        <h4 class="text-xs font-bold uppercase tracking-wider text-green-700">原始数据结构预览</h4>
                        <span class="text-[10px] text-slate-400 font-medium">(${totalRows} 行 × ${originalColumnCount} 列)</span>
                    </div>
                    <div class="overflow-x-auto overflow-y-auto max-h-[300px] border border-slate-200 border-b-0">
                        <table class="w-full text-xs">
                            <thead class="bg-slate-50 sticky top-0">
                                <tr>
                                    <th class="px-3 py-2 text-left font-bold text-slate-600 border-b border-r border-slate-200 w-16">行号</th>
                                    ${previewIndices.map(idx => {
                                        if (idx === -1) return `<th class="px-2 py-2 text-center font-bold text-slate-400 border-b border-r border-slate-200 italic">...</th>`;
                                        return `<th class="px-3 py-2 text-left font-bold text-slate-600 border-b border-r border-slate-200 min-w-[80px]">${getColumnLabel(idx)}</th>`;
                                    }).join('')}
                                </tr>
                            </thead>
                             <tbody>
             `;

             // 渲染行数据的辅助函数
             const renderRow = (row, rowIndex) => {
                 const rowBgClass = rowIndex % 2 === 0 ? 'bg-white' : 'bg-slate-50';
                 return `
                     <tr class="${rowBgClass}">
                         <td class="px-3 py-2 text-slate-400 border-b border-r border-slate-200 font-mono">${rowIndex + 1}</td>
                         ${previewIndices.map(idx => {
                             if (idx === -1) return `<td class="px-2 py-2 text-center text-slate-300 border-b border-r border-slate-200 italic">...</td>`;
                             const cell = row[idx];
                             return `<td class="px-3 py-2 text-slate-700 border-b border-r border-slate-200 truncate max-w-[150px]" title="${cell}">${cell !== undefined && cell !== null ? cell : ''}</td>`;
                         }).join('')}
                     </tr>
                 `;
             };

             if (showEllipsis) {
                 // 显示前25行
                 for (let i = 0; i < headLimit; i++) {
                     previewHtml += renderRow(rawData[i], i);
                 }

                 // 显示省略行
                 const displayColCount = previewIndices.length + 1;
                 previewHtml += `
                     <tr class="bg-slate-100">
                         <td class="px-3 py-2 text-slate-500 border-b border-r border-slate-200 font-bold text-center" colspan="${displayColCount}">
                             ... 省略 ${totalRows - headLimit - tailLimit} 行数据 ...
                         </td>
                     </tr>
                 `;

                 // 显示后25行
                 for (let i = totalRows - tailLimit; i < totalRows; i++) {
                     previewHtml += renderRow(rawData[i], i);
                 }
             } else {
                 // 数据少于50行，全部显示
                 for (let i = 0; i < totalRows; i++) {
                     previewHtml += renderRow(rawData[i], i);
                 }
             }

             previewHtml += `
                            </tbody>
                        </table>
                    </div>
                    <div class="h-px bg-slate-200"></div>
                    ${needsColumnTruncation ? `
                    <div class="px-4 py-2 bg-indigo-50/50 text-indigo-700 text-[10px] border-t border-indigo-100/50 flex items-center gap-2 font-medium">
                        <i class="fa fa-info-circle"></i>
                        表格列数较多 (${originalColumnCount} 列)，预览仅展示前 20 列和最后 20 列以保证性能。
                    </div>
                    ` : ''}
                </div>
            `;

             container.innerHTML = previewHtml;
         }
         
         // 构建数据统计信息
          function buildDataStatistics() {
              if (appData.processedData.length === 0) return '';
              
              const totalRecords = appData.processedData.length;
              const validDataCount = appData.processedData.reduce((sum, row) => {
                  return sum + row.hourlyData.filter(v => v !== null && v !== undefined).length;
              }, 0);
              const totalEnergy = appData.processedData.reduce((sum, row) => sum + (row.dailyTotal || 0), 0);
              const avgDailyEnergy = totalRecords > 0 ? totalEnergy / totalRecords : 0;
              
              // 计算最大日负荷
              let maxDailyLoad = 0;
              let maxDailyDate = '-';
              if (appData.processedData.length > 0) {
                  const maxRow = appData.processedData.reduce((max, row) => (row.dailyTotal > max.dailyTotal ? row : max), appData.processedData[0]);
                  maxDailyLoad = maxRow.dailyTotal;
                  maxDailyDate = maxRow.date;
              }
              
              // 占比计算 (这里简单假设占比为 100% 或者基于某种基准，如果用户没有指定基准，可以显示为有效数据占比)
              // 假设占比是指 有效数据点 / (天数 * 24)
              const totalPoints = totalRecords * 24;
              const dataCoverage = totalPoints > 0 ? ((validDataCount / totalPoints) * 100).toFixed(1) : 0;

              return `
                  <div class="p-8 bg-white">
                      <div class="flex items-center justify-center mb-8">
                          <div class="h-px bg-slate-100 flex-1"></div>
                          <div class="mx-6 flex items-center gap-2">
                              <i class="fa fa-chart-line text-indigo-500"></i>
                              <h4 class="text-sm font-black text-slate-900 uppercase tracking-widest">数据统计摘要</h4>
                          </div>
                          <div class="h-px bg-slate-100 flex-1"></div>
                      </div>
                      <div class="grid grid-cols-2 md:grid-cols-4 gap-8">
                          <div class="text-center group">
                              <div class="text-3xl font-black text-indigo-600 group-hover:scale-110 transition-transform duration-300 inline-block">${totalEnergy.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</div>
                              <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mt-1">时段总用电 (kWh)</div>
                          </div>
                          <div class="text-center group">
                              <div class="text-3xl font-black text-slate-900 group-hover:text-indigo-600 transition-colors">${dataCoverage}%</div>
                              <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mt-1">数据完整率</div>
                          </div>
                          <div class="text-center group">
                              <div class="text-3xl font-black text-indigo-500 group-hover:scale-110 transition-transform duration-300 inline-block">${avgDailyEnergy.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</div>
                              <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mt-1">日均电能 (kWh)</div>
                          </div>
                          <div class="text-center group">
                              <div class="text-3xl font-black text-slate-900 group-hover:text-indigo-600 transition-colors" title="${maxDailyDate}">${maxDailyLoad.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 1})}</div>
                              <div class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mt-1">最大日用电 (${maxDailyDate})</div>
                          </div>
                      </div>
                  </div>
              `;
          }
          
          // 获取数据日期范围
          function getDataDateRange() {
              if (appData.processedData.length === 0) return '无数据';
              const firstDate = appData.processedData[0].date;
              const lastDate = appData.processedData[appData.processedData.length - 1].date;
              return firstDate === lastDate ? firstDate : `${firstDate} ~ ${lastDate}`;
          }
          
          // 获取数据结构标签
          function getDataStructureLabel() {
              const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value;
              return dataStructure === 'columnToRow' ? '列对行 (24小时)' : '列对列 (汇总)';
          }
          
          // 添加数据预览交互事件
          function attachDataPreviewEvents() {
              // 日期点击事件 - 可以扩展为选择日期范围
              const dateCells = document.querySelectorAll('[data-date]');
              dateCells.forEach(cell => {
                  cell.addEventListener('click', function() {
                      const date = this.getAttribute('data-date');
                      appData.__previewSelectedDate = date;
                      document.querySelectorAll('[data-date]').forEach(el => {
                          el.classList.remove('bg-indigo-50', 'text-indigo-700', 'ring-2', 'ring-indigo-500', 'ring-inset');
                      });
                      this.classList.add('bg-indigo-50', 'text-indigo-700', 'ring-2', 'ring-indigo-500', 'ring-inset');
                  });
              });
          }
          
          // 更新按钮状态
          function updatePreviewButtons() {
              if (appData.parsedData.length > 0) {
                  elements.previewFirstHourCalculationBtn.disabled = false;
                  elements.previewFirstDayCalculationBtn.disabled = false;
              } else {
                  elements.previewFirstHourCalculationBtn.disabled = true;
                  elements.previewFirstDayCalculationBtn.disabled = true;
              }
          }
        
          // 更新数据状态显示
          updateDataOverview();

      async function parseWorksheetData() {
    // 只在非追加模式下清除数据
    const appendMode = document.getElementById('appendMode')?.checked || false;
    if (!appendMode) {
        appData.parsedData = [];
        appData.processedData = [];
    } else {
        // 追加模式下不清除现有数据，但确保数组存在
        appData.parsedData = appData.parsedData || [];
        appData.processedData = appData.processedData || [];
    }
    
    // 获取数据结构类型
    const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
    
    for (const worksheet of appData.worksheets) {
        const { data } = worksheet;
        const dateCol = parseInt(appData.config.dateColumn);
        const startCol = parseInt(appData.config.dataStartColumn);
        const endCol = parseInt(appData.config.dataEndColumn);
        const headerRow = Array.isArray(data) ? (data[0] || []) : [];
        
        // 解析用户选择的索引范围，优先级：数据行索引 > 日期范围 > 全部数据
        let startIndex = 0;
        let endIndex = data.length - 1;
        
        console.log(`工作表 ${worksheet.name || '未命名'}:`);
        console.log(`- 原始数据总行数: ${data.length}`);
        console.log(`- 数据行索引范围: ${appData.config.dataStartIndex || 0} 到 ${appData.config.dataEndIndex || -1}`);
        
        // 检查是否设置了非默认的数据行索引范围（优先级最高）
        const userStartIndex = parseInt(appData.config.dataStartIndex) || 0;
        const userEndIndex = appData.config.dataEndIndex === -1 ? data.length - 1 : parseInt(appData.config.dataEndIndex);
        
        // 如果用户设置了非默认的索引范围，优先使用
        const isUsingCustomIndexRange = userStartIndex !== 0 || (appData.config.dataEndIndex !== -1 && userEndIndex !== (data.length - 1));
        
        if (isUsingCustomIndexRange) {
            startIndex = Math.max(0, Math.min(userStartIndex, data.length - 1));
            endIndex = Math.max(startIndex, Math.min(userEndIndex, data.length - 1));
            console.log(`- 使用用户指定的行索引范围: ${startIndex} 到 ${endIndex}`);
            
            // 如果索引范围无效，给出警告
            if (startIndex >= data.length || endIndex < 0 || startIndex > endIndex) {
                showNotification('警告', '指定的行索引范围无效，将使用全部数据', 'warning');
                startIndex = 0;
                endIndex = data.length - 1;
            }
        } else {
            // 默认使用全部数据
            startIndex = 0;
            endIndex = data.length - 1;
            console.log(`- 使用默认行索引范围: ${startIndex} 到 ${endIndex}`);
        }

        // 统一计算数据起始行（跳过表头和说明行）
        let startRow = startIndex;
        let foundDataRow = false;

        // 获取时间列配置
        const configuredTimeCol = parseInt(appData.config.timeColumn);
        let timeCol = !isNaN(configuredTimeCol) ? configuredTimeCol : -1;
        
        // 首先尝试通过日期列识别数据起始行
        for (let i = startIndex; i < Math.min(startIndex + 15, data.length); i++) {
            if (data[i] && data[i][dateCol]) {
                const cellValue = String(data[i][dateCol]).trim();
                // 检查是否为有效的日期格式
                if (isDateLike(cellValue)) {
                    startRow = i;
                    foundDataRow = true;
                    console.log(`- 通过日期列识别到数据起始行: ${i}, 值: ${cellValue}`);
                    break;
                }
            }
        }
        
        // 如果没找到，尝试通过时间列识别（对于纵向格式）
        if (!foundDataRow && timeCol !== -1) {
            const timePattern = /^([01]?\d|2[0-3])[:：][0-5]\d$/;
            for (let i = startIndex; i < Math.min(startIndex + 15, data.length); i++) {
                if (data[i] && data[i][timeCol]) {
                    const cellValue = String(data[i][timeCol]).trim();
                    if (timePattern.test(cellValue)) {
                        startRow = i;
                        foundDataRow = true;
                        console.log(`- 通过时间列识别到数据起始行: ${i}, 值: ${cellValue}`);
                        break;
                    }
                }
            }
        }
        
        // 如果没找到，尝试通过数值列识别（检查数据列是否有数值）
        if (!foundDataRow && startCol !== -1) {
            for (let i = startIndex; i < Math.min(startIndex + 15, data.length); i++) {
                if (data[i] && data[i][startCol] !== undefined) {
                    const cellValue = data[i][startCol];
                    // 检查是否为数值（排除表头说明文字）
                    if (cellValue !== null && cellValue !== '' && !isNaN(parseFloat(cellValue))) {
                        startRow = i;
                        foundDataRow = true;
                        console.log(`- 通过数值列识别到数据起始行: ${i}, 值: ${cellValue}`);
                        break;
                    }
                }
            }
        }
        
        console.log(`- 最终数据起始行: ${startRow}`);

        // 检测时间列（使用 startRow 确保只采样数据行）
        const detectedTimeCol = detectTimeColumnIndex(headerRow, data, dateCol, startRow);
        if (timeCol === -1) timeCol = detectedTimeCol;
        if (timeCol === dateCol) timeCol = -1;

        const meteringPointCol = appData.config.meteringPointColumn ? parseInt(appData.config.meteringPointColumn) : null;
        
        // 倍率列处理：支持单列或 PT/CT 分开
        let multiplierCol = null;
        let ptCol = null;
        let ctCol = null;
        
        if (appData.config.useMultiplierColumn) {
            if (appData.config.multiplierMode === 'ptct') {
                // PT/CT 分开模式
                ptCol = appData.config.ptColumn ? parseInt(appData.config.ptColumn) : null;
                ctCol = appData.config.ctColumn ? parseInt(appData.config.ctColumn) : null;
            } else {
                // 单列倍率模式
                multiplierCol = appData.config.multiplierColumn ? parseInt(appData.config.multiplierColumn) : null;
            }
        }
        
        const selectedMeteringPoint = appData.config.meteringPointFilter;
        
        if (dataStructure === 'columnToRow') {
            // 原有的列对行数据处理逻辑
            const dataColumnCount = endCol - startCol + 1;
            
            // 检查是否启用了跨日期结构数据选项
            const isCrossDateStructure = document.getElementById('crossDateStructure') && document.getElementById('crossDateStructure').checked;
            
            // 增强的智能时间间隔检测
            const detectionResult = detectTimeIntervalIntelligently(data, startIndex, endIndex, dateCol, meteringPointCol, selectedMeteringPoint, isCrossDateStructure, dataColumnCount);
            
            let autoDetectedInterval = detectionResult.interval;
            let intervalDescription = detectionResult.description;
            
            // 显示增强的检测结果通知
            showEnhancedIntervalNotification(detectionResult);
            
            // 如果有候选间隔，显示候选信息
            if (detectionResult.candidates && detectionResult.candidates.length > 0) {
                const candidateInfo = detectionResult.candidates
                    .map(c => `${c.description}(${Math.round(c.confidence * 100)}%)`)
                    .join('、');
                setTimeout(() => {
                    showNotification('🔍 其他候选间隔', `检测到的其他可能间隔：${candidateInfo}\n如果主要检测结果不准确，可手动选择这些候选间隔`, 'info');
                }, 2500);
            }
            
            // 更新时间间隔配置
            appData.config.timeInterval = autoDetectedInterval;
            elements.timeIntervalSelect.value = autoDetectedInterval;
            

            
            // 处理每一行数据
            const targetIndices = [];
            const maxIdx = Math.min(endIndex, data.length - 1);
            for (let k = startRow; k <= maxIdx; k++) targetIndices.push(k);
            
            await processChunksAsync(targetIndices, (i) => {
                const row = data[i];
                if (!row || !row[dateCol]) return;
                
                // 检查计量点编号筛选
                if (meteringPointCol !== null && selectedMeteringPoint) {
                    const rowMeteringPoint = String(row[meteringPointCol]).trim();
                    if (rowMeteringPoint !== selectedMeteringPoint) {
                        return; // 跳过不匹配的行
                    }
                }
                
                // 检查附加筛选项
                let skipRow = false;
                appData.config.additionalFilters.forEach(filter => {
                    if (filter.column && filter.value) {
                        const col = parseInt(filter.column);
                        const rowValue = String(row[col]).trim();
                        if (rowValue !== filter.value) {
                            skipRow = true;
                        }
                    }
                });
                
                if (skipRow) {
                    return; // 跳过不匹配的行
                }
                
                // 解析日期
                const dateStr = String(row[dateCol]).trim();
                const date = parseDate(dateStr, appData.config.dateFormat);
                if (!date) return;
                

                // 获取倍率值（支持单列或 PT/CT 分开）
                let multiplier = appData.config.multiplier;
                
                if (appData.config.multiplierMode === 'ptct' && (ptCol !== null || ctCol !== null)) {
                    // PT/CT 分开模式：倍率 = PT × CT
                    const ptVal = ptCol !== null && row[ptCol] !== undefined ? parseNumber(row[ptCol]) : 1;
                    const ctVal = ctCol !== null && row[ctCol] !== undefined ? parseNumber(row[ctCol]) : 1;
                    if (ptVal !== null && ctVal !== null) {
                        multiplier = ptVal * ctVal;
                    }
                } else if (multiplierCol !== null && row[multiplierCol] !== undefined) {
                    // 单列倍率模式
                    const parsedMultiplier = parseNumber(row[multiplierCol]);
                    if (parsedMultiplier !== null) {
                        multiplier = parsedMultiplier;
                    }
                }
                
                // 提取数据列
                const dataValues = [];
                for (let j = startCol; j <= endCol && j < row.length; j++) {
                    const value = parseNumber(row[j]);
                    dataValues.push(value);
                }
                
                // 验证数据点数量
                const expectedPoints = dataColumnCount; // 使用计算出的数据列数量
                if (dataValues.length < expectedPoints * 0.8) { // 允许20%的数据缺失
                    console.warn(`数据点不足，日期: ${dateStr}, 实际: ${dataValues.length}, 预期: ${expectedPoints}`);
                    return;
                }
                
                // 添加到解析数据
                appData.parsedData.push({
                    date: date,
                    dateStr: dateStr,
                    rawData: dataValues,
                    multiplier: multiplier,
                    sourceFile: worksheet.file.name,
                    meteringPoint: meteringPointCol !== null ? String(row[meteringPointCol]).trim() : null,
                    timeInterval: autoDetectedInterval, // 保存自动检测到的时间间隔
                    dataStructure: 'columnToRow' // 标记数据结构类型
                });
            }, { onProgress: (p, t) => updateGlobalLoading((p/t)*100, '解析数据中...') });
        } else {
            // 新的列对列数据处理逻辑
            // 对于列对列结构，日期在某一列，数据在另一列
            const dataCol = parseInt(appData.config.dataStartColumn); // 只使用起始列作为数据列
            
            // 对于列对列结构，也进行自动时间间隔检测
            // 使用行数特征来检测（长表结构按行存储数据）
            const detectionResult = detectTimeIntervalIntelligently(data, startIndex, endIndex, dateCol, meteringPointCol, selectedMeteringPoint, false, 0);
            
            let autoDetectedInterval = detectionResult.interval;
            let intervalDescription = detectionResult.description;
            
            // 显示增强的检测结果通知
            showEnhancedIntervalNotification(detectionResult);
            
            // 如果自动检测置信度高，使用自动检测结果；否则使用用户选择
            const userTimeInterval = parseInt(appData.config.timeInterval);
            const timeInterval = detectionResult.confidence >= 0.7 ? autoDetectedInterval : userTimeInterval;
            
            // 按日期分组处理数据
            const dateGroups = {};
            
            // 处理每一行数据
            const targetIndicesCol = [];
            const maxIdxCol = Math.min(endIndex, data.length - 1);
            for (let k = startRow; k <= maxIdxCol; k++) targetIndicesCol.push(k);

            await processChunksAsync(targetIndicesCol, (i) => {
                const row = data[i];
                if (!row || !row[dateCol]) return;
                
                // 检查计量点编号筛选
                if (meteringPointCol !== null && selectedMeteringPoint) {
                    const rowMeteringPoint = String(row[meteringPointCol]).trim();
                    if (rowMeteringPoint !== selectedMeteringPoint) {
                        return; // 跳过不匹配的行
                    }
                }
                
                // 检查附加筛选项
                let skipRow = false;
                appData.config.additionalFilters.forEach(filter => {
                    if (filter.column && filter.value) {
                        const col = parseInt(filter.column);
                        const rowValue = String(row[col]).trim();
                        if (rowValue !== filter.value) {
                            skipRow = true;
                        }
                    }
                });
                
                if (skipRow) {
                    return; // 跳过不匹配的行
                }
                
                // 解析日期和时间
                const datePart = String(row[dateCol]).trim();
                const timePart = timeCol !== -1 && row[timeCol] !== undefined ? String(row[timeCol]).trim() : '';
                const dateTimeStr = buildDateTimeString(datePart, timePart);
                const dateTime = parseDateTime(dateTimeStr, appData.config.dateFormat);
                if (!dateTime) return;
                
                // 提取日期部分作为分组键
                const dateKey = formatDateKey(dateTime);
                
                // 获取倍率值（支持单列或 PT/CT 分开）
                let multiplier = appData.config.multiplier;
                
                if (appData.config.multiplierMode === 'ptct' && (ptCol !== null || ctCol !== null)) {
                    // PT/CT 分开模式：倍率 = PT × CT
                    const ptVal = ptCol !== null && row[ptCol] !== undefined ? parseNumber(row[ptCol]) : 1;
                    const ctVal = ctCol !== null && row[ctCol] !== undefined ? parseNumber(row[ctCol]) : 1;
                    if (ptVal !== null && ctVal !== null) {
                        multiplier = ptVal * ctVal;
                    }
                } else if (multiplierCol !== null && row[multiplierCol] !== undefined) {
                    // 单列倍率模式
                    const parsedMultiplier = parseNumber(row[multiplierCol]);
                    if (parsedMultiplier !== null) {
                        multiplier = parsedMultiplier;
                    }
                }
                
                // 提取数据值
                const value = parseNumber(row[dataCol]);
                if (value === null) return;
                
                // 初始化日期分组（如果不存在）
                if (!dateGroups[dateKey]) {
                    dateGroups[dateKey] = {
                        date: new Date(dateTime),
                        dateStr: dateKey,
                        dataPoints: [],
                        multiplier: multiplier,
                        sourceFile: worksheet.file.name,
                        meteringPoint: meteringPointCol !== null ? String(row[meteringPointCol]).trim() : null,
                        timeInterval: timeInterval, // 使用用户选择的时间间隔
                        dataStructure: 'columnToColumn' // 标记数据结构类型
                    };
                }
                
                // 添加数据点到日期分组
                dateGroups[dateKey].dataPoints.push({
                    time: dateTime,
                    value: value
                });
            }, { onProgress: (p, t) => updateGlobalLoading((p/t)*100, '解析列数据...') });
            
            // 将分组数据转换为解析数据格式
            Object.values(dateGroups).forEach(group => {
                // 按时间排序数据点
                group.dataPoints.sort((a, b) => a.time - b.time);
                
                // 添加到解析数据
                appData.parsedData.push({
                    date: group.date,
                    dateStr: group.dateStr,
                    rawData: group.dataPoints.map(p => p.value),
                    multiplier: group.multiplier,
                    sourceFile: group.sourceFile,
                    meteringPoint: group.meteringPoint,
                    timeInterval: group.timeInterval,
                    dataStructure: 'columnToColumn', // 标记数据结构类型
                    dataPoints: group.dataPoints // 保存原始数据点，包含时间信息
                });
            });
        }
    }
    
    // 按日期排序
    appData.parsedData.sort((a, b) => a.date - b.date);
}

        async function processParsedData() {
    appData.processedData = [];
    
    await processChunksAsync(appData.parsedData, parsedRow => {
        const dataStructure = parsedRow.dataStructure || 'columnToRow'; // 默认为列对行结构
        const interval = parsedRow.timeInterval || appData.config.timeInterval;
        
        if (dataStructure === 'columnToRow') {
            // 原有的列对行数据处理逻辑
            let pointsPerHour;
            
            // 根据时间间隔计算每小时数据点数
            if (interval === 1) {
                pointsPerHour = 60; // 每小时60个数据点（每分钟一个）
            } else if (interval === 5) {
                pointsPerHour = 12; // 每小时12个数据点（每5分钟一个）
            } else if (interval === 15) {
                pointsPerHour = 4; // 每小时4个数据点（每15分钟一个）
            } else if (interval === 30) {
                pointsPerHour = 2; // 每小时2个数据点（每30分钟一个）
            } else if (interval === 60) {
                pointsPerHour = 1; // 每小时1个数据点（每60分钟一个）
            } else {
                pointsPerHour = Math.floor(60 / interval); // 计算每小时数据点数
            }
            
            const hourlyData = [];
            let processedRawData = parsedRow.rawData;
            
            // 检查是否为跨日期结构数据
            const isCrossDateStructure = document.getElementById('crossDateStructure') && document.getElementById('crossDateStructure').checked;
            
            // 如果是跨日期结构数据，需要确保时间序列的逻辑正确性
            if (isCrossDateStructure && interval === 15) {
                // 对于15分钟间隔的跨日期结构数据，每4个单元格组合为1小时数据组
                // 需要确保数据的时间顺序正确，避免时间顺序颠倒
                
                // 将数据按小时分组（每4个数据点为一小时）
                const totalHours = Math.floor(processedRawData.length / pointsPerHour);
                
                for (let hour = 0; hour < totalHours; hour++) {
                    const startIdx = hour * pointsPerHour;
                    const endIdx = startIdx + pointsPerHour;
                    const hourData = processedRawData.slice(startIdx, endIdx);
                    
                    // 处理无效数据
                    const fillMethod = appData.config.invalidDataHandling === 'ignore' ? 'ignore' : 'interpolate';
                    const cleanedData = DataUtils.cleanData(hourData, { fillMethod });
                    
                    // 根据数据类型计算小时电能
                    let hourlyValue;
                    
                    if (appData.config.dataType === 'instantPower') {
                        // 瞬时功率数据 (kW)
                        const validData = cleanedData.filter(v => v !== null);
                        if (validData.length === 0) {
                            hourlyValue = null; // 标记为无效
                        } else {
                            const sum = validData.reduce((a, b) => a + b, 0);
                            const avgPower = sum / validData.length; // 计算平均功率
                            hourlyValue = avgPower * 1 * parsedRow.multiplier; // 平均功率×1小时×倍率
                        }
                    } else {
                        // 累积电能数据 (kWh)
                        if (cleanedData.length < 2) {
                            hourlyValue = null; // 数据不足
                        } else {
                            const startValue = cleanedData[0];
                            const endValue = cleanedData[cleanedData.length - 1];
                            
                            if (startValue === null || endValue === null) {
                                hourlyValue = null; // 起始或结束值无效
                            } else if (endValue < startValue) {
                                // 累积值回退，标记为异常
                                console.warn(`累积值异常，日期: ${parsedRow.dateStr}, 小时: ${hour}`);
                                hourlyValue = null;
                            } else if (endValue === startValue && startValue !== 0) {
                                // 累积值相同且不为0，说明可能存在数据未更新或平摊需求
                                // 检查前一个小时是否有有效值，如果相同则可能是数据停滞
                                hourlyValue = 0;
                                console.log(`累积值相同，日期: ${parsedRow.dateStr}, 小时: ${hour}, 值: ${startValue}`);
                            } else {
                                hourlyValue = (endValue - startValue) * parsedRow.multiplier;
                            }
                        }
                    }
                    
                    hourlyData.push(hourlyValue);
                }
                
                // 如果数据不足24小时，用null填充剩余小时
                while (hourlyData.length < 24) {
                    hourlyData.push(null);
                }
            } else {
                // 非跨日期结构数据或其他时间间隔的常规处理
                // 处理每小时数据
                for (let hour = 0; hour < 24; hour++) {
                    const startIdx = hour * pointsPerHour;
                    const endIdx = startIdx + pointsPerHour;
                    const hourData = processedRawData.slice(startIdx, endIdx);
                    
                    // 处理无效数据
                    const fillMethod = appData.config.invalidDataHandling === 'ignore' ? 'ignore' : 'interpolate';
                    const cleanedData = DataUtils.cleanData(hourData, { fillMethod });
                    
                    // 根据数据类型计算小时电能
                    let hourlyValue;
                    
                    if (appData.config.dataType === 'instantPower') {
                        // 瞬时功率数据 (kW)
                        if (interval === 1) {
                            // 1分钟间隔特殊处理：获取第一个小时所有数据点，过滤无效数据后计算总和，通过总和除以有效数据点数得到平均功率，再乘以1小时和倍率
                            const validData = cleanedData.filter(v => v !== null);
                            if (validData.length === 0) {
                                hourlyValue = null; // 标记为无效
                            } else {
                                const sum = validData.reduce((a, b) => a + b, 0);
                                const avgPower = sum / validData.length; // 计算平均功率
                                hourlyValue = avgPower * 1 * parsedRow.multiplier; // 平均功率×1小时×倍率
                            }
                        } else {
                            // 其他间隔：计算(求和结果 ÷ 数据点数) × 1小时 × 倍率 → kWh
                            const validData = cleanedData.filter(v => v !== null);
                            if (validData.length === 0) {
                                hourlyValue = null; // 标记为无效
                            } else {
                                const sum = validData.reduce((a, b) => a + b, 0);
                                const avgPower = sum / validData.length;
                                hourlyValue = avgPower * 1 * parsedRow.multiplier; // 平均功率×1小时×倍率
                            }
                        }
                    } else {
                        // 累积电能数据 (kWh)
                        if (interval === 1) {
                            // 1分钟间隔特殊处理：将每小时内第46-60个单元格数据之和减去第1-15个单元格数据之和，作为1小时一组的数据
                            if (cleanedData.length < 60) {
                                hourlyValue = null; // 数据不足
                            } else {
                                // 第一组：1-15分钟
                                const firstQuarter = cleanedData.slice(0, 15).filter(v => v !== null);
                                // 第二组：46-60分钟
                                const lastQuarter = cleanedData.slice(45, 60).filter(v => v !== null);
                                
                                if (firstQuarter.length === 0 || lastQuarter.length === 0) {
                                    hourlyValue = null; // 数据不足
                                } else {
                                    const firstSum = firstQuarter.reduce((a, b) => a + b, 0);
                                    const lastSum = lastQuarter.reduce((a, b) => a + b, 0);
                                    hourlyValue = (lastSum - firstSum) * parsedRow.multiplier;
                                }
                            }
                        } else {
                            // 其他间隔：计算(当前小时末 - 当前小时初) × 倍率 → kWh
                            if (cleanedData.length < 2) {
                                hourlyValue = null; // 数据不足
                            } else {
                                const startValue = cleanedData[0];
                                const endValue = cleanedData[cleanedData.length - 1];
                                
                                if (startValue === null || endValue === null) {
                                hourlyValue = null; // 起始或结束值无效
                            } else if (endValue < startValue) {
                                // 累积值回退，标记为异常
                                console.warn(`累积值异常，日期: ${parsedRow.dateStr}, 小时: ${hour}`);
                                hourlyValue = null;
                            } else if (endValue === startValue && startValue !== 0) {
                                // 累积值相同且不为0
                                hourlyValue = 0;
                                console.log(`[通用] 累积值相同，日期: ${parsedRow.dateStr}, 小时: ${hour}, 值: ${startValue}`);
                            } else {
                                hourlyValue = (endValue - startValue) * parsedRow.multiplier;
                            }
                            }
                        }
                    }
                    
                    hourlyData.push(hourlyValue);
                }
            }
            
            // 计算日总电能
            const validHourlyData = hourlyData.filter(v => v !== null);
            const dailyTotal = validHourlyData.reduce((a, b) => a + b, 0);
            
            console.log(`=== 日总电能计算调试 (${parsedRow.dateStr}) ===`);
            console.log('有效小时数据点数:', validHourlyData.length);
            console.log('有效小时数据样本:', validHourlyData.slice(0, 5));
            console.log('计算得到的dailyTotal:', dailyTotal);
            
            // 确保日期格式统一为 YYYY-MM-DD
            let formattedDate = parsedRow.dateStr;
            if (typeof parsedRow.dateStr === 'string') {
                const dateStr = parsedRow.dateStr;
                if (dateStr.length === 8 && !dateStr.includes('-')) {
                    // 格式：20240901 -> 2024-09-01
                    formattedDate = `${dateStr.substring(0,4)}-${dateStr.substring(4,6)}-${dateStr.substring(6,8)}`;
                } else if (dateStr.length === 10 && dateStr.includes('/')) {
                    // 格式：2024/09/01 -> 2024-09-01
                    formattedDate = dateStr.replace(/\//g, '-');
                } else {
                    formattedDate = dateStr;
                }
            }
            
            appData.processedData.push({
                date: formattedDate,
                dateObj: parsedRow.date,
                hourlyData: hourlyData,
                dailyTotal: dailyTotal,
                sourceFile: parsedRow.sourceFile,
                meteringPoint: (parsedRow.meteringPoint === undefined || parsedRow.meteringPoint === null || String(parsedRow.meteringPoint).trim() === '') ? '默认计量点' : String(parsedRow.meteringPoint).trim(),
                timeInterval: interval, // 保存时间间隔信息
                dataStructure: 'columnToRow' // 标记数据结构类型
            });
        } else {
            // 按照精确计算规则重构列对列数据处理
            const hourlyData = new Array(24).fill(0); // 仅存储用电量 (kWh)
            const hourlyGeneration = new Array(24).fill(0); // 仅存储发电/反向电量 (kWh)
            
            // ---------------------------------------------------------
            // 1. 数据有效性检测与颗粒度确定
            // ---------------------------------------------------------
            // 即使表格结构是1分钟（1440列），实际数据可能是15分钟一次
            // 我们需要检测实际的"有效数据间隔"
            
            let effectivePointsCount = 0;
            if (parsedRow.dataPoints) {
                effectivePointsCount = parsedRow.dataPoints.filter(p => p.value !== null && !isNaN(p.value) && p.value !== 0).length;
            }
            
            // 计算每小时的平均有效点数
            const avgPointsPerHour = effectivePointsCount / 24;
            
            // 确定能量换算系数 (energyFactor)
            // 逻辑：Energy = Sum(Power) * Factor
            // Factor = Interval_Hours = Interval_Minutes / 60
            let energyFactor = 1; // 默认 1 (1小时数据)
            let detectedInterval = 60; // 默认 60分钟
            
            if (avgPointsPerHour >= 50) {
                // 接近 60 点/小时 -> 1分钟间隔
                detectedInterval = 1;
                energyFactor = 1 / 60; 
            } else if (avgPointsPerHour >= 10) {
                // 接近 12 点/小时 -> 5分钟间隔
                detectedInterval = 5;
                energyFactor = 5 / 60; // 1/12
            } else if (avgPointsPerHour >= 3.5) {
                // 接近 4 点/小时 -> 15分钟间隔 (最常见)
                // 用户需求："正值处理：除以4" -> Sum * (1/4)
                detectedInterval = 15;
                energyFactor = 15 / 60; // 0.25
            } else if (avgPointsPerHour >= 1.5) {
                // 接近 2 点/小时 -> 30分钟间隔
                detectedInterval = 30;
                energyFactor = 30 / 60; // 0.5
            } else {
                // 接近 1 点/小时 -> 60分钟间隔
                detectedInterval = 60;
                energyFactor = 1;
            }
            
            console.log(`[颗粒度检测] 日期: ${parsedRow.dateStr}, 有效点数: ${effectivePointsCount}, 估算间隔: ${detectedInterval}min, 换算系数: ${energyFactor}`);

            // 根据时间间隔计算每小时的理论数据点范围 (用于切片)
            // 注意：这里仍然使用表格结构的间隔来切片，因为 dataPoints 是基于表格列生成的
            let structInterval = interval; // 表格结构的间隔 (如 1)
            let pointsPerHourStruct = Math.floor(60 / structInterval);
            
            // 确保数据点总数正确
            const expectedTotalPoints = 24 * pointsPerHourStruct;
            
            // 处理每个数据点
            if (parsedRow.dataPoints && parsedRow.dataPoints.length >= expectedTotalPoints) {
                const allPoints = parsedRow.dataPoints;
                
                if (appData.config.dataType === 'instantPower') {
                    // ---------------------------------------------------------
                    // 2. 用电量数据处理规则 (瞬时功率 -> 电量)
                    // ---------------------------------------------------------
                    for (let hour = 0; hour < 24; hour++) {
                        // 计算当前小时的数据点范围
                        const startIndex = hour * pointsPerHourStruct;
                        const endIndex = startIndex + pointsPerHourStruct;
                        
                        // 获取当前小时的数据点
                        const hourPoints = allPoints.slice(startIndex, endIndex)
                            .map(p => p.value)
                            .filter(v => v !== null && !isNaN(v)); // 此时保留0和负数以便后续分类
                        
                        if (hourPoints.length > 0) {
                            let posSum = 0;
                            let negSum = 0;
                            
                            hourPoints.forEach(v => {
                                if (v > 0) {
                                    posSum += v;
                                } else if (v < 0) {
                                    // 负值处理：绝对值计入发电
                                    negSum += Math.abs(v);
                                }
                                // v=0 忽略
                            });
                            
                            // 应用公式：Sum * EnergyFactor * Multiplier
                            // 例如 15min间隔: Sum * 0.25 * M
                            hourlyData[hour] = posSum * energyFactor * parsedRow.multiplier;
                            hourlyGeneration[hour] = negSum * energyFactor * parsedRow.multiplier;
                        } else {
                            hourlyData[hour] = null;
                            hourlyGeneration[hour] = null;
                        }
                    }
                } else {
                    // 累积电能数据处理 (保持原有逻辑，但增加负值分离)
                    for (let hour = 0; hour < 24; hour++) {
                        const startIndex = hour * pointsPerHourStruct;
                        const endIndex = startIndex + pointsPerHourStruct;
                        
                        const hourPointsWithTime = allPoints.slice(startIndex, endIndex)
                            .filter(p => p.value !== null && !isNaN(p.value));
                        
                        if (hourPointsWithTime.length >= 2) {
                            hourPointsWithTime.sort((a, b) => a.time - b.time);
                            
                            const firstValue = hourPointsWithTime[0].value;
                            const lastValue = hourPointsWithTime[hourPointsWithTime.length - 1].value;
                            
                            if (lastValue > firstValue) {
                                // 正常用电
                                hourlyData[hour] = (lastValue - firstValue) * parsedRow.multiplier;
                                hourlyGeneration[hour] = 0;
                            } else if (lastValue < firstValue) {
                                // 反向/发电
                                const diff = (firstValue - lastValue) * parsedRow.multiplier; // 取正值
                                
                                // 简单的异常过滤 (如果突变过大)
                                if (diff > firstValue * 0.5 && firstValue > 100) {
                                     hourlyData[hour] = 0;
                                     hourlyGeneration[hour] = 0;
                                } else {
                                     hourlyData[hour] = 0;
                                     hourlyGeneration[hour] = diff;
                                }
                            } else {
                                hourlyData[hour] = 0;
                                hourlyGeneration[hour] = 0;
                            }
                        } else {
                            hourlyData[hour] = null;
                            hourlyGeneration[hour] = null;
                        }
                    }
                }
            } else {
                console.warn(`数据点数量不足`);
            }
            
            // 计算日总电能 (仅正向)
            const validHourlyData = hourlyData.filter(v => v !== null);
            const dailyTotal = validHourlyData.reduce((a, b) => a + b, 0);
            
            // ... (日志)
            
            // 确保日期格式统一为 YYYY-MM-DD
            let formattedDate = parsedRow.dateStr;
            // ... (日期格式化逻辑保持不变)
            if (typeof parsedRow.dateStr === 'string') {
                const dateStr = parsedRow.dateStr;
                if (dateStr.length === 8 && !dateStr.includes('-')) {
                    formattedDate = `${dateStr.substring(0,4)}-${dateStr.substring(4,6)}-${dateStr.substring(6,8)}`;
                } else if (dateStr.length === 10 && dateStr.includes('/')) {
                    formattedDate = dateStr.replace(/\//g, '-');
                } else {
                    formattedDate = dateStr;
                }
            }
            
            appData.processedData.push({
                date: formattedDate,
                dateObj: parsedRow.date,
                hourlyData: hourlyData,         // 纯正向用电
                hourlyGeneration: hourlyGeneration, // 纯反向发电 (新增)
                dailyTotal: dailyTotal,
                sourceFile: parsedRow.sourceFile,
                meteringPoint: (parsedRow.meteringPoint === undefined || parsedRow.meteringPoint === null || String(parsedRow.meteringPoint).trim() === '') ? '默认计量点' : String(parsedRow.meteringPoint).trim(),
                timeInterval: detectedInterval, // 使用检测到的实际间隔
                dataStructure: 'columnToRow'
            });
        }
    }, { onProgress: (p, t) => updateGlobalLoading((p/t)*100, '处理数据结构...') });
    
    // 过滤掉所有数据都为空的日期
    appData.processedData = appData.processedData.filter(row => {
        return row.hourlyData.some(value => value !== null);
    });

    // 针对累积值相同的情况进行数据分摊处理
    await distributeStagnantEnergy();
    
    // 更新数据概览统计信息
    updateDataOverview();
    
    // 初始化可视化界面
    initializeVisualization();
}

/**
 * 针对累积值相同的情况进行数据分摊处理
 * 当发现一段连续的小时内累积值没有变化，但后续突然增加时，
 * 将增加的电量平摊到前面数值停滞的小时段内。
 */
async function distributeStagnantEnergy() {
    if (appData.config.dataType !== 'cumulativeEnergy') return;

    console.log('开始执行累积值停滞电量分摊算法...');
    
    await processChunksAsync(appData.processedData, dayData => {
        const data = dayData.hourlyData;
        let i = 0;
        while (i < 24) {
            // 寻找电量为0且累积值停滞的起点
            if (data[i] === 0) {
                let stagnantStart = i;
                let stagnantEnd = i;
                
                // 寻找停滞期的终点
                while (stagnantEnd < 23 && data[stagnantEnd + 1] === 0) {
                    stagnantEnd++;
                }
                
                // 检查停滞期结束后是否有突发的增量
                if (stagnantEnd < 23 && data[stagnantEnd + 1] > 0) {
                    const totalEnergyToDistribute = data[stagnantEnd + 1];
                    const periodLength = (stagnantEnd - stagnantStart + 1) + 1; // 包含停滞期和突增的小时
                    const averageEnergy = totalEnergyToDistribute / periodLength;
                    
                    console.log(`检测到电量停滞期: ${dayData.date} ${stagnantStart}:00 - ${stagnantEnd + 1}:00, 总电量 ${totalEnergyToDistribute.toFixed(2)}, 参与分摊小时数 ${periodLength}, 平均每小时分摊 ${averageEnergy.toFixed(2)}`);
                    
                    // 执行平摊
                    for (let k = stagnantStart; k <= stagnantEnd + 1; k++) {
                        data[k] = averageEnergy;
                    }
                    
                    // 跳过已处理的区域
                    i = stagnantEnd + 2;
                    continue;
                }
            }
            i++;
        }
        
        // 重新计算日总电能以确保一致性
        dayData.dailyTotal = data.filter(v => v !== null).reduce((a, b) => a + b, 0);
    }, { onProgress: (p, t) => updateGlobalLoading((p/t)*100, '处理累积值...') });
}

        /**
         * 使用新的配置重新处理原始数据
         * 这个函数会在第二步配置发生变化时调用，确保第三步可视化使用最新的配置
         */
        async function reprocessDataWithConfig() {
            console.log('开始使用最新配置重新处理数据...');
            
            // 1. 从界面获取最新配置
            const dataTypeEl = document.querySelector('input[name="dataType"]:checked');
            if (dataTypeEl) appData.config.dataType = dataTypeEl.value;
            
            const invalidDataHandlingEl = document.getElementById('invalidDataHandling');
            if (invalidDataHandlingEl) appData.config.invalidDataHandling = invalidDataHandlingEl.value;
            
            const multiplierEl = document.getElementById('multiplier');
            if (multiplierEl) appData.config.multiplier = parseFloat(multiplierEl.value) || 1;
            
            const timeIntervalEl = document.getElementById('timeInterval');
            if (timeIntervalEl) appData.config.timeInterval = parseInt(timeIntervalEl.value) || 15;
            
            // 2. 重新解析工作表数据（根据最新配置）
            await parseWorksheetData();
            
            // 3. 重新处理解析后的数据
            await processParsedData();
            
            // 4. 更新可视化
            updateChart();
            
            console.log('数据重新处理完成');
        }

        function checkConfigValidity() {
            // 获取数据结构类型
            const dataStructure = document.querySelector('input[name="dataStructure"]:checked')?.value || 'columnToRow';
            
            // 根据数据结构类型验证配置
            let isValid = appData.config.dateColumn !== '' && 
                          appData.config.dataStartColumn !== '';
            
            // 如果是列对行结构，还需要验证结束列
            if (dataStructure === 'columnToRow') {
                isValid = isValid && appData.config.dataEndColumn !== '';
            }

            const hasWorksheets = Array.isArray(appData.worksheets) && appData.worksheets.length > 0;
            const statusBadge = document.getElementById('configStatus');
            const processBtn = elements.processDataBtn;
            
            if (statusBadge) {
                if (!hasWorksheets) {
                    statusBadge.textContent = '等待导入';
                    statusBadge.className = 'inline-flex items-center rounded-full border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-700';
                    if (processBtn) {
                        processBtn.classList.add('opacity-50', 'pointer-events-none');
                        processBtn.classList.add('flex');
                    }
                } else if (isValid) {
                    statusBadge.textContent = '配置完成';
                    statusBadge.className = 'inline-flex items-center rounded-full border border-green-200 bg-green-50 px-3 py-1.5 text-xs font-bold text-green-700 shadow-sm';
                    updateStepStatus(2, true);
                    if (processBtn) {
                        processBtn.classList.remove('opacity-50', 'pointer-events-none');
                        processBtn.classList.add('flex');
                        // 添加脉冲效果提醒
                        processBtn.classList.add('animate-pulse-subtle');
                        setTimeout(() => processBtn.classList.remove('animate-pulse-subtle'), 3000);
                    }
                } else {
                    statusBadge.textContent = '等待配置';
                    statusBadge.className = 'inline-flex items-center rounded-full border border-green-200 bg-green-50 px-3 py-1.5 text-xs font-semibold text-green-700';
                    if (processBtn) {
                        processBtn.classList.add('opacity-50', 'pointer-events-none');
                        processBtn.classList.add('flex');
                    }
                }
            }
            
            // 调试信息：输出配置值
            console.log('checkConfigValidity 调试信息:');
            console.log('数据结构类型:', dataStructure);
            console.log('dateColumn:', appData.config.dateColumn);
            console.log('dataStartColumn:', appData.config.dataStartColumn);
            console.log('dataEndColumn:', appData.config.dataEndColumn);
            console.log('isValid (基本条件):', isValid);
            
            // 智能检测和修正配置（仅作为提示，不再自动覆盖用户选择）
            if (appData.worksheets && appData.worksheets.length > 0) {
                const worksheet = appData.worksheets[0];
                
                // 使用增强版数据结构检测（与 autoDetectDataStructure 一致）
                const detectedStructure = enhancedDetectDataStructure(worksheet.data);
                
                console.log('智能检测到的数据结构:', detectedStructure);
                console.log('当前选择的数据结构:', dataStructure);
                
                // 仅在置信度较低时给出提示，不再自动切换
                if (detectedStructure.structure !== dataStructure && detectedStructure.confidence === 'low') {
                    console.warn(`数据结构可能不匹配！检测到: ${detectedStructure.structure}, 当前选择: ${dataStructure}`);
                    showNotification('提示', `检测到数据可能是${detectedStructure.structure === 'columnToColumn' ? '列对列' : '列对行'}结构，请检查配置`, 'warning');
                }
                
                // 不再自动修正配置 - 尊重用户选择和 autoDetectDataStructure 的结果
                // 如果需要切换，用户可以根据提示手动切换
            }
            
            // 确保结束列不小于起始列（仅对列对行结构）
            if (isValid && dataStructure === 'columnToRow') {
                const startCol = parseInt(appData.config.dataStartColumn);
                const endCol = parseInt(appData.config.dataEndColumn);
                
                console.log('startCol (数字):', startCol);
                console.log('endCol (数字):', endCol);
                console.log('endCol < startCol:', endCol < startCol);
                
                if (endCol < startCol) {
                    showNotification('警告', '结束列不能小于起始列', 'warning');
                    console.log('按钮被禁用：结束列小于起始列');
                    return false;
                }
            }
            
            console.log('最终按钮状态:', !isValid ? '禁用' : '启用');

            return isValid;
        }
        
        // 智能检测数据结构
        function smartDetectDataStructure(data) {
            if (!data || data.length < 2) return 'columnToColumn';
            
            // 检查第一行是否包含时间列标题
            const firstRow = data[0] || [];
            let timeColumnCount = 0;
            
            for (let col = 0; col < Math.min(firstRow.length, 30); col++) {
                const cellValue = firstRow[col];
                if (cellValue) {
                    const headerValue = cellValue.toString().toLowerCase();
                    if (headerValue.includes('时间') || headerValue.includes('time') || 
                        headerValue.includes('hour') || headerValue.includes('小时') ||
                        /^\d{1,2}[:：]\d{2}/.test(headerValue) || /^\d{1,2}[点时]/.test(headerValue) ||
                        /^[0-9]{1,2}$/.test(headerValue.trim()) || // 纯数字小时
                        headerValue.match(/^(0?[0-9]|1[0-9]|2[0-3])$/)) { // 0-23小时格式
                        timeColumnCount++;
                    }
                }
            }
            
            // 检查是否有日期列
            let dateColumnFound = false;
            for (let col = 0; col < Math.min(5, firstRow.length); col++) {
                let dateCount = 0;
                const sampleSize = Math.min(10, data.length - 1);
                
                for (let row = 1; row <= sampleSize; row++) {
                    if (data[row] && data[row][col]) {
                        const cellValue = data[row][col].toString();
                        if (isValidDateString(cellValue)) {
                            dateCount++;
                        }
                    }
                }
                
                if (dateCount >= sampleSize * 0.7) {
                    dateColumnFound = true;
                    break;
                }
            }
            
            // 决策逻辑
            if (timeColumnCount >= 12) {
                return 'columnToRow';
            }
            if (dateColumnFound) {
                return 'columnToColumn';
            }
            
            return 'columnToColumn'; // 默认
        }
        
        // 验证日期字符串
        function isValidDateString(dateStr) {
            if (!dateStr) return false;
            
            const str = dateStr.toString().trim();
            
            // 检查各种日期格式
            const datePatterns = [
                /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/, // YYYY-MM-DD 或 YYYY/MM/DD
                /^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/, // MM-DD-YYYY 或 MM/DD/YYYY
                /^\d{4}\d{2}\d{2}$/, // YYYYMMDD
                /^\d{4}年\d{1,2}月\d{1,2}日?$/, // 中文日期格式
                /^\d{1,2}月\d{1,2}日$/ // 简化中文日期
            ];
            
            // 检查是否匹配任何日期模式
            for (const pattern of datePatterns) {
                if (pattern.test(str)) {
                    // 尝试解析为日期对象验证有效性
                    const date = new Date(str);
                    if (!isNaN(date.getTime())) {
                        return true;
                    }
                }
            }
            
            // 尝试直接解析
            const date = new Date(str);
            return !isNaN(date.getTime()) && date.getFullYear() > 1900 && date.getFullYear() < 2100;
        }

        // 增强版数据结构检测函数（供 checkConfigValidity 使用，与 autoDetectDataStructure 逻辑一致）
        function enhancedDetectDataStructure(jsonData) {
            if (!jsonData || jsonData.length < 2) {
                return { structure: 'columnToRow', confidence: 'low', score: 0 };
            }
            
            const headerRow = jsonData[0];
            const dataRows = jsonData.slice(1, Math.min(51, jsonData.length));
            
            let timeSeriesScore = 0;
            let columnSeriesScore = 0;
            
            // 1. 检查表头中的时间模式
            let timeHeaderCount = 0;
            let timeHeaderIndices = [];
            const timePattern = /^([01]?\d|2[0-3])[:：][0-5]\d$/;
            const hourPattern = /^\d{1,2}(时|点)$/;
            const timePatternWithSeconds = /^([01]?\d|2[0-3])[:：][0-5]\d[:：][0-5]\d$/;
            
            headerRow.forEach((header, idx) => {
                const text = String(header).trim();
                if (timePattern.test(text) || hourPattern.test(text) || timePatternWithSeconds.test(text)) {
                    timeHeaderCount++;
                    timeHeaderIndices.push(idx);
                }
            });
            
            // 检测连续时间列
            let maxConsecutiveTimeHeaders = 0;
            if (timeHeaderIndices.length >= 2) {
                let consecutiveCount = 1;
                maxConsecutiveTimeHeaders = 1;
                for (let i = 1; i < timeHeaderIndices.length; i++) {
                    if (timeHeaderIndices[i] === timeHeaderIndices[i-1] + 1) {
                        consecutiveCount++;
                        maxConsecutiveTimeHeaders = Math.max(maxConsecutiveTimeHeaders, consecutiveCount);
                    } else {
                        consecutiveCount = 1;
                    }
                }
            }
            
            // 评分计算
            if (timeHeaderCount >= 24) timeSeriesScore += 50;
            else if (timeHeaderCount >= 12) timeSeriesScore += 40;
            else if (timeHeaderCount >= 4) timeSeriesScore += 25;
            
            if (maxConsecutiveTimeHeaders >= 48) timeSeriesScore += 25;
            else if (maxConsecutiveTimeHeaders >= 24) timeSeriesScore += 20;
            else if (maxConsecutiveTimeHeaders >= 12) timeSeriesScore += 10;
            
            // 列数分析
            const totalColumns = headerRow.length;
            if (totalColumns >= 50) timeSeriesScore += 20;
            else if (totalColumns >= 24) timeSeriesScore += 15;
            
            const timeColumnRatio = timeHeaderCount / totalColumns;
            if (timeColumnRatio >= 0.5) timeSeriesScore += 20;
            else if (timeColumnRatio >= 0.3) timeSeriesScore += 10;
            
            // 数值列分析
            let numericColumnCount = 0;
            let timeRangeNumericCount = 0;
            const sampleRows = Math.min(10, dataRows.length);
            
            for (let col = 0; col < headerRow.length; col++) {
                let colNumericCount = 0;
                for (let row = 0; row < sampleRows; row++) {
                    const val = dataRows[row]?.[col];
                    if (val !== null && val !== undefined && !isNaN(parseFloat(val)) && isFinite(val)) {
                        colNumericCount++;
                    }
                }
                if (colNumericCount / sampleRows >= 0.8) {
                    numericColumnCount++;
                    if (timeHeaderIndices.includes(col)) timeRangeNumericCount++;
                }
            }
            
            if (numericColumnCount >= 24) timeSeriesScore += 20;
            else if (numericColumnCount >= 12) timeSeriesScore += 15;
            
            if (timeRangeNumericCount >= timeHeaderCount * 0.8 && timeHeaderCount >= 12) {
                timeSeriesScore += 15;
            }
            
            // 特殊模式检测（96个15分钟间隔）
            if (timeHeaderCount >= 90 && timeHeaderCount <= 100) {
                const firstTimeIdx = timeHeaderIndices[0];
                const lastTimeIdx = timeHeaderIndices[timeHeaderIndices.length - 1];
                if (lastTimeIdx - firstTimeIdx + 1 === timeHeaderCount) {
                    timeSeriesScore += 20;
                }
            }
            
            // 综合决策
            let detectedStructure = 'columnToRow';
            let confidence = 'medium';
            const scoreDiff = Math.abs(columnSeriesScore - timeSeriesScore);
            
            if (columnSeriesScore > timeSeriesScore) {
                detectedStructure = 'columnToColumn';
            }
            
            if (scoreDiff > 30) confidence = 'high';
            else if (scoreDiff > 15) confidence = 'medium';
            else confidence = 'low';
            
            // 特殊情况：如果有很多时间列，强制判定为宽表
            if (timeHeaderCount >= 24 && timeSeriesScore > columnSeriesScore * 0.5) {
                detectedStructure = 'columnToRow';
                if (timeHeaderCount >= 48) confidence = 'high';
            }
            
            return {
                structure: detectedStructure,
                confidence: confidence,
                score: timeSeriesScore,
                details: {
                    timeHeaderCount,
                    maxConsecutiveTimeHeaders,
                    totalColumns,
                    timeColumnRatio: timeColumnRatio.toFixed(2),
                    numericColumnCount
                }
            };
        }

        function updateMeteringPointFilterOptions() {
            // 清空现有选项
            elements.meteringPointFilterSelect.innerHTML = '<option value="">计量编号</option>';
            if (elements.meteringPointFilterSelect_v2) {
                elements.meteringPointFilterSelect_v2.innerHTML = '<option value="">计量编号</option>';
            }
            
            // 重置计量点列表
            appData.meteringPoints = [];
            
            // 如果没有选择计量点编号列，则返回
            if (!appData.config.meteringPointColumn) return;
            
            const meteringPointCol = parseInt(appData.config.meteringPointColumn);
            
            // 收集所有唯一的计量点编号
            const meteringPointSet = new Set();
            
            appData.worksheets.forEach(worksheet => {
                const { data } = worksheet;
                
                // 跳过表头行，尝试自动检测数据起始行
                let startRow = 0;
                for (let i = 0; i < Math.min(10, data.length); i++) {
                    if (data[i] && data[i][meteringPointCol] && !isDateLike(data[i][meteringPointCol])) {
                        startRow = i;
                        break;
                    }
                }
                
                // 收集计量点编号
                for (let i = startRow; i < data.length; i++) {
                    const row = data[i];
                    if (!row || !row[meteringPointCol]) continue;
                    
                    const meteringPoint = String(row[meteringPointCol]).trim();
                    if (meteringPoint) {
                        meteringPointSet.add(meteringPoint);
                    }
                }
            });
            
            // 如果数据中存在空/缺失计量点，为其添加统一占位『默认计量点』
            if ([...meteringPointSet].some(v => v === undefined || v === null || String(v).trim() === '')) {
                meteringPointSet.add('默认计量点');
            }
            // 转换为数组并排序
            appData.meteringPoints = Array.from(meteringPointSet).sort();
            
            // 添加到下拉菜单
            appData.meteringPoints.forEach(point => {
                const option1 = document.createElement('option');
                option1.value = point;
                option1.textContent = point;
                elements.meteringPointFilterSelect.appendChild(option1);

                // 同步到 v2
                if (elements.meteringPointFilterSelect_v2) {
                    const option2 = document.createElement('option');
                    option2.value = point;
                    option2.textContent = point;
                    elements.meteringPointFilterSelect_v2.appendChild(option2);
                }
            });
            
            if (appData.config.meteringPointFilter) {
                appData.visualization.selectedMeteringPoints = [appData.config.meteringPointFilter];
                // 恢复选中状态
                elements.meteringPointFilterSelect.value = appData.config.meteringPointFilter;
                if (elements.meteringPointFilterSelect_v2) {
                    elements.meteringPointFilterSelect_v2.value = appData.config.meteringPointFilter;
                }
            } else {
                // 默认不选中任何计量点（即显示空或根据逻辑处理为全部）
                // 修正：用户要求默认空值，这里设置为空数组
                appData.visualization.selectedMeteringPoints = []; 
            }
        }

        function applyVisualizationDateRange() {
            if (!elements.startDateInput || !elements.endDateInput) return;

            const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
            const startDate = elements.startDateInput.value;
            const endDate = elements.endDateInput.value;

            // 如果没有设置日期范围，且不是“显示全部日期”模式，默认只显示第一个日期
            if (!startDate || !endDate || allDates.length === 0) {
                if (appData.visualization.allCurvesMode || appData.visualization.summaryMode) {
                    appData.visualization.selectedDates = allDates;
                } else {
                    appData.visualization.selectedDates = allDates.length > 0 ? [allDates[0]] : [];
                }
                return;
            }

            // 如果处于“显示全部日期”模式，但日期选择器仍然有值，我们尊重选择器的范围
            // 只有当用户没有选择范围（或明确点击了切换按钮）时，我们才可能强制切换

            const startIndex = allDates.indexOf(startDate);
            const endIndex = allDates.indexOf(endDate);

            if (startIndex === -1 || endIndex === -1) {
                if (appData.visualization.allCurvesMode || appData.visualization.summaryMode) {
                    appData.visualization.selectedDates = allDates;
                } else {
                    appData.visualization.selectedDates = allDates.length > 0 ? [allDates[0]] : [];
                }
                return;
            }

            const rangeStart = Math.min(startIndex, endIndex);
            const rangeEnd = Math.max(startIndex, endIndex);
            
            // 如果是单日期模式且没有手动调整结束日期，我们确保只选一天
            if (!appData.visualization.allCurvesMode && !appData.visualization.summaryMode && startDate === endDate) {
                appData.visualization.selectedDates = [startDate];
            } else {
                appData.visualization.selectedDates = allDates.slice(rangeStart, rangeEnd + 1);
            }
        }

        function updateVisualizationDateSelects() {
            if (!elements.startDateInput || !elements.endDateInput) return;

            const dates = [...new Set(appData.processedData.map(d => d.date))].sort();
            const currentStart = elements.startDateInput.value;
            const currentEnd = elements.endDateInput.value;

            elements.startDateInput.innerHTML = '<option value="">起始日期</option>';
            elements.endDateInput.innerHTML = '<option value="">结束日期</option>';

            if (dates.length === 0) {
                appData.visualization.selectedDates = [];
                return;
            }

            dates.forEach(date => {
                const startOption = document.createElement('option');
                startOption.value = date;
                startOption.textContent = date;
                elements.startDateInput.appendChild(startOption);

                const endOption = document.createElement('option');
                endOption.value = date;
                endOption.textContent = date;
                elements.endDateInput.appendChild(endOption);
            });

            elements.startDateInput.value = dates.includes(currentStart) ? currentStart : dates[0];
            
            // 如果不是“显示全部日期”模式，且没有当前选择的结束日期，默认结束日期等于开始日期（即只选一天）
            if (!appData.visualization.allCurvesMode && !appData.visualization.summaryMode && !currentEnd) {
                elements.endDateInput.value = elements.startDateInput.value;
            } else {
                elements.endDateInput.value = dates.includes(currentEnd) ? currentEnd : dates[dates.length - 1];
            }

            if (elements.startDateInput.value && elements.endDateInput.value && elements.startDateInput.value > elements.endDateInput.value) {
                elements.endDateInput.value = elements.startDateInput.value;
            }

            applyVisualizationDateRange();
        }

        // 可视化部分函数
        function initializeVisualization() {
            // 填充日期下拉选择框
            updateVisualizationDateSelects();
            
            // 默认应用日期范围选择逻辑（根据 allCurvesMode 决定是单个还是全部）
            applyVisualizationDateRange();
            console.log('初始化选择日期:', appData.visualization.selectedDates.length, '个日期');
            
            if (!appData.config.meteringPointColumn) {
                appData.meteringPoints = [];
                appData.visualization.selectedMeteringPoints = [];
            } else if (appData.meteringPoints.length > 0) {
                if (appData.config.meteringPointFilter) {
                    appData.visualization.selectedMeteringPoints = [appData.config.meteringPointFilter];
                } else {
                    appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                }
                console.log('初始化选择计量点:', appData.visualization.selectedMeteringPoints.length, '个计量点');
            }
                
                // 初始化图表区域点击事件处理
                document.getElementById('chartContainer').addEventListener('click', function(e) {
                    // 阻止事件冒泡，防止图形消失
                    e.stopPropagation();
                    
                    // 如果点击的是canvas元素，不进行任何操作
                    if (e.target.tagName === 'CANVAS') {
                        return;
                    }
                    
                    // 如果点击的是加载状态或无数据消息区域，也不进行任何操作
                    if (e.target.id === 'chartLoading' || e.target.id === 'noDataMessage' || 
                        e.target.closest('#chartLoading') || e.target.closest('#noDataMessage')) {
                        return;
                    }
                });
                
                // 注：toggleLegend元素已删除，图例显示控制已集成到图表配置中
                
                // 初始化柱状图区域点击事件处理
                document.getElementById('dailyTotalBarChartContainer').addEventListener('click', function(e) {
                    // 阻止事件冒泡，防止图形消失
                    e.stopPropagation();
                    
                    // 如果点击的是canvas元素，不进行任何操作
                    if (e.target.tagName === 'CANVAS') {
                        return;
                    }
                    
                    // 如果点击的是加载状态或无数据消息区域，也不进行任何操作
                    if (e.target.id === 'barChartLoading' || e.target.id === 'barChartNoDataMessage' || 
                        e.target.closest('#barChartLoading') || e.target.closest('#barChartNoDataMessage')) {
                        return;
                    }
                });
                
                // 注：toggleBarLegend元素已删除，柱状图图例显示控制已集成到图表配置中
                
                // 初始化图表
                // 检查是否已存在图表实例，如果存在则先销毁
                if (appData.chart) {
                    appData.chart.destroy();
                    appData.chart = null;
                }
                

                
                const ctx = elements.loadCurveChart.getContext('2d');
                appData.chart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: Array.from({ length: 24 }, (_, i) => `${i}:00`),
                        datasets: []
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        layout: {
                            padding: {
                                top: 10,
                                right: 15,
                                bottom: 10,
                                left: 15
                            }
                        },
                        interaction: {
                            mode: 'index',
                            intersect: false
                        },
                        animation: {
                            duration: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 0 : 1000,
                            easing: 'easeOutCubic',
                            delay: (context) => {
                                if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return 0;
                                let delay = 0;
                                if (context.type === 'data' && context.mode === 'default') {
                                    delay = context.dataIndex * 30;
                                }
                                return delay;
                            }
                        },
                        plugins: {
                            title: {
                                display: false,
                                text: '24小时负荷曲线',
                                font: {
                                    size: 18,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '600'
                                },
                                padding: {
                                    top: 10,
                                    bottom: 15
                                },
                                color: '#1F2937'
                            },
                            legend: {
                                display: false,
                                position: 'top',
                                align: 'center',
                                labels: {
                                    usePointStyle: true,
                                    padding: 8,
                                    font: {
                                        size: 11,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '500'
                                    },
                                    boxWidth: 12,
                                    boxHeight: 12,
                                    // 自动调整图例布局以适应容器宽度
                                    generateLabels: function(chart) {
                                        const datasets = chart.data.datasets;
                                        const originalLabels = Chart.defaults.plugins.legend.labels.generateLabels(chart);
                                        
                                        // 根据数据集数量调整图例显示
                                        if (datasets.length > 20) {
                                            // 数据集较多时，减小字体和间距
                                            originalLabels.forEach(label => {
                                                label.font = {
                                                    size: 10,
                                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                                    weight: '500'
                                                };
                                                label.padding = 6;
                                                label.boxWidth = 6;
                                                label.boxHeight = 6;
                                            });
                                        }
                                        
                                        return originalLabels;
                                    }
                                },
                                // 使用默认的legend点击行为(仅显示/隐藏dataset)
                                onClick: Chart.defaults.plugins.legend.onClick
                            },
                            tooltip: {
                                backgroundColor: 'rgba(15, 23, 42, 0.92)',
                                titleColor: 'rgba(255, 255, 255, 0.98)',
                                titleFont: {
                                    size: 14,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '600'
                                },
                                bodyColor: 'rgba(255, 255, 255, 0.92)',
                                bodyFont: {
                                    size: 13,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif'
                                },
                                borderColor: 'rgba(255, 255, 255, 0.14)',
                                borderWidth: 1,
                                padding: 12,
                                cornerRadius: 8,
                                displayColors: true,
                                boxWidth: 10,
                                boxHeight: 10,
                                boxPadding: 4,
                                usePointStyle: true,
                                caretSize: 8,
                                callbacks: {
                                    title: function(tooltipItems) {
                                        return '时间: ' + tooltipItems[0].label;
                                    },
                                    label: function(context) {
                                        let label = context.dataset.label || '';
                                        if (label) {
                                            label += ': ';
                                        }
                                        if (context.parsed.y !== null) {
                                            // 根据数据集使用的Y轴显示不同的单位
                                            const unit = context.dataset.yAxisID === 'y1' ? 'kW' : 'kWh';
                                            // 使用智能格式化
                                            label += smartFormatNumber(context.parsed.y) + ' ' + unit;
                                            
                                            // 如果是负数，添加额外说明
                                            if (context.parsed.y < 0) {
                                                label += ' (发电/反向)';
                                            }
                                        }
                                        return label;
                                    },
                                    afterLabel: function(context) {
                                        return '';
                                    }
                                }
                            }
                        },
                        scales: {
                            x: {
                                title: {
                                    display: true,
                                    text: '时间',
                                    font: {
                                        size: 14,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '600'
                                    },
                                    padding: {
                                        top: 10,
                                        bottom: 10
                                    }
                                },
                                grid: {
                                    color: 'rgba(0, 0, 0, 0.05)',
                                    lineWidth: 1,
                                    drawBorder: false
                                },
                                ticks: {
                                    font: {
                                        size: 12,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '500'
                                    },
                                    padding: 8,
                                    maxRotation: 0,
                                    autoSkip: false
                                }
                            },
                            y: {
                                title: {
                                    display: true,
                                    text: '用电量 (kWh)',
                                    font: {
                                        size: 14,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '600'
                                    },
                                    padding: {
                                        top: 10,
                                        bottom: 10
                                    }
                                },
                                beginAtZero: true,
                                grid: {
                                    color: 'rgba(0, 0, 0, 0.05)',
                                    lineWidth: 1,
                                    drawBorder: false
                                },
                                ticks: {
                                    font: {
                                        size: 12,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '500'
                                    },
                                    padding: 8,
                                    callback: function(value) {
                                        return smartFormatNumber(value, 1);
                                    }
                                }
                            },
                            y1: {
                                display: appData.visualization.secondaryAxis,
                                position: 'right',
                                title: {
                                    display: true,
                                    text: '功率 (kW)',
                                    font: {
                                        size: 14,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '600'
                                    },
                                    padding: {
                                        top: 10,
                                        bottom: 10
                                    }
                                },
                                beginAtZero: true,
                                grid: {
                                    drawOnChartArea: false, // 不在图表区域绘制网格线
                                    drawBorder: false
                                },
                                ticks: {
                                    font: {
                                        size: 12,
                                        family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                        weight: '500'
                                    },
                                    padding: 8,
                                    callback: function(value) {
                                        return smartFormatNumber(value, 1);
                                    }
                                }
                            }
                        },
                        onClick: (event, elements) => {
                            if (elements.length > 0) {
                                const element = elements[0];
                                const datasetIndex = element.datasetIndex;
                                const index = element.index;
                                const dataset = appData.chart.data.datasets[datasetIndex];
                                const value = dataset.data[index];
                                const label = appData.chart.data.labels[index];
                                
                                showNotification(
                                    '数据点详情', 
                                    `${dataset.label} 在 ${label} 的${dataset.yAxisID === 'y1' ? '功率' : '用电量'}为 ${value ? smartFormatNumber(value) : '无数据'} ${dataset.yAxisID === 'y1' ? 'kW' : 'kWh'}`,
                                    'info'
                                );
                            }
                        },
                        onHover: (event, elements) => {
                            event.native.target.style.cursor = elements.length > 0 ? 'pointer' : 'default';
                            
                            // 添加canvas区域的互动反馈
                            const canvas = event.native.target;
                            if (elements.length > 0) {
                                canvas.style.filter = 'brightness(1.05)';
                            } else {
                                canvas.style.filter = 'brightness(1)';
                            }
                        }
                    }
                });
                
                // 确保图表显示数据
                setTimeout(() => {
                    updateChart();
                }, 100);
            }

        function applyTimeFocus() {
            const startHour = parseInt(elements.focusStartTimeSelect.value);
            const endHour = parseInt(elements.focusEndTimeSelect.value);
            
            if (startHour > endHour) {
                showNotification('警告', '开始时间不能晚于结束时间', 'warning');
                return;
            }
            
            appData.visualization.focusStartTime = startHour;
            appData.visualization.focusEndTime = endHour;
            
            // 更新图表和统计信息
            updateChart();
            
            // 自动应用时不显示通知，保持体验流畅
        }

        function resetTimeFocus() {
            elements.focusStartTimeSelect.value = '0';
            elements.focusEndTimeSelect.value = '23';
            appData.visualization.focusStartTime = 0;
            appData.visualization.focusEndTime = 23;
            
            // 添加视觉反馈
            elements.resetTimeFocusBtn.style.transform = 'scale(1.05)';
            setTimeout(() => {
                elements.resetTimeFocusBtn.style.transform = 'scale(1)';
            }, 200);
            
            // 更新图表和统计信息
            updateChart();
            
            // 显示通知
            showNotification('时间焦段重置', '已重置时间焦段为全天 (0:00 - 23:00)', 'success');
        }

        // 图表性能优化配置
        const CHART_PERFORMANCE = {
            maxDataPoints: 500,       // 最大显示数据点数（支持显示更多曲线）
            pageSize: 100,           // 每页数据量（增加每页显示数量）
            animationDuration: 300,   // 动画持续时间
            enableVirtualization: true // 启用虚拟化
        };
        
        // 数据分页管理
        const chartPagination = {
            currentPage: 0,
            totalPages: 0,
            pageSize: CHART_PERFORMANCE.pageSize,
            totalItems: 0
        };
        
        // 重新设计的24小时负荷数据更新函数
        function updateChart(preFilteredData) {
            console.log('=== 开始更新24小时负荷数据 ===');
            
            // 检查图表实例是否存在，如果不存在则重新初始化
            if (!appData.chart) {
                initializeVisualization();
                return;
            }
            
            // 如果图表Canvas已被替换或不在文档中，重新初始化
            if (!appData.chart.canvas || !document.body.contains(appData.chart.canvas)) {
                try { appData.chart.destroy(); } catch (e) {}
                appData.chart = null;
                initializeVisualization();
                return;
            }
            
            // 显示加载状态
            const chartLoadingElement = document.getElementById('chartLoading');
            const noDataMessageElement = document.getElementById('noDataMessage');
            if (chartLoadingElement && chartLoadingElement.parentNode) {
                chartLoadingElement.classList.remove('hidden');
            }
            if (noDataMessageElement && noDataMessageElement.parentNode) {
                noDataMessageElement.classList.add('hidden');
            }

            // 获取图表配置
            const lineStyle = appData.visualization.lineStyle;
            const showPoints = appData.visualization.showPoints;
            const secondaryAxis = appData.visualization.secondaryAxis;
            
            // 更新X轴标签范围
            const startHour = appData.visualization.focusStartTime;
            const endHour = appData.visualization.focusEndTime;
            
            const newLabels = Array.from({ length: 24 }, (_, i) => `${i}:00`).slice(startHour, endHour + 1);
            appData.chart.data.labels = newLabels;
            
            // ---------------------------------------------------------
            // 3. 构建数据集 (新增发电曲线逻辑)
            // ---------------------------------------------------------
            const datasets = [];
            
            // 如果是汇总模式...
            if (appData.visualization.summaryMode) {
                // ... (汇总模式逻辑保持不变，后续需要确保它使用正向数据)
                // 获取选中的日期范围
                const dates = appData.processedData.map(d => d.date).sort();
                
                if (dates.length > 0) {
                    // 计算每一天的总负荷和最大负荷
                    const days = appData.processedData.map(d => {
                        // 使用正向数据
                        const total = d.hourlyData.filter(v => v !== null).reduce((a, b) => a + b, 0);
                        const max = Math.max(...d.hourlyData.filter(v => v !== null));
                        const min = Math.min(...d.hourlyData.filter(v => v !== null && v > 0)); // 最小正向负荷
                        return { ...d, total, max, min: min === Infinity ? 0 : min };
                    });
                    
                    // 找出最高和最低用电日
                    const maxDay = days.reduce((prev, current) => (prev.total > current.total) ? prev : current);
                    const minDay = days.reduce((prev, current) => (prev.total < current.total && prev.total > 0) ? prev : current);
                    
                    // 计算平均负荷曲线 (仅正向)
                    const avgSeries = (dayList) => {
                        return new Array(24).fill(0).map((_, h) => {
                            let sum = 0;
                            let count = 0;
                            dayList.forEach(d => {
                                const val = d.hourlyData[h];
                                if (val !== null) {
                                    sum += val;
                                    count++;
                                }
                            });
                            return count > 0 ? sum / count : null;
                        });
                    };
                    
                    // 提取曲线数据 (截取选定时段)
                    const sliceData = (arr) => arr.slice(startHour, endHour + 1);
                    
                    const maxDaySeries = sliceData(maxDay.hourlyData);
                    const minDaySeries = sliceData(minDay.hourlyData);
                    
                    // 辅助函数：创建数据集
                    const makeDataset = (label, data, color, borderDash = []) => {
                        const curveCount = appData.processedData.length;
                        let width = 2;
                        if (curveCount > 50) width = 0.8;
                        else if (curveCount > 30) width = 1.2;
                        else if (curveCount > 10) width = 1.6;

                        return {
                            label,
                            data,
                            borderColor: color,
                            backgroundColor: color,
                            borderWidth: width,
                            pointRadius: showPoints ? (curveCount > 30 ? 1 : 3) : 0,
                            pointHoverRadius: 5,
                            fill: false,
                            tension: 0.4,
                            borderDash,
                            yAxisID: 'y'
                        };
                    };

                    // 按月分组数据 (用于月度汇总)
                    const monthGroups = {};
                    const monthColors = [
                        '#ef4444', '#f97316', '#f59e0b', '#eab308', // 春
                        '#84cc16', '#22c55e', '#10b981', '#14b8a6', // 夏
                        '#06b6d4', '#0ea5e9', '#3b82f6', '#6366f1'  // 秋/冬
                    ];
                    days.forEach(d => {
                        const m = d.dateObj.getMonth(); // 0-11
                        if (!monthGroups[m]) monthGroups[m] = [];
                        monthGroups[m].push(d);
                    });

                    // 计算各月平均曲线
                    const monthSeries = new Array(12).fill(null).map((_, i) => 
                        monthGroups[i] ? sliceData(avgSeries(monthGroups[i])) : null
                    );

                    // 计算季度平均
                    const q1Series = sliceData(avgSeries(days.filter(d => d.dateObj.getMonth() <= 2)));
                    const q2Series = sliceData(avgSeries(days.filter(d => { const m = d.dateObj.getMonth(); return m >= 3 && m <= 5; })));
                    const q3Series = sliceData(avgSeries(days.filter(d => { const m = d.dateObj.getMonth(); return m >= 6 && m <= 8; })));
                    const q4Series = sliceData(avgSeries(days.filter(d => d.dateObj.getMonth() >= 9)));

                    const subMode = appData.visualization.summarySubMode || 'all';

                    // 1. 最高/最低用电日 + 平均负荷 (extreme)
                    if (subMode === 'all' || subMode === 'extreme') {
                        datasets.push(makeDataset(`最高用电日 (${formatDateForInput(maxDay.dateObj)})`, maxDaySeries, 'rgb(220, 38, 38)'));  // 红色-最高
                        datasets.push(makeDataset(`最低用电日 (${formatDateForInput(minDay.dateObj)})`, minDaySeries, 'rgb(37, 99, 235)'));   // 蓝色-最低
                        
                        // 计算全局平均负荷
                        const globalAvgSeries = sliceData(avgSeries(days));
                        datasets.push(makeDataset('全周期平均负荷', globalAvgSeries, 'rgb(71, 85, 105)', [8, 4])); // 深灰-平均
                    }
                    
                    // 2. 季度平均 (quarter)
                    if (subMode === 'all' || subMode === 'quarter') {
                        datasets.push(makeDataset('Q1平均(1-3月)', q1Series, 'rgb(5, 150, 105)', [6, 4]));      // 绿色-春季
                        datasets.push(makeDataset('Q2平均(4-6月)', q2Series, 'rgb(245, 158, 11)', [6, 4]));     // 琥珀色-夏季
                        datasets.push(makeDataset('Q3平均(7-9月)', q3Series, 'rgb(234, 88, 12)', [6, 4]));      // 橙色-秋季
                        datasets.push(makeDataset('Q4平均(10-12月)', q4Series, 'rgb(30, 64, 175)', [6, 4]));    // 深蓝-冬季
                    }

                    // 3. 月平均 (month)
                    if (subMode === 'all' || subMode === 'month') {
                        for (let m = 1; m <= 12; m++) {
                            if (monthSeries[m - 1].some(v => v !== null)) {
                                datasets.push(
                                    makeDataset(`${m}月平均`, monthSeries[m - 1], monthColors[m - 1], [4, 2])
                                );
                            }
                        }
                    }
                }
                
            } else {
                // 明细模式：绘制所选日期的曲线
                const selectedDates = appData.visualization.selectedDates || [];
                
                selectedDates.forEach((date, index) => {
                    const dayData = appData.processedData.find(d => d.date === date);
                    if (dayData) {
                        const color = getColorForIndex(index);
                        
                        // 1. 用电曲线 (正向)
                        const hourlyDataSlice = dayData.hourlyData.slice(startHour, endHour + 1);
                        
                        // 根据曲线数量动态调整线宽
                        const curveCount = selectedDates.length;
                        let dynamicBorderWidth = 2.5;
                        let dynamicPointRadius = appData.visualization.showPoints ? 3 : 0;

                        if (curveCount > 30) {
                            dynamicBorderWidth = 1;
                            dynamicPointRadius = appData.visualization.showPoints ? 1.5 : 0;
                        } else if (curveCount > 10) {
                            dynamicBorderWidth = 1.5;
                            dynamicPointRadius = appData.visualization.showPoints ? 2 : 0;
                        }

                        datasets.push({
                            label: `${date} (用电)`,
                            data: hourlyDataSlice,
                            borderColor: color,
                            backgroundColor: color,
                            borderWidth: dynamicBorderWidth,
                            pointRadius: dynamicPointRadius,
                            pointHoverRadius: dynamicPointRadius + 2,
                            fill: false,
                            tension: appData.visualization.smoothCurve ? (appData.visualization.curveTension || 0.4) : 0,
                            borderDash: [], // 实线
                            yAxisID: 'y'
                        });
                        
                        // 2. 发电/反向曲线 (负向绝对值) - 仅当有数据时显示
                        const hasGeneration = dayData.hourlyGeneration && dayData.hourlyGeneration.some(v => v > 0);
                        if (hasGeneration) {
                            const hourlyGenSlice = dayData.hourlyGeneration.slice(startHour, endHour + 1);
                            // 将绝对值转换为负数进行绘制，以便在图表中直观显示在0轴下方
                            const plotData = hourlyGenSlice.map(v => v > 0 ? -v : 0);
                            
                            const genColor = '#10b981'; // 翠绿色
                            
                            datasets.push({
                                label: `${date} (发电/反向)`,
                                data: plotData,
                                borderColor: genColor,
                                backgroundColor: genColor,
                                borderWidth: dynamicBorderWidth,
                                pointRadius: dynamicPointRadius,
                                pointHoverRadius: dynamicPointRadius + 2,
                                fill: {
                                    target: 'origin',
                                    above: 'rgba(16, 185, 129, 0.1)', 
                                    below: 'rgba(16, 185, 129, 0.1)' // 填充0轴与曲线之间的区域
                                },
                                tension: appData.visualization.smoothCurve ? (appData.visualization.curveTension || 0.4) : 0,
                                borderDash: [5, 5], // 虚线区分
                                yAxisID: 'y'
                            });
                        }
                    }
                });
            }
            
            appData.chart.data.datasets = datasets;
            
            // 更新Y轴标题
            if (appData.chart.options.scales.y) {
                appData.chart.options.scales.y.title.text = '用电量 (kWh) / 上网电量 (-kWh)';
            }
            
            appData.chart.update();
            
            // ---------------------------------------------------------
            // 4. 计算并渲染全局汇总 (如果处于明细模式且未开启汇总视图)
            // ---------------------------------------------------------
            // 注意：这里我们使用 computeGlobalSummary 来确保数据的一致性
            // 但如果已经处于 summaryMode，则不需要重复计算，因为 updateDataOverview 会处理
            // 这里主要处理图表更新后的联动（如需要）
            
            // 隐藏加载状态
            if (chartLoadingElement) chartLoadingElement.classList.add('hidden');

            // 更新统计信息
            updateStatistics(getOverviewData(), startHour, endHour);

            // 更新步骤 3 状态
            updateStepStatus(3, true);
        }
            

        
        function updateChartWithPagination() {
            // 获取筛选后的数据
            let filteredData = getFilteredData();
            
            // 更新分页信息
            chartPagination.totalItems = filteredData.length;
            chartPagination.totalPages = Math.ceil(filteredData.length / chartPagination.pageSize);
            
            // 如果数据量超过最大显示限制，启用分页
            if (!appData.visualization.summaryMode && filteredData.length > CHART_PERFORMANCE.maxDataPoints) {
                const startIndex = chartPagination.currentPage * chartPagination.pageSize;
                const endIndex = Math.min(startIndex + chartPagination.pageSize, filteredData.length);
                filteredData = filteredData.slice(startIndex, endIndex);
                
                // 显示分页控件
                showPaginationControls();
            } else {
                // 隐藏分页控件
                hidePaginationControls();
            }
            
            // 根据时间焦段设置准备图表数据
            const startHour = appData.visualization.focusStartTime;
            const endHour = appData.visualization.focusEndTime;
            const hourCount = endHour - startHour + 1;
            const labels = Array.from({ length: hourCount }, (_, i) => `${startHour + i}:00`);
            
            // 创建数据集
            const datasets = appData.visualization.summaryMode
                ? createSummaryDatasets(filteredData, startHour, endHour)
                : createOptimizedDatasets(filteredData, startHour, endHour);
            
            // 更新图表数据和选项
            updateChartDataAndOptions(labels, datasets, filteredData, startHour, endHour);
        }

        /**
         * 更新负荷热力图
         * 展现 24小时 x 日期 的负荷分布
         */
        async function updateHeatmap() {
            const canvas = elements.loadHeatmap;
            if (!canvas) return;

            // 任务 ID，用于防止重叠渲染
            const renderId = Date.now();
            window.currentHeatmapRenderId = renderId;

            const loadingEl = elements.heatmapLoading;
            const noDataEl = elements.heatmapNoDataMessage;

            // 延迟显示加载状态，避免极短时间操作的闪烁
            const loadingTimeout = setTimeout(() => {
                if (window.currentHeatmapRenderId === renderId && loadingEl) {
                    loadingEl.classList.remove('hidden');
                    if (canvas.parentElement) {
                        canvas.parentElement.classList.add('loading-shimmer');
                    }
                }
            }, 100);

            if (noDataEl) noDataEl.classList.add('hidden');
            
            // 不再立即清除画布和设置透明度，保留旧内容直到数据准备就绪

            const renderTask = async () => {
                try {
                    const dataToUse = getOverviewData();
                    if (!dataToUse || dataToUse.length === 0) {
                        clearTimeout(loadingTimeout);
                        if (loadingEl) loadingEl.classList.add('hidden');
                        if (noDataEl) noDataEl.classList.remove('hidden');
                        const ctx = canvas.getContext('2d');
                        ctx.clearRect(0, 0, canvas.width, canvas.height);
                        return;
                    }

                    // 按日期分组数据
                    const dateGroups = {};
                    dataToUse.forEach(item => {
                        if (!dateGroups[item.date]) {
                            dateGroups[item.date] = new Array(24).fill(0);
                        }
                        item.hourlyData.forEach((val, hour) => {
                            dateGroups[item.date][hour] += (val || 0);
                        });
                    });

                    const sortedDates = Object.keys(dateGroups).sort();
                    const ctx = canvas.getContext('2d');
                    canvas.style.cursor = 'pointer';
                    
                    // 设置自适应尺寸
                    const dpr = window.devicePixelRatio || 1;
                    const containerWidth = canvas.parentElement.clientWidth;
                    const margin = { top: 34, right: 16, bottom: 18, left: 96 };
                    const cellWidth = Math.max(16, Math.floor((containerWidth - margin.left - margin.right) / 24));
                    const cellHeight = 32;
                    
                    const totalWidth = margin.left + (24 * cellWidth) + margin.right;
                    const totalHeight = margin.top + (sortedDates.length * cellHeight) + margin.bottom;

                    appData.visualization.heatmapLayout = {
                        sortedDates,
                        margin,
                        cellWidth,
                        cellHeight,
                        dateGroups
                    };

                    if (!canvas.dataset.heatmapClickBound) {
                        canvas.dataset.heatmapClickBound = '1';
                        canvas.addEventListener('click', (e) => {
                            const layout = appData.visualization.heatmapLayout;
                            if (!layout || !Array.isArray(layout.sortedDates) || layout.sortedDates.length === 0) return;
                            const rect = canvas.getBoundingClientRect();
                            const y = e.clientY - rect.top;
                            const rowIndex = Math.floor((y - layout.margin.top) / layout.cellHeight);
                            if (rowIndex < 0 || rowIndex >= layout.sortedDates.length) return;
                            const dateStr = layout.sortedDates[rowIndex];
                            if (!dateStr) return;

                            appData.visualization.summaryMode = false;
                            appData.visualization.allCurvesMode = false;

                            const allCurvesBtn = document.getElementById('toggleAllCurves');
                            if (allCurvesBtn) {
                                const span = allCurvesBtn.querySelector('span');
                                if (span) span.textContent = '显示全部日期';
                                allCurvesBtn.classList.remove('bg-indigo-600', 'text-white', 'border-indigo-600');
                                allCurvesBtn.classList.add('bg-white', 'text-slate-600', 'border-slate-200');
                            }

                            const summaryBtn = document.getElementById('toggleSummaryMode');
                            if (summaryBtn) {
                                const span = summaryBtn.querySelector('span');
                                if (span) span.textContent = '汇总曲线';
                                summaryBtn.classList.remove('bg-gray-600');
                                summaryBtn.classList.add('bg-indigo-600');
                            }

                            if (elements.startDateInput) elements.startDateInput.value = dateStr;
                            if (elements.endDateInput) elements.endDateInput.value = dateStr;
                            appData.visualization.selectedDates = [dateStr];

                            applyVisualizationDateRange();
                            updateChart();
                            showNotification('已跳转', `已切换到 ${dateStr} 明细曲线`, 'success');
                        });
                    }
                    
                    // 准备工作完成后，清除定时器并开始正式渲染
                    clearTimeout(loadingTimeout);
                    
                    // 只有在数据准备好后才调整 canvas 尺寸（这会清除画布）
                    if (canvas.width !== totalWidth * dpr || canvas.height !== totalHeight * dpr) {
                        canvas.width = totalWidth * dpr;
                        canvas.height = totalHeight * dpr;
                        canvas.style.width = totalWidth + 'px';
                        canvas.style.height = totalHeight + 'px';
                    }
                    
                    ctx.setTransform(1, 0, 0, 1, 0, 0); // 重置变换
                    ctx.scale(dpr, dpr);

                    // 初始透明度设为 0，然后开始渲染第一块时再淡入，或者直接渲染
                    // 为了极致流畅，我们直接渲染，不使用 opacity = 0
                    canvas.classList.remove('animate-graph-entry');

                    const allValues = [];
                    sortedDates.forEach(date => {
                        dateGroups[date].forEach(v => {
                            const num = typeof v === 'number' && isFinite(v) ? v : 0;
                            allValues.push(num);
                        });
                    });
                    allValues.sort((a, b) => a - b);
                    const q = (p) => {
                        if (allValues.length === 0) return 0;
                        const idx = Math.min(allValues.length - 1, Math.max(0, Math.round((allValues.length - 1) * p)));
                        return allValues[idx];
                    };

                    // 定义多级渐变色阶
                    const colorStops = [
                        { pos: 0, color: [219, 234, 254] },   // blue-100: #dbeafe
                        { pos: 0.25, color: [147, 197, 253] }, // blue-300: #93c5fd
                        { pos: 0.5, color: [253, 224, 71] },  // yellow-300: #fde047
                        { pos: 0.75, color: [251, 146, 60] }, // orange-400: #fb923c
                        { pos: 1, color: [248, 113, 113] }    // red-400: #f87171
                    ];

                    const getGradientColor = (value, min, max) => {
                        if (max <= min) return `rgb(${colorStops[0].color.join(',')})`;
                        const normalized = Math.min(1, Math.max(0, (value - min) / (max - min)));
                        
                        let lower = colorStops[0];
                        let upper = colorStops[colorStops.length - 1];
                        
                        for (let i = 0; i < colorStops.length - 1; i++) {
                            if (normalized >= colorStops[i].pos && normalized <= colorStops[i + 1].pos) {
                                lower = colorStops[i];
                                upper = colorStops[i + 1];
                                break;
                            }
                        }
                        
                        const range = upper.pos - lower.pos;
                        const factor = range === 0 ? 0 : (normalized - lower.pos) / range;
                        
                        const r = Math.round(lower.color[0] + (upper.color[0] - lower.color[0]) * factor);
                        const g = Math.round(lower.color[1] + (upper.color[1] - lower.color[1]) * factor);
                        const b = Math.round(lower.color[2] + (upper.color[2] - lower.color[2]) * factor);
                        
                        return `rgb(${r},${g},${b})`;
                    };

                    const vMin = q(0.05); // 使用5%分位数作为最小值基准，过滤异常低值
                    const vMax = q(0.95); // 使用95%分位数作为最大值基准，过滤异常高值

                    // 绘制背景
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, totalWidth, totalHeight);

                    // 绘制坐标轴文字样式
                    ctx.fillStyle = '#64748b';
                    ctx.font = '500 12px Inter, system-ui, sans-serif';
                    ctx.textAlign = 'center';

                    // 绘制小时刻度 (X轴)
                    const tickStep = cellWidth < 18 ? 2 : 1;
                    for (let h = 0; h < 24; h += tickStep) {
                        ctx.fillText(h + 'h', margin.left + (h * cellWidth) + cellWidth / 2, margin.top - 14);
                    }

                    // 使用分块渲染优化大数据量
                    const CHUNK_SIZE = 50; // 每批渲染50行
                    let currentChunk = 0;
                    
                    const renderChunk = async () => {
                        // 检查当前任务是否已被新的任务取代
                        if (window.currentHeatmapRenderId !== renderId) {
                            return;
                        }

                        const startIdx = currentChunk * CHUNK_SIZE;
                        const endIdx = Math.min(startIdx + CHUNK_SIZE, sortedDates.length);
                        
                        // 第一块渲染时，如果还显示着加载动画，则关闭它
                        if (currentChunk === 0) {
                            if (loadingEl) loadingEl.classList.add('hidden');
                            if (canvas.parentElement) {
                                canvas.parentElement.classList.remove('loading-shimmer');
                            }
                        }

                        for (let i = startIdx; i < endIdx; i++) {
                            const date = sortedDates[i];
                            // 绘制日期标签 (Y轴)
                            ctx.textAlign = 'right';
                            ctx.fillStyle = '#475569';
                            ctx.fillText(date, margin.left - 15, margin.top + (i * cellHeight) + cellHeight / 2 + 5);

                            const dayData = dateGroups[date];

                            const rowX = margin.left;
                            const rowY = margin.top + (i * cellHeight) + 4;
                            const rowW = 24 * cellWidth;
                            const rowH = cellHeight - 8;

                            dayData.forEach((v, h) => {
                                const num = typeof v === 'number' && isFinite(v) ? v : 0;
                                const fill = getGradientColor(num, vMin, vMax);

                                const rectX = rowX + (h * cellWidth);
                                ctx.fillStyle = fill;
                                ctx.fillRect(rectX, rowY, cellWidth, rowH);
                            });
                        }
                        
                        currentChunk++;
                        if (currentChunk * CHUNK_SIZE < sortedDates.length) {
                            // 继续渲染下一块
                            await new Promise(resolve => requestAnimationFrame(resolve));
                            await renderChunk();
                        } else {
                            // 渲染完全完成
                            if (loadingEl) loadingEl.classList.add('hidden');
                            if (canvas.parentElement) {
                                canvas.parentElement.classList.remove('loading-shimmer');
                            }
                            
                            // 只有在没有被新任务取代时才触发动画
                            if (window.currentHeatmapRenderId === renderId) {
                                requestAnimationFrame(() => {
                                    canvas.style.opacity = '1';
                                    canvas.classList.add('animate-graph-entry');
                                });
                            }
                        }
                    };
                    
                    await renderChunk();
                    
                    // 渲染完成后初始化提示框逻辑
                    const tooltip = document.getElementById('heatmapTooltip');
                    const tooltipTitle = document.getElementById('heatmapTooltipTitle');
                    const tooltipValue = document.getElementById('heatmapTooltipValue');
                    if (tooltip && tooltipTitle && tooltipValue && !canvas.dataset.heatmapHoverBound) {
                        canvas.dataset.heatmapHoverBound = '1';
                        const unit = appData?.config?.dataType === 'cumulativeEnergy' ? 'kWh' : 'kW';
                        const updateTooltip = (clientX, clientY) => {
                            const layout = appData.visualization.heatmapLayout;
                            if (!layout || !layout.dateGroups) return;
                            const rect = canvas.getBoundingClientRect();
                            const x = clientX - rect.left;
                            const y = clientY - rect.top;
                            const inX = x >= layout.margin.left && x <= layout.margin.left + 24 * layout.cellWidth;
                            const inY = y >= layout.margin.top && y <= layout.margin.top + layout.sortedDates.length * layout.cellHeight;
                            if (!inX || !inY) {
                                tooltip.classList.add('hidden');
                                return;
                            }
                            const hour = Math.min(23, Math.max(0, Math.floor((x - layout.margin.left) / layout.cellWidth)));
                            const rowIndex = Math.min(layout.sortedDates.length - 1, Math.max(0, Math.floor((y - layout.margin.top) / layout.cellHeight)));
                            const dateStr = layout.sortedDates[rowIndex];
                            const value = layout.dateGroups?.[dateStr]?.[hour];
                            const displayValue = typeof value === 'number' && isFinite(value) ? value : 0;

                            tooltipTitle.textContent = `${dateStr}  ${hour}:00`;
                            tooltipValue.textContent = `${displayValue.toFixed(2)} ${unit}`;

                            const left = Math.min(rect.width - 10, Math.max(10, x + 12));
                            const top = Math.min(rect.height - 10, Math.max(10, y + 12));
                            
                            // 使用 transition 实现平滑移动
                            tooltip.style.transition = 'left 0.1s ease-out, top 0.1s ease-out, opacity 0.2s ease-out';
                            tooltip.style.left = left + 'px';
                            tooltip.style.top = top + 'px';
                            tooltip.classList.remove('hidden');
                        };
                        let raf = 0;
                        canvas.addEventListener('mousemove', (e) => {
                            if (raf) return;
                            raf = requestAnimationFrame(() => {
                                raf = 0;
                                updateTooltip(e.clientX, e.clientY);
                            });
                        });
                        canvas.addEventListener('mouseleave', () => tooltip.classList.add('hidden'));
                    }
                } catch (error) {
                    console.error('热力图渲染失败:', error);
                    if (loadingEl) loadingEl.classList.add('hidden');
                    if (noDataEl) noDataEl.classList.remove('hidden');
                }
            };
            
            // 使用 requestIdleCallback 或 setTimeout 调度渲染
            if (window.requestIdleCallback) {
                window.requestIdleCallback(renderTask, { timeout: 100 });
            } else {
                setTimeout(renderTask, 16); // 约 60fps
            }
        }
        
        // 优化的数据集创建函数
        function createOptimizedDatasets(filteredData, startHour, endHour) {
            console.log('=== createOptimizedDatasets 调试信息 ===');
            console.log('输入数据总数:', filteredData.length);
            
            const datasets = [];
            
            // 使用 DataUtils 按日期分组数据
            const dateGroupsMap = DataUtils.groupBy(filteredData, row => row.date);
            console.log('分组后的日期数量:', dateGroupsMap.size);
            console.log('分组的日期:', Array.from(dateGroupsMap.keys()));
             
            // 为每个日期创建数据集（优化版本）
            let colorIndex = 0;
            let sortedDates = Array.from(dateGroupsMap.keys()).sort();
            
            // 如果不是“显示全部日期”模式，且不是“汇总模式”，则只显示第一个日期
            if (!appData.visualization.allCurvesMode && !appData.visualization.summaryMode && sortedDates.length > 0) {
                // 如果用户已经手动选择了日期范围，则按照选择的日期显示
                // 否则默认只显示第一个日期
                if (appData.visualization.selectedDates && appData.visualization.selectedDates.length > 0) {
                    sortedDates = sortedDates.filter(d => appData.visualization.selectedDates.includes(d));
                } else {
                    sortedDates = [sortedDates[0]];
                }
            }
            
            sortedDates.forEach(dateKey => {
                const rows = dateGroupsMap.get(dateKey);
                
                // 合并同一日期的所有行数据
                const mergedHourlyData = new Array(24).fill(null);
                
                rows.forEach(row => {
                    for (let hour = 0; hour < 24; hour++) {
                        if (row.hourlyData[hour] !== null) {
                            if (mergedHourlyData[hour] !== null) {
                                // 如果已有数据，取总和
                                mergedHourlyData[hour] = mergedHourlyData[hour] + row.hourlyData[hour];
                            } else {
                                mergedHourlyData[hour] = row.hourlyData[hour];
                            }
                        }
                    }
                });
                
                // 根据时间焦段设置过滤数据
                const filteredHourlyData = mergedHourlyData.slice(startHour, endHour + 1);
                
                // 为每个日期生成不同的颜色
                const color = getColor(colorIndex % getColorCount());
                colorIndex++;
                
                // 根据用户设置确定线条样式
                let borderDash = [];
                if (appData.visualization.lineStyle === 'dashed') {
                    borderDash = [5, 5];
                } else if (appData.visualization.lineStyle === 'dotted') {
                    borderDash = [2, 2];
                }
                
                // 根据用户设置确定是否显示数据点
                const pointRadius = appData.visualization.showPoints ? 4 : 0;
                const pointHoverRadius = 6;
                
                // 根据用户设置确定曲线平滑度
                const tension = appData.visualization.smoothCurve ? appData.visualization.curveTension : 0;
                
                // 创建渐变背景
                const gradient = appData.chart.ctx.createLinearGradient(0, 0, 0, 400);
                gradient.addColorStop(0, color.replace('rgb', 'rgba').replace(')', ', 0.4)'));
                gradient.addColorStop(1, color.replace('rgb', 'rgba').replace(')', ', 0.05)'));
                
                // 创建优化的数据集配置
                const datasetConfig = createDatasetConfig(dateKey, filteredHourlyData, color, borderDash, pointRadius, pointHoverRadius, tension);
                datasets.push(datasetConfig);
            });
             
            return datasets;
        }
        
        // 新增：汇总模式数据集（最高/最低 + 四季度平均 + 1~12月月平均）
        function createSummaryDatasets(filteredData, startHour, endHour) {
            // 按“天”聚合，同时记录每小时是否有数据（作为是否计入该小时分母的依据）
            const dayMap = new Map();
            for (const row of filteredData) {
                const dateObj = row.dateObj || parseDate(row.date, appData.config.dateFormat);
                if (!dateObj || isNaN(dateObj)) continue;
                const dateKey = formatDateForInput(dateObj);
                if (!dayMap.has(dateKey)) {
                    dayMap.set(dateKey, {
                        dateObj,
                        hourly: Array(24).fill(0),  // 当天每小时总量（若同一日期有多条记录则累加）
                        hrCount: Array(24).fill(0), // 当天每小时是否有数据：>0 表示该天该小时有数据
                        dailyTotal: 0
                    });
                }
                const agg = dayMap.get(dateKey);
                for (let h = 0; h < 24; h++) {
                    const v = row.hourlyData[h];
                    if (v != null && !isNaN(v)) {
                        agg.hourly[h] += v;
                        agg.hrCount[h] += 1; // 标记该天该小时有数据
                        agg.dailyTotal += v;
                    }
                }
            }
            if (dayMap.size === 0) return [];
            const days = Array.from(dayMap.values());

            // 最高/最低用电日（以当天24小时总量为判定）
            let maxDay = days[0];
            let minDay = days[0];
            for (const d of days) {
                if (d.dailyTotal > maxDay.dailyTotal) maxDay = d;
                if (d.dailyTotal < minDay.dailyTotal) minDay = d;
            }
            const maxDaySeries = maxDay.hourly.slice(startHour, endHour + 1);
            const minDaySeries = minDay.hourly.slice(startHour, endHour + 1);

            // 季度与月份分桶（按“天”作为单位）
            const quarterBuckets = [[], [], [], []]; // 每个元素是 day 对象
            const monthBuckets = Array.from({ length: 12 }, () => []);
            for (const d of days) {
                const m = d.dateObj.getMonth() + 1; // 1~12
                const q = m <= 3 ? 0 : m <= 6 ? 1 : m <= 9 ? 2 : 3;
                quarterBuckets[q].push(d);
                monthBuckets[m - 1].push(d);
            }

            // 平均函数：以“有数据的天数”为分母（hrCount[h] > 0 视为该天该小时有数据）
            function avgSeries(dayList) {
                const len = endHour - startHour + 1;
                if (dayList.length === 0) return Array(len).fill(null);
                const sum = Array(24).fill(0);
                const cntDays = Array(24).fill(0);
                for (const d of dayList) {
                    for (let h = 0; h < 24; h++) {
                        if (d.hrCount[h] > 0) { // 该天该小时有数据
                            sum[h] += d.hourly[h];
                            cntDays[h] += 1; // 分母按“天数”计数
                        }
                    }
                }
                const avg = sum.map((v, i) => (cntDays[i] > 0 ? v / cntDays[i] : null));
                return avg.slice(startHour, endHour + 1);
            }

            // 四季度平均
            const q1Series = avgSeries(quarterBuckets[0]);
            const q2Series = avgSeries(quarterBuckets[1]);
            const q3Series = avgSeries(quarterBuckets[2]);
            const q4Series = avgSeries(quarterBuckets[3]);

            // 1~12月平均
            const monthSeries = monthBuckets.map(list => avgSeries(list));

            function makeDataset(label, data, color, dash = []) {
                // 让汇总曲线联动样式设置：线型与数据点
                let computedDash = [];
                if (appData.visualization.lineStyle === 'dashed') {
                    computedDash = [5, 5];
                } else if (appData.visualization.lineStyle === 'dotted') {
                    computedDash = [2, 2];
                } else {
                    // UI 选择 solid 时，保留传入的 dash（如季度/月份平均使用虚线的默认样式）
                    computedDash = dash;
                }
                const showPoints = !!appData.visualization.showPoints;
                const pr = showPoints ? 4 : 0;
                return {
                    label,
                    data,
                    borderColor: color,
                    backgroundColor: 'transparent',
                    borderWidth: 3,
                    borderDash: computedDash,
                    borderCapStyle: 'round',
                    borderJoinStyle: 'round',
                    pointRadius: pr,
                    pointHoverRadius: 6,
                    pointStyle: 'circle',
                    pointBackgroundColor: color,
                    pointBorderColor: color,
                    pointBorderWidth: 2,
                    tension: appData.visualization.smoothCurve ? appData.visualization.curveTension : 0,
                    fill: false,
                    yAxisID: 'y',
                    spanGaps: true
                };
            }

            // UI/UX Pro Max 推荐：数据可视化配色方案 - 专业且易于区分
            const monthColors = [
                'rgb(30, 64, 175)',   // 1月 - 深蓝
                'rgb(8, 145, 178)',   // 2月 - 青色
                'rgb(13, 148, 136)',  // 3月 - 青绿
                'rgb(5, 150, 105)',   // 4月 - 绿色
                'rgb(14, 165, 233)',  // 5月 - 天蓝
                'rgb(59, 130, 246)',  // 6月 - 蓝色
                'rgb(99, 102, 241)',  // 7月 - 靛蓝
                'rgb(124, 58, 237)',  // 8月 - 紫色
                'rgb(236, 72, 153)',  // 9月 - 粉色
                'rgb(245, 158, 11)',  // 10月 - 琥珀色
                'rgb(234, 88, 12)',   // 11月 - 橙色
                'rgb(225, 29, 72)'    // 12月 - 玫瑰色
            ];

            const datasets = [];
            const subMode = appData.visualization.summarySubMode || 'all';

            // 1. 最高/最低用电日 + 平均负荷 (extreme)
            if (subMode === 'all' || subMode === 'extreme') {
                datasets.push(makeDataset(`最高用电日 (${formatDateForInput(maxDay.dateObj)})`, maxDaySeries, 'rgb(220, 38, 38)'));  // 红色-最高
                datasets.push(makeDataset(`最低用电日 (${formatDateForInput(minDay.dateObj)})`, minDaySeries, 'rgb(37, 99, 235)'));   // 蓝色-最低
                
                // 计算全局平均负荷
                const globalAvgSeries = avgSeries(days);
                datasets.push(makeDataset('全周期平均负荷', globalAvgSeries, 'rgb(71, 85, 105)', [8, 4])); // 深灰-平均
            }

            // 2. 季度平均 (quarter)
            if (subMode === 'all' || subMode === 'quarter') {
                datasets.push(makeDataset('Q1平均(1-3月)', q1Series, 'rgb(5, 150, 105)', [6, 4]));      // 绿色-春季
                datasets.push(makeDataset('Q2平均(4-6月)', q2Series, 'rgb(245, 158, 11)', [6, 4]));     // 琥珀色-夏季
                datasets.push(makeDataset('Q3平均(7-9月)', q3Series, 'rgb(234, 88, 12)', [6, 4]));      // 橙色-秋季
                datasets.push(makeDataset('Q4平均(10-12月)', q4Series, 'rgb(30, 64, 175)', [6, 4]));    // 深蓝-冬季
            }

            // 3. 月平均 (month)
            if (subMode === 'all' || subMode === 'month') {
                for (let m = 1; m <= 12; m++) {
                    if (monthSeries[m - 1].some(v => v !== null)) {
                        datasets.push(
                            makeDataset(`${m}月平均`, monthSeries[m - 1], monthColors[m - 1], [4, 2])
                        );
                    }
                }
            }
            return datasets;
        }
        
        // 创建数据集配置
        function createDatasetConfig(dateKey, filteredHourlyData, color, borderDash, pointRadius, pointHoverRadius, tension) {
            return {
                label: dateKey,
                data: filteredHourlyData,
                borderColor: color,
                backgroundColor: 'transparent',
                borderWidth: 3,
                borderDash: borderDash,
                borderCapStyle: 'round',
                borderJoinStyle: 'round',
                segment: {
                    borderColor: ctx => {
                        const ds = ctx?.chart?.data?.datasets?.[ctx.datasetIndex];
                        const c = ds && ds.borderColor ? ds.borderColor : color;
                        return ctx.p0.parsed.y !== null && ctx.p1.parsed.y !== null ? c : 'rgba(0,0,0,0)';
                    },
                    borderDash: ctx => {
                        return ctx.p0.parsed.y !== null && ctx.p1.parsed.y !== null ? borderDash : [];
                    }
                },
                pointRadius: pointRadius,
                pointHoverRadius: pointHoverRadius,
                pointStyle: 'circle',
                pointBorderWidth: 2,
                pointBackgroundColor: color,
                pointBorderColor: color,
                pointHoverBackgroundColor: color,
                pointHoverBorderColor: '#fff',
                pointHoverBorderWidth: 0,
                tension: tension,
                fill: false,
                yAxisID: 'y',
                spanGaps: true
            };
        }
        
        // 更新图表数据和选项
        function updateChartDataAndOptions(labels, datasets, filteredData, startHour, endHour) {
            // 更新图表数据
            appData.chart.data.labels = labels;
            appData.chart.data.datasets = datasets;

            if (appData.visualization.summaryMode) {
                for (let i = 0; i < datasets.length; i++) {
                    datasets[i].hidden = false;
                    const meta = appData.chart.getDatasetMeta(i);
                    if (meta) meta.hidden = false;
                }
            } else {
                const curveCount = datasets.length;
                let alpha = 0.7;
                let width = 2.5;
                let pr = appData.visualization.showPoints ? 3 : 0;
                let phr = appData.visualization.showPoints ? 5 : 0;

                if (curveCount >= 50) {
                    alpha = 0.15;
                    width = 0.8;
                    pr = appData.visualization.showPoints ? 1 : 0;
                    phr = 3;
                } else if (curveCount >= 30) {
                    alpha = 0.25;
                    width = 1.2;
                    pr = appData.visualization.showPoints ? 1.5 : 0;
                    phr = 4;
                } else if (curveCount >= 10) {
                    alpha = 0.5;
                    width = 1.8;
                    pr = appData.visualization.showPoints ? 2 : 0;
                    phr = 4.5;
                }

                for (let i = 0; i < datasets.length; i++) {
                    const ds = datasets[i];
                    ds.borderWidth = width;
                    ds.pointRadius = pr;
                    ds.pointHoverRadius = phr;
                    if (typeof ds.borderColor === 'string') {
                        ds.borderColor = toRgba(ds.borderColor, alpha);
                        ds.pointBackgroundColor = toRgba(ds.pointBackgroundColor || ds.borderColor, Math.min(1, alpha + 0.15));
                        ds.pointBorderColor = toRgba(ds.pointBorderColor || ds.borderColor, Math.min(1, alpha + 0.15));
                    }
                }
            }
            

            
            // 更新图表选项
            // 更新图表标题以反映时间焦段和日期选择状态
            const selectedCount = appData.visualization.selectedDates.length;
            const totalCount = appData.processedData.length;
            if (appData.visualization.summaryMode) {
                appData.chart.options.plugins.title.text = `${startHour}:00 - ${endHour}:00 汇总曲线（最高/最低 + 季度平均 + 月平均）`;
            } else {
                appData.chart.options.plugins.title.text = `${startHour}:00 - ${endHour}:00 负荷曲线 (已选择 ${selectedCount}/${totalCount} 个日期)`;
            }
            
            // 折线图模式下的选项
            appData.chart.options.scales.x.stacked = false;
            appData.chart.options.scales.y.stacked = false;
            appData.chart.options.plugins.tooltip.mode = 'index';
            appData.chart.options.plugins.tooltip.intersect = false;
            
            // 更新辅助Y轴
            appData.chart.options.scales.y1 = {
                display: appData.visualization.secondaryAxis,
                position: 'right',
                title: {
                    display: true,
                    text: '功率 (kW)',
                    font: {
                        size: 14
                    }
                },
                beginAtZero: true,
                grid: {
                    drawOnChartArea: false // 不在图表区域绘制网格线
                }
            };
            
            // 使用优化的更新方式
            updateChartWithAnimation();
            
            // 汇总模式下隐藏分页控件
            if (appData.visualization.summaryMode) {
                hidePaginationControls();
            }
            
            // 更新统计信息 - 使用 getOverviewData 以展示全部日期
            updateStatistics(getOverviewData(), startHour, endHour);

            updateSummaryCurveExportAvailability();
            

            
            // 隐藏加载状态
            const chartLoadingFinalElement = document.getElementById('chartLoading');
            if (chartLoadingFinalElement && chartLoadingFinalElement.parentNode) {
                chartLoadingFinalElement.classList.add('hidden');
            }
            
            // 显示数据状态
            updateDataStatus(datasets.length);
        }
        
        // 优化的图表更新函数
        function updateChartWithAnimation() {
            try {
                if (appData.chart && typeof appData.chart.update === 'function') {
                    // 使用平滑动画更新
                    appData.chart.update({
                        duration: CHART_PERFORMANCE.animationDuration,
                        easing: 'easeInOutQuart'
                    });
                }
            } catch (error) {
                console.error('图表更新错误:', error);
                handleChartError(error);
            }
        }
        
        // 图表错误处理
        function handleChartError(error) {
            try {
                if (appData.chart) {
                    appData.chart.destroy();
                    appData.chart = null;
                }
                initializeVisualization();
            } catch (secondError) {
                console.error('图表重新初始化失败:', secondError);
                showNotification('错误', '图表更新失败，请刷新页面重试', 'error');
            }
        }
        
        // 分页控件管理
        function showPaginationControls() {
            let paginationContainer = document.getElementById('chartPagination');
            if (!paginationContainer) {
                paginationContainer = createPaginationContainer();
            }
            
            updatePaginationControls();
            if (paginationContainer && paginationContainer.parentNode) {
                paginationContainer.classList.remove('hidden');
            }
        }
        
        function hidePaginationControls() {
            const paginationContainer = document.getElementById('chartPagination');
            if (paginationContainer && paginationContainer.parentNode) {
                paginationContainer.classList.add('hidden');
            }
        }
        
        function createPaginationContainer() {
            const container = document.createElement('div');
            container.id = 'chartPagination';
            container.className = 'flex flex-col gap-3 mt-4 p-4 rounded-2xl border border-slate-200 bg-white/80 shadow-sm backdrop-blur md:flex-row md:items-center md:justify-between';
            
            container.innerHTML = `
                <div class="flex flex-wrap items-center gap-x-3 gap-y-1">
                    <span class="text-sm font-semibold text-slate-700">第 <span id="currentPageInfo">1</span> / <span id="totalPagesInfo">1</span> 页</span>
                    <span class="text-xs font-semibold text-slate-500">共 <span id="totalItemsInfo">0</span> 条</span>
                </div>
                <div class="flex items-center gap-2">
                    <button id="prevPageBtn" class="btn-secondary px-4 py-2 text-xs" onclick="changePage(-1)">
                        <i class="fa fa-chevron-left"></i> 上一页
                    </button>
                    <button id="nextPageBtn" class="btn-secondary px-4 py-2 text-xs" onclick="changePage(1)">
                        下一页 <i class="fa fa-chevron-right"></i>
                    </button>
                </div>
            `;
            
            // 插入到图表容器后面
            const chartContainer = document.querySelector('.chart-container');
            if (chartContainer) {
                chartContainer.parentNode.insertBefore(container, chartContainer.nextSibling);
            }
            
            return container;
        }
        
        function updatePaginationControls() {
            // 检查元素是否存在，避免null错误
            const currentPageInfo = document.getElementById('currentPageInfo');
            const totalPagesInfo = document.getElementById('totalPagesInfo');
            const totalItemsInfo = document.getElementById('totalItemsInfo');
            const prevPageBtn = document.getElementById('prevPageBtn');
            const nextPageBtn = document.getElementById('nextPageBtn');
            
            if (currentPageInfo && currentPageInfo.parentNode) currentPageInfo.textContent = chartPagination.currentPage + 1;
            if (totalPagesInfo && totalPagesInfo.parentNode) totalPagesInfo.textContent = chartPagination.totalPages;
            if (totalItemsInfo && totalItemsInfo.parentNode) totalItemsInfo.textContent = chartPagination.totalItems;
            
            if (prevPageBtn) prevPageBtn.disabled = chartPagination.currentPage === 0;
            if (nextPageBtn) nextPageBtn.disabled = chartPagination.currentPage >= chartPagination.totalPages - 1;
        }
        
        // 分页切换函数（支持相对/绝对跳转）
        function changePage(directionOrPage, isAbsolute = false) {
            let newPage;
            if (isAbsolute) {
                newPage = directionOrPage;
            } else {
                newPage = chartPagination.currentPage + directionOrPage;
            }
            if (newPage >= 0 && newPage < chartPagination.totalPages) {
                chartPagination.currentPage = newPage;
                updateChart();
            }
        }
        
        // 数据状态更新
        function updateDataStatus(datasetCount) {
            // 柱状图始终显示（增加空值保护）
            const barCanvas = document.getElementById('dailyTotalBarChart');
            if (barCanvas && barCanvas.parentElement && barCanvas.parentElement.parentElement) {
                const barChartContainer = barCanvas.parentElement.parentElement;
                barChartContainer.style.display = 'block';
            }
            updateDailyTotalBarChart(getFilteredData());
            
            // 如果没有数据，显示无数据消息，否则隐藏
            const noDataMsgElement = document.getElementById('noDataMessage');
            if (datasetCount === 0) {
                if (noDataMsgElement && noDataMsgElement.parentNode) {
                    noDataMsgElement.classList.remove('hidden');
                }
            } else {
                if (noDataMsgElement && noDataMsgElement.parentNode) {
                    noDataMsgElement.classList.add('hidden');
                }
            }
        }

        /**
         * 更新步骤指示灯状态
         * @param {number} step 步骤 1, 2, 3
         * @param {boolean} active 是否激活（变绿）
         */
        function updateStepStatus(step, active) {
            const indicator = document.getElementById(`step${step}Indicator`);
            if (indicator) {
                if (active) {
                    indicator.classList.remove('bg-slate-400');
                    indicator.classList.add('bg-green-500', 'animate-pulse');
                } else {
                    indicator.classList.add('bg-slate-400');
                    indicator.classList.remove('bg-green-500', 'animate-pulse');
                }
            }
        }

        function updateStatistics(unused_data, startHour, endHour) {
            // 强制使用全量数据，不再受曲线筛选联动
            const displayData = getOverviewData();
            
            if (!displayData || displayData.length === 0) {
                const unifiedStats = document.getElementById('unifiedStats');
                if (unifiedStats) {
                    unifiedStats.innerHTML = `
                        <div class="flex flex-col items-center justify-center py-20 text-slate-400">
                            <div class="mb-4 flex h-16 w-16 items-center justify-center rounded-2xl bg-slate-50 border border-slate-100">
                                <i class="fa fa-chart-pie text-2xl text-slate-300"></i>
                            </div>
                            <p class="text-sm font-medium">暂无数据，请先导入表格</p>
                        </div>
                    `;
                }
                return;
            }

            // 如果没有传入时间焦段设置，使用appData中的设置
            if (typeof startHour !== 'number' || typeof endHour !== 'number') {
                startHour = appData.visualization.focusStartTime;
                endHour = appData.visualization.focusEndTime;
            }

            const periodHours = endHour - startHour + 1;
            const formatNum = (n, fraction = 2) => smartFormatNumber(n, fraction);
            const formatPct = (n) => isFinite(n) ? `${Number(n).toFixed(1)}%` : '-';

            // 1. 数据预处理
            const meteringPointSet = new Set();
            displayData.forEach(d => {
                if (d && d.meteringPoint !== null && d.meteringPoint !== undefined) {
                    const t = String(d.meteringPoint).trim();
                    if (t) meteringPointSet.add(t);
                }
            });
            const showMeteringPointColumn = meteringPointSet.size > 1;

            let totalPeriodAll = 0;
            let totalPeriodGenerationAll = 0;
            let totalDailyAll = 0;
            let maxPeriodTotal = -Infinity;
            let maxPeriodKey = '';
            let minPeriodTotal = Infinity;
            let minPeriodKey = '';

            const perRowStats = displayData.map((dateData) => {
                let periodTotal = 0;
                let periodGeneration = 0;
                let validHours = 0;
                
                // 统计时段内的正向用电和反向上网
                for (let hour = startHour; hour <= endHour; hour++) {
                    const v = dateData.hourlyData?.[hour];
                    if (v !== null && v !== undefined) {
                        if (v > 0) periodTotal += v;
                        validHours++;
                    }
                    
                    const g = dateData.hourlyGeneration?.[hour];
                    if (g !== null && g !== undefined && g > 0) {
                        periodGeneration += g;
                    }
                }
                
                totalPeriodAll += periodTotal;
                totalPeriodGenerationAll += periodGeneration;
                totalDailyAll += dateData.dailyTotal;
                
                const key = showMeteringPointColumn ? `${dateData.date} · ${String(dateData.meteringPoint ?? '').trim()}` : dateData.date;
                if (periodTotal > maxPeriodTotal) {
                    maxPeriodTotal = periodTotal;
                    maxPeriodKey = key;
                }
                if (periodTotal < minPeriodTotal) {
                    minPeriodTotal = periodTotal;
                    minPeriodKey = key;
                }
                
                return {
                    date: dateData.date,
                    meteringPoint: String(dateData.meteringPoint ?? '').trim(),
                    periodTotal,
                    periodGeneration,
                    dailyPercentage: dateData.dailyTotal > 0 ? (periodTotal / dateData.dailyTotal) * 100 : 0,
                    hourlyAverage: validHours > 0 ? periodTotal / validHours : 0,
                    dailyTotal: dateData.dailyTotal
                };
            });

            // 2. 排序处理
            const sortVal = window._lastStatsSort || 'date_asc';
            perRowStats.sort((a, b) => {
                if (sortVal === 'date_asc') return a.date.localeCompare(b.date);
                if (sortVal === 'date_desc') return b.date.localeCompare(a.date);
                if (sortVal === 'period_desc') return b.periodTotal - a.periodTotal;
                if (sortVal === 'daily_desc') return b.dailyTotal - a.dailyTotal;
                return 0;
            });

            // 3. 全局统计
            const overallPeriodPercentage = totalDailyAll > 0 ? (totalPeriodAll / totalDailyAll) * 100 : 0;
            const avgHourlyInPeriod = (displayData.length * periodHours) > 0 ? (totalPeriodAll / (displayData.length * periodHours)) : 0;
            const hasGenerationData = totalPeriodGenerationAll > 0;

            // 4. 构建 HTML 内容
            const gridColsClass = hasGenerationData ? 'xl:grid-cols-7' : 'xl:grid-cols-6';
            
            let html = `
                <div class="flex flex-col gap-6 p-6">
                    <!-- 顶部标题与筛选 -->
                    <div class="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                        <div class="flex items-center gap-4">
                            <div class="flex h-12 w-12 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600 shadow-sm border border-indigo-100/50">
                                <i class="fa fa-chart-pie text-xl"></i>
                            </div>
                            <div>
                                <h3 class="text-lg font-bold text-slate-900">统计分析面板</h3>
                                <div class="flex items-center gap-2 mt-0.5">
                                    <span class="inline-flex items-center rounded-md bg-slate-100 px-2 py-0.5 text-[10px] font-bold text-slate-600 uppercase tracking-wider">${startHour}:00 - ${endHour}:00</span>
                                    <span class="text-[10px] font-medium text-slate-400">基于全量数据统计</span>
                                </div>
                            </div>
                        </div>
                        
                        <div class="flex flex-wrap items-center gap-3">
                            <div class="flex items-center gap-2 rounded-xl bg-slate-50 p-1 border border-slate-200/60 shadow-sm">
                                <span class="pl-2 text-[11px] font-bold text-slate-500 uppercase">排序</span>
                                <select id="statsSortSelect" class="rounded-lg border-none bg-transparent px-2 py-1 text-xs font-bold text-slate-700 outline-none cursor-pointer focus:ring-0">
                                    <option value="date_asc">日期升序</option>
                                    <option value="date_desc">日期降序</option>
                                    <option value="period_desc">时段用电降序</option>
                                    <option value="daily_desc">日总用电降序</option>
                                </select>
                            </div>
                            <div class="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-indigo-50/50 border border-indigo-100/50 text-indigo-600 font-bold text-xs">
                                <i class="fa fa-calendar-check opacity-70"></i>
                                <span>${displayData.length} 天</span>
                            </div>
                        </div>
                    </div>

                    <!-- 核心指标卡片 (Glassmorphism) -->
                    <div class="grid grid-cols-2 gap-4 md:grid-cols-3 lg:grid-cols-4 ${gridColsClass}">
                        <div class="group relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:border-blue-200/60">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-blue-500 shadow-[0_0_8px_rgba(59,130,246,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">时段总用电</div>
                                </div>
                                <div class="mt-2 flex items-baseline gap-1">
                                    <span class="text-xl font-black text-slate-900">${formatNum(totalPeriodAll)}</span>
                                    <span class="text-[10px] font-bold text-slate-400 uppercase">kWh</span>
                                </div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">选定时段内的累计用电总量</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-blue-50/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-bolt"></i>
                            </div>
                        </div>

                        ${hasGenerationData ? `
                        <div class="group relative overflow-hidden rounded-2xl border border-emerald-200 bg-emerald-50/30 p-5 shadow-sm transition-all hover:shadow-md hover:border-emerald-300">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-emerald-600 uppercase tracking-wider">上网电量 (估算)</div>
                                    <div class="group/note relative">
                                        <i class="fa fa-circle-info text-[10px] text-emerald-400 cursor-help"></i>
                                        <div class="absolute bottom-full left-0 mb-2 hidden group-hover/note:block w-48 p-2 bg-slate-800 text-white text-[10px] rounded-lg shadow-xl z-50 leading-relaxed">
                                            <p class="font-bold mb-1">上网电量说明：</p>
                                            上网电量（发电量）统计来源于数据源中的负值（反向负载），计算逻辑为：筛选负值得到。
                                        </div>
                                    </div>
                                </div>
                                <div class="mt-2 flex items-baseline gap-1">
                                    <span class="text-xl font-black text-emerald-700">${formatNum(totalPeriodGenerationAll)}</span>
                                    <span class="text-[10px] font-bold text-emerald-500 uppercase">kWh</span>
                                </div>
                                <div class="mt-1 text-[9px] text-emerald-600/60 font-medium">选定时段内的累计上网电量</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-emerald-200/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-solar-panel"></i>
                            </div>
                        </div>` : ''}

                        <div class="group relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:border-indigo-200/60">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-indigo-500 shadow-[0_0_8px_rgba(99,102,241,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">全部日总用电</div>
                                </div>
                                <div class="mt-2 flex items-baseline gap-1">
                                    <span class="text-xl font-black text-slate-900">${formatNum(totalDailyAll)}</span>
                                    <span class="text-[10px] font-bold text-slate-400 uppercase">kWh</span>
                                </div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">所选日期范围内的全天总用电</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-indigo-50/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-calendar-alt"></i>
                            </div>
                        </div>
                    `;

                    // 时段占比卡片 - 语义化颜色
                    const getPctCardStyles = (pct) => {
                        if (pct >= 80) return { border: 'hover:border-rose-200', dot: 'bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]', icon: 'text-rose-50/40', text: 'text-rose-600' };
                        if (pct >= 50) return { border: 'hover:border-orange-200', dot: 'bg-orange-500 shadow-[0_0_8px_rgba(249,115,22,0.5)]', icon: 'text-orange-50/40', text: 'text-orange-600' };
                        if (pct >= 20) return { border: 'hover:border-amber-200', dot: 'bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]', icon: 'text-amber-50/40', text: 'text-amber-600' };
                        return { border: 'hover:border-emerald-200', dot: 'bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]', icon: 'text-emerald-50/40', text: 'text-emerald-600' };
                    };
                    const pctStyles = getPctCardStyles(overallPeriodPercentage);

                    html += `
                        <div class="group relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:shadow-md ${pctStyles.border}">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full ${pctStyles.dot}"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">时段占比</div>
                                </div>
                                <div class="mt-2">
                                    <span class="text-xl font-black ${pctStyles.text}">${formatPct(overallPeriodPercentage)}</span>
                                </div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">时段用电占全天总用电的比例</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 ${pctStyles.icon} text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-percentage"></i>
                            </div>
                        </div>

                        <div class="group relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:border-cyan-200/60">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-cyan-500 shadow-[0_0_8px_rgba(6,182,212,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">平均每小时</div>
                                </div>
                                <div class="mt-2 flex items-baseline gap-1">
                                    <span class="text-xl font-black text-slate-900">${formatNum(avgHourlyInPeriod)}</span>
                                    <span class="text-[10px] font-bold text-slate-400 uppercase">kWh</span>
                                </div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">选定时段内平均每小时的负荷</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-cyan-50/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-clock"></i>
                            </div>
                        </div>

                        <div class="group relative overflow-hidden rounded-2xl border border-slate-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:border-rose-200">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">最大时段日</div>
                                </div>
                                <div class="mt-2 truncate text-sm font-black text-slate-900" title="${maxPeriodKey}">${maxPeriodKey}</div>
                                <div class="text-[10px] font-bold text-rose-500">${formatNum(maxPeriodTotal)} kWh</div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">时段内用电量最高的一天</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-rose-50/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-arrow-trend-up"></i>
                            </div>
                        </div>

                        <div class="group relative overflow-hidden rounded-2xl border border-amber-200 bg-white p-5 shadow-sm transition-all hover:shadow-md hover:border-amber-200">
                            <div class="relative z-10">
                                <div class="flex items-center gap-2 mb-2">
                                    <span class="h-1.5 w-1.5 rounded-full bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]"></span>
                                    <div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider">最小时段日</div>
                                </div>
                                <div class="mt-2 truncate text-sm font-black text-slate-900" title="${minPeriodKey}">${minPeriodKey}</div>
                                <div class="text-[10px] font-bold text-amber-500">${formatNum(minPeriodTotal)} kWh</div>
                                <div class="mt-1 text-[9px] text-slate-400 font-medium">时段内用电量最低的一天</div>
                            </div>
                            <div class="absolute -right-2 -bottom-2 text-amber-50/40 text-4xl transform -rotate-12 transition-transform group-hover:scale-110 group-hover:rotate-0">
                                <i class="fa fa-arrow-trend-down"></i>
                            </div>
                        </div>
                    </div>

                    <!-- 数据详情表格 -->
                    <div class="relative isolate overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm transition-all transform-gpu" style="-webkit-mask-image: -webkit-radial-gradient(white, black); mask-image: radial-gradient(white, black); backface-visibility: hidden;">
                        <div class="overflow-x-auto scrollbar-thin" style="max-height: 520px;">
                            <table class="w-full border-collapse">
                                <thead class="sticky top-0 z-30">
                                    <tr class="bg-slate-50/95 backdrop-blur-md">
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">日期</th>
                                        ${showMeteringPointColumn ? `<th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">计量点</th>` : ''}
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">时段用电 (kWh)</th>
                                        ${hasGenerationData ? `<th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-emerald-600">上网电量 (kWh)</th>` : ''}
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">占全天</th>
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">占全年</th>
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">平均每小时</th>
                                        <th class="border-b border-slate-200 px-4 py-4 text-center text-[10px] font-black uppercase tracking-widest text-slate-500">日总用电 (kWh)</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100">
            `;

            try {
                const totalYearlyAll = displayData.reduce((acc, curr) => acc + (curr.dailyTotal || 0), 0);

                perRowStats.forEach((row, index) => {
                    const yearlyPercentage = totalYearlyAll > 0 ? (row.periodTotal / totalYearlyAll) * 100 : 0;
                    
                    // 根据比例大小决定颜色：小的时候绿色，大的时候红色
                    const getPctColor = (pct) => {
                        if (pct >= 80) return 'text-rose-600 font-black';
                        if (pct >= 50) return 'text-orange-500 font-bold';
                        if (pct >= 20) return 'text-amber-500 font-bold';
                        return 'text-emerald-500 font-bold';
                    };

                    const getBarColor = (pct) => {
                        if (pct >= 80) return 'bg-rose-500';
                        if (pct >= 50) return 'bg-orange-400';
                        if (pct >= 20) return 'bg-amber-400';
                        return 'bg-emerald-400';
                    };

                    html += `
                        <tr class="group transition-colors hover:bg-slate-50/80">
                            <td class="px-4 py-4 text-center text-xs font-bold text-slate-600 sticky left-0 z-10 bg-inherit group-hover:bg-slate-50/80">${row.date}</td>
                            ${showMeteringPointColumn ? `<td class="px-4 py-4 text-center text-[10px] font-medium text-slate-400">${row.meteringPoint}</td>` : ''}
                            <td class="px-4 py-4 text-center text-xs font-black text-slate-900">${formatNum(row.periodTotal)}</td>
                            ${hasGenerationData ? `<td class="px-4 py-4 text-center text-xs font-bold text-emerald-600 bg-emerald-50/20">${row.periodGeneration > 0 ? formatNum(row.periodGeneration) : '-'}</td>` : ''}
                            <td class="px-4 py-4 text-center">
                                <div class="inline-flex flex-col items-center">
                                    <span class="text-xs ${getPctColor(row.dailyPercentage)}">${formatPct(row.dailyPercentage)}</span>
                                    <div class="mt-1 h-1 w-12 overflow-hidden rounded-full bg-slate-100">
                                        <div class="h-full ${getBarColor(row.dailyPercentage)} transition-all duration-500" style="width: ${row.dailyPercentage}%"></div>
                                    </div>
                                </div>
                            </td>
                            <td class="px-4 py-4 text-center text-[10px] ${getPctColor(yearlyPercentage)}">${formatPct(yearlyPercentage)}</td>
                            <td class="px-4 py-4 text-center text-xs font-medium text-slate-500">${formatNum(row.hourlyAverage)}</td>
                            <td class="px-4 py-4 text-center text-xs font-bold text-slate-700 bg-slate-50/30">${formatNum(row.dailyTotal)}</td>
                        </tr>
                    `;
                });

                // 计算统计表的汇总和平均
                const statsSummary = perRowStats.reduce((acc, curr) => {
                    acc.totalPeriod += curr.periodTotal;
                    acc.totalGeneration += curr.periodGeneration;
                    acc.totalDaily += curr.dailyTotal;
                    return acc;
                }, { totalPeriod: 0, totalGeneration: 0, totalDaily: 0 });

                const statsAvg = {
                    periodTotal: perRowStats.length > 0 ? statsSummary.totalPeriod / perRowStats.length : 0,
                    periodGeneration: perRowStats.length > 0 ? statsSummary.totalGeneration / perRowStats.length : 0,
                    dailyTotal: perRowStats.length > 0 ? statsSummary.totalDaily / perRowStats.length : 0,
                    dailyPercentage: totalDailyAll > 0 ? (statsSummary.totalPeriod / totalDailyAll) * 100 : 0,
                    yearlyPercentage: totalYearlyAll > 0 ? (statsSummary.totalPeriod / totalYearlyAll) * 100 : 0,
                    hourlyAverage: avgHourlyInPeriod
                };

                // 添加汇总行
                html += `
                    <tr class="bg-slate-100 font-bold sticky bottom-[38px] z-40 backdrop-blur-md border-t-2 border-slate-200">
                        <td class="px-4 py-3 text-center text-xs text-slate-900 sticky left-0 z-50 bg-inherit">累计汇总</td>
                        ${showMeteringPointColumn ? `<td class="px-4 py-3 text-center text-[10px] text-slate-400">-</td>` : ''}
                        <td class="px-4 py-3 text-center text-xs text-slate-900 font-black">${formatNum(statsSummary.totalPeriod)}</td>
                        ${hasGenerationData ? `<td class="px-4 py-3 text-center text-xs font-bold text-emerald-700 bg-emerald-100/30">${formatNum(statsSummary.totalGeneration)}</td>` : ''}
                        <td class="px-4 py-3 text-center text-xs text-slate-900">${formatPct(overallPeriodPercentage)}</td>
                        <td class="px-4 py-3 text-center text-[10px] text-slate-400">100.0%</td>
                        <td class="px-4 py-3 text-center text-xs text-slate-700">-</td>
                        <td class="px-4 py-3 text-center text-xs font-black text-slate-900 bg-slate-100/20">${formatNum(statsSummary.totalDaily)}</td>
                    </tr>
                `;

                // 添加平均行
                html += `
                    <tr class="bg-slate-50/95 font-bold sticky bottom-0 z-40 backdrop-blur-md border-t border-slate-200 shadow-[0_-1px_2px_rgba(0,0,0,0.05)]">
                        <td class="px-4 py-3 text-center text-xs text-slate-700 sticky left-0 z-50 bg-inherit">平均数值</td>
                        ${showMeteringPointColumn ? `<td class="px-4 py-3 text-center text-[10px] text-slate-400">-</td>` : ''}
                        <td class="px-4 py-3 text-center text-xs text-slate-900 font-black">${formatNum(statsAvg.periodTotal)}</td>
                        ${hasGenerationData ? `<td class="px-4 py-3 text-center text-xs font-bold text-emerald-600 bg-emerald-50/10">${formatNum(statsAvg.periodGeneration)}</td>` : ''}
                        <td class="px-4 py-3 text-center text-xs text-slate-700">${formatPct(statsAvg.dailyPercentage)}</td>
                        <td class="px-4 py-3 text-center text-[10px] text-slate-400">${formatPct(statsAvg.yearlyPercentage)}</td>
                        <td class="px-4 py-3 text-center text-xs text-slate-600">${formatNum(statsAvg.hourlyAverage)}</td>
                        <td class="px-4 py-3 text-center text-xs font-black text-slate-700 bg-slate-100/20">${formatNum(statsAvg.dailyTotal)}</td>
                    </tr>
                `;

                html += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            `;
            } catch (e) {
                console.error('统计处理出错:', e);
            }

            const unifiedStats = document.getElementById('unifiedStats');
            if (unifiedStats) {
                unifiedStats.innerHTML = html;
                
                // 绑定排序事件
                const sortSelect = document.getElementById('statsSortSelect');
                if (sortSelect) {
                    sortSelect.value = window._lastStatsSort || 'date_asc';
                    sortSelect.addEventListener('change', (e) => {
                        window._lastStatsSort = e.target.value;
                        updateStatistics(displayData, startHour, endHour);
                    });
                }
            }

            try {
                updateGlobalSummary(displayData, startHour, endHour);
            } catch (e) {
                console.error('更新全局统计失败:', e);
            }
        }

        function updateGlobalSummary(displayData, startHour, endHour) {
            const summary = computeGlobalSummary(displayData, startHour, endHour);
            renderGlobalSummary(summary);
        }

        function computeGlobalSummary(displayData, startHour, endHour) {
            const safeData = Array.isArray(displayData) ? displayData : [];
            const hours = Math.max(0, (endHour - startHour + 1));
            const dateRange = (() => {
                if (!safeData.length) return '--';
                const dates = safeData.map(d => d.date).filter(Boolean).sort();
                if (!dates.length) return '--';
                return dates[0] === dates[dates.length - 1] ? dates[0] : `${dates[0]} ~ ${dates[dates.length - 1]}`;
            })();

            const meteringPointSet = new Set();
            safeData.forEach(d => {
                const t = String(d?.meteringPoint ?? '').trim();
                if (t) meteringPointSet.add(t);
            });
            const meteringPointLabel = meteringPointSet.size === 0
                ? '默认计量点'
                : meteringPointSet.size === 1
                    ? Array.from(meteringPointSet)[0]
                    : `${meteringPointSet.size} 个计量点`;

            let totalEnergy = 0; // 这里的totalEnergy现在仅代表正向用电量
            let totalPositiveEnergy = 0; // 正向电量（用电）
            let totalNegativeEnergy = 0; // 反向电量（发电）
            let hasNegativeValues = false; // 是否存在负值
            let validPoints = 0;
            let maxTotal = -Infinity;
            let maxDay = '';
            let minTotal = Infinity;
            let minDay = '';

            safeData.forEach(d => {
                let periodPositiveTotal = 0; // 单日正向总和
                
                // 1. 统计正向用电 (d.hourlyData 仅包含正值或0)
                for (let h = startHour; h <= endHour; h++) {
                    const v = d.hourlyData?.[h];
                    if (v !== null && v !== undefined) {
                        validPoints += 1;
                        if (v > 0) {
                            periodPositiveTotal += v;
                            totalPositiveEnergy += v;
                        }
                    }
                }
                
                // 2. 统计反向/发电 (d.hourlyGeneration 包含绝对值)
                if (d.hourlyGeneration) {
                    for (let h = startHour; h <= endHour; h++) {
                        const g = d.hourlyGeneration[h];
                        if (g !== null && g !== undefined && g > 0) {
                            totalNegativeEnergy += g;
                            hasNegativeValues = true;
                        }
                    }
                }
                
                // 总用电量只统计正向部分
                totalEnergy += periodPositiveTotal;
                
                // 极值判断：通常基于"用电负荷"判断高低
                if (periodPositiveTotal > maxTotal) {
                    maxTotal = periodPositiveTotal;
                    maxDay = d.date;
                }
                if (periodPositiveTotal < minTotal) {
                    minTotal = periodPositiveTotal;
                    minDay = d.date;
                }
            });

            const totalPoints = safeData.length * hours;
            const coverage = totalPoints > 0 ? (validPoints / totalPoints) * 100 : 0;
            
            // 平均负荷 = 总正向用电量 / 有效数据点数
            // 逻辑说明：
            // 1. 分子：totalEnergy (仅累加正向电量，负数计为0)
            // 2. 分母：validPoints (包含正数、0、负数的所有有效记录点)
            // 3. 结果：符合 (10+20+30+0)/4 的逻辑，且排除了 null 值的影响
            const avgLoad = validPoints > 0 ? (totalEnergy / validPoints) : 0;

            return {
                dateRange,
                hoursLabel: hours > 0 ? `${startHour}:00 - ${endHour}:00` : '--',
                meteringPointLabel,
                totalEnergy,
                totalPositiveEnergy,
                totalNegativeEnergy,
                hasNegativeValues,
                avgLoad,
                maxDay,
                maxTotal,
                minDay,
                minTotal,
                coverage
            };
        }

        function renderGlobalSummary(summary) {
            const s = summary || {};
            const fmt2 = (n) => smartFormatNumber(n, 2);
            const fmtPct = (n) => (typeof n === 'number' && isFinite(n)) ? Number(n).toFixed(1) : '—';

            const dr = document.getElementById('gsMetaDateRange');
            const hr = document.getElementById('gsMetaHours');
            if (dr) dr.textContent = s.dateRange || '--';
            if (hr) hr.textContent = s.hoursLabel || '--';

            const te = document.getElementById('gsTotalEnergy');
            const al = document.getElementById('gsAvgLoad');
            const md = document.getElementById('gsMaxDay');
            const mv = document.getElementById('gsMaxValue');
            const nd = document.getElementById('gsMinDay');
            const nv = document.getElementById('gsMinValue');

            // 1. 颜色指示器 (图片反馈修改)
            // 总用电量指示灯 - 保持紫色/蓝色
            const teIndicator = te?.closest('.group')?.querySelector('.rounded-full');
            if (teIndicator) {
                teIndicator.className = 'h-1.5 w-1.5 rounded-full bg-violet-500 shadow-[0_0_8px_rgba(139,92,246,0.5)]';
            }

            // 平均负荷指示灯 - 青色/蓝色
            const alIndicator = al?.closest('.group')?.querySelector('.rounded-full');
            if (alIndicator) {
                alIndicator.className = 'h-1.5 w-1.5 rounded-full bg-cyan-500 shadow-[0_0_8px_rgba(6,182,212,0.5)]';
            }

            // 最小负荷指示灯 - 琥珀色/黄色
            const ndIndicator = nd?.closest('.group')?.querySelector('.rounded-full');
            if (ndIndicator) {
                ndIndicator.className = 'h-1.5 w-1.5 rounded-full bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]';
            }

            // 最大负荷指示灯 - 玫瑰红/红色
            const mdIndicator = md?.closest('.group')?.querySelector('.rounded-full');
            if (mdIndicator) {
                mdIndicator.className = 'h-1.5 w-1.5 rounded-full bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.5)]';
            }

            // 智能显示总电量：如果有负值，显示更详细的信息
            if (te) {
                if (s.hasNegativeValues) {
                    te.innerHTML = `
                        <div class="flex flex-col">
                            <span class="cursor-help border-b border-dashed border-slate-300" title="仅统计正向用电，负值(发电)计为0">${fmt2(s.totalPositiveEnergy)}</span>
                            <span class="text-[10px] text-slate-400 font-normal mt-0.5" title="发电/反向: ${fmt2(s.totalNegativeEnergy)}">
                                (发电 ${fmt2(s.totalNegativeEnergy)})
                            </span>
                        </div>
                    `;
                } else {
                    te.textContent = fmt2(s.totalEnergy);
                }
            }
            
            if (al) {
                al.textContent = fmt2(s.avgLoad);
                if (s.hasNegativeValues) {
                    al.title = "计算公式：正向总用电量 / 有效时长 (负荷<0的时段按0计算)";
                    al.classList.add("cursor-help", "border-b", "border-dashed", "border-slate-300");
                } else {
                    al.title = "";
                    al.classList.remove("cursor-help", "border-b", "border-dashed", "border-slate-300");
                }
            }
            
            if (md) md.textContent = s.maxDay || '—';
            if (mv) mv.textContent = fmt2(s.maxTotal);
            if (nd) nd.textContent = s.minDay || '—';
            if (nv) nv.textContent = fmt2(s.minTotal);
        }

        // 导出部分函数
        let pendingExportRequest = null;

        function setExportProgress(percent, statusText = '导出中') {
            const wrap = document.getElementById('exportProgressWrap');
            const bar = document.getElementById('exportProgressBar');
            const text = document.getElementById('exportProgressText');
            const pillEl = document.getElementById('exportStatusPill');

            const safe = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
            if (wrap) wrap.classList.remove('hidden');
            if (bar) bar.style.width = `${safe}%`;
            if (text) text.textContent = `${safe}%`;
            if (pillEl) {
                pillEl.textContent = `${statusText} ${safe}%`;
                pillEl.classList.remove('border-indigo-200', 'bg-indigo-50', 'text-indigo-700');
                pillEl.classList.add('border-indigo-200', 'bg-indigo-50', 'text-indigo-700');
            }
        }

        function clearExportProgress() {
            const wrap = document.getElementById('exportProgressWrap');
            const bar = document.getElementById('exportProgressBar');
            const text = document.getElementById('exportProgressText');
            if (wrap) wrap.classList.add('hidden');
            if (bar) bar.style.width = '0%';
            if (text) text.textContent = '0%';
        }

        function updateExportStatus(title, fileName) {
            const panel = document.getElementById('exportStatusPanel');
            const titleEl = document.getElementById('exportStatusTitle');
            const metaEl = document.getElementById('exportStatusMeta');
            const timeEl = document.getElementById('exportStatusTime');
            const pillEl = document.getElementById('exportStatusPill');

            const now = new Date();
            const timeText = now.toLocaleString();

            clearExportProgress();
            if (titleEl) titleEl.textContent = title;
            if (metaEl) metaEl.textContent = fileName ? `文件：${fileName}` : '已生成导出文件';
            if (timeEl) timeEl.textContent = timeText;

            if (pillEl) {
                pillEl.textContent = '完成';
                pillEl.classList.remove('border-indigo-200', 'bg-indigo-50', 'text-indigo-700');
                pillEl.classList.add('border-indigo-200', 'bg-indigo-50', 'text-indigo-700');
            }

            if (panel) {
                panel.classList.remove('animate-scale-in');
                void panel.offsetWidth;
                panel.classList.add('animate-scale-in');
            }
        }

        function closeExportPreview() {
            if (!elements.exportPreviewModal) return;
            elements.exportPreviewModal.classList.add('hidden');
            elements.exportPreviewModal.setAttribute('aria-hidden', 'true');
            pendingExportRequest = null;
        }

        function confirmExportPreview() {
            const req = pendingExportRequest;
            closeExportPreview();
            if (req && typeof req.execute === 'function') {
                req.execute();
            }
        }

        function getExportDateRangeLabel() {
            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            
            if (exportAllDates) {
                const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                if (allDates.length === 0) return '--';
                return allDates[0] === allDates[allDates.length - 1] ? allDates[0] : `${allDates[0]} ~ ${allDates[allDates.length - 1]}`;
            }
            
            if (startDate && endDate) return `${startDate} ~ ${endDate}`;
            
            const selectedDates = Array.isArray(appData?.visualization?.selectedDates) ? appData.visualization.selectedDates.slice() : [];
            selectedDates.sort();
            if (selectedDates.length === 0) return '--';
            return selectedDates[0] === selectedDates[selectedDates.length - 1]
                ? selectedDates[0]
                : `${selectedDates[0]} ~ ${selectedDates[selectedDates.length - 1]}`;
        }

        function formatExportTimestamp() {
            const now = new Date();
            const pad2 = (n) => String(n).padStart(2, '0');
            return `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}${pad2(now.getSeconds())}`;
        }

        function resolveExportNameTemplateContext(baseName, extension) {
            const prefix = String(elements.exportFileNameInput?.value || '负荷曲线数据').trim() || '负荷曲线数据';
            const dateRange = getExportDateRangeLabel();
            const exportTime = formatExportTimestamp();
            return {
                prefix,
                baseName,
                dateRange,
                exportTime,
                ext: String(extension || '').replace(/^\./, '')
            };
        }

        function applyNameTemplate(template, ctx) {
            const safe = String(template || '').trim();
            if (!safe) return '';
            return safe
                .replaceAll('{prefix}', ctx.prefix)
                .replaceAll('{baseName}', ctx.baseName)
                .replaceAll('{dateRange}', ctx.dateRange)
                .replaceAll('{exportTime}', ctx.exportTime)
                .replaceAll('{ext}', ctx.ext);
        }

        function getExportFileName(baseName, extension) {
            const ext = String(extension || '').replace(/^\./, '');
            const template = String(elements.exportNameTemplateInput?.value || '').trim();
            const ctx = resolveExportNameTemplateContext(baseName, ext);

            if (template) {
                const templated = applyNameTemplate(template, ctx);
                if (templated) return templated;
            }

            let fileName = `${ctx.prefix}_${baseName}`;

            if (elements.includeFiltersInNameCheckbox && elements.includeFiltersInNameCheckbox.checked) {
                const startDate = elements.startDateInput?.value;
                const endDate = elements.endDateInput?.value;
                if (startDate && endDate) {
                    fileName += `_${startDate}_to_${endDate}`;
                }
            }

            if (elements.includeTimestampCheckbox && elements.includeTimestampCheckbox.checked) {
                fileName += `_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}`;
            }

            fileName += `.${ext}`;
            return fileName;
        }

        function openExportPreview(payload) {
            if (!elements.exportPreviewModal) {
                if (payload && typeof payload.execute === 'function') payload.execute();
                return;
            }

            pendingExportRequest = payload;
            if (elements.exportPreviewFileName) elements.exportPreviewFileName.textContent = payload.fileName || '—';
            if (elements.exportPreviewSummary) elements.exportPreviewSummary.textContent = payload.summary || '—';
            if (elements.exportPreviewRange) elements.exportPreviewRange.textContent = payload.range || '—';

            const showChart = !!payload.chartPreviewUrl;
            if (elements.exportPreviewChartWrap) {
                elements.exportPreviewChartWrap.classList.toggle('hidden', !showChart);
            }
            if (showChart && elements.exportPreviewChartImg) {
                elements.exportPreviewChartImg.src = payload.chartPreviewUrl;
            }
            
            // 显示柱状图预览（如果有）
            const showBarChart = !!payload.barChartPreviewUrl;
            if (elements.exportPreviewBarChartWrap) {
                elements.exportPreviewBarChartWrap.classList.toggle('hidden', !showBarChart);
            }
            if (showBarChart && elements.exportPreviewBarChartImg) {
                elements.exportPreviewBarChartImg.src = payload.barChartPreviewUrl;
            }

            elements.exportPreviewModal.classList.remove('hidden');
            elements.exportPreviewModal.setAttribute('aria-hidden', 'false');
        }

        // 更新图表导出预览
        function updateChartExportPreview() {
            if (!elements.exportPreviewModal || elements.exportPreviewModal.classList.contains('hidden')) {
                return; // 如果预览模态框没打开，不更新
            }

            const chartType = elements.exportChartSelect?.value || 'loadCurve';
            const previewUrls = getChartPreviewUrls(chartType);
            
            // 更新预览图片
            if (previewUrls.chartPreviewUrl && elements.exportPreviewChartImg) {
                elements.exportPreviewChartImg.src = previewUrls.chartPreviewUrl;
                if (elements.exportPreviewChartWrap) {
                    elements.exportPreviewChartWrap.classList.remove('hidden');
                }
            }
            
            if (previewUrls.barChartPreviewUrl && elements.exportPreviewBarChartImg) {
                elements.exportPreviewBarChartImg.src = previewUrls.barChartPreviewUrl;
                if (elements.exportPreviewBarChartWrap) {
                    elements.exportPreviewBarChartWrap.classList.remove('hidden');
                }
            }
            
            // 根据选择的图表类型显示/隐藏对应的预览
            if (chartType === 'loadCurve') {
                if (elements.exportPreviewChartWrap) elements.exportPreviewChartWrap.classList.remove('hidden');
                if (elements.exportPreviewBarChartWrap) elements.exportPreviewBarChartWrap.classList.add('hidden');
            } else if (chartType === 'dailyBar') {
                if (elements.exportPreviewChartWrap) elements.exportPreviewChartWrap.classList.add('hidden');
                if (elements.exportPreviewBarChartWrap) elements.exportPreviewBarChartWrap.classList.remove('hidden');
            } else if (chartType === 'both') {
                if (elements.exportPreviewChartWrap) elements.exportPreviewChartWrap.classList.remove('hidden');
                if (elements.exportPreviewBarChartWrap) elements.exportPreviewBarChartWrap.classList.remove('hidden');
            }
        }

        // ==================== 多页报告预览功能 ====================
        let currentReportPage = 1;
        const totalReportPages = 4;

        function openReportPreview() {
            console.log('打开报告预览');
            // 检查是否有数据
            if (!appData.processedData || appData.processedData.length === 0) {
                showNotification('提示', '请先导入并处理数据后再查看报告', 'warning');
                return;
            }

            // 填充报告数据
            console.log('开始填充报告数据');
            fillReportData();

            // 显示模态框
            const panel = document.getElementById('reportPreviewPanel');
            const content = document.getElementById('reportPreviewContent');

            if (!panel || !content) {
                console.error('报告预览元素未找到');
                return;
            }

            console.log('显示报告预览模态框');

            // 清除所有可能冲突的类
            panel.classList.remove('hidden', 'opacity-0');
            content.classList.remove('scale-95', 'opacity-0', 'hidden');
            
            // 显示面板
            panel.style.display = 'flex';
            panel.style.position = 'fixed';
            panel.style.top = '0';
            panel.style.left = '0';
            panel.style.width = '100vw';
            panel.style.height = '100vh';
            panel.style.backgroundColor = 'rgba(15, 23, 42, 0.8)';
            panel.style.zIndex = '9999';
            panel.style.alignItems = 'center';
            panel.style.justifyContent = 'center';
            panel.style.overflow = 'hidden';
            
            // 设置内容样式
            content.style.display = 'flex';
            content.style.flexDirection = 'column';
            content.style.width = '90%';
            content.style.maxWidth = '1200px';
            content.style.height = '90vh';
            content.style.maxHeight = 'none';
            content.style.backgroundColor = 'white';
            content.style.borderRadius = '1rem';
            content.style.boxShadow = '0 25px 50px -12px rgba(0, 0, 0, 0.25)';
            content.style.transform = 'none';
            content.style.opacity = '1';
            content.style.margin = '0';
            content.style.position = 'relative';
            content.style.overflow = 'hidden'; // 确保内容区域有滚动条

            // 强制显示第一页
            console.log('切换到第一页');
            goToReportPage(1);
            
            // 检查第一页是否显示
            setTimeout(() => {
                const page1 = document.getElementById('reportPage1');
                console.log('第一页元素:', page1);
                console.log('第一页隐藏状态:', page1?.classList.contains('hidden'));
                console.log('第一页display:', page1?.style.display);
            }, 100);
            
            console.log('报告预览已打开');
        }

        function closeReportPreview() {
            const panel = document.getElementById('reportPreviewPanel');
            const content = document.getElementById('reportPreviewContent');

            if (panel) {
                panel.style.display = 'none';
                panel.classList.add('hidden');
            }
            
            if (content) {
                content.style.display = 'none';
                content.style.transform = '';
                content.style.opacity = '';
            }
        }

        function navigateReportPage(direction) {
            const newPage = currentReportPage + direction;
            if (newPage >= 1 && newPage <= totalReportPages) {
                goToReportPage(newPage);
            }
        }

        function goToReportPage(page) {
            currentReportPage = page;
            console.log(`切换到第${page}页`);

            // 隐藏所有页面
            for (let i = 1; i <= totalReportPages; i++) {
                const pageEl = document.getElementById(`reportPage${i}`);
                if (pageEl) {
                    pageEl.classList.add('hidden');
                    pageEl.style.display = 'none'; // 同时设置display
                    console.log(`隐藏页面${i}`);
                }
            }

            // 显示当前页面
            const currentPageEl = document.getElementById(`reportPage${page}`);
            if (currentPageEl) {
                currentPageEl.classList.remove('hidden');
                currentPageEl.style.display = 'block'; // 同时设置display
                console.log(`显示页面${page}`);
            } else {
                console.error(`找不到页面${page}的元素`);
            }

            // 更新页面指示器
            const indicator = document.getElementById('reportPageIndicator');
            if (indicator) {
                indicator.textContent = `第 ${page} / ${totalReportPages} 页`;
            }

            // 更新导航点
            document.querySelectorAll('.report-dot').forEach((dot, index) => {
                if (index + 1 === page) {
                    dot.classList.remove('bg-slate-300');
                    dot.classList.add('bg-indigo-600');
                } else {
                    dot.classList.remove('bg-indigo-600');
                    dot.classList.add('bg-slate-300');
                }
            });

            // 更新按钮状态
            const prevBtn = document.getElementById('reportPrevBtn');
            const nextBtn = document.getElementById('reportNextBtn');

            if (prevBtn) {
                prevBtn.disabled = page === 1;
            }
            if (nextBtn) {
                nextBtn.disabled = page === totalReportPages;
            }
        }

        function fillReportData() {
            const data = appData.processedData;
            console.log('fillReportData 被调用，数据:', data);
            if (!data || data.length === 0) {
                console.warn('没有数据可填充报告');
                return;
            }

            console.log('填充报告数据，数据条数:', data.length);

            // 计算关键指标 - 使用正确的数据字段
            // processedData 结构: { date, hourlyData[], dailyTotal, ... }
            const totalEnergy = data.reduce((sum, d) => sum + (d.dailyTotal || 0), 0);
            
            // 计算所有小时数据的平均值和最大/最小值
            let allHourlyValues = [];
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    // 过滤掉null、NaN和0值
                    allHourlyValues = allHourlyValues.concat(
                        d.hourlyData.filter(v => v !== null && !isNaN(v) && v > 0)
                    );
                }
            });
            
            const avgLoad = allHourlyValues.length > 0 
                ? allHourlyValues.reduce((sum, v) => sum + v, 0) / allHourlyValues.length 
                : 0;
            const maxLoad = allHourlyValues.length > 0 ? Math.max(...allHourlyValues) : 0;
            const minLoad = allHourlyValues.length > 0 ? Math.min(...allHourlyValues) : 0;
            const loadFactor = maxLoad > 0 ? (avgLoad / maxLoad * 100).toFixed(1) : '0.0';

            console.log('计算结果:', { totalEnergy, avgLoad, maxLoad, minLoad, loadFactor });

            // 第1页：概览与关键指标
            const reportTotalEnergy = document.getElementById('reportTotalEnergy');
            const reportAvgLoad = document.getElementById('reportAvgLoad');
            const reportMaxLoad = document.getElementById('reportMaxLoad');
            const reportMinLoad = document.getElementById('reportMinLoad');
            
            console.log('填充第1页数据:', { totalEnergy, avgLoad, maxLoad, minLoad });
            console.log('元素存在性:', { 
                reportTotalEnergy: !!reportTotalEnergy,
                reportAvgLoad: !!reportAvgLoad,
                reportMaxLoad: !!reportMaxLoad,
                reportMinLoad: !!reportMinLoad
            });
            
            if (reportTotalEnergy) reportTotalEnergy.textContent = totalEnergy.toFixed(2);
            if (reportAvgLoad) reportAvgLoad.textContent = avgLoad.toFixed(2);
            if (reportMaxLoad) reportMaxLoad.textContent = maxLoad.toFixed(2);
            if (reportMinLoad) reportMinLoad.textContent = minLoad.toFixed(2);

            // 日期范围
            const dates = [...new Set(data.map(d => d.date))].sort();
            const dateRange = dates.length > 0
                ? (dates[0] === dates[dates.length - 1] ? dates[0] : `${dates[0]} ~ ${dates[dates.length - 1]}`)
                : '--';
            const reportDateRange = document.getElementById('reportDateRange');
            if (reportDateRange) reportDateRange.textContent = dateRange;

            // 数据摘要
            const reportTotalDays = document.getElementById('reportTotalDays');
            const reportTotalPoints = document.getElementById('reportTotalPoints');
            const reportLoadFactor = document.getElementById('reportLoadFactor');
            const reportPeakValleyDiff = document.getElementById('reportPeakValleyDiff');
            
            if (reportTotalDays) reportTotalDays.textContent = `${dates.length} 天`;
            if (reportTotalPoints) reportTotalPoints.textContent = `${data.length} 个`;
            if (reportLoadFactor) reportLoadFactor.textContent = `${loadFactor} %`;
            if (reportPeakValleyDiff) reportPeakValleyDiff.textContent = `${(maxLoad - minLoad).toFixed(2)} kW`;

            // 生成时间
            const reportGenerateTime = document.getElementById('reportGenerateTime');
            if (reportGenerateTime) reportGenerateTime.textContent = new Date().toLocaleString('zh-CN');

            // 填充时段分析表
            fillTimeTable(data);

            // 填充月度汇总表
            fillMonthlyTable(data);

            // 填充极值统计
            fillExtremeValues(data);

            // 填充第2页详细分析数据
            fillDetailedAnalysis(data);

            // 设置图表图片
            const chartCanvas = document.getElementById('loadCurveChart');
            const reportChartImage = document.getElementById('reportChartImage');
            if (chartCanvas && reportChartImage) {
                try {
                    reportChartImage.src = chartCanvas.toDataURL('image/png');
                } catch (e) {
                    console.warn('无法获取图表图片:', e);
                }
            }
        }

        // 填充第2页详细分析数据
        function fillDetailedAnalysis(data) {
            console.log('填充详细分析数据');

            // 计算日间(06:00-18:00)和夜间(18:00-06:00)数据
            let dayPowerSum = 0, dayCount = 0;
            let nightPowerSum = 0, nightCount = 0;
            let dayEnergy = 0, nightEnergy = 0;

            // 季度数据
            const quarterlyData = { Q1: [], Q2: [], Q3: [], Q4: [] };

            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    // 获取月份用于季度计算
                    const month = parseInt(d.date?.substring(5, 7)) || 1;
                    const quarter = month <= 3 ? 'Q1' : month <= 6 ? 'Q2' : month <= 9 ? 'Q3' : 'Q4';

                    d.hourlyData.forEach((val, hour) => {
                        if (val !== null && !isNaN(val)) {
                            if (hour >= 6 && hour < 18) {
                                // 日间
                                dayPowerSum += val;
                                dayCount++;
                                dayEnergy += val;
                            } else {
                                // 夜间
                                nightPowerSum += val;
                                nightCount++;
                                nightEnergy += val;
                            }
                            // 季度数据
                            quarterlyData[quarter].push(val);
                        }
                    });
                }
            });

            const dayAvg = dayCount > 0 ? dayPowerSum / dayCount : 0;
            const nightAvg = nightCount > 0 ? nightPowerSum / nightCount : 0;
            const totalEnergy = dayEnergy + nightEnergy;
            const dayRatio = totalEnergy > 0 ? (dayEnergy / totalEnergy * 100).toFixed(1) : 0;
            const nightRatio = totalEnergy > 0 ? (nightEnergy / totalEnergy * 100).toFixed(1) : 0;

            console.log('日间/夜间分析:', { dayAvg, nightAvg, dayRatio, nightRatio });

            // 收集所有有效值用于计算负荷率
            let allPowerValues = [];
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach(val => {
                        if (val !== null && !isNaN(val) && val > 0) {
                            allPowerValues.push(val);
                        }
                    });
                }
            });

            // 计算负荷率
            const avgLoad = allPowerValues.length > 0 ? allPowerValues.reduce((sum, v) => sum + v, 0) / allPowerValues.length : 0;
            const maxLoad = allPowerValues.length > 0 ? Math.max(...allPowerValues) : 0;
            const loadFactor = maxLoad > 0 ? (avgLoad / maxLoad * 100).toFixed(1) : '0.0';

            // 填充第2页数据
            const reportDayAvg = document.getElementById('reportDayAvg');
            const reportNightAvg = document.getElementById('reportNightAvg');
            const reportDayRatio = document.getElementById('reportDayRatio');
            const reportNightRatio = document.getElementById('reportNightRatio');

            if (reportDayAvg) reportDayAvg.textContent = `${dayAvg.toFixed(2)} kW`;
            if (reportNightAvg) reportNightAvg.textContent = `${nightAvg.toFixed(2)} kW`;
            if (reportDayRatio) reportDayRatio.textContent = `${dayRatio} %`;
            if (reportNightRatio) reportNightRatio.textContent = `${nightRatio} %`;

            // 填充季度数据
            ['Q1', 'Q2', 'Q3', 'Q4'].forEach((q, index) => {
                const values = quarterlyData[q];
                const avg = values.length > 0 ? values.reduce((sum, v) => sum + v, 0) / values.length : 0;
                const element = document.getElementById(`report${q}Avg`);
                if (element) {
                    element.textContent = `${avg.toFixed(2)} kW`;
                }
            });

            // 计算峰值和谷值出现时段
            let peakHour = 0, peakValue = 0;
            let valleyHour = 0, valleyValue = Infinity;

            // 按小时统计平均负荷
            const hourlyAvg = new Array(24).fill(0).map(() => ({ sum: 0, count: 0 }));
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach((val, hour) => {
                        if (val !== null && !isNaN(val)) {
                            hourlyAvg[hour].sum += val;
                            hourlyAvg[hour].count++;
                        }
                    });
                }
            });

            hourlyAvg.forEach((h, hour) => {
                const avg = h.count > 0 ? h.sum / h.count : 0;
                if (avg > peakValue) {
                    peakValue = avg;
                    peakHour = hour;
                }
                if (avg < valleyValue && avg > 0) {
                    valleyValue = avg;
                    valleyHour = hour;
                }
            });

            const reportDayPeakTime = document.getElementById('reportDayPeakTime');
            const reportNightValleyTime = document.getElementById('reportNightValleyTime');
            const reportPeakTime = document.getElementById('reportPeakTime');
            const reportValleyTime = document.getElementById('reportValleyTime');

            const peakTimeStr = `${String(peakHour).padStart(2, '0')}:00-${String(peakHour + 1).padStart(2, '0')}:00`;
            const valleyTimeStr = `${String(valleyHour).padStart(2, '0')}:00-${String(valleyHour + 1).padStart(2, '0')}:00`;

            if (reportDayPeakTime) reportDayPeakTime.textContent = peakTimeStr;
            if (reportNightValleyTime) reportNightValleyTime.textContent = valleyTimeStr;
            if (reportPeakTime) reportPeakTime.textContent = peakTimeStr;
            if (reportValleyTime) reportValleyTime.textContent = valleyTimeStr;

            console.log('详细分析数据填充完成');

            // 根据分析结果生成优化建议
            generateOptimizationSuggestions(data, {
                dayRatio: parseFloat(dayRatio),
                nightRatio: parseFloat(nightRatio),
                loadFactor: parseFloat(loadFactor),
                peakHour,
                valleyHour,
                avgLoad,
                maxLoad
            });
        }

        // 根据实际数据动态生成优化建议
        function generateOptimizationSuggestions(data, stats) {
            console.log('生成优化建议，统计数据:', stats);

            const suggestions = {
                peakShifting: '',
                loadBalancing: '',
                renewableEnergy: ''
            };

            // 1. 错峰用电建议
            if (stats.peakHour >= 9 && stats.peakHour <= 17) {
                // 峰值出现在日间工作时段
                suggestions.peakShifting = `根据分析，当前负荷峰值出现在 ${String(stats.peakHour).padStart(2, '0')}:00 左右。建议将部分非关键生产任务调整至 ${String((stats.peakHour + 12) % 24).padStart(2, '0')}:00-${String((stats.peakHour + 13) % 24).padStart(2, '0')}:00 时段运行，可有效降低峰值负荷，预计可节省10-15%的需量电费。`;
            } else if (stats.peakHour >= 18 && stats.peakHour <= 21) {
                // 峰值出现在晚间
                suggestions.peakShifting = `当前负荷峰值出现在晚间 ${String(stats.peakHour).padStart(2, '0')}:00，建议在日间 10:00-16:00 安排大功率设备运行，充分利用电网低谷时段，可降低约15-20%的用电成本。`;
            } else {
                suggestions.peakShifting = `建议在低谷时段（22:00-06:00）安排大功率设备运行，可有效利用分时电价政策，预计可降低约15-20%的用电成本。`;
            }

            // 2. 负荷平衡建议
            const peakValleyDiff = stats.maxLoad > 0 ? ((stats.maxLoad - (stats.avgLoad * 0.5)) / stats.maxLoad * 100).toFixed(1) : 0;
            
            if (parseFloat(peakValleyDiff) > 50) {
                suggestions.loadBalancing = `当前峰谷差较大（${peakValleyDiff}%），负荷率较低。建议引入储能系统，在低谷时段充电、高峰时段放电，可平滑负荷曲线约30-40%，同时提高变压器利用率。`;
            } else if (parseFloat(peakValleyDiff) > 30) {
                suggestions.loadBalancing = `当前负荷有一定波动，建议通过调整生产排班，将部分工序从峰值时段转移至谷值时段，优化设备启停顺序，可提高设备利用率约15-20%。`;
            } else {
                suggestions.loadBalancing = `当前负荷曲线较为平稳，设备利用率良好。建议继续保持，定期监测负荷变化，及时调整生产计划。`;
            }

            // 3. 新能源配置建议
            if (stats.dayRatio > 60) {
                // 日间用电量大，适合光伏
                suggestions.renewableEnergy = `当前日间用电占比高达${stats.dayRatio}%，非常适合配置光伏发电系统。建议安装装机容量约 ${(stats.avgLoad * 8 / 1000).toFixed(1)} kWp 的光伏电站，日均可自发自用约 ${(stats.dayRatio / 100 * stats.avgLoad * 6).toFixed(0)} kWh，年均可节省电费约15-25%。`;
            } else if (stats.dayRatio > 40) {
                suggestions.renewableEnergy = `当前日间用电占比${stats.dayRatio}%，建议配置中小型光伏发电系统（约 ${(stats.avgLoad * 5 / 1000).toFixed(1)} kWp），配合储能设备，可在日间高峰期提供部分电力，减少电网购电。`;
            } else {
                suggestions.renewableEnergy = `当前夜间用电占比较高，日间用电相对较少。如需配置新能源，建议优先考虑小型光伏或风力发电，主要用于办公区域照明和辅助设备，可作为备用电源使用。`;
            }

            console.log('生成的优化建议:', suggestions);

            // 填充到页面
            const suggestion1 = document.getElementById('reportSuggestion1');
            const suggestion2 = document.getElementById('reportSuggestion2');
            const suggestion3 = document.getElementById('reportSuggestion3');

            if (suggestion1) suggestion1.textContent = suggestions.peakShifting;
            if (suggestion2) suggestion2.textContent = suggestions.loadBalancing;
            if (suggestion3) suggestion3.textContent = suggestions.renewableEnergy;

            // 同时更新AI分析的优化建议
            updateAISuggestions(suggestions);
        }

        // 更新AI分析的优化建议
        function updateAISuggestions(suggestions) {
            const suggestionsContainer = document.getElementById('aiSuggestions');
            if (!suggestionsContainer) return;

            const allSuggestions = [
                suggestions.peakShifting,
                suggestions.loadBalancing,
                suggestions.renewableEnergy
            ];

            // 更新AI分析的建议
            const existingSuggestions = window.aiAnalysisResult?.suggestions || [];
            if (existingSuggestions.length > 0) {
                // 合并AI建议和优化建议
                window.aiAnalysisResult.suggestions = [...existingSuggestions, ...allSuggestions];
            } else {
                // 使用优化建议
                window.aiAnalysisResult = window.aiAnalysisResult || {};
                window.aiAnalysisResult.suggestions = allSuggestions;
            }
        }

        function fillTimeTable(data) {
            const tbody = document.getElementById('reportTimeTable');
            if (!tbody) return;

            console.log('填充时段分析表，数据条数:', data.length);

            // 定义时段
            const timeSlots = [
                { name: '凌晨', range: '00:00-06:00', hours: [0, 1, 2, 3, 4, 5] },
                { name: '上午', range: '06:00-12:00', hours: [6, 7, 8, 9, 10, 11] },
                { name: '下午', range: '12:00-18:00', hours: [12, 13, 14, 15, 16, 17] },
                { name: '晚上', range: '18:00-24:00', hours: [18, 19, 20, 21, 22, 23] }
            ];

            // 计算总用电量（使用dailyTotal）
            const totalEnergy = data.reduce((sum, d) => sum + (d.dailyTotal || 0), 0);
            console.log('总用电量:', totalEnergy);

            let html = '';

            timeSlots.forEach(slot => {
                // 从hourlyData中提取对应时段的数据
                let slotPowerSum = 0;
                let slotDataCount = 0;
                let slotEnergy = 0; // 时段总用电量

                data.forEach(d => {
                    if (d.hourlyData && Array.isArray(d.hourlyData)) {
                        slot.hours.forEach(hour => {
                            const value = d.hourlyData[hour];
                            if (value !== null && !isNaN(value) && value > 0) {
                                slotPowerSum += value;
                                slotDataCount++;
                                // 用电量 = 功率(kW) × 时间(h) = kWh
                                // 假设每个数据点代表1小时
                                slotEnergy += value;
                            }
                        });
                    }
                });

                // 计算平均负荷
                const slotAvgLoad = slotDataCount > 0 ? slotPowerSum / slotDataCount : 0;
                const ratio = totalEnergy > 0 ? (slotEnergy / totalEnergy * 100).toFixed(1) : 0;

                console.log(`时段 ${slot.name}: 用电量=${slotEnergy.toFixed(2)}, 占比=${ratio}%, 平均负荷=${slotAvgLoad.toFixed(2)}`);

                let level = '低';
                let levelClass = 'text-indigo-600 bg-indigo-50 ring-1 ring-indigo-100';
                if (ratio > 30) {
                    level = '高';
                    levelClass = 'text-indigo-900 bg-indigo-100 ring-1 ring-indigo-200';
                } else if (ratio > 20) {
                    level = '中';
                    levelClass = 'text-indigo-700 bg-indigo-50/50 ring-1 ring-indigo-100/50';
                }

                html += `
                    <tr>
                        <td class="px-4 py-3 font-medium text-slate-800">${slot.name}</td>
                        <td class="px-4 py-3 text-center text-slate-600">${slot.range}</td>
                        <td class="px-4 py-3 text-right font-medium text-slate-800">${slotAvgLoad.toFixed(2)} kW</td>
                        <td class="px-4 py-3 text-right text-slate-600">${ratio}%</td>
                        <td class="px-4 py-3 text-center">
                            <span class="px-2 py-1 rounded-full text-xs font-bold ${levelClass}">${level}</span>
                        </td>
                    </tr>
                `;
            });

            tbody.innerHTML = html;
            console.log('时段分析表填充完成');
        }

        function fillMonthlyTable(data) {
            const tbody = document.getElementById('reportMonthlyTable');
            if (!tbody) return;

            console.log('填充月度汇总表，数据条数:', data.length);

            // 按月份分组
            const monthlyData = {};
            data.forEach(d => {
                const month = d.date?.substring(0, 7) || '未知';
                if (!monthlyData[month]) {
                    monthlyData[month] = { energy: 0, allPowerValues: [], count: 0 };
                }
                // 使用dailyTotal作为用电量
                monthlyData[month].energy += d.dailyTotal || 0;
                
                // 收集所有小时功率值
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach(val => {
                        if (val !== null && !isNaN(val)) {
                            monthlyData[month].allPowerValues.push(val);
                        }
                    });
                }
                monthlyData[month].count++;
            });

            console.log('月度数据:', monthlyData);

            let html = '';
            Object.keys(monthlyData).sort().forEach(month => {
                const m = monthlyData[month];
                // 过滤掉0值
                const validPowerValues = m.allPowerValues.filter(v => v > 0);
                const avgLoad = validPowerValues.length > 0 
                    ? validPowerValues.reduce((sum, v) => sum + v, 0) / validPowerValues.length 
                    : 0;
                const maxLoad = validPowerValues.length > 0 ? Math.max(...validPowerValues) : 0;
                const minLoad = validPowerValues.length > 0 ? Math.min(...validPowerValues) : 0;
                const loadFactor = maxLoad > 0 ? (avgLoad / maxLoad * 100).toFixed(1) : '0.0';

                html += `
                    <tr>
                        <td class="px-4 py-3 font-medium text-slate-800">${month}</td>
                        <td class="px-4 py-3 text-right text-slate-700">${m.energy.toFixed(2)}</td>
                        <td class="px-4 py-3 text-right text-slate-700">${avgLoad.toFixed(2)}</td>
                        <td class="px-4 py-3 text-right text-slate-700">${maxLoad.toFixed(2)}</td>
                        <td class="px-4 py-3 text-right text-slate-700">${minLoad.toFixed(2)}</td>
                        <td class="px-4 py-3 text-right text-slate-700">${loadFactor}%</td>
                    </tr>
                `;
            });

            tbody.innerHTML = html;
            console.log('月度汇总表填充完成');
        }

        function fillExtremeValues(data) {
            console.log('填充极值统计，数据条数:', data.length);

            // 按日期汇总 - 使用dailyTotal
            const dailyData = {};
            data.forEach(d => {
                if (!dailyData[d.date]) {
                    dailyData[d.date] = 0;
                }
                dailyData[d.date] += d.dailyTotal || 0;
            });

            console.log('每日数据汇总:', dailyData);

            const sortedDays = Object.entries(dailyData).sort((a, b) => b[1] - a[1]);
            console.log('排序后的数据:', sortedDays);

            // 高峰日 TOP 5
            const peakTbody = document.getElementById('reportTopPeakDays');
            if (peakTbody) {
                const top5 = sortedDays.slice(0, 5);
                if (top5.length > 0) {
                    peakTbody.innerHTML = top5.map(([date, energy]) => `
                        <tr class="hover:bg-indigo-50/30 transition-colors group">
                            <td class="py-2 text-slate-700 font-bold group-hover:text-indigo-600 transition-colors">${date}</td>
                            <td class="py-2 text-right font-black text-indigo-900">${energy.toFixed(2)}</td>
                        </tr>
                    `).join('');
                } else {
                    peakTbody.innerHTML = '<tr><td colspan="2" class="py-2 text-center text-slate-400">暂无数据</td></tr>';
                }
            }

            // 低谷日 TOP 5
            const valleyTbody = document.getElementById('reportTopValleyDays');
            if (valleyTbody) {
                const bottom5 = sortedDays.slice(-5).reverse();
                if (bottom5.length > 0 && bottom5[0][1] > 0) {
                    valleyTbody.innerHTML = bottom5.map(([date, energy]) => `
                        <tr class="hover:bg-indigo-50/30 transition-colors group">
                            <td class="py-2 text-slate-700 font-bold group-hover:text-indigo-600 transition-colors">${date}</td>
                            <td class="py-2 text-right font-black text-indigo-500">${energy.toFixed(2)}</td>
                        </tr>
                    `).join('');
                } else {
                    valleyTbody.innerHTML = '<tr><td colspan="2" class="py-2 text-center text-slate-400">暂无数据</td></tr>';
                }
            }

            console.log('极值统计填充完成');
        }

        function getCanvasPreviewUrl(canvasEl) {
            try {
                if (!canvasEl || typeof canvasEl.toDataURL !== 'function') return '';
                return canvasEl.toDataURL('image/png', 0.9);
            } catch (e) {
                return '';
            }
        }

        function getExportRuntimeOptions() {
            return {
                weekdaysOnly: !!elements.exportWeekdaysOnlyCheckbox?.checked,
                includeAnomalyMarks: elements.exportAnomalyMarksCheckbox ? !!elements.exportAnomalyMarksCheckbox.checked : false,
                samplingInterval: String(elements.exportSamplingIntervalSelect?.value || 'auto'),
                samplingAgg: String(elements.exportSamplingAggSelect?.value || 'sum'),
                chartType: String(elements.exportChartSelect?.value || 'loadCurve') // 添加图表选择
            };
        }

        // 获取当前要导出的图表
        function getSelectedChartCanvas() {
            const chartType = elements.exportChartSelect?.value || 'loadCurve';
            switch (chartType) {
                case 'loadCurve':
                    return elements.loadCurveChart;
                case 'dailyBar':
                    return appData.dailyTotalBarChart?.canvas;
                case 'both':
                    return elements.loadCurveChart; // 默认返回主图表
                default:
                    return elements.loadCurveChart;
            }
        }

        // 获取图表预览URL（支持多图表）
        function getChartPreviewUrls(chartType) {
            const urls = {};
            if (chartType === 'loadCurve' || chartType === 'both') {
                urls.chartPreviewUrl = getCanvasPreviewUrl(elements.loadCurveChart);
            }
            if (chartType === 'dailyBar' || chartType === 'both') {
                urls.barChartPreviewUrl = getCanvasPreviewUrl(appData.dailyTotalBarChart?.canvas);
            }
            return urls;
        }

        // 根据选择的图表类型生成导出描述
        function getChartExportSummary(format) {
            const chartType = elements.exportChartSelect?.value || 'loadCurve';
            switch (chartType) {
                case 'loadCurve':
                    return `导出24小时负荷曲线图（${format}）`;
                case 'dailyBar':
                    return `导出日期总用电量对比图（${format}）`;
                case 'both':
                    return `导出24小时负荷曲线图和日期总用电量对比图（${format}）`;
                default:
                    return `导出当前图表（${format}）`;
            }
        }

        function requestExportAction(actionKey) {
            const opts = getExportRuntimeOptions();
            const range = getExportDateRangeLabel();

            const buildAndOpen = (cfg) => {
                const fileName = cfg.fileName || getExportFileName(cfg.baseName, cfg.ext);
                const chartType = opts.chartType || 'loadCurve';
                const previewUrls = getChartPreviewUrls(chartType);
                
                const payload = {
                    fileName,
                    summary: cfg.summary,
                    range,
                    chartPreviewUrl: cfg.chartPreview ? previewUrls.chartPreviewUrl || '' : '',
                    barChartPreviewUrl: cfg.chartPreview ? previewUrls.barChartPreviewUrl || '' : '',
                    execute: () => cfg.execute({ fileName, opts })
                };
                openExportPreview(payload);
            };

            switch (actionKey) {
                case 'chart_png':
                    return buildAndOpen({
                        baseName: '图表',
                        ext: 'png',
                        chartPreview: true,
                        summary: getChartExportSummary('PNG'),
                        execute: ({ fileName }) => exportAsPNG({ fileName })
                    });
                case 'chart_jpg':
                    return buildAndOpen({
                        baseName: '图表',
                        ext: 'jpg',
                        chartPreview: true,
                        summary: getChartExportSummary('JPG'),
                        execute: ({ fileName }) => exportAsJPG({ fileName })
                    });
                case 'chart_svg':
                    return buildAndOpen({
                        baseName: '图表',
                        ext: 'svg',
                        chartPreview: false,
                        summary: getChartExportSummary('SVG'),
                        execute: ({ fileName }) => exportAsSVG({ fileName })
                    });
                case 'report_pdf':
                    return buildAndOpen({
                        baseName: '专业报告',
                        ext: 'pdf',
                        chartPreview: true,
                        barChartPreview: true,
                        summary: '生成PDF专业报告（会打开打印窗口，可另存为PDF）',
                        execute: ({ fileName }) => exportPDFReport({ fileName })
                    });
                case 'data_15min': {
                    const fmt = String(elements.exportFormatSelect?.value || 'xlsx');
                    const intervalForName = opts.samplingInterval === 'auto'
                        ? (appData?.config?.timeInterval || 15)
                        : (parseInt(String(opts.samplingInterval), 10) || (appData?.config?.timeInterval || 15));
                    const baseName = intervalForName === 15 ? '15分钟明细' : `${intervalForName}分钟明细`;
                    return buildAndOpen({
                        baseName,
                        ext: fmt,
                        chartPreview: false,
                        summary: `导出明细数据（采样：${opts.samplingInterval === 'auto' ? '自动' : opts.samplingInterval + '分钟'}，聚合：${opts.samplingAgg === 'sum' ? '求和' : '平均'}）`,
                        execute: ({ fileName, opts }) => export15MinData({ fileName, opts })
                    });
                }
                case 'data_24h': {
                    const fmt = String(elements.exportFormatSelect?.value || 'xlsx');
                    return buildAndOpen({
                        baseName: '24小时电能明细',
                        ext: fmt,
                        chartPreview: false,
                        summary: '导出24小时汇总数据',
                        execute: ({ fileName, opts }) => exportHourlyData({ fileName, opts })
                    });
                }
                case 'data_daily_stats': {
                    const fmt = String(elements.exportFormatSelect?.value || 'xlsx');
                    return buildAndOpen({
                        baseName: '日统计数据',
                        ext: fmt,
                        chartPreview: false,
                        summary: '导出日统计与峰谷指标',
                        execute: ({ fileName, opts }) => exportDailyStats({ fileName, opts })
                    });
                }
                case 'data_summary_curve': {
                    const fmt = String(elements.exportFormatSelect?.value || 'xlsx');
                    return buildAndOpen({
                        baseName: '汇总曲线数据',
                        ext: fmt,
                        chartPreview: false,
                        summary: '导出当前汇总曲线（各曲线列）',
                        execute: ({ fileName, opts }) => exportSummaryCurveData({ fileName, opts })
                    });
                }
                case 'data_pvsyst':
                    return buildAndOpen({
                        baseName: 'PVsyst月度重排8760小时数据',
                        ext: 'csv',
                        chartPreview: false,
                        summary: '导出PVsyst标准格式（8760h）',
                        execute: ({ fileName, opts }) => exportPVsystData({ fileName, opts })
                    });
                case 'package_full':
                    return buildAndOpen({
                        baseName: '完整数据包',
                        ext: 'xlsx',
                        chartPreview: false,
                        summary: '导出概况/明细/24h/日统计（多Sheet XLSX）',
                        execute: ({ fileName, opts }) => exportFullPackage({ fileName, opts })
                    });
                case 'bundle_zip':
                    return buildAndOpen({
                        baseName: '打包导出',
                        ext: 'zip',
                        chartPreview: true,
                        summary: '打包导出（ZIP：图表+完整数据包）',
                        execute: ({ fileName, opts }) => exportZipBundle({ fileName, opts })
                    });
                default:
                    showNotification('提示', '暂不支持该导出类型', 'info');
                    return;
            }
        }

        function exportAsPNG(options = {}) {
            const chartCanvas = getSelectedChartCanvas();
            if (!chartCanvas) {
                showNotification('错误', '没有可导出的图表', 'error');
                return;
            }
            
            const quality = parseFloat(elements.pngQualitySelect.value);
            setExportProgress(10, '导出中');
            chartCanvas.toBlob(blob => {
                const fileName = options.fileName || getExportFileName('图表', 'png');
                saveAs(blob, fileName);
                updateExportStatus('图表 PNG 已导出', fileName);
                showNotification('成功', '图表已导出为PNG', 'success');
            }, 'image/png', quality);
        }

        function exportAsJPG(options = {}) {
            const chartCanvas = getSelectedChartCanvas();
            if (!chartCanvas) {
                showNotification('错误', '没有可导出的图表', 'error');
                return;
            }
            
            setExportProgress(10, '导出中');
            chartCanvas.toBlob(blob => {
                const fileName = options.fileName || getExportFileName('图表', 'jpg');
                saveAs(blob, fileName);
                updateExportStatus('图表 JPG 已导出', fileName);
                showNotification('成功', '图表已导出为JPG', 'success');
            }, 'image/jpeg', 0.9);
        }

        function exportAsSVG(options = {}) {
            const chartCanvas = getSelectedChartCanvas();
            if (!chartCanvas) {
                showNotification('错误', '没有可导出的图表', 'error');
                return;
            }
            
            // 获取对应的Chart实例
            let chartInstance = appData.chart;
            if (chartCanvas === appData.dailyTotalBarChart?.canvas) {
                chartInstance = appData.dailyTotalBarChart;
            }
            
            if (!chartInstance) {
                showNotification('错误', '图表实例不存在', 'error');
                return;
            }
            
            // 生成SVG
            const svg = chartToSVG(chartInstance);
            const blob = new Blob([svg], { type: 'image/svg+xml' });
            setExportProgress(40, '导出中');
            const fileName = options.fileName || getExportFileName('图表', 'svg');
            saveAs(blob, fileName);
            updateExportStatus('图表 SVG 已导出', fileName);
            showNotification('成功', '图表已导出为SVG', 'success');
        }

        function updateSummaryCurveExportAvailability() {
            if (!elements.exportSummaryCurveDataBtn) return;
            const enabled = !!appData.chart && !!appData.visualization?.summaryMode;
            elements.exportSummaryCurveDataBtn.disabled = !enabled;
            elements.exportSummaryCurveDataBtn.classList.toggle('bg-indigo-50', enabled);
            elements.exportSummaryCurveDataBtn.classList.toggle('border-indigo-200', enabled);
            elements.exportSummaryCurveDataBtn.classList.toggle('text-indigo-700', enabled);
        }

        function exportSummaryCurveData(params = {}) {
            if (!appData.chart) {
                showNotification('错误', '没有可导出的图表', 'error');
                return;
            }
            if (!appData.visualization?.summaryMode) {
                showNotification('提示', '请先切换到“汇总曲线”模式再导出', 'info');
                return;
            }

            const labels = Array.isArray(appData.chart.data?.labels) ? appData.chart.data.labels : [];
            const datasets = Array.isArray(appData.chart.data?.datasets) ? appData.chart.data.datasets : [];
            if (labels.length === 0 || datasets.length === 0) {
                showNotification('错误', '当前汇总曲线数据为空，无法导出', 'error');
                return;
            }

            const header = ['系列', ...labels];
            const rows = [header];
            datasets.forEach(ds => {
                const name = String(ds?.label ?? '');
                const data = Array.isArray(ds?.data) ? ds.data : [];
                const row = [name, ...labels.map((_, i) => {
                    const v = data[i];
                    return (typeof v === 'number' && isFinite(v)) ? Number(v.toFixed(3)) : '';
                })];
                rows.push(row);
            });

            setExportProgress(30, '导出中');
            exportDataToFile(rows, '汇总曲线数据', { fileName: params.fileName });
        }

        function toProcessedDateString(dateStr) {
            if (typeof dateStr !== 'string') return String(dateStr || '');
            const s = dateStr.trim();
            if (s.length === 8 && !s.includes('-') && !s.includes('/')) {
                return `${s.substring(0, 4)}-${s.substring(4, 6)}-${s.substring(6, 8)}`;
            }
            if (s.length === 10 && s.includes('/')) {
                return s.replace(/\//g, '-');
            }
            return s;
        }

        function normalizeMeteringPoint(v) {
            const t = String(v ?? '').trim();
            return t ? t : '默认计量点';
        }

        function isWeekdayDate(dateStr) {
            const d = new Date(`${dateStr}T00:00:00`);
            if (isNaN(d.getTime())) return true;
            const day = d.getDay();
            return day !== 0 && day !== 6;
        }

        function buildTimeLabels(intervalMinutes) {
            const labels = [];
            for (let hour = 0; hour < 24; hour++) {
                for (let minute = 0; minute < 60; minute += intervalMinutes) {
                    // 使用简洁格式：0:00, 0:15, 0:30 等（去掉前导零）
                    const hourStr = hour.toString();
                    const minuteStr = minute.toString().padStart(2, '0');
                    labels.push(`${hourStr}:${minuteStr}`);
                }
            }
            return labels;
        }

        function resampleSeries(values, baseIntervalMinutes, targetIntervalMinutes, agg) {
            const base = Math.max(1, Number(baseIntervalMinutes) || 15);
            const target = Math.max(1, Number(targetIntervalMinutes) || base);
            if (target === base) {
                return {
                    intervalMinutes: base,
                    labels: buildTimeLabels(base),
                    values: Array.isArray(values) ? values.slice() : []
                };
            }
            if (target < base || target % base !== 0) {
                return {
                    intervalMinutes: base,
                    labels: buildTimeLabels(base),
                    values: Array.isArray(values) ? values.slice() : []
                };
            }

            const groupSize = target / base;
            const groups = Math.floor((24 * 60) / target);
            const out = [];
            for (let i = 0; i < groups; i++) {
                const startIdx = i * groupSize;
                const seg = values.slice(startIdx, startIdx + groupSize);
                const nums = seg.filter(v => v !== null && v !== undefined && typeof v === 'number' && isFinite(v));
                if (nums.length === 0) {
                    out.push(null);
                    continue;
                }
                const sum = nums.reduce((a, b) => a + b, 0);
                out.push(agg === 'avg' ? sum / nums.length : sum);
            }
            return {
                intervalMinutes: target,
                labels: buildTimeLabels(target),
                values: out
            };
        }

        function computeAnomalySummary(values, labels) {
            const nums = [];
            let missing = 0;
            (values || []).forEach(v => {
                if (v === null || v === undefined || !isFinite(v)) {
                    missing += 1;
                } else {
                    nums.push(v);
                }
            });
            if (nums.length === 0) {
                return { missingCount: missing, outlierCount: 0, summary: missing ? '全部缺失' : '' };
            }

            const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
            const variance = nums.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / nums.length;
            const std = Math.sqrt(variance);

            const outlierIdx = [];
            if (std > 0) {
                (values || []).forEach((v, i) => {
                    if (v === null || v === undefined || !isFinite(v)) return;
                    if (Math.abs(v - mean) >= 3 * std) outlierIdx.push(i);
                });
            }

            const outlierCount = outlierIdx.length;
            const outlierTimes = outlierIdx.slice(0, 6).map(i => labels[i]).filter(Boolean);
            const outlierText = outlierTimes.length ? `异常点：${outlierTimes.join(', ')}${outlierIdx.length > outlierTimes.length ? '...' : ''}` : '';
            const missingText = missing ? `缺失点数：${missing}` : '';
            const summary = [missingText, outlierText].filter(Boolean).join('；');
            return { missingCount: missing, outlierCount, summary };
        }

        function build15MinRows(params = {}) {
            const opts = params.opts || {};
            if (!Array.isArray(appData.parsedData) || appData.parsedData.length === 0) {
                return { rows: [], baseName: '', error: '没有可导出的数据' };
            }

            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            const selectedDates = new Set(appData?.visualization?.selectedDates || []);

            const dailyTotalMap = new Map();
            (appData.processedData || []).forEach((row) => {
                const key = `${row.date}||${normalizeMeteringPoint(row.meteringPoint)}`;
                dailyTotalMap.set(key, row.dailyTotal);
            });

            const start = startDate ? new Date(`${startDate}T00:00:00`) : null;
            const end = endDate ? new Date(`${endDate}T00:00:00`) : null;

            const candidates = [];
            for (const parsedRow of appData.parsedData) {
                const date = toProcessedDateString(parsedRow.dateStr);
                
                // 联动修复：如果不是“导出全部”，且没有设置起止日期，则回退到“当前视图”
                // 如果勾选了“导出全部”，则跳过 selectedDates 检查
                if (!exportAllDates) {
                    if (start && end) {
                        const d = new Date(`${date}T00:00:00`);
                        if (!(d >= start && d <= end)) continue;
                    } else {
                        // 回退到当前视图筛选
                        if (!selectedDates.has(date)) continue;
                    }
                }
                
                if (opts.weekdaysOnly && !isWeekdayDate(date)) continue;
                
                candidates.push({
                    date,
                    meteringPoint: normalizeMeteringPoint(parsedRow.meteringPoint),
                    rawData: Array.isArray(parsedRow.rawData) ? parsedRow.rawData : [],
                    multiplier: typeof parsedRow.multiplier === 'number' && isFinite(parsedRow.multiplier) ? parsedRow.multiplier : (appData.config.multiplier || 1),
                    timeInterval: typeof parsedRow.timeInterval === 'number' && isFinite(parsedRow.timeInterval) ? parsedRow.timeInterval : (appData.config.timeInterval || 15)
                });
            }

            if (candidates.length === 0) {
                return { rows: [], baseName: '', error: '没有符合筛选条件的数据可导出' };
            }

            const baseIntervalMinutes = Math.max(1, parseInt(String(candidates[0].timeInterval || 15), 10) || 15);
            const requestedInterval = opts.samplingInterval === 'auto' ? baseIntervalMinutes : Math.max(1, parseInt(String(opts.samplingInterval), 10) || baseIntervalMinutes);
            const effectiveInterval = (requestedInterval >= baseIntervalMinutes && requestedInterval % baseIntervalMinutes === 0) ? requestedInterval : baseIntervalMinutes;
            const pointsPerDay = Math.max(1, Math.floor(24 * 60 / baseIntervalMinutes));

            const timeLabels = buildTimeLabels(effectiveInterval);
            const includeAnomaly = !!opts.includeAnomalyMarks;

            const header = includeAnomaly
                ? ['日期', '计量点', ...timeLabels, '日总电能 (kWh)', '缺失点数', '异常点数', '异常摘要']
                : ['日期', '计量点', ...timeLabels, '日总电能 (kWh)'];
            const rows = [header];

            const fillMethod = appData.config.invalidDataHandling === 'ignore' ? 'ignore' : 'interpolate';
            const agg = opts.samplingAgg === 'avg' ? 'avg' : 'sum';

            candidates.forEach(({ date, meteringPoint, rawData, multiplier }) => {
                const series = rawData.slice(0, pointsPerDay);
                const cleaned = DataUtils.cleanData(series, { fillMethod });
                const scaledBase = cleaned.map(v => (v === null || v === undefined) ? null : v * multiplier);
                while (scaledBase.length < pointsPerDay) scaledBase.push(null);

                const resampled = resampleSeries(scaledBase, baseIntervalMinutes, effectiveInterval, agg);
                const dailyTotalKey = `${date}||${meteringPoint}`;
                const dailyTotal = dailyTotalMap.has(dailyTotalKey) ? dailyTotalMap.get(dailyTotalKey) : '';

                if (includeAnomaly) {
                    const a = computeAnomalySummary(resampled.values, resampled.labels);
                    rows.push([date, meteringPoint, ...resampled.values, dailyTotal, a.missingCount, a.outlierCount, a.summary]);
                } else {
                    rows.push([date, meteringPoint, ...resampled.values, dailyTotal]);
                }
            });

            return { rows, baseName: effectiveInterval === 15 ? '15分钟明细' : `${effectiveInterval}分钟明细` };
        }

        function export15MinData(params = {}) {
            const opts = params.opts || getExportRuntimeOptions();
            const built = build15MinRows({ opts });
            if (built.error) {
                showNotification('错误', built.error, 'error');
                return;
            }
            setExportProgress(30, '导出中');
            exportDataToFile(built.rows, built.baseName, { fileName: params.fileName });
        }

        function buildHourlyRows(params = {}) {
            const opts = params.opts || {};
            if (!Array.isArray(appData.processedData) || appData.processedData.length === 0) {
                return { rows: [], baseName: '', error: '没有可导出的数据' };
            }

            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            const selected = Array.isArray(appData?.visualization?.selectedDates) ? appData.visualization.selectedDates : [];

            let exportData = [...appData.processedData];

            // 联动修复
            if (!exportAllDates) {
                if (startDate && endDate) {
                    const start = new Date(`${startDate}T00:00:00`);
                    const end = new Date(`${endDate}T00:00:00`);
                    exportData = exportData.filter(d => {
                        const currentDate = new Date(`${d.date}T00:00:00`);
                        return currentDate >= start && currentDate <= end;
                    });
                } else {
                    // 回退到当前视图
                    exportData = exportData.filter(d => selected.includes(d.date));
                }
            }

            if (opts.weekdaysOnly) {
                exportData = exportData.filter(d => isWeekdayDate(d.date));
            }

            if (exportData.length === 0) {
                return { rows: [], baseName: '', error: '没有符合筛选条件的数据可导出' };
            }

            const includeAnomaly = !!opts.includeAnomalyMarks;
            const header = includeAnomaly
                ? ['日期', ...Array.from({ length: 24 }, (_, h) => `${h}:00`), '日总电能 (kWh)', '缺失点数', '异常点数', '异常摘要']
                : ['日期', ...Array.from({ length: 24 }, (_, h) => `${h}:00`), '日总电能 (kWh)'];
            const rows = [header];

            const hourLabels = Array.from({ length: 24 }, (_, h) => `${h}:00`);
            exportData.forEach(dateData => {
                const series = Array.from({ length: 24 }, (_, h) => {
                    const v = dateData.hourlyData?.[h];
                    return (v === null || v === undefined || !isFinite(v)) ? null : v;
                });
                const dailyTotal = series.filter(v => v !== null).reduce((a, b) => a + b, 0);

                if (includeAnomaly) {
                    const a = computeAnomalySummary(series, hourLabels);
                    rows.push([dateData.date, ...series.map(v => v === null ? '' : v), dailyTotal, a.missingCount, a.outlierCount, a.summary]);
                } else {
                    rows.push([dateData.date, ...series.map(v => v === null ? '' : v), dailyTotal]);
                }
            });

            return { rows, baseName: '24小时电能明细' };
        }

        function exportHourlyData(params = {}) {
            const opts = params.opts || getExportRuntimeOptions();
            const built = buildHourlyRows({ opts });
            if (built.error) {
                showNotification('错误', built.error, 'error');
                return;
            }
            setExportProgress(30, '导出中');
            exportDataToFile(built.rows, built.baseName, { fileName: params.fileName });
        }

        function buildDailyStatsRows(params = {}) {
            const opts = params.opts || {};
            if (!Array.isArray(appData.processedData) || appData.processedData.length === 0) {
                return { rows: [], baseName: '', error: '没有可导出的数据' };
            }
            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            const selected = Array.isArray(appData?.visualization?.selectedDates) ? appData.visualization.selectedDates : [];

            let exportData = [...appData.processedData];

            // 联动修复
            if (!exportAllDates) {
                if (startDate && endDate) {
                    const start = new Date(`${startDate}T00:00:00`);
                    const end = new Date(`${endDate}T00:00:00`);
                    exportData = exportData.filter(d => {
                        const currentDate = new Date(`${d.date}T00:00:00`);
                        return currentDate >= start && currentDate <= end;
                    });
                } else {
                    // 回退到当前视图
                    exportData = exportData.filter(d => selected.includes(d.date));
                }
            }

            if (opts.weekdaysOnly) exportData = exportData.filter(d => isWeekdayDate(d.date));
            if (exportData.length === 0) {
                return { rows: [], baseName: '', error: '没有符合筛选条件的数据可导出' };
            }

            const includeAnomaly = !!opts.includeAnomalyMarks;
            const header = includeAnomaly
                ? ['日期', '日总用电量 (kWh)', '平均小时用电量 (kWh)', '峰谷差值 (kWh)', '缺失小时数', '异常小时数', '异常摘要']
                : ['日期', '日总用电量 (kWh)', '平均小时用电量 (kWh)', '峰谷差值 (kWh)'];
            const rows = [header];

            const hourLabels = Array.from({ length: 24 }, (_, h) => `${h}:00`);
            exportData.forEach(dateData => {
                const series = Array.from({ length: 24 }, (_, h) => {
                    const v = dateData.hourlyData?.[h];
                    return (v === null || v === undefined || !isFinite(v)) ? null : v;
                });
                const validHours = series.filter(v => v !== null);
                const hourlyAverage = validHours.length > 0 ? validHours.reduce((a, b) => a + b, 0) / validHours.length : 0;
                const peak = validHours.length > 0 ? Math.max(...validHours) : 0;
                const valley = validHours.length > 0 ? Math.min(...validHours) : 0;
                const peakValleyDiff = peak - valley;
                const dailyTotal = validHours.reduce((a, b) => a + b, 0);

                if (includeAnomaly) {
                    const a = computeAnomalySummary(series, hourLabels);
                    rows.push([dateData.date, dailyTotal.toFixed(2), hourlyAverage.toFixed(2), peakValleyDiff.toFixed(2), a.missingCount, a.outlierCount, a.summary]);
                } else {
                    rows.push([dateData.date, dailyTotal.toFixed(2), hourlyAverage.toFixed(2), peakValleyDiff.toFixed(2)]);
                }
            });

            return { rows, baseName: '日统计数据' };
        }

        function exportDailyStats(params = {}) {
            const opts = params.opts || getExportRuntimeOptions();
            const built = buildDailyStatsRows({ opts });
            if (built.error) {
                showNotification('错误', built.error, 'error');
                return;
            }
            setExportProgress(30, '导出中');
            exportDataToFile(built.rows, built.baseName, { fileName: params.fileName });
        }

        function exportPVsystData(params = {}) {
            if (appData.processedData.length === 0) {
                showNotification('错误', '没有可导出的数据', 'error');
                return;
            }

            setExportProgress(5, '导出中');

            // 始终使用所有处理后的数据
            let exportData = [...appData.processedData];
            console.log('=== PVSyst数据导出（按月份重新排列）===');
            console.log('原始数据条数:', exportData.length);

            // 检查所有日期的有效性
            console.log('检查所有日期的有效性...');
            let validDataCount = 0;
            
            exportData.forEach((row, index) => {
                let dateObj = null;
                
                // 尝试多种方式获取有效日期
                if (row.dateObj && !isNaN(row.dateObj.getTime())) {
                    dateObj = row.dateObj;
                } else if (row.date) {
                    // 增强的日期解析
                    dateObj = parseDate(row.date, appData.config.dateFormat);
                    if (!dateObj) {
                        // 尝试自动检测格式
                        dateObj = parseDate(row.date, 'auto');
                    }
                }
                
                if (!dateObj || isNaN(dateObj.getTime())) {
                    console.error(`无效日期在第${index}行: "${row.date}"`);
                    // 标记为无效
                    row._invalidDate = true;
                } else {
                    console.log(`有效日期在第${index}行: "${row.date}" -> ${dateObj.toISOString()}`);
                    // 确保dateObj存在
                    row.dateObj = dateObj;
                    validDataCount++;
                }
            });
            
            console.log(`有效数据: ${validDataCount}/${exportData.length}`);

            if (exportData.length === 0) {
                showNotification('错误', '没有可导出的数据', 'error');
                return;
            }

            // 创建基于真实日期的数据映射
            const realDateDataMap = new Map();
            
            if (exportData.length === 0) {
                showNotification('错误', '没有可导出的数据', 'error');
                return;
            }
            
            // 过滤掉无效日期的数据
            const validExportData = exportData.filter(dateData => !dateData._invalidDate);

            if (validExportData.length === 0) {
                showNotification('错误', '所有数据的日期格式都无效，无法导出', 'error');
                return;
            }
            
            const sortedExportData = [...validExportData].sort((a, b) => {
                return a.dateObj - b.dateObj;
            });

            // 构建真实日期的数据映射
            sortedExportData.forEach(dateData => {
                // 使用本地日期格式，避免UTC偏移
                const dateKey = formatDateForInput(dateData.dateObj); // YYYY-MM-DD格式
                
                if (dateData.hourlyData && Array.isArray(dateData.hourlyData) && dateData.hourlyData.length === 24) {
                    realDateDataMap.set(dateKey, dateData.hourlyData);
                } else {
                    realDateDataMap.set(dateKey, Array(24).fill(0));
                }
            });

            // 确定数据的实际年份范围
            const allRealDates = [...realDateDataMap.keys()].sort();
            let startRealDate, endRealDate;
            
            if (allRealDates.length > 0) {
                startRealDate = new Date(allRealDates[0]);
                endRealDate = new Date(allRealDates[allRealDates.length - 1]);
                console.log('实际数据日期范围:', formatDateForInput(startRealDate), '到', formatDateForInput(endRealDate));
            } else {
                console.log('警告：没有可用的实际数据');
                showNotification('警告', '没有可用的实际数据，将使用零值填充', 'warning');
            }
            console.log('实际数据天数:', realDateDataMap.size);

            // 创建月份顺序的日期映射（1月1日到12月31日）
            const monthOrderDateMap = new Map();
            
            // 如果没有实际数据，创建一个默认的零值数据集
            if (realDateDataMap.size === 0) {
                console.log('创建默认零值数据集');
                for (let targetMonth = 1; targetMonth <= 12; targetMonth++) {
                    const daysInMonth = new Date(2000, targetMonth, 0).getDate();
                    for (let targetDay = 1; targetDay <= daysInMonth; targetDay++) {
                        const monthDayKey = `${String(targetMonth).padStart(2, '0')}-${String(targetDay).padStart(2, '0')}`;
                        monthOrderDateMap.set(monthDayKey, {
                            data: Array(24).fill(0),
                            source: '零值填充'
                        });
                    }
                }
            } else {
                // 按月份从1月到12月的顺序处理
                for (let targetMonth = 1; targetMonth <= 12; targetMonth++) {
                    const daysInMonth = new Date(2000, targetMonth, 0).getDate();
                    
                    for (let targetDay = 1; targetDay <= daysInMonth; targetDay++) {
                        const monthDayKey = `${String(targetMonth).padStart(2, '0')}-${String(targetDay).padStart(2, '0')}`;
                        
                        // 查找对应的真实数据
                        let matchedData = null;
                        let dataSource = '';
                        
                        // 优先查找同月同日的数据
                        for (const [realDate, data] of realDateDataMap) {
                            const realDateObj = new Date(realDate);
                            if (isNaN(realDateObj.getTime())) continue;
                            
                            if (realDateObj.getMonth() + 1 === targetMonth && realDateObj.getDate() === targetDay) {
                                matchedData = [...data];
                                dataSource = `实际数据-${realDate}`;
                                break;
                            }
                        }
                        
                        // 如果没有找到完全匹配的，使用最接近的数据
                        if (!matchedData) {
                            // 查找相邻年份的同月同日数据
                            const nearbyData = findNearestMonthDayData(targetMonth, targetDay, realDateDataMap);
                            if (nearbyData) {
                                matchedData = [...nearbyData.data];
                                dataSource = nearbyData.source;
                            } else {
                                // 使用前一天的数据，或查找相近月份的数据
                                const fallbackData = findFallbackData(targetMonth, targetDay, monthOrderDateMap, realDateDataMap);
                                matchedData = [...fallbackData.data];
                                dataSource = fallbackData.source;
                            }
                        }
                        
                        monthOrderDateMap.set(monthDayKey, {
                            data: matchedData,
                            source: dataSource
                        });
                    }
                }
            }

            // 生成CSV内容
            let csvContent = '';
            csvContent += '# 自定义负载曲线 - 按月份重新排列\n';
            csvContent += '# 生成时间: ' + new Date().toISOString() + '\n';
            csvContent += '# 数据来源: 负荷数据处理工具\n';
            csvContent += '# 说明: 按月份从低到高排列，去除年份影响\n';
            csvContent += '# 规则: 按月份和日匹配，缺失数据用最近有效数据补全\n';
            csvContent += '# 格式: DD/MM/YY hh:mm，年份固定为99\n';
            csvContent += '\n';
            csvContent += 'Date;P Load\n';
            csvContent += ';kW\n';

            // 按月份顺序生成8760小时数据
            let dataUsageStats = {
                actual: 0,
                fallback: 0,
                details: []
            };

            for (let month = 1; month <= 12; month++) {
                setExportProgress(Math.round((month - 1) / 12 * 80) + 10, '导出中');
                const daysInMonth = new Date(2000, month, 0).getDate();
                
                for (let day = 1; day <= daysInMonth; day++) {
                    const monthDayKey = `${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const dayData = monthOrderDateMap.get(monthDayKey);
                    
                    if (dayData.source.startsWith('实际数据')) {
                        dataUsageStats.actual++;
                    } else {
                        dataUsageStats.fallback++;
                    }
                    
                    dataUsageStats.details.push({
                        monthDay: monthDayKey,
                        source: dayData.source,
                        total: dayData.data.reduce((a, b) => a + b, 0)
                    });

                    // 为每个小时生成数据行
                    for (let hour = 0; hour < 24; hour++) {
                        const value = dayData.data[hour] || 0;
                        const dayStr = String(day).padStart(2, '0');
                        const monthStr = String(month).padStart(2, '0');
                        const hourStr = String(hour).padStart(2, '0');
                        const dateTimeStr = `${dayStr}/${monthStr}/99 ${hourStr}:00`;
                        
                        csvContent += `${dateTimeStr};${value.toFixed(3)}\n`;
                    }
                }
            }

            // 输出统计信息
            console.log('=== 数据重新排列完成 ===');
            console.log('实际使用数据天数:', dataUsageStats.actual);
            console.log('补全数据天数:', dataUsageStats.fallback);
            console.log('总天数:', dataUsageStats.actual + dataUsageStats.fallback);
            console.log('数据覆盖率:', ((dataUsageStats.actual / (dataUsageStats.actual + dataUsageStats.fallback)) * 100).toFixed(1) + '%');

            // 导出CSV文件
            setExportProgress(95, '导出中');
            const fileName = params.fileName || getExportFileName('PVsyst月度重排8760小时数据', 'csv');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            saveAs(blob, fileName);
            updateExportStatus('PVsyst 数据已导出', fileName);

            showNotification('成功', 'PVsyst月度重排8760小时数据已导出', 'success');
        }


        function exportDataToFile(data, baseName, options = {}) {
            const format = String(options.formatOverride || elements.exportFormatSelect?.value || 'xlsx');
            const ext = format === 'csv' ? 'csv' : 'xlsx';
            const fileName = String(options.fileName || getExportFileName(baseName, ext));

            setExportProgress(60, '导出中');
            const ws = XLSX.utils.aoa_to_sheet(Array.isArray(data) ? data : []);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, String(baseName || 'Sheet1'));

            if (format === 'csv') {
                const csv = XLSX.utils.sheet_to_csv(ws);
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                saveAs(blob, fileName);
            } else {
                XLSX.writeFile(wb, fileName);
            }

            updateExportStatus(`${baseName} 已导出`, fileName);
            showNotification('成功', `${baseName}已导出为${format.toUpperCase()}`, 'success');
        }

        function exportPDFReport(params = {}) {
            if (!appData.processedData || appData.processedData.length === 0) {
                showNotification('错误', '请先导入并处理数据后再生成报告', 'error');
                return;
            }

            const opts = params.opts || getExportRuntimeOptions();
            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            
            let data = [...appData.processedData];
            
            // 联动修复：PDF报告也应遵循导出范围设置
            if (!exportAllDates) {
                if (startDate && endDate) {
                    const start = new Date(`${startDate}T00:00:00`);
                    const end = new Date(`${endDate}T00:00:00`);
                    data = data.filter(d => {
                        const currentDate = new Date(`${d.date}T00:00:00`);
                        return currentDate >= start && currentDate <= end;
                    });
                } else {
                    // 回退到当前视图
                    data = typeof getFilteredData === 'function' ? getFilteredData() : data;
                }
            }
            
            if (opts.weekdaysOnly) {
                data = data.filter(d => isWeekdayDate(d.date));
            }

            const startHour = Number(appData?.visualization?.focusStartTime ?? 0);
            const endHour = Number(appData?.visualization?.focusEndTime ?? 23);
            const summary = computeGlobalSummary(data, startHour, endHour);
            
            // 生成多页PDF报告
            generateMultiPagePDFReport(data, summary, params);
        }
        
        // 生成多页PDF专业报告
        function generateMultiPagePDFReport(data, summary, params = {}) {
            const fileName = params.fileName || getExportFileName('专业报告', 'pdf');
            const reportDate = new Date().toLocaleString('zh-CN');
            setExportProgress(30, '导出中');

            // 获取图表图片
            const chartUrl = getCanvasPreviewUrl(elements.loadCurveChart);
            const barChartUrl = getCanvasPreviewUrl(elements.dailyTotalBarChart);
            
            // 计算统计数据
            const allPowerValues = [];
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach(val => {
                        if (val !== null && !isNaN(val) && val > 0) {
                            allPowerValues.push(val);
                        }
                    });
                }
            });
            
            const avgLoad = allPowerValues.length > 0 ? allPowerValues.reduce((sum, v) => sum + v, 0) / allPowerValues.length : 0;
            const maxLoad = allPowerValues.length > 0 ? Math.max(...allPowerValues) : 0;
            const minLoad = allPowerValues.length > 0 ? Math.min(...allPowerValues) : 0;
            const loadFactor = maxLoad > 0 ? (avgLoad / maxLoad * 100).toFixed(1) : '0.0';
            
            // 计算时段分布
            const timeSlots = [
                { name: '凌晨', range: '00:00-06:00', hours: [0, 1, 2, 3, 4, 5] },
                { name: '上午', range: '06:00-12:00', hours: [6, 7, 8, 9, 10, 11] },
                { name: '下午', range: '12:00-18:00', hours: [12, 13, 14, 15, 16, 17] },
                { name: '晚上', range: '18:00-24:00', hours: [18, 19, 20, 21, 22, 23] }
            ];
            
            const totalEnergy = data.reduce((sum, d) => sum + (d.dailyTotal || 0), 0);
            
            // 计算日间/夜间数据
            let dayPowerSum = 0, dayCount = 0, dayEnergy = 0;
            let nightPowerSum = 0, nightCount = 0, nightEnergy = 0;
            
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach((val, hour) => {
                        if (val !== null && !isNaN(val) && val > 0) {
                            if (hour >= 6 && hour < 18) {
                                dayPowerSum += val;
                                dayCount++;
                                dayEnergy += val;
                            } else {
                                nightPowerSum += val;
                                nightCount++;
                                nightEnergy += val;
                            }
                        }
                    });
                }
            });
            
            const dayAvg = dayCount > 0 ? dayPowerSum / dayCount : 0;
            const nightAvg = nightCount > 0 ? nightPowerSum / nightCount : 0;
            const dayRatio = totalEnergy > 0 ? (dayEnergy / totalEnergy * 100).toFixed(1) : 0;
            const nightRatio = totalEnergy > 0 ? (nightEnergy / totalEnergy * 100).toFixed(1) : 0;
            
            // 计算峰值和谷值时段
            let peakHour = 0, peakValue = 0;
            let valleyHour = 0, valleyValue = Infinity;
            const hourlyAvg = new Array(24).fill(0).map(() => ({ sum: 0, count: 0 }));
            
            data.forEach(d => {
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach((val, hour) => {
                        if (val !== null && !isNaN(val) && val > 0) {
                            hourlyAvg[hour].sum += val;
                            hourlyAvg[hour].count++;
                        }
                    });
                }
            });
            
            hourlyAvg.forEach((h, hour) => {
                const avg = h.count > 0 ? h.sum / h.count : 0;
                if (avg > peakValue) {
                    peakValue = avg;
                    peakHour = hour;
                }
                if (avg < valleyValue && avg > 0) {
                    valleyValue = avg;
                    valleyHour = hour;
                }
            });
            
            // 月度数据
            const monthlyData = {};
            data.forEach(d => {
                const month = d.date?.substring(0, 7) || '未知';
                if (!monthlyData[month]) {
                    monthlyData[month] = { energy: 0, allPowerValues: [], count: 0 };
                }
                monthlyData[month].energy += d.dailyTotal || 0;
                if (d.hourlyData && Array.isArray(d.hourlyData)) {
                    d.hourlyData.forEach(val => {
                        if (val !== null && !isNaN(val) && val > 0) {
                            monthlyData[month].allPowerValues.push(val);
                        }
                    });
                }
                monthlyData[month].count++;
            });
            
            // 高峰/低谷日
            const dailyData = {};
            data.forEach(d => {
                if (!dailyData[d.date]) {
                    dailyData[d.date] = 0;
                }
                dailyData[d.date] += d.dailyTotal || 0;
            });
            const sortedDays = Object.entries(dailyData).sort((a, b) => b[1] - a[1]);
            
            const fmt2 = (n) => smartFormatNumber(n, 2);
            const fmt0 = (n) => smartFormatNumber(n, 0);
            
            // 生成优化建议
            let suggestion1 = '', suggestion2 = '', suggestion3 = '';
            
            // 错峰用电建议
            if (peakHour >= 9 && peakHour <= 17) {
                suggestion1 = `根据分析，当前负荷峰值出现在 ${String(peakHour).padStart(2, '0')}:00 左右。建议将部分非关键生产任务调整至 ${String((peakHour + 12) % 24).padStart(2, '0')}:00-${String((peakHour + 13) % 24).padStart(2, '0')}:00 时段运行，可有效降低峰值负荷，预计可节省10-15%的需量电费。`;
            } else if (peakHour >= 18 && peakHour <= 21) {
                suggestion1 = `当前负荷峰值出现在晚间 ${String(peakHour).padStart(2, '0')}:00，建议在日间 10:00-16:00 安排大功率设备运行，充分利用电网低谷时段，可降低约15-20%的用电成本。`;
            } else {
                suggestion1 = `建议在低谷时段（22:00-06:00）安排大功率设备运行，可有效利用分时电价政策，预计可降低约15-20%的用电成本。`;
            }
            
            // 负荷平衡建议
            const peakValleyDiff = maxLoad > 0 ? ((maxLoad - minLoad) / maxLoad * 100).toFixed(1) : 0;
            if (parseFloat(peakValleyDiff) > 50) {
                suggestion2 = `当前峰谷差较大（${peakValleyDiff}%），负荷率较低。建议引入储能系统，在低谷时段充电、高峰时段放电，可平滑负荷曲线约30-40%，同时提高变压器利用率。`;
            } else if (parseFloat(peakValleyDiff) > 30) {
                suggestion2 = `当前负荷有一定波动，建议通过调整生产排班，将部分工序从峰值时段转移至谷值时段，优化设备启停顺序，可提高设备利用率约15-20%。`;
            } else {
                suggestion2 = `当前负荷曲线较为平稳，设备利用率良好。建议继续保持，定期监测负荷变化，及时调整生产计划。`;
            }
            
            // 新能源配置建议
            if (parseFloat(dayRatio) > 60) {
                suggestion3 = `当前日间用电占比高达${dayRatio}%，非常适合配置光伏发电系统。建议安装装机容量约 ${(avgLoad * 8 / 1000).toFixed(1)} kWp 的光伏电站，日均可自发自用约 ${(parseFloat(dayRatio) / 100 * avgLoad * 6).toFixed(0)} kWh，年均可节省电费约15-25%。`;
            } else if (parseFloat(dayRatio) > 40) {
                suggestion3 = `当前日间用电占比${dayRatio}%，建议配置中小型光伏发电系统（约 ${(avgLoad * 5 / 1000).toFixed(1)} kWp），配合储能设备，可在日间高峰期提供部分电力，减少电网购电。`;
            } else {
                suggestion3 = `当前夜间用电占比较高，日间用电相对较少。如需配置新能源，建议优先考虑小型光伏或风力发电，主要用于办公区域照明和辅助设备，可作为备用电源使用。`;
            }
            
            setExportProgress(50, '导出中');
            
            // 生成时段分布表格HTML
            let timeTableHtml = '';
            timeSlots.forEach(slot => {
                let slotPowerSum = 0;
                let slotDataCount = 0;
                let slotEnergy = 0;
                
                data.forEach(d => {
                    if (d.hourlyData && Array.isArray(d.hourlyData)) {
                        slot.hours.forEach(hour => {
                            const value = d.hourlyData[hour];
                            if (value !== null && !isNaN(value) && value > 0) {
                                slotPowerSum += value;
                                slotDataCount++;
                                slotEnergy += value;
                            }
                        });
                    }
                });
                
                const slotAvgLoad = slotDataCount > 0 ? slotPowerSum / slotDataCount : 0;
                const ratio = totalEnergy > 0 ? (slotEnergy / totalEnergy * 100).toFixed(1) : 0;
                
                let level = '低';
                let levelColor = '#10b981';
                if (ratio > 30) {
                    level = '高';
                    levelColor = '#ef4444';
                } else if (ratio > 20) {
                    level = '中';
                    levelColor = '#f59e0b';
                }
                
                timeTableHtml += `
                    <tr>
                        <td style="padding:12px 16px;font-weight:600;color:#0f172a;">${slot.name}</td>
                        <td style="padding:12px 16px;text-align:center;color:#475569;">${slot.range}</td>
                        <td style="padding:12px 16px;text-align:right;font-weight:600;color:#0f172a;">${slotAvgLoad.toFixed(2)} kW</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${ratio}%</td>
                        <td style="padding:12px 16px;text-align:center;">
                            <span style="padding:4px 12px;border-radius:20px;font-size:12px;font-weight:600;background:${levelColor}20;color:${levelColor};">${level}</span>
                        </td>
                    </tr>
                `;
            });
            
            // 生成月度汇总表格HTML
            let monthlyTableHtml = '';
            Object.keys(monthlyData).sort().forEach(month => {
                const m = monthlyData[month];
                const validValues = m.allPowerValues;
                const mAvgLoad = validValues.length > 0 ? validValues.reduce((sum, v) => sum + v, 0) / validValues.length : 0;
                const mMaxLoad = validValues.length > 0 ? Math.max(...validValues) : 0;
                const mMinLoad = validValues.length > 0 ? Math.min(...validValues) : 0;
                const mLoadFactor = mMaxLoad > 0 ? (mAvgLoad / mMaxLoad * 100).toFixed(1) : '0.0';
                
                monthlyTableHtml += `
                    <tr>
                        <td style="padding:12px 16px;font-weight:600;color:#0f172a;">${month}</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${fmt2(m.energy)}</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${fmt2(mAvgLoad)}</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${fmt2(mMaxLoad)}</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${fmt2(mMinLoad)}</td>
                        <td style="padding:12px 16px;text-align:right;color:#475569;">${mLoadFactor}%</td>
                    </tr>
                `;
            });
            
            // 生成高峰/低谷日HTML
            let peakDaysHtml = sortedDays.slice(0, 5).map(([date, energy]) => `
                <tr>
                    <td style="padding:8px 0;color:#475569;">${date}</td>
                    <td style="padding:8px 0;text-align:right;font-weight:600;color:#ef4444;">${fmt2(energy)}</td>
                </tr>
            `).join('');
            
            let valleyDaysHtml = sortedDays.slice(-5).reverse().map(([date, energy]) => `
                <tr>
                    <td style="padding:8px 0;color:#475569;">${date}</td>
                    <td style="padding:8px 0;text-align:right;font-weight:600;color:#3b82f6;">${fmt2(energy)}</td>
                </tr>
            `).join('');
            
            // 获取 AI 分析内容
            let aiAnalysisHtml = '';
            const aiContentEl = document.getElementById('aiAnalysisContent');
            if (aiContentEl && aiContentEl.innerText && aiContentEl.innerText.trim().length > 10) {
                // 简单的 Markdown 转 HTML 处理
                const rawText = aiContentEl.innerText;
                const lines = rawText.split('\n');
                let processedHtml = '';
                
                lines.forEach(line => {
                    line = line.trim();
                    if (!line) return;
                    
                    if (line.startsWith('### ')) {
                        processedHtml += `<h4>${line.substring(4)}</h4>`;
                    } else if (line.startsWith('## ')) {
                        processedHtml += `<h3>${line.substring(3)}</h3>`;
                    } else if (line.startsWith('**') && line.endsWith('**')) {
                        processedHtml += `<p><strong>${line.substring(2, line.length-2)}</strong></p>`;
                    } else if (line.startsWith('- ')) {
                        processedHtml += `<li>${line.substring(2)}</li>`;
                    } else if (/^\d+\./.test(line)) {
                        processedHtml += `<p class="list-item"><strong>${line}</strong></p>`;
                    } else {
                        // 处理行内加粗
                        line = line.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
                        processedHtml += `<p>${line}</p>`;
                    }
                });
                
                // 包装列表
                processedHtml = processedHtml.replace(/<li>/g, '<ul style="margin:8px 0;padding-left:20px;"><li>').replace(/<\/li>(?!<ul)/g, '</li></ul>');
                // 修复连续列表的标签
                processedHtml = processedHtml.replace(/<\/ul><ul[^>]*>/g, '');
                
                aiAnalysisHtml = `
                <div class="page">
                    <h2>🤖 AI 智能分析报告</h2>
                    <div class="highlight-box" style="background:linear-gradient(135deg, #eef2ff 0%, #fff 100%);border-color:#c7d2fe;">
                        <div style="font-size:14px;color:#312e81;line-height:1.8;">
                            ${processedHtml}
                        </div>
                    </div>
                    <div class="footer">
                        <p>本报告由 24小时负荷曲线生成工具 自动生成</p>
                    </div>
                    <div class="page-number">- 附录：AI 分析 -</div>
                </div>
                `;
            }
            
            setExportProgress(70, '导出中');
            
            // CSS样式
            const css = `
                :root{--bg:#f8fafc;--text:#0f172a;--muted:#475569;--line:#e2e8f0;--card:#fff;--primary:#4f46e5;--primary-light:#e0e7ff;--success:#10b981;--warning:#f59e0b;--danger:#ef4444;}
                *{box-sizing:border-box;}
                body{margin:0;padding:0;font-family:ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Helvetica Neue,Arial; background:var(--bg);color:var(--text);line-height:1.6;}
                .page{width:210mm;min-height:297mm;padding:20mm;margin:0 auto;background:#fff;box-shadow:0 0 20px rgba(0,0,0,0.1);page-break-after:always;}
                .page:last-child{page-break-after:auto;}
                h1{font-size:28px;margin:0 0 8px;font-weight:800;letter-spacing:-0.5px;color:var(--text);}
                h2{font-size:20px;margin:0 0 16px;font-weight:700;color:var(--text);display:flex;align-items:center;gap:8px;}
                h2::before{content:'';width:4px;height:20px;background:var(--primary);border-radius:2px;}
                h3{font-size:16px;margin:0 0 12px;font-weight:600;color:var(--text);}
                .sub{color:var(--muted);font-size:14px;margin-bottom:24px;}
                .header{text-align:center;padding:40px 0;border-bottom:2px solid var(--primary-light);margin-bottom:32px;}
                .header-icon{width:64px;height:64px;background:linear-gradient(135deg, var(--primary) 0%, #6366f1 100%);border-radius:16px;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;box-shadow:0 8px 24px rgba(79,70,229,0.3);}
                .header-icon svg{width:32px;height:32px;color:#fff;}
                .badge{display:inline-flex;align-items:center;gap:4px;padding:6px 12px;border-radius:20px;font-size:12px;font-weight:600;background:#f1f5f9;color:#475569;margin:0 4px;}
                .badge-primary{background:var(--primary-light);color:var(--primary);}
                .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:24px 0;}
                .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,0.05);}
                .card-header{display:flex;align-items:center;gap:12px;margin-bottom:12px;}
                .card-icon{width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;}
                .card-icon.blue{background:#e0e7ff;color:#4f46e5;}
                .card-icon.green{background:#e0e7ff;color:#4338ca;}
                .card-icon.amber{background:#e0e7ff;color:#3730a3;}
                .card-icon.rose{background:#e0e7ff;color:#312e81;}
                .card-label{font-size:12px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;}
                .card-value{font-size:24px;font-weight:800;color:var(--text);margin-top:4px;}
                .card-unit{font-size:14px;color:var(--muted);font-weight:600;}
                .section{margin-top:32px;padding-top:24px;border-top:1px solid var(--line);}
                .stats-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:16px 0;}
                .stat-item{background:#f8fafc;border-radius:10px;padding:16px;}
                .stat-label{font-size:12px;color:var(--muted);margin-bottom:4px;}
                .stat-value{font-size:18px;font-weight:700;color:var(--text);}
                .highlight-box{background:linear-gradient(135deg, var(--primary-light) 0%, #fff 100%);border:1px solid #c7d2fe;border-radius:12px;padding:20px;margin:16px 0;}
                .highlight-box h4{font-size:14px;margin:0 0 12px;color:var(--primary);font-weight:700;}
                .highlight-box p{margin:8px 0;font-size:13px;color:#3730a3;line-height:1.7;}
                .highlight-box ul{margin:8px 0;padding-left:20px;}
                .highlight-box li{margin:6px 0;font-size:13px;color:#3730a3;}
                table{width:100%;border-collapse:collapse;margin:16px 0;}
                th{background:#f8fafc;padding:12px 16px;text-align:left;font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:0.5px;border-bottom:2px solid var(--line);}
                td{border-bottom:1px solid var(--line);}
                tr:hover td{background:#f8fafc;}
                .suggestion-card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px;margin:16px 0;display:flex;gap:16px;}
                .suggestion-icon{width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
                .suggestion-icon.green{background:#e0e7ff;color:#4f46e5;}
                .suggestion-icon.blue{background:#e0e7ff;color:#4f46e5;}
                .suggestion-icon.amber{background:#e0e7ff;color:#4f46e5;}
                .suggestion-content{flex:1;}
                .suggestion-title{font-size:16px;font-weight:700;margin-bottom:8px;}
                .suggestion-title.green{color:#312e81;}
                .suggestion-title.blue{color:#312e81;}
                .suggestion-title.amber{color:#312e81;}
                .suggestion-text{font-size:14px;color:var(--muted);line-height:1.6;}
                .chart-container{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px;margin:16px 0;text-align:center;}
                .chart-container img{max-width:100%;border-radius:8px;}
                .chart-caption{font-size:12px;color:var(--muted);margin-top:12px;font-weight:500;}
                .two-col{display:grid;grid-template-columns:1fr 1fr;gap:24px;}
                .footer{text-align:center;padding:24px;color:var(--muted);font-size:12px;border-top:1px solid var(--line);margin-top:32px;}
                .page-number{text-align:center;padding:16px;color:var(--muted);font-size:12px;}
                @media print{
                    body{background:#fff;}
                    .page{width:100%;min-height:auto;padding:15mm;box-shadow:none;page-break-after:always;}
                    .no-print{display:none;}
                    .page:last-child{page-break-after:auto;}
                }
            `;
            
            // 生成HTML
            const html = `
                <!doctype html>
                <html lang="zh-CN">
                <head>
                    <meta charset="utf-8" />
                    <meta name="viewport" content="width=device-width,initial-scale=1" />
                    <title>电力负荷数据分析报告</title>
                    <style>${css}</style>
                </head>
                <body>
                    <!-- 第1页：概览与关键指标 -->
                    <div class="page">
                        <div class="header">
                            <div class="header-icon">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
                            </div>
                            <h1>电力负荷数据分析报告</h1>
                            <p class="sub">Load Data Analysis Report</p>
                            <div style="margin-top:16px;">
                                <span class="badge badge-primary">📅 ${summary.dateRange || '--'}</span>
                                <span class="badge">📊 ${data.length} 天数据</span>
                                <span class="badge">⏰ ${summary.hoursLabel || '--'}</span>
                            </div>
                        </div>
                        
                        <div class="grid">
                            <div class="card">
                                <div class="card-header">
                                    <div class="card-icon blue"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v20M2 12h20"/></svg></div>
                                    <div>
                                        <div class="card-label">总用电量</div>
                                        <div class="card-value">${fmt2(totalEnergy)} <span class="card-unit">kWh</span></div>
                                    </div>
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-header">
                                    <div class="card-icon green"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12h18M3 6h18M3 18h18"/></svg></div>
                                    <div>
                                        <div class="card-label">平均负荷</div>
                                        <div class="card-value">${fmt2(avgLoad)} <span class="card-unit">kW</span></div>
                                    </div>
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-header">
                                    <div class="card-icon amber"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg></div>
                                    <div>
                                        <div class="card-label">最大负荷</div>
                                        <div class="card-value">${fmt2(maxLoad)} <span class="card-unit">kW</span></div>
                                    </div>
                                </div>
                            </div>
                            <div class="card">
                                <div class="card-header">
                                    <div class="card-icon rose"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v20M2 12h20"/></svg></div>
                                    <div>
                                        <div class="card-label">最小负荷</div>
                                        <div class="card-value">${fmt2(minLoad)} <span class="card-unit">kW</span></div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="section">
                            <h2>📊 数据摘要</h2>
                            <div class="stats-grid">
                                <div class="stat-item">
                                    <div class="stat-label">数据天数</div>
                                    <div class="stat-value">${data.length} 天</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-label">峰值出现时间</div>
                                    <div class="stat-value">${String(peakHour).padStart(2, '0')}:00-${String(peakHour + 1).padStart(2, '0')}:00</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-label">负荷率</div>
                                    <div class="stat-value" style="color:${parseFloat(loadFactor) > 70 ? '#10b981' : '#f59e0b'};">${loadFactor}%</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-label">峰谷差</div>
                                    <div class="stat-value">${fmt2(maxLoad - minLoad)} kW</div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="highlight-box">
                            <h4>💡 初步分析结论</h4>
                            <p>根据数据分析，该负荷曲线呈现${parseFloat(dayRatio) > 60 ? '典型的工作日用电模式，日间用电量占比较高' : '较为均衡的用电模式'}。建议在低谷时段（${String(valleyHour).padStart(2, '0')}:00-${String(valleyHour + 1).padStart(2, '0')}:00）安排大功率设备运行，以降低用电成本。</p>
                        </div>
                        
                        <div class="footer">
                            <p>本报告由 24小时负荷曲线生成工具 自动生成</p>
                            <p>生成时间：${reportDate}</p>
                        </div>
                        <div class="page-number">- 第 1 页 -</div>
                    </div>
                    
                    ${aiAnalysisHtml}
                    
                    <!-- 第2页：24小时负荷曲线分析 -->
                    <div class="page">
                        <h2>📈 24小时负荷曲线分析</h2>
                        
                        <div class="chart-container">
                            ${chartUrl ? `<img src="${chartUrl}" alt="24小时负荷曲线图" style="max-height:300px;" />` : '<p>未能生成图表</p>'}
                            <div class="chart-caption">图1：24小时负荷曲线对比图</div>
                        </div>
                        
                        <div class="two-col" style="margin-top:24px;">
                            <div class="card">
                                <h3 style="color:#f59e0b;">☀️ 日间负荷特征 (06:00-18:00)</h3>
                                <div style="margin-top:16px;">
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">日间平均负荷：</strong>${dayAvg.toFixed(2)} kW</p>
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">日间用电量占比：</strong>${dayRatio} %</p>
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">峰值出现时段：</strong>${String(peakHour).padStart(2, '0')}:00-${String(peakHour + 1).padStart(2, '0')}:00</p>
                                </div>
                            </div>
                            <div class="card">
                                <h3 style="color:#6366f1;">🌙 夜间负荷特征 (18:00-06:00)</h3>
                                <div style="margin-top:16px;">
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">夜间平均负荷：</strong>${nightAvg.toFixed(2)} kW</p>
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">夜间用电量占比：</strong>${nightRatio} %</p>
                                    <p style="margin:8px 0;color:#475569;"><strong style="color:#0f172a;">谷值出现时段：</strong>${String(valleyHour).padStart(2, '0')}:00-${String(valleyHour + 1).padStart(2, '0')}:00</p>
                                </div>
                            </div>
                        </div>
                        
                        <div class="footer">
                            <p>本报告由 24小时负荷曲线生成工具 自动生成</p>
                        </div>
                        <div class="page-number">- 第 2 页 -</div>
                    </div>
                    
                    <!-- 第3页：时段分析与优化建议 -->
                    <div class="page">
                        <h2>⏰ 时段分析与优化建议</h2>
                        
                        <div class="section" style="margin-top:0;border-top:none;">
                            <h3>各时段用电分布</h3>
                            <table>
                                <thead>
                                    <tr>
                                        <th>时段</th>
                                        <th style="text-align:center;">时间范围</th>
                                        <th style="text-align:right;">平均负荷</th>
                                        <th style="text-align:right;">用电量占比</th>
                                        <th style="text-align:center;">负荷等级</th>
                                    </tr>
                                </thead>
                                <tbody>${timeTableHtml}</tbody>
                            </table>
                        </div>
                        
                        <div class="section">
                            <h3>优化建议</h3>
                            
                            <div class="suggestion-card">
                                <div class="suggestion-icon green">
                                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>
                                </div>
                                <div class="suggestion-content">
                                    <div class="suggestion-title green">错峰用电建议</div>
                                    <div class="suggestion-text">${suggestion1}</div>
                                </div>
                            </div>
                            
                            <div class="suggestion-card">
                                <div class="suggestion-icon blue">
                                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="10" rx="2"/><path d="M6 7v10M18 7v10"/></svg>
                                </div>
                                <div class="suggestion-content">
                                    <div class="suggestion-title blue">负荷平衡建议</div>
                                    <div class="suggestion-text">${suggestion2}</div>
                                </div>
                            </div>
                            
                            <div class="suggestion-card">
                                <div class="suggestion-icon amber">
                                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>
                                </div>
                                <div class="suggestion-content">
                                    <div class="suggestion-title amber">新能源配置建议</div>
                                    <div class="suggestion-text">${suggestion3}</div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="footer">
                            <p>本报告由 24小时负荷曲线生成工具 自动生成</p>
                        </div>
                        <div class="page-number">- 第 3 页 -</div>
                    </div>
                    
                    <!-- 第4页：详细数据汇总 -->
                    <div class="page">
                        <h2>📋 详细数据汇总</h2>
                        
                        <div class="section" style="margin-top:0;border-top:none;">
                            <h3>月度用电汇总</h3>
                            <table>
                                <thead>
                                    <tr>
                                        <th>月份</th>
                                        <th style="text-align:right;">总用电量(kWh)</th>
                                        <th style="text-align:right;">平均负荷(kW)</th>
                                        <th style="text-align:right;">最大负荷(kW)</th>
                                        <th style="text-align:right;">最小负荷(kW)</th>
                                        <th style="text-align:right;">负荷率(%)</th>
                                    </tr>
                                </thead>
                                <tbody>${monthlyTableHtml}</tbody>
                            </table>
                        </div>
                        
                        <div class="two-col" style="margin-top:24px;">
                            <div class="card" style="border-left:4px solid #ef4444;">
                                <h3 style="color:#ef4444;">🔺 用电高峰日 TOP 5</h3>
                                <table style="margin-top:12px;">
                                    <thead>
                                        <tr>
                                            <th style="padding:8px 0;">日期</th>
                                            <th style="padding:8px 0;text-align:right;">用电量(kWh)</th>
                                        </tr>
                                    </thead>
                                    <tbody>${peakDaysHtml}</tbody>
                                </table>
                            </div>
                            <div class="card" style="border-left:4px solid #3b82f6;">
                                <h3 style="color:#3b82f6;">🔻 用电低谷日 TOP 5</h3>
                                <table style="margin-top:12px;">
                                    <thead>
                                        <tr>
                                            <th style="padding:8px 0;">日期</th>
                                            <th style="padding:8px 0;text-align:right;">用电量(kWh)</th>
                                        </tr>
                                    </thead>
                                    <tbody>${valleyDaysHtml}</tbody>
                                </table>
                            </div>
                        </div>
                        
                        ${barChartUrl ? `
                        <div class="chart-container" style="margin-top:24px;">
                            <img src="${barChartUrl}" alt="日期总用电量柱状图" style="max-height:250px;" />
                            <div class="chart-caption">图2：日期总用电量对比图</div>
                        </div>
                        ` : ''}
                        
                        <div class="footer">
                            <p>本报告由 24小时负荷曲线生成工具 自动生成</p>
                            <p>生成时间：${reportDate}</p>
                        </div>
                        <div class="page-number">- 第 4 页 -</div>
                    </div>
                    
                    <div class="no-print" style="position:fixed;bottom:20px;right:20px;display:flex;gap:12px;">
                        <button onclick="window.print()" style="border:none;background:var(--primary);color:#fff;border-radius:10px;padding:12px 24px;font-weight:700;cursor:pointer;box-shadow:0 4px 12px rgba(79,70,229,0.3);">
                            🖨️ 打印 / 另存为PDF
                        </button>
                    </div>
                </body>
                </html>
            `;
            
            setExportProgress(90, '导出中');
            
            // 打开新窗口显示报告
            const w = window.open('', '_blank');
            if (!w) {
                showNotification('错误', '浏览器拦截了弹窗，请允许后重试', 'error');
                return;
            }
            
            w.document.open();
            w.document.write(html);
            w.document.close();
            
            setExportProgress(100, '完成');
            updateExportStatus('PDF 专业报告已生成', fileName);
            showNotification('成功', '多页PDF专业报告已生成，请查看', 'success');
        }
        
        // 生成PDF分析报告文本
        function generatePDFAnalysisText(data, summary) {
            const findings = [];
            const suggestions = [];
            
            // 基于数据生成发现
            if (summary.maxTotal && summary.minTotal) {
                const ratio = summary.maxTotal / summary.minTotal;
                if (ratio > 3) {
                    findings.push(`负荷波动较大，最大日用电量是最小日的 ${ratio.toFixed(1)} 倍，建议关注用电稳定性`);
                } else if (ratio < 1.5) {
                    findings.push(`负荷较为平稳，日用电量波动控制在 ${ratio.toFixed(1)} 倍以内`);
                }
            }
            
            if (summary.totalEnergy) {
                const avgDaily = summary.totalEnergy / (data.length || 1);
                findings.push(`平均日用电量为 ${avgDaily.toFixed(2)} kWh`);
            }
            
            // 生成建议
            suggestions.push('建议定期监控负荷曲线，及时发现异常用电模式');
            suggestions.push('可考虑实施峰谷电价策略，降低用电成本');
            
            if (summary.avgLoad && summary.maxTotal) {
                const loadFactor = (summary.avgLoad * 24 / summary.maxTotal);
                if (loadFactor < 0.6) {
                    suggestions.push('负荷率偏低，建议优化设备运行时段，提高用电效率');
                }
            }
            
            return {
                overview: `本报告分析了 ${data.length} 天的负荷数据，总用电量 ${summary.totalEnergy?.toFixed(2) || '—'} kWh。数据显示${findings.length > 0 ? '存在明显的用电特征' : '整体用电平稳'}，建议根据分析结果优化用电策略。`,
                findings: findings.length > 0 ? findings : ['数据正常，未发现明显异常'],
                suggestions: suggestions,
                peakHours: '08:00-10:00, 18:00-21:00',
                valleyHours: '23:00-07:00'
            };
        }

        function buildOverviewRows(dataInput) {
            const startHour = Number(appData?.visualization?.focusStartTime ?? 0);
            const endHour = Number(appData?.visualization?.focusEndTime ?? 23);
            const data = dataInput || (typeof getFilteredData === 'function' ? getFilteredData() : (appData.processedData || []));
            const summary = computeGlobalSummary(data, startHour, endHour);
            const rows = [];
            rows.push(['字段', '值']);
            rows.push(['日期范围', summary.dateRange || '--']);
            rows.push(['统计时段', summary.hoursLabel || '--']);
            rows.push(['计量点', summary.meteringPointLabel || '--']);
            rows.push(['总用电量(kWh)', summary.totalEnergy]);
            rows.push(['平均负荷(kWh/h)', summary.avgLoad]);
            rows.push(['最大负荷日', summary.maxDay || '']);
            rows.push(['最大负荷(kWh)', summary.maxTotal]);
            rows.push(['最小负荷日', summary.minDay || '']);
            rows.push(['最小负荷(kWh)', summary.minTotal]);
            return rows;
        }

        function buildFullPackageWorkbook(opts = {}) {
            const wb = XLSX.utils.book_new();
            
            // 联动修复：多表导出也应遵循范围设置
            const exportAllDates = document.getElementById('exportAllDates')?.checked ?? true;
            const startDate = document.getElementById('exportStartDate')?.value || '';
            const endDate = document.getElementById('exportEndDate')?.value || '';
            const selected = Array.isArray(appData?.visualization?.selectedDates) ? appData.visualization.selectedDates : [];
            
            let dataForOverview = [...appData.processedData];
            if (!exportAllDates) {
                if (startDate && endDate) {
                    const start = new Date(`${startDate}T00:00:00`);
                    const end = new Date(`${endDate}T00:00:00`);
                    dataForOverview = dataForOverview.filter(d => {
                        const currentDate = new Date(`${d.date}T00:00:00`);
                        return currentDate >= start && currentDate <= end;
                    });
                } else {
                    dataForOverview = dataForOverview.filter(d => selected.includes(d.date));
                }
            }
            if (opts.weekdaysOnly) {
                dataForOverview = dataForOverview.filter(d => isWeekdayDate(d.date));
            }

            const overview = XLSX.utils.aoa_to_sheet(buildOverviewRows(dataForOverview));
            XLSX.utils.book_append_sheet(wb, overview, '概况');

            const d15 = build15MinRows({ opts });
            if (d15.rows && d15.rows.length) {
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(d15.rows), d15.baseName || '15min明细');
            }

            const d24 = buildHourlyRows({ opts });
            if (d24.rows && d24.rows.length) {
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(d24.rows), '24h汇总');
            }

            const ds = buildDailyStatsRows({ opts });
            if (ds.rows && ds.rows.length) {
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ds.rows), '日统计');
            }

            return wb;
        }

        function exportFullPackage(params = {}) {
            if (!appData.processedData || appData.processedData.length === 0) {
                showNotification('错误', '没有可导出的数据', 'error');
                return;
            }
            const opts = params.opts || getExportRuntimeOptions();
            const fileName = params.fileName || getExportFileName('完整数据包', 'xlsx');

            setExportProgress(10, '导出中');
            const wb = buildFullPackageWorkbook(opts);
            setExportProgress(70, '导出中');
            XLSX.writeFile(wb, fileName);
            updateExportStatus('完整数据包已导出', fileName);
            showNotification('成功', '已导出多表合一Excel数据包', 'success');
        }

        function canvasToBlobPromise(canvas, type, quality) {
            return new Promise((resolve) => {
                if (!canvas || typeof canvas.toBlob !== 'function') {
                    resolve(null);
                    return;
                }
                canvas.toBlob((blob) => resolve(blob), type, quality);
            });
        }

        async function exportZipBundle(params = {}) {
            const Zip = window.JSZip;
            if (!Zip) {
                showNotification('错误', '未能加载JSZip，无法打包ZIP（请检查网络或稍后重试）', 'error');
                return;
            }
            if (!appData.processedData || appData.processedData.length === 0) {
                showNotification('错误', '没有可导出的数据', 'error');
                return;
            }

            const opts = params.opts || getExportRuntimeOptions();
            const zipName = params.fileName || getExportFileName('打包导出', 'zip');
            const zip = new Zip();

            setExportProgress(10, '导出中');
            const wb = buildFullPackageWorkbook(opts);
            const wbArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            zip.file(getExportFileName('完整数据包', 'xlsx'), wbArray);

            setExportProgress(40, '导出中');
            const chartBlob = await canvasToBlobPromise(elements.loadCurveChart, 'image/png', 0.92);
            if (chartBlob) {
                zip.file(getExportFileName('图表', 'png'), chartBlob);
            }

            setExportProgress(70, '导出中');
            const zipBlob = await zip.generateAsync({ type: 'blob' }, (meta) => {
                const p = Math.round(70 + (meta.percent || 0) * 0.25);
                setExportProgress(p, '导出中');
            });

            saveAs(zipBlob, zipName);
            updateExportStatus('ZIP 打包已导出', zipName);
            showNotification('成功', '已打包导出ZIP（图表+完整数据包）', 'success');
        }

        function findNearbyData(month, day, dataMap) {
    // 寻找相近月份的数据
    const searchOrder = [
        [month, day],       // 当前月份
        [month-1, day],     // 前一个月
        [month+1, day],     // 后一个月
        [month, day-1],     // 前一天
        [month, day+1],     // 后一天
        [month-2, day],     // 前两个月
        [month+2, day]      // 后两个月
    ];
    
    for (const [m, d] of searchOrder) {
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            const key = `${m}-${d}`;
            if (dataMap.has(key)) {
                return key;
            }
        }
    }
    
    // 如果找不到相近日期，尝试基于季节查找数据
    // 根据月份确定季节
    let season;
    if (month >= 3 && month <= 5) {
        season = 'spring'; // 春季
    } else if (month >= 6 && month <= 8) {
        season = 'summer'; // 夏季
    } else if (month >= 9 && month <= 11) {
        season = 'autumn'; // 秋季
    } else {
        season = 'winter'; // 冬季
    }
    
    // 根据季节查找对应月份的数据
    let seasonMonths = [];
    switch(season) {
        case 'spring':
            seasonMonths = [3, 4, 5];
            break;
        case 'summer':
            seasonMonths = [6, 7, 8];
            break;
        case 'autumn':
            seasonMonths = [9, 10, 11];
            break;
        case 'winter':
            seasonMonths = [12, 1, 2];
            break;
    }
    
    // 在同季节中查找数据
    for (const m of seasonMonths) {
        // 优先查找月中数据（15号左右）
        for (const d of [15, 14, 16, 13, 17, 12, 18, 11, 19, 10, 20]) {
            if (d <= new Date(2000, m, 0).getDate()) { // 确保日期有效
                const key = `${m}-${d}`;
                if (dataMap.has(key)) {
                    return key;
                }
            }
        }
    }
    
    // 如果还是找不到，查找任何可用的数据
    const allKeys = [...dataMap.keys()];
    if (allKeys.length > 0) {
        // 优先选择同季节的数据
        const seasonKeys = allKeys.filter(key => {
            const keyMonth = parseInt(key.split('-')[0]);
            return seasonMonths.includes(keyMonth);
        });
        
        if (seasonKeys.length > 0) {
            // 随机选择一个同季节的数据，避免总是使用同一个
            const randomIndex = Math.floor(Math.random() * seasonKeys.length);
            return seasonKeys[randomIndex];
        } else {
            // 如果没有同季节的数据，随机选择一个可用数据
            const randomIndex = Math.floor(Math.random() * allKeys.length);
            return allKeys[randomIndex];
        }
    }
    
    return null;
}

        function findNearestMonthDayData(targetMonth, targetDay, realDateDataMap) {
            // 查找最接近的月日数据
            if (!realDateDataMap || realDateDataMap.size === 0) {
                return null;
            }
            
            let bestMatch = null;
            let minDistance = Infinity;
            
            for (const [realDate, data] of realDateDataMap) {
                const realDateObj = new Date(realDate);
                if (isNaN(realDateObj.getTime())) {
                    console.warn('无效的日期在findNearestMonthDayData中:', realDate);
                    continue;
                }
                
                const realMonth = realDateObj.getMonth() + 1;
                const realDay = realDateObj.getDate();
                
                // 计算月日差异
                const monthDiff = Math.abs(realMonth - targetMonth);
                const dayDiff = Math.abs(realDay - targetDay);
                
                // 考虑月份循环（12月到1月只算1个月差异）
                const adjustedMonthDiff = Math.min(monthDiff, 12 - monthDiff);
                
                // 计算综合距离
                const distance = adjustedMonthDiff * 31 + dayDiff;
                
                if (distance < minDistance) {
                    minDistance = distance;
                    bestMatch = { data, source: `最近匹配-${realDate}` };
                }
            }
            
            return bestMatch;
        }

        function findFallbackData(targetMonth, targetDay, monthOrderDateMap, realDateDataMap) {
            // 查找回退数据
            
            // 1. 尝试使用前一天的数据
            if (targetDay > 1) {
                const prevDayKey = `${String(targetMonth).padStart(2, '0')}-${String(targetDay - 1).padStart(2, '0')}`;
                const prevData = monthOrderDateMap.get(prevDayKey);
                if (prevData) {
                    return {
                        data: [...prevData.data],
                        source: `前一天数据-${prevDayKey}`
                    };
                }
            }
            
            // 2. 尝试使用同月的其他日数据
            if (realDateDataMap && realDateDataMap.size > 0) {
                const nearestMonthDay = findNearestMonthDayData(targetMonth, targetDay, realDateDataMap);
                if (nearestMonthDay) {
                    return nearestMonthDay;
                }
            }
            
            // 3. 使用任意可用数据
            if (realDateDataMap && realDateDataMap.size > 0) {
                const anyData = [...realDateDataMap.values()][0];
                if (anyData) {
                    return {
                        data: [...anyData],
                        source: '默认数据'
                    };
                }
            }
            
            // 4. 返回零值
            return {
                data: Array(24).fill(0),
                source: '零值填充'
            };
        }

        function updateDataOverview() {
            const dataStatus = document.getElementById('dataStatus');
            const currentDataCount = document.getElementById('currentDataCount');
            const currentDateRange = document.getElementById('currentDateRange');
            const currentMeteringPointCount = document.getElementById('currentMeteringPointCount');
            
            if (appData.parsedData && appData.parsedData.length > 0) {
                if (dataStatus && dataStatus.parentNode) {
                    dataStatus.classList.remove('hidden');
                }
                
                // 显示当前数据量
                if (currentDataCount && currentDataCount.parentNode) {
                    currentDataCount.textContent = appData.parsedData.length.toLocaleString();
                }
                
                // 计算日期范围
                const dates = [...new Set(appData.parsedData.map(item => item.date))].sort();
                if (dates.length > 0) {
                    const startDate = dates[0];
                    const endDate = dates[dates.length - 1];
                    // 使用formatDateRangeDisplay函数格式化日期范围
                    const formattedStartDate = formatDateRangeDisplay(startDate);
                    const formattedEndDate = formatDateRangeDisplay(endDate);
                    if (currentDateRange && currentDateRange.parentNode) {
                        currentDateRange.textContent = `${formattedStartDate} ~ ${formattedEndDate}`;
                    }
                } else {
                    if (currentDateRange && currentDateRange.parentNode) {
                        currentDateRange.textContent = '-';
                    }
                }
                
                // 计算计量点数量
                const uniqueMeteringPoints = [...new Set(appData.parsedData.map(item => item.meteringPoint))];
                if (currentMeteringPointCount && currentMeteringPointCount.parentNode) {
                    currentMeteringPointCount.textContent = uniqueMeteringPoints.length;
                }
            } else {
                if (dataStatus && dataStatus.parentNode) {
                    dataStatus.classList.add('hidden');
                }
            }
        }
        
        function startOver() {
            // 重置应用数据
            appData.files = [];
            appData.worksheets = [];
            appData.parsedData = [];
            appData.processedData = [];
            appData.config = {
                dataType: 'instantPower',
                multiplier: 1.0,
                useMultiplierColumn: false,
                multiplierColumn: '',
                dateColumn: '',

                dataStartColumn: '',
                dataEndColumn: '',
                dataStartIndex: 0,
                dataEndIndex: -1,
                meteringPointColumn: '',
                meteringPointFilter: '',
                dateFormat: 'YYYY-MM-DD',
                invalidDataHandling: 'ignore',
                timeInterval: 15,
                additionalFilters: []
            };
            appData.visualization = {
                selectedDates: [],
                selectedMeteringPoints: [],
                focusStartTime: 0,
                focusEndTime: 23,
                lineStyle: 'solid',
                showPoints: false,
                smoothCurve: true,
                curveTension: 0.4,
                secondaryAxis: false,
                showDailyTotalBarChart: true,
                // 新增：汇总模式（最高/最低 + 季度平均）
                summaryMode: false
            };
            appData.meteringPoints = [];
            
            // 重置UI
            elements.fileInput.value = '';
            elements.importedFiles.innerHTML = '';
            // 文件列表现在固定展开，无需隐藏
            if (elements.removeAllFiles && elements.removeAllFiles.parentNode) {
                elements.removeAllFiles.classList.add('hidden');
            }

            
            // 重置配置表单
            elements.dataTypeRadios[0].checked = true;
            elements.multiplierInput.value = '1.0';
            
            // 重置倍率配置
            elements.useMultiplierColumnCheckbox.checked = false;
            if (elements.multiplierOptions) {
                elements.multiplierOptions.classList.remove('hidden');
            }
            if (elements.multiplierColumnOption) {
                elements.multiplierColumnOption.classList.add('hidden');
            }
            elements.dateColumnSelect.innerHTML = '<option value="">选择列</option>';
            elements.dataStartColumnSelect.innerHTML = '<option value="">起始列</option>';
            elements.dataEndColumnSelect.innerHTML = '<option value="">结束列</option>';
            elements.meteringPointColumnSelect.innerHTML = '<option value="">计量点列</option>';
            elements.meteringPointFilterSelect.innerHTML = '<option value="">全部计量点</option>';
            elements.multiplierColumnSelect.innerHTML = '<option value="">倍率列</option>';

            elements.invalidDataHandlingSelect.value = 'ignore';
            elements.timeIntervalSelect.value = '15';
            
            // 重置预览区域
            if (elements.previewButtons) {
                elements.previewButtons.classList.add('hidden');
            }
            if (elements.dataPreview) {
                elements.dataPreview.classList.add('bg-white', 'border', 'border-slate-200', 'shadow-sm');
                elements.dataPreview.innerHTML = `
                    <div class="flex flex-col items-center justify-center py-24 text-center">
                        <div class="w-20 h-20 mb-6 rounded-3xl bg-slate-50 border border-slate-100 flex items-center justify-center text-slate-300">
                            <i class="fa fa-file-import text-3xl"></i>
                        </div>
                        <h3 class="text-base font-black text-slate-900 mb-2">等待数据导入</h3>
                        <p class="text-xs font-medium text-slate-500 max-w-[240px] leading-relaxed">
                            导入数据后，系统将自动识别并在此展示实时预览。
                        </p>
                    </div>
                `;
            }
            
            // 重置可视化表单
            elements.startDateInput.innerHTML = '<option value="">起始日期</option>';
            elements.endDateInput.innerHTML = '<option value="">结束日期</option>';
            elements.focusStartTimeSelect.value = '0';
            elements.focusEndTimeSelect.value = '23';
            elements.lineStyleSelect.value = 'solid';
            elements.showPointsSelect.value = 'false';
            elements.smoothCurveSelect.value = 'true';
            elements.curveTensionSelect.value = '0.4';
            elements.secondaryAxisCheckbox.checked = false;
            elements.dailyStats.innerHTML = '<p class="text-neutral">请先生成图表以查看统计信息</p>';
            elements.periodStats.innerHTML = '<p class="text-neutral">请先生成图表以查看统计信息</p>';
            
            // 重置日期选择
            appData.visualization.selectedDates = [];
            if (appData.chart) {
                appData.chart.destroy();
                appData.chart = null;
            }
            
            // 返回到导入部分
            goToImportSection();
            
            showNotification('信息', '已重置所有数据，可以重新开始', 'info');
        }

        // 导航函数
        function goToImportSection() {
            // 使用平滑过渡效果
            const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            allSections.forEach(id => {
                const section = document.getElementById(id);
                if (section) {
                    if (!section.classList.contains('section-container')) {
                        section.classList.add('section-container');
                    }
                    if (id === 'importSection') {
                        section.classList.remove('section-hidden');
                        section.classList.add('section-visible');
                    } else {
                        section.classList.remove('section-visible');
                        section.classList.add('section-hidden');
                    }
                }
            });
            
            // 更新步骤指示器（添加安全检查）
            if (elements.step2Indicator) {
                elements.step2Indicator.className = 'w-12 h-12 rounded-full bg-gray-300 text-gray-600 flex items-center justify-center font-semibold text-sm shadow-md hover:shadow-lg transition-all duration-300 cursor-pointer hover:bg-gray-400';
            }
            if (elements.step2Line) {
                elements.step2Line.className = 'w-8 h-0.5 bg-gray-300 rounded-full';
            }
            if (elements.step3Indicator) {
                elements.step3Indicator.className = 'w-12 h-12 rounded-full bg-gray-300 text-gray-600 flex items-center justify-center font-semibold text-sm shadow-md hover:shadow-lg transition-all duration-300 cursor-pointer hover:bg-gray-400';
            }
            if (elements.step3Line) {
                elements.step3Line.className = 'w-8 h-0.5 bg-gray-300 rounded-full';
            }
            if (elements.step4Indicator) {
                elements.step4Indicator.className = 'w-12 h-12 rounded-full bg-gray-300 text-gray-600 flex items-center justify-center font-semibold text-sm shadow-md hover:shadow-lg transition-all duration-300 cursor-pointer hover:bg-gray-400';
            }
        }

        function goToConfigSection() {
            // 使用平滑过渡效果
            const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            allSections.forEach(id => {
                const section = document.getElementById(id);
                if (section) {
                    if (!section.classList.contains('section-container')) {
                        section.classList.add('section-container');
                    }
                    if (id === 'configSection') {
                        section.classList.remove('section-hidden');
                        section.classList.add('section-visible');
                    } else {
                        section.classList.remove('section-visible');
                        section.classList.add('section-hidden');
                    }
                }
            });
            
            // 更新步骤导航状态
            updateStepNavigation('configSection');
            
            // 标记第一步为已完成
            completeStep(1);
            
            // 自动触发数据解析和预览更新
            console.log('进入配置步骤，自动触发数据解析...');
            setTimeout(async () => {
                if (appData.worksheets && appData.worksheets.length > 0) {
                    showGlobalLoading('正在解析数据...');
                    try {
                        // 解析工作表数据
                        await parseWorksheetData();
                        
                        // 处理解析后的数据
                        await processParsedData();
                        
                        // 更新计量点筛选选项
                        updateMeteringPointFilterOptions();
                        
                        // 更新数据预览
                        updateDataPreview();
                        
                        console.log('数据解析和预览更新完成，共处理数据点:', appData.processedData.length);
                        
                        // 显示成功通知
                        if (appData.processedData.length > 0) {
                            showNotification('成功', `数据解析完成，共处理 ${appData.processedData.length} 条记录`, 'success');
                        } else {
                            showNotification('信息', '数据解析完成，但未找到有效数据', 'warning');
                        }
                        
                    } catch (error) {
                        console.error('数据解析出错:', error);
                        showNotification('错误', '数据解析失败: ' + error.message, 'error');
                    } finally {
                        hideGlobalLoading();
                    }
                } else {
                    console.log('没有工作表数据可供解析');
                }
            }, 100);
        }

        async function goToVisualizationSection() {
    console.log('进入可视化步骤，使用最新配置重新处理数据...');
    
    // 使用最新配置重新处理数据，确保第三步可视化使用第二步的最新配置
    await reprocessDataWithConfig();
    
    // 更新计量点编号筛选选项（reprocessDataWithConfig已调用，但再调用一次确保最新）
    updateMeteringPointFilterOptions();
    
    // 更新步骤导航状态
    updateStepNavigation('visualizationSection');
    
    // 标记第二步为已完成
    completeStep(2);
    
    // 切换显示部分 - 使用平滑过渡效果
    const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
    allSections.forEach(id => {
        const section = document.getElementById(id);
        if (section) {
            if (!section.classList.contains('section-container')) {
                section.classList.add('section-container');
            }
            if (id === 'visualizationSection') {
                section.classList.remove('section-hidden');
                section.classList.add('section-visible');
            } else {
                section.classList.remove('section-visible');
                section.classList.add('section-hidden');
            }
        }
    });
    
    // 检查是否已存在图表实例，如果存在则先销毁
    if (appData.chart) {
        appData.chart.destroy();
        appData.chart = null;
    }
    
    // 初始化可视化
            initializeVisualization();
            
            // 初始化日期总用电量柱状图
            initDailyTotalBarChart();
            
            // 确保图表显示数据（reprocessDataWithConfig已包含此逻辑，但再调用一次确保）
            setTimeout(() => {
                updateChart();
                
                // 添加调试信息检查数据状态
                console.log('=== 页面初始化完成后的数据状态 ===');
                console.log('processedData长度:', appData.processedData.length);
                console.log('柱状图实例存在:', !!appData.dailyTotalBarChart);
                
                if (appData.processedData.length > 0) {
                    console.log('第一条数据:', appData.processedData[0]);
                    console.log('第一条数据dailyTotal:', appData.processedData[0].dailyTotal);
                    console.log('第一条数据类型:', typeof appData.processedData[0].dailyTotal);
                    
                    // 验证数据格式并准备柱状图数据
                    const chartData = appData.processedData.map(item => ({
                        date: String(item.date || '').replace(/-/g, ''),
                        dailyTotal: Number(item.dailyTotal) || 0,
                        hourlyData: item.hourlyData || []
                    })).filter(item => item.dailyTotal > 0);
                    
                    console.log('格式化后的图表数据:', chartData);
                    
                    if (chartData.length > 0 && appData.dailyTotalBarChart) {
                        console.log('使用实际数据更新柱状图...');
                        updateDailyTotalBarChart(chartData);
                    } else {
                        console.log('数据格式有问题或柱状图未初始化');
                    }
                } else {
                    console.log('没有数据，创建测试数据来验证柱状图功能...');
                    // 创建测试数据
                    const testData = [
                        { date: '2024-09-01', dailyTotal: 2354.28, hourlyData: Array(24).fill(0).map((_, i) => 98) },
                        { date: '2024-09-02', dailyTotal: 2213.70, hourlyData: Array(24).fill(0).map((_, i) => 92) },
                        { date: '2024-09-03', dailyTotal: 2079.54, hourlyData: Array(24).fill(0).map((_, i) => 87) },
                        { date: '2024-09-04', dailyTotal: 2171.19, hourlyData: Array(24).fill(0).map((_, i) => 90) },
                        { date: '2024-09-05', dailyTotal: 2188.41, hourlyData: Array(24).fill(0).map((_, i) => 91) }
                    ];
                    console.log('测试数据:', testData);
                    
                    // 测试柱状图更新
                    if (appData.dailyTotalBarChart) {
                        console.log('使用测试数据更新柱状图...');
                        updateDailyTotalBarChart(testData);
                    }
                }
                console.log('===============================');
            }, 500);
            
}

        function goToExportSection() {
            // 使用平滑过渡效果
            const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            allSections.forEach(id => {
                const section = document.getElementById(id);
                if (section) {
                    if (!section.classList.contains('section-container')) {
                        section.classList.add('section-container');
                    }
                    if (id === 'exportSection') {
                        section.classList.remove('section-hidden');
                        section.classList.add('section-visible');
                    } else {
                        section.classList.remove('section-visible');
                        section.classList.add('section-hidden');
                    }
                }
            });
            
            // 更新步骤导航状态
            updateStepNavigation('exportSection');
            
            // 标记第三步为已完成
            completeStep(3);
        }

        // 通知和模态框函数
        function showNotification(title, message, type = 'info', duration = 5000) {
            const iconMap = {
                success: 'fa-check-circle',
                error: 'fa-times-circle',
                warning: 'fa-exclamation-triangle',
                info: 'fa-info-circle'
            };
            
            // 防御性检查所有通知相关元素
            if (!elements.notification || !elements.notificationTitle || !elements.notificationMessage || !elements.notificationIcon || !elements.notificationIconContainer) {
                console.warn('通知元素不存在，无法显示通知:', title, message);
                return;
            }
            
            elements.notificationTitle.textContent = title;
            elements.notificationMessage.textContent = message;
            
            // 设置图标类名
            const iconClass = iconMap[type] || iconMap.info;
            elements.notificationIcon.className = `fa ${iconClass}`;
            
            if (elements.notification && elements.notification.parentNode) {
                // 重置容器样式
                elements.notification.classList.remove('hidden');
                
                // 设置图标容器的背景色，保持主容器为纯白
                const typeColorMap = {
                    success: 'bg-indigo-50 text-indigo-600 border-indigo-100',
                    error: 'bg-slate-50 text-slate-600 border-slate-100',
                    warning: 'bg-indigo-50 text-indigo-500 border-indigo-100',
                    info: 'bg-indigo-50 text-indigo-600 border-indigo-100'
                };
                
                const typeStyle = typeColorMap[type] || typeColorMap.info;
                elements.notificationIconContainer.className = `mr-4 flex h-10 w-10 items-center justify-center rounded-xl shadow-inner border transition-all duration-300 ${typeStyle}`;
            }
            
            // 显示动画
            setTimeout(() => {
                if (elements.notification && elements.notification.parentNode) {
                    elements.notification.classList.remove('translate-y-20', 'opacity-0', 'scale-95');
                }
            }, 10);
            
            // 自动隐藏（使用传入的时长）
            setTimeout(hideNotification, duration);
        }

        function hideNotification() {
            if (!elements.notification || !elements.notification.parentNode) {
                return;
            }
            
            elements.notification.classList.add('translate-y-20', 'opacity-0', 'scale-95');
            setTimeout(() => {
                if (elements.notification && elements.notification.parentNode) {
                    elements.notification.classList.add('hidden');
                }
            }, 300);
        }
        
        // 增强的时间间隔检测结果通知显示
        function showEnhancedIntervalNotification(detectionResult) {
            const confidence = Math.round(detectionResult.confidence * 100);
            let title, message, type;
            
            // 根据置信度和检测方法确定通知类型和内容
            if (detectionResult.confidence >= 0.9) {
                title = '🎯 高精度检测';
                type = 'success';
                message = `${detectionResult.message}\n检测方法：${detectionResult.method}\n置信度：${confidence}%（非常可靠）`;
            } else if (detectionResult.confidence >= 0.8) {
                title = '✅ 检测成功';
                type = 'success';
                message = `${detectionResult.message}\n检测方法：${detectionResult.method}\n置信度：${confidence}%（可靠）`;
            } else if (detectionResult.confidence >= 0.6) {
                title = '⚠️ 检测结果';
                type = 'warning';
                message = `${detectionResult.message}\n检测方法：${detectionResult.method}\n置信度：${confidence}%（建议手动确认）`;
            } else if (detectionResult.confidence >= 0.4) {
                title = '❓ 低置信度检测';
                type = 'warning';
                message = `${detectionResult.message}\n检测方法：${detectionResult.method}\n置信度：${confidence}%（建议手动设置）`;
            } else {
                title = '🔧 使用默认值';
                type = 'info';
                message = `${detectionResult.message}\n检测方法：${detectionResult.method}\n置信度：${confidence}%（请手动设置正确的时间间隔）`;
            }
            
            // 添加详细信息（如果有）
            if (detectionResult.details) {
                message += `\n详细信息：${detectionResult.details}`;
            }
            
            showNotification(title, message, type);
        }

        // 曲线样式弹窗
        function openCurveStyleModal() {
            if (elements.curveStyleModal) {
                elements.curveStyleModal.classList.remove('hidden');
                setTimeout(() => {
                    const modalContent = elements.curveStyleModal.querySelector('.bg-white');
                    if (modalContent) {
                        modalContent.classList.remove('scale-95', 'opacity-0');
                        modalContent.classList.add('scale-100', 'opacity-100');
                    }
                }, 10);
            }
        }

        function closeCurveStyleModal() {
            if (elements.curveStyleModal) {
                const modalContent = elements.curveStyleModal.querySelector('.bg-white');
                if (modalContent) {
                    modalContent.classList.remove('scale-100', 'opacity-100');
                    modalContent.classList.add('scale-95', 'opacity-0');
                }
                setTimeout(() => {
                    elements.curveStyleModal.classList.add('hidden');
                }, 300);
            }
        }

        function showModal(title, content) {
            if (elements.modalTitle) {
                elements.modalTitle.textContent = title;
            }
            if (elements.modalContent) {
                elements.modalContent.innerHTML = content;
            }
            
            if (elements.modalOverlay && elements.modalOverlay.parentNode) {
                elements.modalOverlay.classList.remove('hidden');
            }
            setTimeout(() => {
                if (elements.modal && elements.modal.parentNode) {
                    elements.modal.classList.remove('scale-95', 'opacity-0');
                    elements.modal.classList.add('scale-100', 'opacity-100');
                }
            }, 10);
        }

        function closeModal() {
            if (elements.modal && elements.modal.parentNode) {
                elements.modal.classList.remove('scale-100', 'opacity-100');
                elements.modal.classList.add('scale-95', 'opacity-0');
            }
            setTimeout(() => {
                if (elements.modalOverlay && elements.modalOverlay.parentNode) {
                    elements.modalOverlay.classList.add('hidden');
                }
            }, 300);
        }

        function showHelpModal() {
            const helpContent = `
                <div class="space-y-4">
                    <div>
                        <h4 class="font-medium text-primary mb-2">如何使用本工具？</h4>
                        <p class="text-sm">本工具分为四个步骤：导入数据 → 数据配置 → 可视化 → 导出结果。请按照步骤完成操作。</p>
                    </div>
                    
                    <div>
                        <h4 class="font-medium text-primary mb-2">支持的数据格式</h4>
                        <p class="text-sm">支持Excel (.xlsx, .xls) 和CSV (.csv) 格式的用电数据表格。</p>
                    </div>
                    
                    <div>
                        <h4 class="font-medium text-primary mb-2">数据类型说明</h4>
                        <ul class="list-disc pl-5 text-sm space-y-1">
                            <li><strong>15分钟瞬时功率</strong>：每15分钟一个功率值(kW)，工具会计算每小时的电能</li>
                            <li><strong>15分钟累积电能</strong>：每15分钟一个累积电能值(kWh)，工具会计算每小时的电能差值</li>
                        </ul>
                    </div>
                    
                    <div>
                        <h4 class="font-medium text-primary mb-2">常见问题</h4>
                        <div class="space-y-2 text-sm">
                            <p><strong>Q: 为什么导入文件后没有数据？</strong></p>
                            <p>A: 请检查数据配置是否正确，特别是日期列和数据列的选择是否准确。</p>
                            
                            <p><strong>Q: 如何处理无效数据？</strong></p>
                            <p>A: 工具提供两种处理方式：忽略无效数据或用相邻有效数据填充。</p>
                            
                            <p><strong>Q: 可以同时导入多个文件吗？</strong></p>
                            <p>A: 可以，工具支持多文件导入并合并处理数据。</p>
                        </div>
                    </div>
                </div>
            `;
            
            showModal('使用帮助', helpContent);
        }

        function showAboutModal() {
            const aboutContent = `
                <div class="space-y-4">
                    <div class="text-center">
                        <i class="fa fa-line-chart text-4xl text-primary mb-2"></i>
                        <h4 class="text-xl font-bold text-primary">24小时负荷曲线生成工具</h4>
                        <p class="text-neutral">版本 1.0.0</p>
                    </div>
                    
                    <div>
                        <p class="text-sm">本工具用于处理用电数据表格，生成24小时负荷曲线图表，支持按日期区分曲线类别，适配不同格式的原始数据。</p>
                    </div>
                    
                    <div>
                        <h4 class="font-medium mb-2">主要功能</h4>
                        <ul class="list-disc pl-5 text-sm space-y-1">
                            <li>支持多文件导入处理汇总</li>
                            <li>处理15分钟瞬时功率和累积电能数据</li>
                            <li>按日期范围筛选显示曲线</li>
                            <li>自定义显示时段，聚焦用电高峰</li>
                            <li>计算并展示日总用电量等统计信息</li>
                            <li>导出图表为PNG/JPG/SVG，导出数据为Excel/CSV</li>
                        </ul>
                    </div>
                </div>
            `;
            
            showModal('关于工具', aboutContent);
        }

        // 工具函数
        function parseNumber(value) {
            if (value === null || value === undefined || value === '') return null;
            
            // 检查是否为无效数据标记
            const strValue = String(value).trim().toLowerCase();
            if (strValue === '未采集' || strValue === '故障' || strValue === '无数据' || strValue === 'n/a' || strValue === 'null') {
                return null;
            }
            
            // 处理千分位分隔符（如1,234.56或1.234,56）
            let processedStr = strValue;
            
            // 检测并处理千分位分隔符
            // 情况1：1,234.56（逗号为千分位，点号为小数点）
            if (processedStr.includes(',') && processedStr.includes('.')) {
                // 如果逗号在点号前面，且逗号后不是3位数字，则可能是欧洲格式（1.234,56）
                const commaPos = processedStr.indexOf(',');
                const dotPos = processedStr.indexOf('.');
                
                if (commaPos < dotPos && (dotPos - commaPos) !== 4) {
                    // 可能是欧洲格式，将点号替换为空，逗号替换为点号
                    processedStr = processedStr.replace(/\./g, '').replace(',', '.');
                } else {
                    // 标准格式，移除逗号
                    processedStr = processedStr.replace(/,/g, '');
                }
            } 
            // 情况2：1.234,56（点号为千分位，逗号为小数点）
            else if (processedStr.includes(',') && processedStr.lastIndexOf(',') > processedStr.length - 4) {
                // 逗号在最后3位之前，可能是欧洲格式的小数点
                processedStr = processedStr.replace(/\./g, '').replace(',', '.');
            }
            // 情况3：只有逗号，可能是千分位或小数点
            else if (processedStr.includes(',')) {
                const commaPos = processedStr.indexOf(',');
                const afterComma = processedStr.substring(commaPos + 1);
                
                if (afterComma.length === 3 && commaPos > 0) {
                    // 可能是千分位分隔符
                    processedStr = processedStr.replace(/,/g, '');
                } else {
                    // 可能是小数点
                    processedStr = processedStr.replace(',', '.');
                }
            }
            
            // 处理百分比（如12.34%或12,34%）
            if (processedStr.includes('%')) {
                processedStr = processedStr.replace('%', '');
                const num = parseFloat(processedStr);
                return isNaN(num) ? null : num / 100;
            }
            
            // 处理科学计数法（如1.23e+4或1.23E-4）
            if (/[eE][+-]?\d+$/.test(processedStr)) {
                const num = parseFloat(processedStr);
                return isNaN(num) ? null : num;
            }
            
            // 处理货币符号（如$123.45或¥123.45）
            processedStr = processedStr.replace(/[$￥€£¥]/g, '');
            
            // 处理单位（如123.45kW或123.45kWh）
            const unitMatch = processedStr.match(/^([\d.]+)([a-zA-Z]+)$/);
            if (unitMatch) {
                const numStr = unitMatch[1];
                const unit = unitMatch[2].toLowerCase();
                
                const num = parseFloat(numStr);
                if (isNaN(num)) return null;
                
                // 处理常见单位
                switch (unit) {
                    case 'k':
                        return num * 1000;
                    case 'm':
                        return num * 1000000;
                    case 'g':
                        return num * 1000000000;
                    case 't':
                        return num * 1000;
                    default:
                        // 其他单位不转换，只返回数字部分
                        return num;
                }
            }
            
            // 尝试解析为数字
            const num = parseFloat(processedStr);
            return isNaN(num) ? null : num;
        }

        function isDateLike(value) {
            if (!value) return false;
            
            const strValue = String(value).trim();
            if (!strValue) return false;
            
            // 扩展的日期格式检测模式
            const datePatterns = [
                /^\d{4}-\d{1,2}-\d{1,2}$/, // YYYY-MM-DD 或 YYYY-M-D
                /^\d{4}\/\d{1,2}\/\d{1,2}$/, // YYYY/MM/DD 或 YYYY/M/D
                /^\d{1,2}\/\d{1,2}\/\d{4}$/, // MM/DD/YYYY 或 DD/MM/YYYY
                /^\d{1,2}\/\d{1,2}\/\d{2}$/, // MM/DD/YY 或 M/D/YY
                /^\d{1,2}-\d{1,2}-\d{4}$/, // DD-MM-YYYY 或 D-M-YYYY
                /^\d{1,2}-\d{1,2}-\d{2}$/, // DD-MM-YY 或 D-M-YY
                /^\d{4}\.\d{1,2}\.\d{1,2}$/, // YYYY.MM.DD
                /^\d{1,2}\.\d{1,2}\.\d{4}$/, // DD.MM.YYYY
                /^\d{1,2}\.\d{1,2}\.\d{2}$/, // DD.MM.YY
                /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}/, // MM/DD/YYYY HH:MM
                /^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}/, // YYYY-MM-DD HH:MM
                /^\d{1,2}\/\d{1,2}\/\d{2}\s+\d{1,2}:\d{2}/, // MM/DD/YY HH:MM
                /^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}:\d{2}/, // YYYY-MM-DD HH:MM:SS
                /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}:\d{2}/, // MM/DD/YYYY HH:MM:SS
                /^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}:\d{2}\.\d{3}/, // YYYY-MM-DD HH:MM:SS.mmm
                /^\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{2}:\d{2}/, // ISO格式 YYYY-MM-DDTHH:MM:SS
                /^\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{2}:\d{2}\.\d{3}/, // ISO格式 YYYY-MM-DDTHH:MM:SS.mmm
                /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M/, // MM/DD/YYYY HH:MM AM/PM
                /^\d{1,2}-\d{1,2}-\d{4}\s+\d{1,2}:\d{2}\s+[AP]M/ // DD-MM-YYYY HH:MM AM/PM
            ];
            
            if (/^\d{8}$/.test(strValue)) {
                const year = parseInt(strValue.substring(0, 4));
                const month = parseInt(strValue.substring(4, 6));
                const day = parseInt(strValue.substring(6, 8));
                if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return true;
                }
            }

            // 检查是否匹配任何日期格式
            for (let pattern of datePatterns) {
                if (pattern.test(strValue)) {
                    return true;
                }
            }
            
            // 检查是否为Excel日期序列号（数字或带小数）
            if (/^\d+(\.\d+)?$/.test(strValue)) {
                const numericValue = parseFloat(strValue);
                if (numericValue >= 10000 && numericValue <= 70000) {
                    return true;
                }
            }
            
            // 检查是否为日期对象可解析的格式
            if (!/^\d+$/.test(strValue)) {
                const testDate = new Date(strValue);
                if (!isNaN(testDate.getTime()) && strValue.length > 5) {
                    return true;
                }
            }
            
            return false;
        }

        function parseDate(dateStr, format) {
            try {
                let year, month, day, hour = 0, minute = 0, second = 0;
                
                dateStr = String(dateStr).trim();
                if (!dateStr) return null;
                
                // 清理日期字符串，移除多余的空格
                dateStr = dateStr.replace(/\s+/g, ' ').trim();
                
                // 处理带时间的日期格式
                let datePart = dateStr;
                let timePart = null;
                
                // 分离日期和时间部分
                const timeMatch = dateStr.match(/(\d{1,2}:\d{2}(?::\d{2})?(?:\.\d{3})?\s*(?:[AP]M)?)$/i);
                if (timeMatch) {
                    timePart = timeMatch[1];
                    datePart = dateStr.substring(0, timeMatch.index).trim();
                }
                
                // 解析时间部分
                if (timePart) {
                    const timeStr = timePart.trim();
                    let timeHour, timeMinute, timeSecond = 0;
                    
                    // 处理AM/PM格式
                    const ampmMatch = timeStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)/i);
                    if (ampmMatch) {
                        timeHour = parseInt(ampmMatch[1]);
                        timeMinute = parseInt(ampmMatch[2]);
                        const ampm = ampmMatch[3].toUpperCase();
                        
                        if (ampm === 'PM' && timeHour < 12) timeHour += 12;
                        if (ampm === 'AM' && timeHour === 12) timeHour = 0;
                    } else {
                        // 处理24小时格式
                        const timeParts = timeStr.split(':');
                        timeHour = parseInt(timeParts[0]) || 0;
                        timeMinute = parseInt(timeParts[1]) || 0;
                        if (timeParts[2]) {
                            const secondPart = timeParts[2].replace(/\.\d{3}$/, '');
                            timeSecond = parseInt(secondPart) || 0;
                        }
                    }
                    
                    hour = timeHour;
                    minute = timeMinute;
                    second = timeSecond;
                }
                
                // 根据指定格式解析日期
                if (format === 'YYYY-MM-DD') {
                    const parts = datePart.split('-');
                    if (parts.length >= 3) {
                        year = parseInt(parts[0]);
                        month = parseInt(parts[1]) - 1;
                        day = parseInt(parts[2]);
                    }
                } else if (format === 'YYYY/MM/DD') {
                    const parts = datePart.split('/');
                    if (parts.length >= 3) {
                        year = parseInt(parts[0]);
                        month = parseInt(parts[1]) - 1;
                        day = parseInt(parts[2]);
                    }
                } else if (format === 'MM/DD/YYYY') {
                    const parts = datePart.split('/');
                    if (parts.length >= 3) {
                        month = parseInt(parts[0]) - 1;
                        day = parseInt(parts[1]);
                        year = parseInt(parts[2]);
                    }
                } else if (format === 'DD/MM/YYYY') {
                    const parts = datePart.split('/');
                    if (parts.length >= 3) {
                        day = parseInt(parts[0]);
                        month = parseInt(parts[1]) - 1;
                        year = parseInt(parts[2]);
                    }
                }
                
                // 如果指定格式未成功解析，尝试自动检测格式
                if (!year || month === undefined || !day) {
                    // 处理YYYY-MM-DD格式
                    const ymdMatch = datePart.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
                    if (ymdMatch) {
                        year = parseInt(ymdMatch[1]);
                        month = parseInt(ymdMatch[2]) - 1;
                        day = parseInt(ymdMatch[3]);
                    } else {
                        // 处理YYYY/MM/DD格式
                        const ymdSlashMatch = datePart.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
                        if (ymdSlashMatch) {
                            year = parseInt(ymdSlashMatch[1]);
                            month = parseInt(ymdSlashMatch[2]) - 1;
                            day = parseInt(ymdSlashMatch[3]);
                        } else {
                            // 处理MM/DD/YYYY格式
                            const mdyMatch = datePart.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
                            if (mdyMatch) {
                                month = parseInt(mdyMatch[1]) - 1;
                                day = parseInt(mdyMatch[2]);
                                year = parseInt(mdyMatch[3]);
                            } else {
                                // 处理DD/MM/YYYY格式
                                const dmyMatch = datePart.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
                                if (dmyMatch) {
                                    day = parseInt(dmyMatch[1]);
                                    month = parseInt(dmyMatch[2]) - 1;
                                    year = parseInt(dmyMatch[3]);
                                } else {
                                    // 处理YYYYMMDD格式
                                    const yyyymmddMatch = datePart.match(/^(\d{4})(\d{2})(\d{2})$/);
                                    if (yyyymmddMatch) {
                                        year = parseInt(yyyymmddMatch[1]);
                                        month = parseInt(yyyymmddMatch[2]) - 1;
                                        day = parseInt(yyyymmddMatch[3]);
                                    } else {
                                        // 处理DDMMYYYY格式
                                        const ddmmyyyyMatch = datePart.match(/^(\d{2})(\d{2})(\d{4})$/);
                                        if (ddmmyyyyMatch) {
                                            day = parseInt(ddmmyyyyMatch[1]);
                                            month = parseInt(ddmmyyyyMatch[2]) - 1;
                                            year = parseInt(ddmmyyyyMatch[3]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                // 检查是否为有效日期
                if (year && month !== undefined && day) {
                    // 验证月份和日期范围
                    if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
                        const date = new Date(year, month, day, hour, minute, second);
                        
                        // 验证日期是否有效
                        if (date.getFullYear() === year && 
                            date.getMonth() === month && 
                            date.getDate() === day) {
                            return date;
                        }
                    }
                }
                
                // 尝试让Date对象自动解析（处理更多格式）
                try {
                    const autoDate = new Date(dateStr);
                    if (!isNaN(autoDate.getTime())) {
                        if (timePart) {
                            autoDate.setHours(hour, minute, second);
                        }
                        return autoDate;
                    }
                } catch (e) {
                    // 忽略解析错误
                }
                
                // 尝试处理Excel日期序列号（数字格式）
                if (/^\d+(\.\d+)?$/.test(dateStr) && !isNaN(parseFloat(dateStr))) {
                    const excelDate = parseFloat(dateStr);
                    // Excel日期序列号从1900-01-01开始，但需要考虑闰年bug
                    if (excelDate >= 1 && excelDate <= 2958465) { // Excel有效日期范围
                        const baseDate = new Date(1900, 0, 1);
                        const wholeDays = Math.floor(excelDate);
                        const dayFraction = excelDate - wholeDays;
                        baseDate.setDate(baseDate.getDate() + wholeDays - 1);
                        if (dayFraction) {
                            baseDate.setTime(baseDate.getTime() + Math.round(dayFraction * 86400000));
                        } else if (timePart) {
                            baseDate.setHours(hour, minute, second);
                        }
                        return baseDate;
                    }
                }
                
                return null;
            } catch (error) {
                console.error('Error parsing date:', error);
                return null;
            }
        }

        function formatDateForInput(date) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        // 格式化日期范围显示
        function formatDateRangeDisplay(dateString) {
            if (!dateString) return '';
            
            // 直接使用Date对象解析日期字符串
            const date = new Date(dateString);
            if (isNaN(date.getTime())) {
                // 如果直接解析失败，尝试使用parseDate函数
                const parsedDate = parseDate(dateString);
                if (!parsedDate) return dateString;
                return formatDateForInput(parsedDate);
            }
            
            // 使用formatDateForInput函数格式化日期
            return formatDateForInput(date);
        }

        function getColorForIndex(index) {
            // 优化的专业配色方案 - 基于Tailwind CSS色彩系统
            const colors = [
                '#3B82F6', '#EF4444', '#10B981', '#F59E0B', '#8B5CF6',
                '#06B6D4', '#F97316', '#84CC16', '#EC4899', '#6366F1',
                '#14B8A6', '#F43F5E', '#A855F7', '#0EA5E9', '#EAB308',
                '#22C55E', '#D946EF', '#F87171', '#34D399', '#60A5FA',
                '#FBBF24', '#A78BFA', '#FB923C', '#67E8F9', '#FDA4AF'
            ];
            
            return colors[index % colors.length];
        }

        function toRgba(color, alpha) {
            const a = Math.max(0, Math.min(1, Number(alpha)));
            const m = String(color).match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
            if (m) {
                return `rgba(${m[1]}, ${m[2]}, ${m[3]}, ${a})`;
            }
            return color;
        }

        function getLineDash(style) {
            switch (style) {
                case 'solid':
                    return [];
                case 'dashed':
                    return [5, 5];
                case 'dotted':
                    return [2, 2];
                default:
                    return [];
            }
        }

        function chartToSVG(chart) {
            // 将Chart.js图表转换为SVG
            const canvas = chart.canvas;
            const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            svg.setAttribute('width', canvas.width);
            svg.setAttribute('height', canvas.height);
            svg.setAttribute('viewBox', `0 0 ${canvas.width} ${canvas.height}`);
            
            const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
            image.setAttribute('width', canvas.width);
            image.setAttribute('height', canvas.height);
            image.setAttribute('href', canvas.toDataURL('image/png'));
            
            svg.appendChild(image);
            
            return new XMLSerializer().serializeToString(svg);
        }
        
        // 辅助函数：将间隔四舍五入到最接近的标准值
        function roundToNearestStandardInterval(interval) {
            const standardIntervals = [1, 5, 15, 30, 60];
            
            // 找到最接近的标准间隔
            let closestInterval = standardIntervals[0];
            let minDiff = Math.abs(interval - closestInterval);
            
            for (let i = 1; i < standardIntervals.length; i++) {
                const diff = Math.abs(interval - standardIntervals[i]);
                if (diff < minDiff) {
                    minDiff = diff;
                    closestInterval = standardIntervals[i];
                }
            }
            
            return closestInterval;
        }
        
        // 时间戳间隔分析函数
        function analyzeTimestampIntervals(timestamps, dateKey) {
            if (!timestamps || timestamps.length < 3) return null;
            
            // 排序时间戳
            const sortedTimestamps = [...timestamps].sort((a, b) => a - b);
            const intervals = [];
            const intervalCounts = {};
            
            // 计算所有相邻时间戳的间隔
            for (let i = 1; i < sortedTimestamps.length; i++) {
                const intervalMs = sortedTimestamps[i] - sortedTimestamps[i-1];
                const intervalMinutes = intervalMs / (1000 * 60);
                
                // 过滤合理的间隔（1分钟到2小时）
                if (intervalMinutes >= 1 && intervalMinutes <= 120) {
                    intervals.push(intervalMinutes);
                    
                    // 统计间隔频次（四舍五入到最近的整数）
                    const roundedInterval = Math.round(intervalMinutes);
                    intervalCounts[roundedInterval] = (intervalCounts[roundedInterval] || 0) + 1;
                }
            }
            
            if (intervals.length === 0) return null;
            
            // 找到最频繁的间隔
            let mostFrequentInterval = null;
            let maxCount = 0;
            for (const [interval, count] of Object.entries(intervalCounts)) {
                if (count > maxCount) {
                    maxCount = count;
                    mostFrequentInterval = parseInt(interval);
                }
            }
            
            // 计算统计信息
            const avgInterval = intervals.reduce((a, b) => a + b, 0) / intervals.length;
            const medianInterval = intervals.sort((a, b) => a - b)[Math.floor(intervals.length / 2)];
            const stdDev = Math.sqrt(intervals.reduce((sum, val) => sum + Math.pow(val - avgInterval, 2), 0) / intervals.length);
            
            // 计算一致性分数
            const consistencyScore = maxCount / intervals.length;
            const variabilityScore = 1 - Math.min(stdDev / avgInterval, 1);
            
            // 选择最佳间隔（优先考虑最频繁的，然后是中位数）
            let detectedInterval = mostFrequentInterval;
            if (consistencyScore < 0.6) {
                detectedInterval = Math.round(medianInterval);
            }
            
            const closestStandard = roundToNearestStandardInterval(detectedInterval);
            
            // 计算置信度
            let confidence = 0.5;
            if (consistencyScore >= 0.8 && variabilityScore >= 0.8) {
                confidence = 0.95;
            } else if (consistencyScore >= 0.6 && variabilityScore >= 0.6) {
                confidence = 0.85;
            } else if (consistencyScore >= 0.4 || variabilityScore >= 0.4) {
                confidence = 0.7;
            }
            
            // 如果检测到的间隔与标准间隔完全匹配，提高置信度
            if (detectedInterval === closestStandard) {
                confidence = Math.min(0.98, confidence + 0.1);
            }
            
            return {
                interval: closestStandard,
                description: `${closestStandard}分钟`,
                confidence: confidence,
                method: '时间戳间隔分析',
                details: `${dateKey}日期分析${intervals.length}个间隔，平均${avgInterval.toFixed(1)}分钟，最频繁${mostFrequentInterval}分钟（${maxCount}次），一致性${(consistencyScore*100).toFixed(1)}%`
            };
        }
        
        // 跨日期时间戳模式分析函数
        function analyzeCrossDateTimestampPattern(timeStampAnalysis) {
            const dateKeys = Object.keys(timeStampAnalysis);
            if (dateKeys.length < 2) return null;
            
            const allIntervalResults = [];
            const intervalFrequency = {};
            
            // 分析每个日期的时间戳模式
            for (const dateKey of dateKeys) {
                const timestamps = timeStampAnalysis[dateKey];
                if (timestamps.length < 3) continue;
                
                const result = analyzeTimestampIntervals(timestamps, dateKey);
                if (result && result.confidence >= 0.6) {
                    allIntervalResults.push(result);
                    intervalFrequency[result.interval] = (intervalFrequency[result.interval] || 0) + 1;
                }
            }
            
            if (allIntervalResults.length === 0) return null;
            
            // 找到最一致的间隔
            let mostConsistentInterval = null;
            let maxFrequency = 0;
            for (const [interval, frequency] of Object.entries(intervalFrequency)) {
                if (frequency > maxFrequency) {
                    maxFrequency = frequency;
                    mostConsistentInterval = parseInt(interval);
                }
            }
            
            // 计算跨日期一致性
            const consistencyRatio = maxFrequency / allIntervalResults.length;
            const avgConfidence = allIntervalResults
                .filter(r => r.interval === mostConsistentInterval)
                .reduce((sum, r) => sum + r.confidence, 0) / maxFrequency;
            
            // 计算最终置信度
            let finalConfidence = avgConfidence * consistencyRatio;
            if (consistencyRatio >= 0.8) {
                finalConfidence = Math.min(0.95, finalConfidence + 0.1);
            }
            
            return {
                interval: mostConsistentInterval,
                description: `${mostConsistentInterval}分钟`,
                confidence: finalConfidence,
                method: '跨日期模式分析',
                details: `分析${dateKeys.length}个日期，${maxFrequency}个日期检测到${mostConsistentInterval}分钟间隔，一致性${(consistencyRatio*100).toFixed(1)}%`
            };
        }
        
        // 增强的智能时间间隔检测函数 - 多维度综合分析
        function detectTimeIntervalIntelligently(data, startIndex, endIndex, dateCol, meteringPointCol, selectedMeteringPoint, isCrossDateStructure, dataColumnCount) {
            const standardIntervals = [
                { interval: 1, points: 1440, description: '1分钟', tolerance: 0.01, weight: 1.0 },
                { interval: 5, points: 288, description: '5分钟', tolerance: 0.02, weight: 1.0 },
                { interval: 15, points: 96, description: '15分钟', tolerance: 0.01, weight: 1.0 }, // 96行数据精确匹配15分钟
                { interval: 30, points: 48, description: '30分钟', tolerance: 0.02, weight: 0.8 },
                { interval: 60, points: 24, description: '60分钟', tolerance: 0.05, weight: 1.0 }
            ];
            
            let detectionResults = [];
            let analysisLog = [];
            
            // 数据预处理 - 采样优化：对于超大规模数据集，只采样前 1000 行
            const maxSampleRows = 1000;
            const actualEndIndex = Math.min(endIndex, startIndex + maxSampleRows);
            
            const processedData = [];
            const dateCellCounts = {};
            const timestamps = [];
            
            for (let i = startIndex; i <= Math.min(actualEndIndex, data.length - 1); i++) {
                const row = data[i];
                if (!row || !row[dateCol]) continue;
                
                // 检查计量点编号筛选
                if (meteringPointCol !== null && selectedMeteringPoint) {
                    const rowMeteringPoint = String(row[meteringPointCol]).trim();
                    if (rowMeteringPoint !== selectedMeteringPoint) continue;
                }
                
                // 解析日期
                const dateStr = String(row[dateCol]).trim();
                const date = parseDate(dateStr, appData.config.dateFormat);
                if (!date) continue;
                
                const dateKey = formatDateKey(date);
                const timestamp = date.getTime();
                
                processedData.push({
                    date: dateKey,
                    timestamp: timestamp,
                    rawData: row
                });
                
                dateCellCounts[dateKey] = (dateCellCounts[dateKey] || 0) + 1;
                timestamps.push(timestamp);
            }
            
            // 96行数据特殊处理 - 直接识别为15分钟间隔（适用于所有结构）
            const totalRows = processedData.length;
            const uniqueDates = Object.keys(dateCellCounts).length;
            const avgPerDay = totalRows / Math.max(1, uniqueDates);
            
            // 检查是否接近96行/天（15分钟间隔）
            if (Math.abs(avgPerDay - 96) <= 2) {
                return {
                    interval: 15,
                    description: '15分钟（96行/天精确匹配）',
                    confidence: 0.98,
                    method: '96行/天数据精确匹配',
                    message: `检测到平均${avgPerDay.toFixed(1)}行/天，精确识别为15分钟间隔`
                };
            }
            
            // 检查是否接近48行/天（30分钟间隔）
            if (Math.abs(avgPerDay - 48) <= 1) {
                return {
                    interval: 30,
                    description: '30分钟（48行/天精确匹配）',
                    confidence: 0.95,
                    method: '48行/天数据精确匹配',
                    message: `检测到平均${avgPerDay.toFixed(1)}行/天，识别为30分钟间隔`
                };
            }
            
            // 检查是否接近24行/天（60分钟间隔）
            if (Math.abs(avgPerDay - 24) <= 1) {
                return {
                    interval: 60,
                    description: '60分钟（24行/天精确匹配）',
                    confidence: 0.95,
                    method: '24行/天数据精确匹配',
                    message: `检测到平均${avgPerDay.toFixed(1)}行/天，识别为60分钟间隔`
                };
            }
            
            if (processedData.length === 0) {
                return {
                    interval: 15,
                    description: '15分钟（默认）',
                    confidence: 0.3,
                    method: '默认值',
                    message: '数据为空，使用默认15分钟间隔'
                };
            }
            
            // 方法1: 基于单元格数量的分析（跨日期结构）
            if (isCrossDateStructure) {
                const counts = Object.values(dateCellCounts);
                const uniqueDates = Object.keys(dateCellCounts).length;
                
                if (counts.length > 0) {
                    const stats = calculateEnhancedStatistics(counts);
                    
                    standardIntervals.forEach(standard => {
                        const scores = calculateEnhancedCellCountScores(stats, standard, uniqueDates);
                        if (scores.totalScore > 0.5) {
                            detectionResults.push({
                                interval: standard.interval,
                                description: standard.description,
                                confidence: scores.totalScore,
                                method: '单元格数量分析',
                                details: `平均${stats.mean.toFixed(1)}个/天，${uniqueDates}个日期`
                            });
                        }
                    });
                }
                
                // 方法2: 基于时间戳间隔分析
                if (timestamps.length >= 3) {
                    const sortedTimestamps = [...timestamps].sort((a, b) => a - b);
                    const intervals = [];
                    
                    for (let i = 1; i < sortedTimestamps.length; i++) {
                        const interval = (sortedTimestamps[i] - sortedTimestamps[i-1]) / (1000 * 60);
                        if (interval >= 0.5 && interval <= 120) {
                            intervals.push(interval);
                        }
                    }
                    
                    if (intervals.length > 0) {
                        const stats = calculateEnhancedStatistics(intervals);
                        const filtered = filterOutliers(intervals);
                        
                        standardIntervals.forEach(standard => {
                            const scores = calculateEnhancedTimestampScores(stats, filtered, standard);
                            if (scores.totalScore > 0.5) {
                                detectionResults.push({
                                    interval: standard.interval,
                                    description: standard.description,
                                    confidence: scores.totalScore,
                                    method: '时间戳间隔分析',
                                    details: `平均${stats.mean.toFixed(1)}分钟，变异系数${stats.cv.toFixed(3)}`
                                });
                            }
                        });
                    }
                }
                
            } else {
                // 非跨日期结构：基于数据列数量的检测
                if (dataColumnCount > 0) {
                    standardIntervals.forEach(standard => {
                        const tolerance = Math.max(1, Math.floor(standard.points * 0.05));
                        const diff = Math.abs(dataColumnCount - standard.points);
                        
                        if (diff <= tolerance) {
                            const score = 1.0 - (diff / standard.points);
                            detectionResults.push({
                                interval: standard.interval,
                                description: standard.description,
                                confidence: Math.max(0.6, score),
                                method: '数据列数量分析',
                                details: `检测到${dataColumnCount}个数据列`
                            });
                        }
                    });
                }
            }
            
            // 方法3: 基于统计一致性的分析
            const consistencyResult = analyzeStatisticalConsistency(processedData, standardIntervals);
            if (consistencyResult) {
                detectionResults.push(consistencyResult);
            }
            
            // 综合评估和选择
            return synthesizeEnhancedResults(detectionResults);
        }
        
        // 增强统计计算函数
        function calculateEnhancedStatistics(arr) {
            if (!arr || arr.length === 0) return null;
            
            const sorted = [...arr].sort((a, b) => a - b);
            const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
            const median = sorted.length % 2 === 0 ? 
                (sorted[Math.floor(sorted.length/2) - 1] + sorted[Math.floor(sorted.length/2)]) / 2 : 
                sorted[Math.floor(sorted.length/2)];
            
            const counts = {};
            arr.forEach(val => {
                const key = Math.round(val);
                counts[key] = (counts[key] || 0) + 1;
            });
            
            const mode = parseInt(Object.keys(counts).reduce((a, b) => counts[a] > counts[b] ? a : b));
            const stdDev = Math.sqrt(arr.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / arr.length);
            const cv = stdDev / mean;
            
            return { mean, median, mode, stdDev, cv, min: sorted[0], max: sorted[sorted.length - 1] };
        }
        
        // 计算单元格数量评分
        function calculateEnhancedCellCountScores(stats, standard, uniqueDates) {
            const tolerance = Math.max(1, Math.floor(standard.points * standard.tolerance));
            const exactScore = stats.mean === standard.points ? 1.0 : 0;
            const closeScore = Math.max(0, 1 - Math.abs(stats.mean - standard.points) / standard.points);
            const consistencyScore = Math.max(0, 1 - stats.cv);
            const dateFactor = Math.min(1, uniqueDates / 3); // 至少3天数据
            
            return {
                exact: exactScore,
                close: closeScore,
                consistency: consistencyScore,
                dateFactor: dateFactor,
                totalScore: (exactScore * 0.3 + closeScore * 0.4 + consistencyScore * 0.2 + dateFactor * 0.1)
            };
        }
        
        // 计算时间戳评分
        function calculateEnhancedTimestampScores(stats, filtered, standard) {
            const tolerance = standard.interval * standard.tolerance;
            const modeScore = Math.max(0, 1 - Math.abs(stats.mode - standard.interval) / standard.interval);
            const meanScore = Math.max(0, 1 - Math.abs(stats.mean - standard.interval) / standard.interval);
            const consistencyScore = Math.max(0, 1 - stats.cv);
            const outlierFactor = filtered.length / (filtered.length + 1); // 考虑异常值过滤
            
            return {
                mode: modeScore,
                mean: meanScore,
                consistency: consistencyScore,
                outlier: outlierFactor,
                totalScore: (modeScore * 0.4 + meanScore * 0.3 + consistencyScore * 0.2 + outlierFactor * 0.1)
            };
        }
        
        // 异常值过滤
        function filterOutliers(arr) {
            if (!arr || arr.length < 4) return arr;
            
            const sorted = [...arr].sort((a, b) => a - b);
            const q1 = sorted[Math.floor(sorted.length * 0.25)];
            const q3 = sorted[Math.floor(sorted.length * 0.75)];
            const iqr = q3 - q1;
            const lowerBound = q1 - 1.5 * iqr;
            const upperBound = q3 + 1.5 * iqr;
            
            return sorted.filter(val => val >= lowerBound && val <= upperBound);
        }
        
        // 统计一致性分析
        function analyzeStatisticalConsistency(data, standardIntervals) {
            if (!data || data.length < 10) return null;
            
            const timestamps = data.map(d => d.timestamp).sort((a, b) => a - b);
            const intervals = [];
            
            for (let i = 1; i < timestamps.length; i++) {
                const interval = (timestamps[i] - timestamps[i-1]) / (1000 * 60);
                intervals.push(interval);
            }
            
            if (intervals.length === 0) return null;
            
            const stats = calculateEnhancedStatistics(intervals);
            const standardIntervalsList = standardIntervals.map(s => s.interval);
            
            // 计算与标准间隔的匹配度
            let bestMatch = null;
            let bestScore = 0;
            
            standardIntervals.forEach(standard => {
                const matches = intervals.filter(i => Math.abs(i - standard.interval) <= standard.interval * 0.1);
                const matchRate = matches.length / intervals.length;
                
                if (matchRate > 0.6) { // 60%以上匹配
                    const score = matchRate * (1 - stats.cv);
                    if (score > bestScore) {
                        bestScore = score;
                        bestMatch = {
                            interval: standard.interval,
                            description: standard.description,
                            confidence: score,
                            method: '统计一致性分析',
                            details: `匹配率${(matchRate*100).toFixed(1)}%，一致性${((1-stats.cv)*100).toFixed(1)}%`
                        };
                    }
                }
            });
            
            return bestMatch;
        }
        
        // 综合结果评估
        function synthesizeEnhancedResults(results) {
            if (!results || results.length === 0) {
                return {
                    interval: 15,
                    description: '15分钟（默认）',
                    confidence: 0.3,
                    method: '默认值',
                    message: '未能检测到明确的时间间隔模式，使用默认15分钟间隔'
                };
            }
            
            // 按置信度排序
            results.sort((a, b) => b.confidence - a.confidence);
            
            // 检查一致性
            const intervals = results.map(r => r.interval);
            const uniqueIntervals = [...new Set(intervals)];
            
            let bestResult = results[0];
            
            // 如果多个方法指向同一间隔，提高置信度
            const supportingMethods = results.filter(r => r.interval === bestResult.interval);
            if (supportingMethods.length > 1) {
                const avgConfidence = supportingMethods.reduce((sum, r) => sum + r.confidence, 0) / supportingMethods.length;
                bestResult.confidence = Math.min(0.99, avgConfidence + 0.1);
                bestResult.message = `通过${supportingMethods.length}种方法一致检测到：${bestResult.description}`;
            } else {
                bestResult.message = `${bestResult.method}检测到：${bestResult.description}，${bestResult.details}`;
            }
            
            // 生成候选方案
            const candidates = results
                .filter(r => r.interval !== bestResult.interval && r.confidence >= 0.4)
                .slice(0, 3)
                .map(r => ({
                    interval: r.interval,
                    confidence: r.confidence,
                    method: r.method
                }));
            
            bestResult.candidates = candidates;
            
            return bestResult;
        }

        // 初始化日期总用电量柱状图
        function initDailyTotalBarChart() {
            console.log('=== 创建每日总用电量柱状图 ===');
            
            // 检查Chart.js是否加载
            if (typeof Chart === 'undefined') {
                console.error('Chart.js库未加载！等待500ms后重试...');
                setTimeout(initDailyTotalBarChart, 500);
                return;
            }
            
            const canvasElement = document.getElementById('dailyTotalBarChart');
            if (!canvasElement) {
                console.error('找不到dailyTotalBarChart canvas元素');
                return;
            }
            
            // 确保canvas元素可见且有尺寸（若为0则强制设置容器尺寸并继续）
            if (canvasElement.offsetWidth === 0 || canvasElement.offsetHeight === 0) {
                console.warn('Canvas元素尺寸为0，尝试强制设置容器尺寸并继续初始化...');
                const barChartContainer = document.getElementById('dailyTotalBarChartContainer');
                if (barChartContainer) {
                    barChartContainer.style.display = 'block';
                    if (!barChartContainer.style.height) {
                        barChartContainer.style.height = '500px';
                    }
                }
                if (!canvasElement.style.height) {
                    canvasElement.style.height = '400px';
                }
                // 不再直接return，继续创建图表，由Chart.js自适应
            }
            
            const ctx = canvasElement.getContext('2d');
            if (!ctx) {
                console.error('无法获取canvas 2D上下文');
                return;
            }
            
            // 如果已存在图表实例，先销毁它
            if (appData.dailyTotalBarChart) {
                console.log('销毁现有图表实例');
                try {
                    appData.dailyTotalBarChart.destroy();
                } catch (e) {
                    console.warn('销毁图表实例时出错:', e);
                }
                appData.dailyTotalBarChart = null;
            }
            
            // 创建渐变色
            const gradient = ctx.createLinearGradient(0, 0, 0, 400);
            gradient.addColorStop(0, 'rgba(59, 130, 246, 0.8)');
            gradient.addColorStop(1, 'rgba(147, 197, 253, 0.4)');
            
            // 创建柱状图
            appData.dailyTotalBarChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: [],
                    datasets: [{
                        label: '日总用电量',
                        data: [],
                        backgroundColor: gradient,
                        borderColor: 'rgba(59, 130, 246, 1)',
                        borderWidth: 2,
                        borderRadius: 8,
                        borderSkipped: false,
                        hoverBackgroundColor: 'rgba(59, 130, 246, 0.9)',
                        hoverBorderWidth: 3
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {
                        padding: {
                            top: 20,
                            right: 30,
                            bottom: 20,
                            left: 20
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: '每日总用电量统计',
                            font: {
                                size: 22,
                                family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                weight: '600'
                            },
                            padding: {
                                top: 10,
                                bottom: 30
                            },
                            color: '#1F2937'
                        },
                        legend: {
                            display: false
                        },
                        tooltip: {
                            backgroundColor: 'rgba(255, 255, 255, 0.98)',
                            titleColor: '#1F2937',
                            titleFont: {
                                size: 16,
                                weight: '600'
                            },
                            bodyColor: '#4B5563',
                            bodyFont: {
                                size: 14,
                                family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif'
                            },
                            borderColor: 'rgba(0, 0, 0, 0.15)',
                            borderWidth: 1,
                            padding: 15,
                            cornerRadius: 10,
                            displayColors: false,
                            callbacks: {
                                title: function(tooltipItems) {
                                    return '日期: ' + tooltipItems[0].label;
                                },
                                label: function(context) {
                                    const value = context.parsed.y;
                                    return `总用电量: ${value.toFixed(2)} kWh`;
                                },
                                afterLabel: function(context) {
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = total > 0 ? ((context.parsed.y / total) * 100).toFixed(1) : 0;
                                    return `占比: ${percentage}%`;
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: '日期',
                                font: {
                                    size: 16,
                                    weight: '500'
                                }
                            },
                            ticks: {
                                maxRotation: 45,
                                minRotation: 45,
                                font: {
                                    size: 13
                                }
                            },
                            grid: {
                                display: false
                            }
                        },
                        y: {
                            title: {
                                display: true,
                                text: '总用电量 (kWh)',
                                font: {
                                    size: 16,
                                    weight: '500'
                                }
                            },
                            beginAtZero: true,
                            ticks: {
                                font: {
                                    size: 13
                                },
                                callback: function(value) {
                                    return value.toFixed(0);
                                }
                            },
                            grid: {
                                color: 'rgba(0, 0, 0, 0.05)',
                                drawBorder: false
                            }
                        }
                    },
                    interaction: {
                        mode: 'index',
                        intersect: false
                    },
                    animation: {
                        duration: 1200,
                        easing: 'easeOutQuart'
                    }
                }
            });
            
            console.log('每日总用电量柱状图创建完成');
        } // 结束 initEventListeners 函数
        
        // 更新日期总用电量柱状图（支持按日/按月聚合）
        function updateDailyTotalBarChart(filteredData) {
            console.log('=== 更新每日/每月总用电量柱状图 ===', {
                接收数据量: filteredData ? filteredData.length : 0,
                数据样本: filteredData ? filteredData.slice(0, 2) : '无数据',
                柱状图实例: !!appData.dailyTotalBarChart,
                聚合方式: appData?.visualization?.barChartAggregation || 'daily'
            });

            const targetData = getOverviewData();
            
            // 使用 targetData 替代 filteredData
            const dataToUse = targetData;
            
            console.log('=== 柱状图数据源重置 ===', {
                原始filteredData长度: filteredData ? filteredData.length : 0,
                重置后targetData长度: targetData.length,
                说明: '柱状图默认展示全部天数（仅受计量点筛选影响）'
            });
            
            if (!appData.dailyTotalBarChart) {
                console.error('柱状图实例不存在，尝试重新初始化...');
                initDailyTotalBarChart();
                
                // 延迟重试，最多重试3次
                let retryCount = 0;
                const retryInit = () => {
                    retryCount++;
                    if (retryCount <= 3 && !appData.dailyTotalBarChart) {
                        console.log(`重试初始化柱状图，第${retryCount}次...`);
                        initDailyTotalBarChart();
                        setTimeout(retryInit, 200);
                    } else if (appData.dailyTotalBarChart) {
                        updateDailyTotalBarChart(dataToUse);
                    } else {
                        console.error('柱状图初始化失败，请检查Chart.js和canvas元素');
                    }
                };
                setTimeout(retryInit, 100);
                return;
            }
            
            // 显示加载状态
            const loadingEl = document.getElementById('barChartLoading');
            const noDataEl = document.getElementById('barChartNoDataMessage');
            
            if (loadingEl && loadingEl.parentNode) loadingEl.classList.remove('hidden');
            if (noDataEl && noDataEl.parentNode) noDataEl.classList.add('hidden');
            
            try {
                // 数据验证和过滤
                if (!dataToUse || !Array.isArray(dataToUse)) {
                    console.warn('数据格式无效:', typeof dataToUse);
                    throw new Error('数据格式不正确');
                }
                
                if (dataToUse.length === 0) {
                    console.warn('无数据可显示');
                    appData.dailyTotalBarChart.data.labels = [];
                    appData.dailyTotalBarChart.data.datasets[0].data = [];
                    appData.dailyTotalBarChart.update('none');
                    if (noDataEl && noDataEl.parentNode) noDataEl.classList.remove('hidden');
                    return;
                }
                
                // 验证并清理数据
                const validData = dataToUse.filter(row => 
                    row && 
                    typeof row === 'object' && 
                    row.date && 
                    typeof row.dailyTotal === 'number' && 
                    !isNaN(row.dailyTotal) && 
                    row.dailyTotal >= 0
                );
                
                console.log('数据验证结果:', {
                    原始数据量: dataToUse.length,
                    有效数据量: validData.length,
                    样本数据: validData.slice(0, 3)
                });
                
                if (validData.length === 0) {
                    console.warn('无有效数据可显示');
                    appData.dailyTotalBarChart.data.labels = [];
                    appData.dailyTotalBarChart.data.datasets[0].data = [];
                    appData.dailyTotalBarChart.update('none');
                    if (noDataEl && noDataEl.parentNode) noDataEl.classList.remove('hidden');
                    return;
                }
                
                // 读取聚合方式：daily 或 monthly
                const aggregation = (appData?.visualization?.barChartAggregation === 'monthly') ? 'monthly' : 'daily';
                
                // 按聚合方式分组并累加总用电量
                const groupMap = new Map();
                validData.forEach(row => {
                    if (!row.date || typeof row.dailyTotal !== 'number') return;
                    let dateStr = String(row.date);
                    // 标准化为 YYYY-MM-DD
                    if (dateStr.length === 8 && !dateStr.includes('-')) {
                        dateStr = `${dateStr.substring(0,4)}-${dateStr.substring(4,6)}-${dateStr.substring(6,8)}`;
                    }
                    let key;
                    if (aggregation === 'monthly') {
                        // 提取 YYYY-MM 作为月份键
                        if (dateStr.length >= 7 && dateStr.includes('-')) {
                            key = dateStr.slice(0, 7);
                        } else {
                            // 回退：尝试用 Date 解析并格式化
                            const d = new Date(dateStr);
                            if (!isNaN(d.getTime())) {
                                const y = d.getFullYear();
                                const m = String(d.getMonth() + 1).padStart(2, '0');
                                key = `${y}-${m}`;
                            } else {
                                // 无法解析则原样使用（可能会导致排序以字符串比较）
                                key = dateStr;
                            }
                        }
                    } else {
                        // 每日
                        key = dateStr;
                    }
                    const cur = groupMap.get(key) || 0;
                    groupMap.set(key, cur + row.dailyTotal);
                });
                
                // 排序（尽量按时间顺序）
                const sortedEntries = Array.from(groupMap.entries()).sort((a, b) => {
                    try {
                        const keyA = a[0];
                        const keyB = b[0];
                        // 对于月份，补全为该月第一天以便比较
                        const dateA = new Date(aggregation === 'monthly' ? `${keyA}-01` : keyA);
                        const dateB = new Date(aggregation === 'monthly' ? `${keyB}-01` : keyB);
                        if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
                            return keyA.localeCompare(keyB);
                        }
                        return dateA - dateB;
                    } catch (e) {
                        return a[0].localeCompare(b[0]);
                    }
                });
                
                const labels = sortedEntries.map(([k]) => k);
                const data = sortedEntries.map(([, v]) => v);
                
                console.log('图表数据准备完成:', {
                    维度: aggregation === 'monthly' ? '月份' : '日期',
                    维度数量: labels.length,
                    范围: labels.length > 0 ? `${labels[0]} 至 ${labels[labels.length-1]}` : '无',
                    数据样本: data.slice(0, 5),
                    数据总和: data.reduce((a, b) => a + b, 0)
                });
                
                if (labels.length === 0) {
                    console.warn('无有效数据');
                    if (noDataEl && noDataEl.parentNode) noDataEl.classList.remove('hidden');
                    return;
                }
                
                // 更新图表标题/坐标轴/提示框
                const rangeText = labels.length > 0 ? `${labels[0]} 至 ${labels[labels.length - 1]}` : '无';
                const titlePrefix = aggregation === 'monthly' ? '每月总用电量' : '每日总用电量';
                appData.dailyTotalBarChart.options.plugins.title.text = `${titlePrefix} (${rangeText})`;
                // X轴标题
                if (appData.dailyTotalBarChart.options.scales && appData.dailyTotalBarChart.options.scales.x && appData.dailyTotalBarChart.options.scales.x.title) {
                    appData.dailyTotalBarChart.options.scales.x.title.text = aggregation === 'monthly' ? '月份' : '日期';
                }
                // Tooltip 标题
                if (appData.dailyTotalBarChart.options.plugins && appData.dailyTotalBarChart.options.plugins.tooltip && appData.dailyTotalBarChart.options.plugins.tooltip.callbacks) {
                    appData.dailyTotalBarChart.options.plugins.tooltip.callbacks.title = function(tooltipItems) {
                        const dim = aggregation === 'monthly' ? '月份' : '日期';
                        return `${dim}: ${tooltipItems[0].label}`;
                    };
                }
                
                // 更新图表数据
                appData.dailyTotalBarChart.data.labels = labels;
                appData.dailyTotalBarChart.data.datasets[0].data = data;
                
                // 使用动画更新
                appData.dailyTotalBarChart.update('normal');
                
                // 触发图表入场动画
                const barCanvas = appData.dailyTotalBarChart.canvas;
                if (barCanvas) {
                    barCanvas.classList.remove('animate-graph-entry');
                    void barCanvas.offsetWidth; // 触发重绘
                    barCanvas.classList.add('animate-graph-entry');
                }
                
            } catch (error) {
                console.error('更新柱状图时出错:', error);
                if (noDataEl && noDataEl.parentNode) noDataEl.classList.remove('hidden');
            } finally {
                if (loadingEl && loadingEl.parentNode) loadingEl.classList.add('hidden');
            }
        }

        // 初始化月份总用电量柱状图





        // 初始化应用
        function initApp() {
            // 强制滚动到顶部，修复初始加载时位置偏移问题
            window.scrollTo(0, 0);
            
            initEventListeners();
            setImportStatus('idle');
            
            // 初始化导出按钮状态为禁用

            
            // 初始化步骤导航
            initializeStepNavigation();
            
            // 初始化日期总用电量柱状图
            initDailyTotalBarChart();
            
            // 初始化24小时负荷曲线图表（即使没有数据也显示XY轴）
            initEmptyChart();
            
            // 初始化section显示状态，避免页面闪烁
            initializeSectionVisibility();
        }
        
        // 初始化section显示状态
        function initializeSectionVisibility() {
            const allSections = ['importSection', 'configSection', 'visualizationSection', 'exportSection'];
            
            // 默认显示导入section，其他隐藏
            allSections.forEach(id => {
                const section = document.getElementById(id);
                if (section) {
                    section.classList.add('section-container');
                    if (id === 'importSection') {
                        section.classList.add('section-visible');
                        section.classList.remove('section-hidden');
                    } else {
                        section.classList.add('section-hidden');
                        section.classList.remove('section-visible');
                    }
                }
            });
        }
        
        // 初始化空图表，显示XY轴
        function initEmptyChart() {
            if (appData.chart) {
                appData.chart.destroy();
            }
            
            const ctx = elements.loadCurveChart.getContext('2d');
            appData.chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: Array.from({ length: 24 }, (_, i) => `${i}:00`),
                    datasets: [{
                        label: '示例数据',
                        data: [],
                        borderColor: 'rgba(59, 130, 246, 0.3)',
                        backgroundColor: 'rgba(59, 130, 246, 0.1)',
                        borderWidth: 1,
                        pointRadius: 0,
                        fill: false
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {
                        padding: {
                            top: 10,
                            right: 15,
                            bottom: 10,
                            left: 15
                        }
                    },
                    plugins: {
                        title: {
                            display: true,
                            text: '24小时负荷曲线',
                            font: {
                                size: 18,
                                family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                weight: '600'
                            },
                            padding: {
                                top: 10,
                                bottom: 15
                            },
                            color: '#1F2937'
                        },
                        legend: {
                            display: false
                        },
                        tooltip: {
                            enabled: true
                        }
                    },
                    scales: {
                        x: {
                            display: true,
                            title: {
                                display: true,
                                text: '时间',
                                font: {
                                    size: 14,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '600'
                                },
                                padding: {
                                    top: 10,
                                    bottom: 10
                                }
                            },
                            grid: {
                                color: 'rgba(0, 0, 0, 0.05)',
                                lineWidth: 1,
                                drawBorder: false
                            },
                            ticks: {
                                font: {
                                    size: 12,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '500'
                                },
                                padding: 8,
                                maxRotation: 0,
                                autoSkip: false
                            }
                        },
                        y: {
                            display: true,
                            title: {
                                display: true,
                                text: '用电量 (kWh)',
                                font: {
                                    size: 14,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '600'
                                },
                                padding: {
                                    top: 10,
                                    bottom: 10
                                }
                            },
                            beginAtZero: true,
                            grid: {
                                color: 'rgba(0, 0, 0, 0.05)',
                                lineWidth: 1,
                                drawBorder: false
                            },
                            ticks: {
                                font: {
                                    size: 12,
                                    family: '"Fira Sans", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
                                    weight: '500'
                                },
                                padding: 8,
                                callback: function(value) {
                                    return value.toFixed(1);
                                }
                            }
                        }
                    }
                }
            });
            
            // 隐藏加载状态
            const chartLoadingElement = document.getElementById('chartLoading');
            if (chartLoadingElement && chartLoadingElement.parentNode) {
                chartLoadingElement.classList.add('hidden');
            }
        }

        // 初始化步骤导航
        function initializeStepNavigation() {
            // 不默认激活任何步骤，避免初始化时的动画效果
            // updateStepNavigation('importSection');
            
            // 延迟添加滚动监听，避免初始化时的滚动触发
            setTimeout(() => {
                // 监听滚动事件，根据当前视图区域更新步骤状态
                window.addEventListener('scroll', debounce(() => {
                    updateActiveStepBasedOnScroll();
                }, 100));
                
                // 页面初始化完成后，允许导航状态更新
                isPageInitializing = false;
                console.log('页面初始化完成，导航状态更新已启用');
            }, 1000);
            
            // 添加调试日志，帮助排查导航问题
            console.log('步骤导航初始化完成，无默认激活步骤');
        }

        // 根据滚动位置更新激活步骤
        function updateActiveStepBasedOnScroll() {
            // 如果页面正在初始化，不更新导航状态
            if (isPageInitializing) {
                return;
            }
            
            const sections = [
                { id: 'importSection', threshold: 0.5 },
                { id: 'configSection', threshold: 0.5 },
                { id: 'visualizationSection', threshold: 0.5 },
                { id: 'exportSection', threshold: 0.5 }
            ];
            
            const scrollPosition = window.scrollY + window.innerHeight / 2;
            
            for (let section of sections) {
                const element = document.getElementById(section.id);
                if (element) {
                    const rect = element.getBoundingClientRect();
                    const elementTop = rect.top + window.scrollY;
                    const elementBottom = elementTop + rect.height;
                    
                    if (scrollPosition >= elementTop && scrollPosition <= elementBottom) {
                        updateStepNavigation(section.id);
                        break;
                    }
                }
            }
        }

        // 防抖函数
        function debounce(func, wait) {
            let timeout;
            return function executedFunction(...args) {
                const later = () => {
                    clearTimeout(timeout);
                    func(...args);
                };
                clearTimeout(timeout);
                timeout = setTimeout(later, wait);
            };
        }

        function openAdvancedFiltersModal() {
            if (appData.worksheets.length === 0) {
                showNotification('提示', '请先导入数据后再设置高级过滤规则', 'info');
                return;
            }

            const existingFilters = Array.isArray(appData.config.additionalFilters) ? appData.config.additionalFilters.map(f => ({
                column: f && f.column ? String(f.column) : '',
                value: f && f.value ? String(f.value) : ''
            })) : [];

            const initialCount = Math.max(1, existingFilters.length || 1);
            const maxCount = 10;
            const optionsHtml = Array.from({ length: maxCount }, (_, i) => {
                const v = i + 1;
                return `<option value="${v}" ${v === initialCount ? 'selected' : ''}>${v}</option>`;
            }).join('');

            const content = `
                <div class="space-y-4">
                    <div class="flex items-center justify-between">
                        <div class="text-sm font-semibold text-slate-700">规则数量</div>
                        <select id="advancedFilterCount" class="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm font-bold text-slate-800 outline-none transition-all focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10">
                            ${optionsHtml}
                        </select>
                    </div>
                    <div id="advancedFilterOptions" class="space-y-3"></div>
                    <div class="flex justify-end gap-3 pt-2">
                        <button id="advancedFiltersCancel" class="btn-secondary">取消</button>
                        <button id="advancedFiltersApply" class="btn-rose">应用</button>
                    </div>
                </div>
            `;

            showModal('更多过滤规则', content);

            const countSelect = document.getElementById('advancedFilterCount');
            const optionsContainer = document.getElementById('advancedFilterOptions');
            const cancelBtn = document.getElementById('advancedFiltersCancel');
            const applyBtn = document.getElementById('advancedFiltersApply');

            if (!countSelect || !optionsContainer || !cancelBtn || !applyBtn) return;

            let tempFilters = Array.from({ length: initialCount }, (_, i) => existingFilters[i] || ({ column: '', value: '' }));

            const headerRow = appData.worksheets[0].data[0] || [];

            const buildColumnOptions = (selected) => {
                const opts = ['<option value="">选择列</option>'];
                for (let i = 0; i < headerRow.length; i++) {
                    let columnLetter = "";
                    let columnNum = i;
                    while (columnNum >= 0) {
                        columnLetter = String.fromCharCode(65 + (columnNum % 26)) + columnLetter;
                        columnNum = Math.floor(columnNum / 26) - 1;
                        if (columnNum < 0) break;
                    }
                    const headerText = headerRow[i] ? String(headerRow[i]).substring(0, 20) : `列 ${columnLetter}`;
                    const value = String(i);
                    opts.push(`<option value="${value}" ${selected === value ? 'selected' : ''}>${columnLetter}: ${headerText}</option>`);
                }
                return opts.join('');
            };

            const fillValueOptions = (selectEl, columnIndex, selectedValue) => {
                selectEl.innerHTML = '<option value="">全部</option>';
                if (!columnIndex) return;
                const col = parseInt(columnIndex);
                if (Number.isNaN(col)) return;

                const valueSet = new Set();
                appData.worksheets.forEach(worksheet => {
                    const { data } = worksheet;
                    for (let i = 1; i < data.length; i++) {
                        const row = data[i];
                        if (!row || row[col] === undefined || row[col] === null) continue;
                        const value = String(row[col]).trim();
                        if (value) valueSet.add(value);
                    }
                });

                Array.from(valueSet).sort().forEach(v => {
                    const option = document.createElement('option');
                    option.value = v;
                    option.textContent = v;
                    selectEl.appendChild(option);
                });

                if (selectedValue) {
                    selectEl.value = selectedValue;
                }
            };

            const render = () => {
                optionsContainer.innerHTML = tempFilters.map((f, i) => {
                    return `
                        <div class="rounded-2xl border border-slate-200 bg-slate-50/70 p-4">
                            <div class="flex items-center justify-between">
                                <div class="text-xs font-extrabold uppercase tracking-wider text-slate-500">规则 ${i + 1}</div>
                                <button type="button" data-index="${i}" class="advanced-filter-clear text-xs font-bold text-indigo-600 transition-all hover:text-indigo-700 hover:underline">清除</button>
                            </div>
                            <div class="mt-3 grid grid-cols-1 gap-3 md:grid-cols-2">
                                <div>
                                    <label class="mb-1 block text-xs font-semibold text-slate-600">列</label>
                                    <select id="advancedFilterColumn${i}" class="w-full appearance-none rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-bold text-slate-800 outline-none transition-all focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10">
                                        ${buildColumnOptions(f.column || '')}
                                    </select>
                                </div>
                                <div>
                                    <label class="mb-1 block text-xs font-semibold text-slate-600">值</label>
                                    <select id="advancedFilterValue${i}" class="w-full appearance-none rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-bold text-slate-800 outline-none transition-all focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10">
                                        <option value="">全部</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('');

                tempFilters.forEach((f, i) => {
                    const colSelect = document.getElementById(`advancedFilterColumn${i}`);
                    const valSelect = document.getElementById(`advancedFilterValue${i}`);
                    if (!colSelect || !valSelect) return;

                    fillValueOptions(valSelect, f.column, f.value);

                    colSelect.addEventListener('change', (e) => {
                        tempFilters[i].column = e.target.value;
                        tempFilters[i].value = '';
                        fillValueOptions(valSelect, e.target.value, '');
                    });

                    valSelect.addEventListener('change', (e) => {
                        tempFilters[i].value = e.target.value;
                    });
                });
            };

            optionsContainer.addEventListener('click', (e) => {
                const btn = e.target.closest('.advanced-filter-clear');
                if (!btn) return;
                const index = parseInt(btn.dataset.index);
                if (Number.isNaN(index) || !tempFilters[index]) return;
                tempFilters[index].column = '';
                tempFilters[index].value = '';
                render();
            });

            countSelect.addEventListener('change', (e) => {
                const newCount = Math.max(1, Math.min(maxCount, parseInt(e.target.value) || 1));
                if (newCount > tempFilters.length) {
                    while (tempFilters.length < newCount) tempFilters.push({ column: '', value: '' });
                } else if (newCount < tempFilters.length) {
                    tempFilters = tempFilters.slice(0, newCount);
                }
                render();
            });

            cancelBtn.addEventListener('click', closeModal);
            applyBtn.addEventListener('click', () => {
                const normalized = tempFilters.map(f => ({
                    column: f.column ? String(f.column) : '',
                    value: f.value ? String(f.value) : ''
                }));

                const active = normalized.filter(f => f.column && f.value);
                const inactive = normalized.filter(f => !(f.column && f.value));
                appData.config.additionalFilters = [...active, ...inactive];

                updateDataPreview();
                setTimeout(() => {
                    if (appData.processedData.length > 0) {
                        updateChart();
                        updateDailyTotalBarChart(getFilteredData());
                    }
                }, 0);

                closeModal();
                showNotification('已应用', `已设置 ${active.length} 条过滤规则`, 'success');
            });

            render();
        }

        // 预览第一个小时计算过程
        function previewFirstHourCalculation() {
            if (appData.parsedData.length === 0) {
                showNotification('警告', '没有可预览的数据，请先导入并配置数据', 'warning');
                return;
            }
            
            // 获取第一行解析后的数据
            const firstParsedData = appData.parsedData[0];
            const interval = firstParsedData.timeInterval || appData.config.timeInterval;
            
            // 1. 获取第一个小时的原始数据 (前60分钟)
            let pointsPerHourRaw = 60; // 默认为1分钟间隔的60个点
            if (interval === 5) pointsPerHourRaw = 12;
            if (interval === 15) pointsPerHourRaw = 4;
            if (interval === 30) pointsPerHourRaw = 2;
            if (interval === 60) pointsPerHourRaw = 1;
            
            // 如果是1分钟间隔，我们要看前60个点。如果是15分钟间隔，我们要看前4个点。
            // 为了通用性，我们先获取足够的点来分析
            const checkPointsCount = interval === 1 ? 60 : pointsPerHourRaw;
            const hourData = firstParsedData.rawData.slice(0, checkPointsCount);
            
            // 2. 数据清洗与有效性检测
            let validPointsCount = 0;
            let positiveSum = 0;
            let negativeSum = 0;
            let negativeCount = 0;
            let processedValues = [];
            
            hourData.forEach(val => {
                if (val !== null && val !== undefined && val !== '') {
                    validPointsCount++;
                    const numVal = parseFloat(val);
                    if (!isNaN(numVal)) {
                        if (numVal < 0) {
                            // 负数处理：记入发电，用电视为0
                            negativeSum += Math.abs(numVal);
                            negativeCount++;
                            processedValues.push({ original: numVal, forConsumption: 0, forGeneration: Math.abs(numVal), note: '负值(发电)' });
                        } else {
                            // 正数处理：记入用电
                            positiveSum += numVal;
                            processedValues.push({ original: numVal, forConsumption: numVal, forGeneration: 0, note: '正值(用电)' });
                        }
                    } else {
                        processedValues.push({ original: val, forConsumption: 0, forGeneration: 0, note: '无效数值' });
                    }
                } else {
                    processedValues.push({ original: null, forConsumption: 0, forGeneration: 0, note: '空值' });
                }
            });
            
            // 3. 确定计算逻辑
            // 逻辑复刻 parseWorksheetData 中的核心逻辑
            let calculationProcess = '';
            let formula = '';
            let hourlyConsumption = 0;
            
            // 判定有效数据点密度
            // 假设我们期望的是每小时4个点（15分钟间隔）
            // 如果有效点数接近4个（比如3-5个），或者虽然是1分钟间隔但每15分钟才有一个数据
            
            let detectedInterval = interval;
            let energyFactor = 1;
            
            // 简单的密度检测逻辑展示
            // 如果实际有效点数 <= 4，且原间隔是1分钟，则可能实际上是15分钟或60分钟数据
            let effectiveDivisor = 1;
            
            if (interval === 1 && validPointsCount > 0 && validPointsCount <= 4) {
                // 1分钟间隔文件，但每小时只有稀疏几个点，判定为15分钟或60分钟数据
                 calculationProcess += `检测到每小时有效数据点为 ${validPointsCount} 个 (稀疏数据)。\n`;
                 // 这里简化逻辑，只展示最终的计算结果，假设 parseWorksheetData 已经正确处理了倍率
                 // 在预览中，我们演示 "正值之和 / 4" (如果是15分钟) 或 "正值之和 / 1"
                 
                 // 按照用户最新要求：如果判断为15分钟间隔，则除以4
                 effectiveDivisor = 4; // 默认假设15分钟
                 if (validPointsCount === 1) effectiveDivisor = 1; // 只有一个点，可能是整点数据
            } else if (interval === 15) {
                 effectiveDivisor = 4;
            } else if (interval === 60) {
                 effectiveDivisor = 1;
            } else {
                // 其他情况，按点数归一化
                // 比如 1分钟间隔，60个点，sum / 60 * 1 (如果单位是kW) 
                // 但用户现在的逻辑是 "正值直接/4" (针对15分钟间隔)
                // 这里我们展示通用的 "加权平均" 或 "累加后调整"
                if (validPointsCount > 0) {
                     // 这里的逻辑需要与 parseWorksheetData 保持一致。
                     // 假设 parseWorksheetData 中：
                     // if (pointsPerHour === 4) factor = 0.25;
                     // if (pointsPerHour === 1) factor = 1;
                     // if (pointsPerHour === 60) factor = 1/60; 
                     // 但用户特别指出 "正值除以4"
                     effectiveDivisor = 4; 
                     if (interval === 60) effectiveDivisor = 1;
                     if (interval === 1) effectiveDivisor = 60; // 瞬时功率平均
                }
            }

            // 根据用户最新强规则：
            // "一旦确定了实际的数据记录间隔（例如15分钟）... 正值数据除以4"
            // 我们在预览中模拟这个判定
            let detectedGranularity = "未知";
            if (interval === 1) {
                if (validPointsCount > 45) {
                    detectedGranularity = "1分钟 (连续)";
                    energyFactor = 1/60; // 功率 -> 电量
                } else if (validPointsCount <= 5 && validPointsCount > 1) {
                    detectedGranularity = "15分钟 (稀疏)";
                    energyFactor = 0.25; // /4
                } else if (validPointsCount === 1) {
                    detectedGranularity = "60分钟 (整点)";
                    energyFactor = 1;
                } else {
                     detectedGranularity = `1分钟 (密度: ${validPointsCount}/60)`;
                     energyFactor = 1/60; 
                }
            } else if (interval === 15) {
                detectedGranularity = "15分钟";
                energyFactor = 0.25;
            } else if (interval === 60) {
                detectedGranularity = "60分钟";
                energyFactor = 1;
            }

            // 计算
            hourlyConsumption = positiveSum * energyFactor * firstParsedData.multiplier;
            
            // 构建计算过程描述
            calculationProcess = `<div class="space-y-2 text-sm">`;
            calculationProcess += `<div class="font-bold text-gray-700">1. 数据分析</div>`;
            calculationProcess += `<ul class="list-disc pl-5 text-gray-600">`;
            calculationProcess += `<li>原始间隔: ${interval} 分钟</li>`;
            calculationProcess += `<li>检测样本: 前 ${checkPointsCount} 个时间点</li>`;
            calculationProcess += `<li>有效数值: ${validPointsCount} 个</li>`;
            calculationProcess += `<li><span class="text-blue-600 font-semibold">判定颗粒度: ${detectedGranularity}</span></li>`;
            calculationProcess += `</ul>`;

            calculationProcess += `<div class="font-bold text-gray-700 mt-2">2. 数值分离</div>`;
            calculationProcess += `<ul class="list-disc pl-5 text-gray-600">`;
            calculationProcess += `<li>正值总和 (用电): <span class="font-mono text-blue-600">${positiveSum.toFixed(2)}</span></li>`;
            calculationProcess += `<li>负值总和 (发电): <span class="font-mono text-green-600">${negativeSum.toFixed(2)}</span> (已排除)</li>`;
            calculationProcess += `</ul>`;

            calculationProcess += `<div class="font-bold text-gray-700 mt-2">3. 计算公式</div>`;
            calculationProcess += `<div class="bg-gray-50 p-2 rounded border border-gray-200 font-mono text-xs">`;
            calculationProcess += `用电量 = 正值总和 × 时间系数 × 倍率<br>`;
            calculationProcess += `= ${positiveSum.toFixed(2)} × ${energyFactor} × ${firstParsedData.multiplier}<br>`;
            calculationProcess += `= <span class="text-blue-600 font-bold text-base">${hourlyConsumption.toFixed(4)} kWh</span>`;
            calculationProcess += `</div>`;
            calculationProcess += `</div>`;

            formula = `Σ(正值) × ${energyFactor} × 倍率`;
            
            // 更新模态框内容
            if (elements.previewDate && elements.previewDate.parentNode) {
                elements.previewDate.textContent = firstParsedData.dateStr;
            }
            if (elements.previewInterval && elements.previewInterval.parentNode) {
                elements.previewInterval.textContent = `${interval} 分钟`;
            }
            if (elements.previewDataType && elements.previewDataType.parentNode) {
                elements.previewDataType.textContent = detectedGranularity;
            }
            if (elements.previewMultiplier && elements.previewMultiplier.parentNode) {
                elements.previewMultiplier.textContent = firstParsedData.multiplier;
            }
            
            // 显示原始数据详情
            const rawDataHtml = processedValues.map((item, idx) => {
                let colorClass = 'text-gray-500';
                if (item.note.includes('用电')) colorClass = 'text-blue-600 font-semibold';
                if (item.note.includes('发电')) colorClass = 'text-green-600 font-semibold';
                
                return `<div class="flex justify-between text-xs border-b border-gray-100 py-1">
                    <span class="w-12 text-gray-400">#${idx + 1}</span>
                    <span class="w-20 font-mono">${item.original !== null ? item.original : '-'}</span>
                    <span class="${colorClass}">${item.note}</span>
                </div>`;
            }).join('');
            
            if (elements.previewRawData && elements.previewRawData.parentNode) {
                // 如果是textarea，只能显示文本；如果是div，可以显示HTML
                if (elements.previewRawData.tagName === 'DIV' || elements.previewRawData.tagName === 'SPAN') {
                    elements.previewRawData.innerHTML = `<div class="max-h-40 overflow-y-auto">${rawDataHtml}</div>`;
                } else {
                    elements.previewRawData.textContent = processedValues.map((v, i) => `#${i+1}: ${v.original} (${v.note})`).join('\n');
                }
            }
            
            // 隐藏 "清洗后数据" 区域，因为我们现在是一步到位的逻辑，或者用它来显示发电数据
            if (elements.previewCleanedData && elements.previewCleanedData.parentNode) {
                 // 复用此区域显示发电数据预览
                 const genDataHtml = negativeCount > 0 
                    ? `<div class="text-green-600">检测到 ${negativeCount} 个上网电量数据点 (负值)<br>上网电量估算: ${negativeSum.toFixed(2)} × ${energyFactor} × ${firstParsedData.multiplier} = ${(negativeSum * energyFactor * firstParsedData.multiplier).toFixed(4)} kWh</div>`
                    : `<div class="text-gray-400">本时段无发电数据 (无负值)</div>`;
                 
                 if (elements.previewCleanedData.tagName === 'DIV' || elements.previewCleanedData.tagName === 'SPAN') {
                    elements.previewCleanedData.innerHTML = genDataHtml;
                 } else {
                    elements.previewCleanedData.textContent = "详见上方发电数据说明";
                 }
            }
            
            // 显示计算过程
            elements.previewCalculationProcess.innerHTML = calculationProcess;
            
            // 显示计算结果
            if (elements.previewResult && elements.previewResult.parentNode) {
                elements.previewResult.textContent = `${hourlyConsumption.toFixed(4)} kWh`;
            }
            if (elements.previewFormula && elements.previewFormula.parentNode) {
                elements.previewFormula.textContent = formula;
            }
            
            // 显示模态框
            if (elements.firstHourCalculationModal && elements.firstHourCalculationModal.parentNode) {
                elements.firstHourCalculationModal.classList.remove('hidden');
            }
            setTimeout(() => {
                if (elements.firstHourCalculationModal && elements.firstHourCalculationModal.parentNode) {
                    const modalContent = elements.firstHourCalculationModal.querySelector('.bg-white');
                    if (modalContent) {
                        modalContent.classList.remove('scale-95', 'opacity-0');
                        modalContent.classList.add('scale-100', 'opacity-100');
                    }
                }
            }, 10);
        }
        
        // 关闭第一个小时计算过程预览模态框
        function closeFirstHourCalculationModal() {
            if (elements.firstHourCalculationModal && elements.firstHourCalculationModal.parentNode) {
                const modalContent = elements.firstHourCalculationModal.querySelector('.bg-white');
                if (modalContent) {
                    modalContent.classList.remove('scale-100', 'opacity-100');
                    modalContent.classList.add('scale-95', 'opacity-0');
                }
                
                setTimeout(() => {
                    elements.firstHourCalculationModal.classList.add('hidden');
                }, 300);
            }
        }
        
        // 预览第一天的计算过程
        function previewFirstDayCalculation() {
            if (!appData.processedData || appData.processedData.length === 0) {
                showNotification('错误', '没有可用的数据，请先导入并处理数据', 'error');
                return;
            }
            
            // 获取第一天的数据 - 对于列对列结构，只取第一个元素
            let firstDayData;
            if (appData.processedData[0].dataStructure === 'columnToColumn') {
                // 列对列结构：每个元素代表一个完整日期的数据，只取第一个
                firstDayData = [appData.processedData[0]];
            } else {
                // 列对行结构：按日期筛选
                const firstDate = appData.processedData[0].dateObj || parseDate(appData.processedData[0].date);
                firstDayData = appData.processedData.filter(item => {
                    const itemDate = item.dateObj || parseDate(item.date);
                    return itemDate && firstDate && itemDate.toDateString() === firstDate.toDateString();
                });
            }
            
            if (firstDayData.length === 0) {
                showNotification('错误', '没有找到第一天的数据', 'error');
                return;
            }
            
            // 获取配置信息
            const interval = appData.config.timeInterval;
            const multiplier = appData.config.multiplier;
            
            // 提取24小时数据
            let hourlyValues = [];
            let hourlyGenValues = [];
            
            // 适配不同的数据结构存储方式
            if (firstDayData[0].hourlyData && Array.isArray(firstDayData[0].hourlyData)) {
                hourlyValues = firstDayData[0].hourlyData;
                hourlyGenValues = firstDayData[0].hourlyGeneration || new Array(24).fill(0);
            } else {
                 // 可能是行数据聚合
                 hourlyValues = new Array(24).fill(null);
                 firstDayData.forEach(d => {
                     // 这里简化处理，假设processedData已经整理好了
                 });
            }

            // 过滤有效值
            const validConsumption = hourlyValues.map(v => (v !== null && !isNaN(v)) ? parseFloat(v) : 0);
            const validGeneration = hourlyGenValues.map(v => (v !== null && !isNaN(v)) ? parseFloat(v) : 0);
            
            const totalConsumption = validConsumption.reduce((a, b) => a + b, 0);
            const totalGeneration = validGeneration.reduce((a, b) => a + b, 0);
            
            // 构建计算过程
            let calculationProcess = `<div class="space-y-2 text-sm">`;
            
            calculationProcess += `<div class="font-bold text-gray-700">1. 数据汇总</div>`;
            calculationProcess += `<ul class="list-disc pl-5 text-gray-600">`;
            calculationProcess += `<li>日期: ${firstDayData[0].date || '未知'}</li>`;
            calculationProcess += `<li>有效小时数: ${hourlyValues.filter(v => v !== null).length}/24</li>`;
            calculationProcess += `</ul>`;

            calculationProcess += `<div class="font-bold text-gray-700 mt-2">2. 用电量计算 (Consumption)</div>`;
            calculationProcess += `<div class="text-xs text-gray-500 mb-1">各小时正向电量累加</div>`;
            calculationProcess += `<div class="bg-blue-50 p-2 rounded border border-blue-100 font-mono text-xs overflow-x-auto whitespace-nowrap">`;
            calculationProcess += validConsumption.map(v => v.toFixed(2)).join(' + ') + ' = ';
            calculationProcess += `<span class="font-bold text-blue-700 text-base">${totalConsumption.toFixed(2)} kWh</span>`;
            calculationProcess += `</div>`;

            if (totalGeneration > 0) {
                calculationProcess += `<div class="font-bold text-gray-700 mt-2">3. 上网电量计算 (Export Power)</div>`;
                calculationProcess += `<div class="text-xs text-gray-500 mb-1">各小时反向电量累加 (独立统计)</div>`;
                calculationProcess += `<div class="bg-green-50 p-2 rounded border border-green-100 font-mono text-xs overflow-x-auto whitespace-nowrap">`;
                calculationProcess += validGeneration.map(v => v.toFixed(2)).join(' + ') + ' = ';
                calculationProcess += `<span class="font-bold text-green-700 text-base">${totalGeneration.toFixed(2)} kWh</span>`;
                calculationProcess += `</div>`;
            } else {
                calculationProcess += `<div class="font-bold text-gray-700 mt-2">3. 上网电量计算</div>`;
                calculationProcess += `<div class="text-gray-500 italic">本日无发电数据</div>`;
            }
            
            calculationProcess += `</div>`;

            let formula = 'Σ(每小时用电量)';

            // 更新模态框内容
            if (elements.previewFirstDayDate && elements.previewFirstDayDate.parentNode) {
                elements.previewFirstDayDate.textContent = firstDayData[0] && firstDayData[0].date ? firstDayData[0].date.split(' ')[0] : '未知日期';
            }
            if (elements.previewFirstDayInterval && elements.previewFirstDayInterval.parentNode) {
                elements.previewFirstDayInterval.textContent = `${interval} 分钟`;
            }
            if (elements.previewFirstDayDataType && elements.previewFirstDayDataType.parentNode) {
                elements.previewFirstDayDataType.textContent = '聚合日电量';
            }
            if (elements.previewFirstDayMultiplier && elements.previewFirstDayMultiplier.parentNode) {
                elements.previewFirstDayMultiplier.textContent = multiplier;
            }
            
            // 显示原始数据 (这里展示小时数据)
            const rawDataHtml = hourlyValues.map((val, idx) => {
                 const genVal = hourlyGenValues[idx];
                 let display = `<div class="flex justify-between text-xs border-b border-gray-100 py-1">
                    <span class="w-10 text-gray-400">${idx}:00</span>
                    <span class="w-24 font-mono text-blue-600">用: ${val !== null ? val.toFixed(2) : '-'}</span>`;
                 
                 if (genVal > 0) {
                     display += `<span class="w-24 font-mono text-green-600">发: ${genVal.toFixed(2)}</span>`;
                 } else {
                     display += `<span class="w-24 text-gray-300">-</span>`;
                 }
                 display += `</div>`;
                 return display;
            }).join('');
            
            if (elements.previewFirstDayRawData && elements.previewFirstDayRawData.parentNode) {
                 if (elements.previewFirstDayRawData.tagName === 'DIV' || elements.previewFirstDayRawData.tagName === 'SPAN') {
                    elements.previewFirstDayRawData.innerHTML = `<div class="max-h-40 overflow-y-auto">${rawDataHtml}</div>`;
                 } else {
                    elements.previewFirstDayRawData.textContent = "请查看HTML视图";
                 }
            }
            
            // 隐藏 "清洗后数据" 区域
            if (elements.previewFirstDayCleanedData && elements.previewFirstDayCleanedData.parentNode) {
                elements.previewFirstDayCleanedData.textContent = "详见上方明细数据";
            }
            
            // 显示计算过程
            if (elements.previewFirstDayCalculationProcess && elements.previewFirstDayCalculationProcess.parentNode) {
                elements.previewFirstDayCalculationProcess.innerHTML = calculationProcess;
            }
            
            // 显示计算结果
            if (elements.previewFirstDayResult && elements.previewFirstDayResult.parentNode) {
                elements.previewFirstDayResult.textContent = `${totalConsumption.toFixed(2)} kWh`;
            }
            if (elements.previewFirstDayFormula && elements.previewFirstDayFormula.parentNode) {
                elements.previewFirstDayFormula.textContent = formula;
            }
            
            // 显示模态框
            if (elements.firstDayCalculationModal && elements.firstDayCalculationModal.parentNode) {
                elements.firstDayCalculationModal.classList.remove('hidden');
            }
            setTimeout(() => {
                if (elements.firstDayCalculationModal && elements.firstDayCalculationModal.parentNode) {
                    const modalContent = elements.firstDayCalculationModal.querySelector('.bg-white');
                    if (modalContent) {
                        modalContent.classList.remove('scale-95', 'opacity-0');
                        modalContent.classList.add('scale-100', 'opacity-100');
                    }
                }
            }, 10);
        }
        
        // 关闭第一天计算过程预览模态框
        function closeFirstDayCalculationModal() {
            if (elements.firstDayCalculationModal && elements.firstDayCalculationModal.parentNode) {
                const modalContent = elements.firstDayCalculationModal.querySelector('.bg-white');
                if (modalContent) {
                    modalContent.classList.remove('scale-100', 'opacity-100');
                    modalContent.classList.add('scale-95', 'opacity-0');
                }
                
                setTimeout(() => {
                    elements.firstDayCalculationModal.classList.add('hidden');
                }, 300);
            }
        }
        
        // 更新数据统计信息
        function updateDataStats() {
            const totalFiles = appData.files.length;
            const totalRecords = appData.worksheets.reduce((sum, ws) => sum + (ws.data.length - 1), 0);
            
            const totalFilesElement = document.getElementById('totalFiles');
            const totalRecordsElement = document.getElementById('totalRecords');
            if (totalFilesElement && totalFilesElement.parentNode) {
                totalFilesElement.textContent = totalFiles;
            }
            if (totalRecordsElement && totalRecordsElement.parentNode) {
                totalRecordsElement.textContent = totalRecords;
            }
            
            // 更新数据概览统计信息
            updateDataOverview();
            
            // 更新文件列表显示
            renderFileList();
        }
        
        // 更新数据概览div中的统计信息
        function updateDataOverview() {
            const fileCountElement = document.getElementById('fileCount');
            const timeRangeElement = document.getElementById('timeRange');
            const fileBadgeCount = document.getElementById('fileBadgeCount');
            
            // 更新文件数量
            if (fileCountElement) {
                fileCountElement.textContent = appData.files.length;
            }
            if (fileBadgeCount) {
                fileBadgeCount.textContent = appData.files.length;
            }
            
            // 更新时间范围
            if (timeRangeElement && timeRangeElement.parentNode) {
                if (appData.processedData.length > 0) {
                    const dates = appData.processedData.map(d => d.date).sort();
                    const startDate = dates[0];
                    const endDate = dates[dates.length - 1];
                    if (startDate === endDate) {
                        timeRangeElement.textContent = startDate;
                    } else {
                        timeRangeElement.textContent = `${startDate} 至 ${endDate}`;
                    }
                    
                    // 更新导出日期下拉框
                    updateExportDateDropdowns(dates);
                    updateVisualizationDateSelects();
                } else {
                    timeRangeElement.textContent = '-';
                    updateExportDateDropdowns([]);
                    updateVisualizationDateSelects();
                }
            }
        }
        
        /**
         * 更新导出设置中的日期下拉框
         * @param {string[]} dates 已排序的日期字符串数组
         */
        function updateExportDateDropdowns(dates) {
            const startSelect = document.getElementById('exportStartDate');
            const endSelect = document.getElementById('exportEndDate');
            
            if (!startSelect || !endSelect) return;
            
            // 记录当前选中的值
            const currentStart = startSelect.value;
            const currentEnd = endSelect.value;
            
            // 清空并重新填充
            startSelect.innerHTML = '<option value="">开始日期</option>';
            endSelect.innerHTML = '<option value="">结束日期</option>';
            
            if (dates.length === 0) return;
            
            // 去重（虽然 dates 应该是已经按日期排序的，但 map(d => d.date) 可能会有重复日期，如果一个日期有多个记录）
            const uniqueDates = [...new Set(dates)];
            
            uniqueDates.forEach(date => {
                const optStart = document.createElement('option');
                optStart.value = date;
                optStart.textContent = date;
                startSelect.appendChild(optStart);
                
                const optEnd = document.createElement('option');
                optEnd.value = date;
                optEnd.textContent = date;
                endSelect.appendChild(optEnd);
            });
            
            // 恢复选中值，如果还在列表中；否则默认选中第一个和最后一个
            if (uniqueDates.includes(currentStart)) {
                startSelect.value = currentStart;
            } else {
                startSelect.value = uniqueDates[0];
            }
            
            if (uniqueDates.includes(currentEnd)) {
                endSelect.value = currentEnd;
            } else {
                endSelect.value = uniqueDates[uniqueDates.length - 1];
            }
        }
        
        function renderFileList() {
            const fileListElement = document.getElementById('fileListContent');
            if (!fileListElement) return;
            
            fileListElement.innerHTML = '';
            
            if (appData.files.length === 0) {
                fileListElement.innerHTML = '<div class="text-center py-4 text-gray-500">暂无导入的文件</div>';
                return;
            }
            
            appData.files.forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'flex items-center justify-between p-3 bg-gray-50 rounded-lg mb-2';
                
                const worksheet = appData.worksheets.find(ws => ws.file === file);
                const recordCount = worksheet ? worksheet.data.length - 1 : 0;
                
                fileItem.innerHTML = `
                    <div class="flex items-center space-x-3">
                        <div class="w-8 h-8 bg-indigo-50 rounded-full flex items-center justify-center ring-1 ring-indigo-100">
                            <svg class="w-4 h-4 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                            </svg>
                        </div>
                        <div>
                            <div class="text-sm font-bold text-slate-800">${file.name}</div>
                            <div class="text-[10px] font-bold text-indigo-600/70 uppercase tracking-wider">${recordCount} 条记录</div>
                        </div>
                    </div>
                    <button onclick="removeFile(${index})" class="text-slate-400 hover:text-rose-600 transition-all p-1.5 rounded-xl hover:bg-rose-50 group">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                        </svg>
                    </button>
                `;
                
                fileListElement.appendChild(fileItem);
            });
        }
        
        function updateUploadProgress(processed, total) {
            const progressBar = document.getElementById('progressBar');
            const progressText = document.getElementById('progressText');
            
            // 确保进度不超过100%
            const actualProcessed = Math.min(processed, total);
            const percentage = Math.round((actualProcessed / total) * 100);
            
            if (progressBar && progressBar.parentNode) {
                progressBar.style.width = `${percentage}%`;
            }
            if (progressText && progressText.parentNode) {
                progressText.textContent = `${percentage}% (${actualProcessed}/${total})`;
            }
            
            if (processed >= total) {
                setTimeout(() => {
                    if (elements.uploadProgress && elements.uploadProgress.parentNode) {
                        elements.uploadProgress.classList.add('hidden');
                    }
                }, 1000);
            }
        }

        function setImportStatus(state, text) {
            const importStatusElement = document.getElementById('importStatus');
            if (!importStatusElement || !importStatusElement.parentNode) return;

            const variants = {
                idle: {
                    className: 'inline-flex items-center rounded-full border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-700',
                    text: '等待文件上传'
                },
                processing: {
                    className: 'inline-flex items-center rounded-full border border-green-200 bg-green-50 px-3 py-1.5 text-xs font-semibold text-green-700',
                    text: '正在处理文件...'
                },
                done: {
                    className: 'inline-flex items-center rounded-full border border-green-200 bg-green-50 px-3 py-1.5 text-xs font-semibold text-green-700',
                    text: '文件导入完成'
                },
                error: {
                    className: 'inline-flex items-center rounded-full border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-700',
                    text: '导入失败'
                }
            };

            const variant = variants[state] || variants.idle;
            importStatusElement.className = variant.className;
            importStatusElement.textContent = text || variant.text;
        }
        








        // 辅助函数：解析日期时间
function parseDateTime(dateTimeStr, format) {
    if (!dateTimeStr || typeof dateTimeStr !== 'string') {
        return null;
    }
    
    // 尝试多种日期时间格式
    const formats = [
        { regex: /\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/, id: 'yyyy-mm-dd-hh-mm-ss' }, // YYYY-MM-DD HH:MM:SS
        { regex: /\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}/, id: 'yyyy-mm-dd-hh-mm' },     // YYYY/MM/DD HH:MM
        { regex: /\d{4}\d{2}\d{2}\d{2}\d{2}\d{2}/, id: 'yyyymmddhhmmss' },     // YYYYMMDDHHMMSS
        { regex: /\d{4}\/\d{2}\/\d{2}\/\d{2}:\d{2}/, id: 'yyyy-mm-dd-hh-mm-slash' },    // YYYY/MM/DD/HH:MM
        { regex: /\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}/, id: 'mm-dd-yyyy-hh-mm' },     // MM/DD/YYYY HH:MM
        { regex: /\d{2}-\d{2}-\d{4} \d{2}:\d{2}/, id: 'mm-dd-yyyy-hh-mm-dash' },      // MM-DD-YYYY HH:MM
        { regex: /\d{4}-\d{2}-\d{2} \d{2}:\d{2}/, id: 'yyyy-mm-dd-hh-mm-dash' },      // YYYY-MM-DD HH:MM
        { regex: /\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}:\d{2}/, id: 'yyyy-mm-dd-hh-mm-ss-slash' }, // YYYY/MM/DD HH:MM:SS
        { regex: /\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}/, id: 'mm-dd-yyyy-hh-mm-ss' }  // MM/DD/YYYY HH:MM:SS
    ];
    
    for (const fmt of formats) {
        if (fmt.regex.test(dateTimeStr)) {
            let parts;
            let date;
            
            // 根据匹配的格式解析日期时间
            if (fmt.id === 'yyyy-mm-dd-hh-mm-ss') {
                parts = dateTimeStr.match(/(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/);
                if (parts && parts.length >= 7) {
                    date = new Date(parts[1], parts[2] - 1, parts[3], parts[4], parts[5], parts[6]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'yyyy-mm-dd-hh-mm') {
                parts = dateTimeStr.match(/(\d{4})\/(\d{2})\/(\d{2}) (\d{2}):(\d{2})/);
                if (parts && parts.length >= 6) {
                    date = new Date(parts[1], parts[2] - 1, parts[3], parts[4], parts[5]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'yyyymmddhhmmss') {
                if (dateTimeStr.length >= 14) {
                    const year = parseInt(dateTimeStr.substring(0, 4));
                    const month = parseInt(dateTimeStr.substring(4, 6)) - 1;
                    const day = parseInt(dateTimeStr.substring(6, 8));
                    const hour = parseInt(dateTimeStr.substring(8, 10));
                    const minute = parseInt(dateTimeStr.substring(10, 12));
                    const second = parseInt(dateTimeStr.substring(12, 14));
                    date = new Date(year, month, day, hour, minute, second);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'yyyy-mm-dd-hh-mm-slash') {
                parts = dateTimeStr.match(/(\d{4})\/(\d{2})\/(\d{2})\/(\d{2}):(\d{2})/);
                if (parts && parts.length >= 6) {
                    date = new Date(parts[1], parts[2] - 1, parts[3], parts[4], parts[5]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'mm-dd-yyyy-hh-mm') {
                parts = dateTimeStr.match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
                if (parts && parts.length >= 6) {
                    date = new Date(parts[3], parts[1] - 1, parts[2], parts[4], parts[5]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'mm-dd-yyyy-hh-mm-dash') {
                parts = dateTimeStr.match(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/);
                if (parts && parts.length >= 6) {
                    date = new Date(parts[3], parts[1] - 1, parts[2], parts[4], parts[5]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'yyyy-mm-dd-hh-mm-dash') {
                parts = dateTimeStr.match(/(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
                if (parts && parts.length >= 6) {
                    date = new Date(parts[1], parts[2] - 1, parts[3], parts[4], parts[5]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'yyyy-mm-dd-hh-mm-ss-slash') {
                parts = dateTimeStr.match(/(\d{4})\/(\d{2})\/(\d{2}) (\d{2}):(\d{2}):(\d{2})/);
                if (parts && parts.length >= 7) {
                    date = new Date(parts[1], parts[2] - 1, parts[3], parts[4], parts[5], parts[6]);
                    if (!isNaN(date.getTime())) return date;
                }
            } else if (fmt.id === 'mm-dd-yyyy-hh-mm-ss') {
                parts = dateTimeStr.match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})/);
                if (parts && parts.length >= 7) {
                    date = new Date(parts[3], parts[1] - 1, parts[2], parts[4], parts[5], parts[6]);
                    if (!isNaN(date.getTime())) return date;
                }
            }
        }
    }
    
    // 如果没有匹配的格式，尝试使用Date对象解析
    const date = new Date(dateTimeStr);
    return isNaN(date.getTime()) ? null : date;
}

// 辅助函数：格式化日期键
        function formatDateKey(date) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}${month}${day}`;
        }
        
        // 统一计算小时功率的函数
        function calculateHourlyPower(dataPoints, multiplier, timeInterval = 15, isDebug = false) {
            // 过滤无效数据
            const validData = dataPoints.filter(v => v !== null);
            
            if (validData.length === 0) {
                if (isDebug) {
                    console.log('计算小时功率调试信息：');
                    console.log('- 原始数据点数量:', dataPoints.length);
                    console.log('- 有效数据点数量: 0');
                    console.log('- 计算结果: null (无有效数据)');
                }
                return null;
            }
            
            // 根据时间间隔计算每小时应有的数据点数量
            const expectedPointsPerHour = 60 / timeInterval;
            
            // 计算总和
            const sum = validData.reduce((a, b) => a + b, 0);
            
            // 计算平均功率
            const avgPower = sum / validData.length;
            
            // 计算小时值：平均功率 × 1 × 倍率
            const hourlyValue = avgPower * 1 * multiplier;
            
            // 检查数据点数量是否匹配预期
            const isPointsCountValid = Math.abs(validData.length - expectedPointsPerHour) / expectedPointsPerHour <= 0.2; // 允许20%的误差
            
            if (isDebug) {
                console.log('计算小时功率调试信息：');
                console.log('- 原始数据点数量:', dataPoints.length);
                console.log('- 有效数据点数量:', validData.length);
                console.log('- 时间间隔:', timeInterval, '分钟');
                console.log('- 预期每小时数据点数量:', expectedPointsPerHour);
                console.log('- 数据点数量是否匹配:', isPointsCountValid ? '是' : '否');
                console.log('- 数据点总和:', sum.toFixed(2), 'kW');
                console.log('- 平均功率:', avgPower.toFixed(2), 'kW');
                console.log('- 倍率:', multiplier);
                console.log('- 计算结果:', hourlyValue.toFixed(2), 'kWh');
            }
            
            return hourlyValue;
        }
        
        // 辅助函数：获取筛选后的数据
        function getFilteredData() {
            console.log('=== getFilteredData 调试信息 ===');
            console.log('原始数据总数:', appData.processedData.length);
            console.log('选中日期:', appData.visualization.selectedDates);
            console.log('选中计量点:', appData.visualization.selectedMeteringPoints);
            
            // 如果没有选择任何日期，默认根据模式选择
            if (appData.visualization.selectedDates.length === 0 && appData.processedData.length > 0) {
                const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                if (appData.visualization.allCurvesMode) {
                    // 默认选择全部日期
                    appData.visualization.selectedDates = allDates;
                    console.log('自动选择全部日期:', appData.visualization.selectedDates.length);
                } else {
                    // 默认选择第一天
                    appData.visualization.selectedDates = [allDates[0]];
                    console.log('自动选择第一天:', appData.visualization.selectedDates[0]);
                    
                    // 同步更新日期选择器的值
                    if (elements.startDateInput) elements.startDateInput.value = allDates[0];
                    if (elements.endDateInput) elements.endDateInput.value = allDates[0];
                }
            }
            
            // 如果没有选择任何计量点，默认选择所有计量点
            if (appData.visualization.selectedMeteringPoints.length === 0 && appData.meteringPoints.length > 0) {
                appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                console.log('自动选择所有计量点:', appData.visualization.selectedMeteringPoints);
            }
            
            // 当已选择集合与新数据可用集合无交集时，自动回退为全选
            const availableDateSet = new Set(appData.processedData.map(d => d.date));
            if (appData.visualization.selectedDates.length > 0) {
                const hasDateOverlap = appData.visualization.selectedDates.some(d => availableDateSet.has(d));
                if (!hasDateOverlap && appData.processedData.length > 0) {
                    const allDates = [...new Set(appData.processedData.map(d => d.date))].sort();
                    appData.visualization.selectedDates = allDates;
                    console.log('检测到选择的日期与新数据不重叠，已自动重置为全部日期');
                }
            }
            const availableMPSet = new Set(appData.meteringPoints);
            if (appData.visualization.selectedMeteringPoints.length > 0) {
                const hasMPOverlap = appData.visualization.selectedMeteringPoints.some(mp => availableMPSet.has(mp));
                if (!hasMPOverlap && appData.meteringPoints.length > 0) {
                    appData.visualization.selectedMeteringPoints = [...appData.meteringPoints];
                    console.log('检测到选择的计量点与新数据不重叠，已自动重置为全部计量点');
                }
            }
            
            const filteredData = DataUtils.filterData(appData.processedData, {
                dates: appData.visualization.selectedDates,
                meteringPoints: appData.visualization.selectedMeteringPoints.length > 0 ? 
                    appData.visualization.selectedMeteringPoints : null
            });
            
            console.log('筛选后数据总数:', filteredData.length);
            console.log('筛选后数据日期:', filteredData.map(d => d.date));
            console.log('筛选后数据计量点:', [...new Set(filteredData.map(d => d.meteringPoint))]);
            if (filteredData.length > 0) {
                console.log('第一条数据结构:', Object.keys(filteredData[0]));
                console.log('第一条数据的dailyTotal:', filteredData[0].dailyTotal);
            }
            console.log('===============================');
            
            return filteredData;
        }

        function getOverviewData() {
            const baseData = Array.isArray(appData.processedData) ? appData.processedData : [];
            const selected = Array.isArray(appData.visualization?.selectedMeteringPoints)
                ? appData.visualization.selectedMeteringPoints
                : [];

            if (selected.length === 0) return baseData;
            const selectedSet = new Set(selected);
            const allMP = Array.isArray(appData.meteringPoints) ? appData.meteringPoints : [];
            if (allMP.length > 0 && selectedSet.size >= allMP.length) return baseData;
            return baseData.filter(d => selectedSet.has(d.meteringPoint));
        }
        
        // 性能优化配置
        const PERFORMANCE_CONFIG = {
            BATCH_SIZE: 1000,           // 批处理大小
            CACHE_ENABLED: true,        // 启用缓存
            WORKER_ENABLED: false,      // Web Worker支持（暂时禁用）
            DEBOUNCE_DELAY: 300         // 防抖延迟
        };
        
        // 缓存管理器
        const CacheManager = {
            cache: new Map(),
            
            get: function(key) {
                return this.cache.get(key);
            },
            
            set: function(key, value, ttl = 300000) { // 默认5分钟TTL
                const item = {
                    value,
                    timestamp: Date.now(),
                    ttl
                };
                this.cache.set(key, item);
                
                // 清理过期缓存
                this.cleanup();
            },
            
            has: function(key) {
                const item = this.cache.get(key);
                if (!item) return false;
                
                if (Date.now() - item.timestamp > item.ttl) {
                    this.cache.delete(key);
                    return false;
                }
                return true;
            },
            
            cleanup: function() {
                const now = Date.now();
                for (const [key, item] of this.cache.entries()) {
                    if (now - item.timestamp > item.ttl) {
                        this.cache.delete(key);
                    }
                }
            },
            
            clear: function() {
                this.cache.clear();
            }
        };
        
        // 数据处理工具类
        const DataUtils = {
            // 通用数据过滤器
            filterData: function(data, filters = {}) {
                // 预处理过滤器为 Set 以提升查找性能
                const dateSet = filters.dates && filters.dates.length > 0 ? new Set(filters.dates) : null;
                const mpSet = filters.meteringPoints && filters.meteringPoints.length > 0 ? new Set(filters.meteringPoints) : null;

                return data.filter(item => {
                    // 日期过滤
                    if (dateSet) {
                        if (!dateSet.has(item.date)) return false;
                    }
                    
                    // 计量点过滤
                    if (mpSet) {
                        // 与下拉选项和分组逻辑保持一致：将空值统一映射为『默认计量点』
                        const rawMp = item.meteringPoint;
                        const mpValue = (rawMp === undefined || rawMp === null || String(rawMp).trim() === '')
                            ? '默认计量点'
                            : String(rawMp).trim();
                        if (!mpSet.has(mpValue)) return false;
                    }
                    
                    // 自定义过滤器
                    if (filters.custom && typeof filters.custom === 'function') {
                        if (!filters.custom(item)) return false;
                    }
                    
                    return true;
                });
            },
            
            // 数据分组工具
            groupBy: function(data, keyFn) {
                const groups = new Map();
                data.forEach(item => {
                    const key = keyFn(item);
                    if (!groups.has(key)) {
                        groups.set(key, []);
                    }
                    groups.get(key).push(item);
                });
                return groups;
            },
            
            // 数据聚合工具
            aggregate: function(data, aggregateFn) {
                if (!data || data.length === 0) return null;
                return aggregateFn(data);
            },
            
            // 数据清洗工具
            cleanData: function(data, options = {}) {
                const { fillMethod = 'ignore' } = options;
                
                return data.map((value, index, array) => {
                    if (value === null || value === undefined || value === '' || isNaN(value)) {
                        if (fillMethod === 'ignore') {
                            return null;
                        } else if (fillMethod === 'interpolate') {
                            // 用相邻数据填充
                            let prevValue = null;
                            for (let i = index - 1; i >= 0; i--) {
                                if (array[i] !== null && !isNaN(array[i])) {
                                    prevValue = parseFloat(array[i]);
                                    break;
                                }
                            }
                            
                            let nextValue = null;
                            for (let i = index + 1; i < array.length; i++) {
                                if (array[i] !== null && !isNaN(array[i])) {
                                    nextValue = parseFloat(array[i]);
                                    break;
                                }
                            }
                            
                            if (prevValue !== null && nextValue !== null) {
                                return (prevValue + nextValue) / 2;
                            } else if (prevValue !== null) {
                                return prevValue;
                            } else if (nextValue !== null) {
                                return nextValue;
                            } else {
                                return null;
                            }
                        }
                        return null;
                    }
                    return parseFloat(value);
                });
            },
            
            // 数据验证工具
            validateData: function(data, rules = {}) {
                const errors = [];
                
                if (rules.required && (!data || data.length === 0)) {
                    errors.push('数据不能为空');
                }
                
                if (rules.minLength && data.length < rules.minLength) {
                    errors.push(`数据长度不能少于${rules.minLength}`);
                }
                
                if (rules.maxLength && data.length > rules.maxLength) {
                    errors.push(`数据长度不能超过${rules.maxLength}`);
                }
                
                return { isValid: errors.length === 0, errors };
            },
            
            // 批处理工具
            processBatch: function(data, processFn, batchSize = PERFORMANCE_CONFIG.BATCH_SIZE) {
                const results = [];
                
                for (let i = 0; i < data.length; i += batchSize) {
                    const batch = data.slice(i, i + batchSize);
                    const batchResult = processFn(batch);
                    
                    if (Array.isArray(batchResult)) {
                        results.push(...batchResult);
                    } else {
                        results.push(batchResult);
                    }
                    
                    // 让出控制权，避免阻塞UI
                    if (i + batchSize < data.length) {
                        setTimeout(() => {}, 0);
                    }
                }
                
                return results;
            },
            
            // 缓存装饰器
            withCache: function(fn, keyGenerator) {
                return function(...args) {
                    if (!PERFORMANCE_CONFIG.CACHE_ENABLED) {
                        return fn.apply(this, args);
                    }
                    
                    const cacheKey = keyGenerator ? keyGenerator(...args) : JSON.stringify(args);
                    
                    if (CacheManager.has(cacheKey)) {
                        return CacheManager.get(cacheKey).value;
                    }
                    
                    const result = fn.apply(this, args);
                    CacheManager.set(cacheKey, result);
                    
                    return result;
                };
            },
            
            // 防抖工具
            debounce: function(fn, delay = PERFORMANCE_CONFIG.DEBOUNCE_DELAY) {
                let timeoutId;
                return function(...args) {
                    clearTimeout(timeoutId);
                    timeoutId = setTimeout(() => fn.apply(this, args), delay);
                };
            },
            
            // 优化的forEach（支持早期退出）
            fastForEach: function(array, callback, breakCondition) {
                for (let i = 0; i < array.length; i++) {
                    const result = callback(array[i], i, array);
                    if (breakCondition && breakCondition(result, i)) {
                        break;
                    }
                }
            },
            
            // 内存友好的大数组处理
            processLargeArray: function(array, processFn, options = {}) {
                const { 
                    batchSize = PERFORMANCE_CONFIG.BATCH_SIZE,
                    onProgress = null,
                    onComplete = null 
                } = options;
                
                let processedCount = 0;
                const totalCount = array.length;
                const results = [];
                
                const processBatch = (startIndex) => {
                    const endIndex = Math.min(startIndex + batchSize, totalCount);
                    const batch = array.slice(startIndex, endIndex);
                    
                    const batchResults = batch.map(processFn);
                    results.push(...batchResults);
                    
                    processedCount += batch.length;
                    
                    if (onProgress) {
                        onProgress(processedCount, totalCount);
                    }
                    
                    if (endIndex < totalCount) {
                        // 使用 requestAnimationFrame 或 setTimeout 来避免阻塞
                        if (window.requestAnimationFrame) {
                            requestAnimationFrame(() => processBatch(endIndex));
                        } else {
                            setTimeout(() => processBatch(endIndex), 0);
                        }
                    } else if (onComplete) {
                        onComplete(results);
                    }
                };
                
                processBatch(0);
                return results;
            }
        };
        
        // 辅助函数：获取颜色
        function getColor(index) {
            const colors = [
                '#6366F1',   // 现代靛蓝色
                '#8B5CF6',   // 现代紫罗兰色
                '#EC4899',   // 现代粉红色
                '#10B981',   // 现代翠绿色
                '#F59E0B',   // 现代琥珀色
                '#EF4444',   // 现代红色
                '#06B6D4',   // 现代青色
                '#F97316',   // 现代橙色
                '#84CC16',   // 现代酸橙绿
                '#A855F7'    // 现代紫色
            ];
            
            return colors[index % colors.length];
        }
        
        // 辅助函数：获取颜色数量
        function getColorCount() {
            return 10; // 颜色数组中的颜色数量
        }



// 启动应用
document.addEventListener('DOMContentLoaded', initApp);



        // 时段统计对比窗口功能
        let timeSlotComparisonDragging = false;
        let timeSlotComparisonDragOffset = { x: 0, y: 0 };
        let timeSlotComparisonMinimized = false;

        // 显示时段统计对比窗口
        window.showTimeSlotComparison = function() {
            const window = document.getElementById('timeSlotComparisonWindow');
            if (window && window.parentNode) {
                window.classList.remove('hidden');
            }
            updateTimeSlotComparison();
            
            // 添加拖动事件监听
            setupTimeSlotComparisonDrag();
        }

        // 关闭时段统计对比窗口
        window.closeTimeSlotComparison = function() {
            const window = document.getElementById('timeSlotComparisonWindow');
            if (window && window.parentNode) {
                window.classList.add('hidden');
            }
        }

        // 最小化/恢复时段统计对比窗口
        window.toggleTimeSlotComparisonMinimize = function() {
            const content = document.getElementById('timeSlotComparisonContent');
            const minimizeBtn = document.getElementById('minimizeTimeSlotComparison');
            
            if (timeSlotComparisonMinimized) {
                content.style.display = 'block';
                minimizeBtn.innerHTML = '<i class="fa fa-window-minimize"></i>';
                timeSlotComparisonMinimized = false;
            } else {
                content.style.display = 'none';
                minimizeBtn.innerHTML = '<i class="fa fa-window-maximize"></i>';
                timeSlotComparisonMinimized = true;
            }
        }

        // 设置拖动功能
        function setupTimeSlotComparisonDrag() {
            const header = document.getElementById('timeSlotComparisonHeader');
            const window = document.getElementById('timeSlotComparisonWindow');

            header.addEventListener('mousedown', function(e) {
                if (e.target.tagName === 'BUTTON' || e.target.tagName === 'I') return;
                
                timeSlotComparisonDragging = true;
                const rect = window.getBoundingClientRect();
                timeSlotComparisonDragOffset.x = e.clientX - rect.left;
                timeSlotComparisonDragOffset.y = e.clientY - rect.top;
                
                document.addEventListener('mousemove', timeSlotComparisonMouseMove);
                document.addEventListener('mouseup', timeSlotComparisonMouseUp);
            });
        }

        function timeSlotComparisonMouseMove(e) {
            if (!timeSlotComparisonDragging) return;
            
            const window = document.getElementById('timeSlotComparisonWindow');
            const newX = e.clientX - timeSlotComparisonDragOffset.x;
            const newY = e.clientY - timeSlotComparisonDragOffset.y;
            
            // 确保窗口不超出屏幕边界
            const maxX = window.innerWidth - window.offsetWidth;
            const maxY = window.innerHeight - window.offsetHeight;
            
            window.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
            window.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
        }

        function timeSlotComparisonMouseUp() {
            timeSlotComparisonDragging = false;
            document.removeEventListener('mousemove', timeSlotComparisonMouseMove);
            document.removeEventListener('mouseup', timeSlotComparisonMouseUp);
        }

        // 更新时段统计对比数据
        window.updateTimeSlotComparison = function() {
            const startHour = parseInt(document.getElementById('timeSlotStart').value);
            const endHour = parseInt(document.getElementById('timeSlotEnd').value);
            
            if (startHour > endHour) {
                showNotification('提示', '开始时段不能大于结束时段', 'warning');
                return;
            }

            const filteredData = getFilteredData();
            if (filteredData.length === 0) {
                document.getElementById('timeSlotComparisonTableBody').innerHTML = '<tr><td colspan="6" class="border border-gray-200 p-2 text-center text-gray-500 text-sm">暂无数据</td></tr>';
                return;
            }

            // 定义颜色数组
            const colors = [
                '#3B82F6', '#EF4444', '#10B981', '#F59E0B', '#8B5CF6', '#06B6D4', '#F97316', '#84CC16',
                '#EC4899', '#6366F1', '#14B8A6', '#F43F5E', '#A855F7', '#0EA5E9', '#EAB308', '#22C55E',
                '#D946EF', '#F87171', '#34D399', '#60A5FA', '#FBBF24', '#A78BFA', '#FB923C', '#67E8F9', '#FDA4AF'
            ];

            let totalPeriodEnergy = 0;
            let maxEnergy = 0;
            let minEnergy = Infinity;
            const periodData = [];
            let periodHours = endHour - startHour + 1;

            // 计算每个日期在指定时段的用电量
            filteredData.forEach(row => {
                let periodEnergy = 0;
                let validHours = 0;
                
                for (let hour = startHour; hour <= endHour; hour++) {
                    if (row.hourlyData[hour] !== null && row.hourlyData[hour] !== undefined) {
                        periodEnergy += row.hourlyData[hour];
                        validHours++;
                    }
                }
                
                if (validHours > 0) {
                    periodData.push({
                        date: row.date,
                        energy: periodEnergy,
                        totalEnergy: row.dailyTotal || 0,
                        hourlyAverage: validHours > 0 ? periodEnergy / validHours : 0
                    });
                    
                    totalPeriodEnergy += periodEnergy;
                    maxEnergy = Math.max(maxEnergy, periodEnergy);
                    minEnergy = Math.min(minEnergy, periodEnergy);
                }
            });

            // 计算汇总统计
            const avgEnergy = periodData.length > 0 ? totalPeriodEnergy / periodData.length : 0;
            const totalDailyEnergy = periodData.reduce((sum, item) => sum + item.totalEnergy, 0);
            const overallPercentage = totalDailyEnergy > 0 ? (totalPeriodEnergy / totalDailyEnergy) * 100 : 0;
            const avgHourlyOverall = totalDailyEnergy > 0 ? totalDailyEnergy / (periodData.length * 24) : 0;
            const avgHourlyInPeriod = totalPeriodEnergy > 0 ? totalPeriodEnergy / (periodData.length * periodHours) : 0;

            // 更新汇总卡片
            document.getElementById('timeSlotTotal').textContent = totalPeriodEnergy.toFixed(2) + ' kWh';
            document.getElementById('timeSlotAvg').textContent = avgEnergy.toFixed(2) + ' kWh';
            document.getElementById('timeSlotMax').textContent = maxEnergy.toFixed(2) + ' kWh';
            document.getElementById('timeSlotMin').textContent = minEnergy === Infinity ? '0.00 kWh' : minEnergy.toFixed(2) + ' kWh';

            // 添加新的汇总信息
            const summaryHtml = `
                <div class="grid grid-cols-2 md:grid-cols-5 gap-4 mt-4">
                    <div class="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100/50 group hover:bg-indigo-50 transition-colors">
                        <div class="text-indigo-600/70 text-[10px] font-bold uppercase tracking-wider mb-1">占全天比例</div>
                        <div class="text-indigo-900 text-lg font-black tracking-tight">${overallPercentage.toFixed(1)}%</div>
                    </div>
                    <div class="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100/50 group hover:bg-indigo-50 transition-colors">
                        <div class="text-indigo-600/70 text-[10px] font-bold uppercase tracking-wider mb-1">时段平均小时</div>
                        <div class="text-indigo-900 text-lg font-black tracking-tight">${avgHourlyInPeriod.toFixed(2)} <span class="text-xs font-medium opacity-60">kWh</span></div>
                    </div>
                    <div class="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100/50 group hover:bg-indigo-50 transition-colors">
                        <div class="text-indigo-600/70 text-[10px] font-bold uppercase tracking-wider mb-1">用电量最大日</div>
                        <div class="text-indigo-900 text-sm font-bold truncate" title="${(periodData.sort((a,b)=>b.energy-a.energy)[0]?.date) || '-'}">${(periodData.sort((a,b)=>b.energy-a.energy)[0]?.date) || '-'}</div>
                        <div class="text-indigo-600 text-xs font-black mt-0.5">${(periodData.length?Math.max(...periodData.map(p=>p.energy)).toFixed(2):'0.00')} <span class="text-[10px] font-medium opacity-60">kWh</span></div>
                    </div>
                    <div class="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100/50 group hover:bg-indigo-50 transition-colors">
                        <div class="text-indigo-600/70 text-[10px] font-bold uppercase tracking-wider mb-1">用电量最小日</div>
                        <div class="text-indigo-900 text-sm font-bold truncate" title="${(periodData.sort((a,b)=>a.energy-b.energy)[0]?.date) || '-'}">${(periodData.sort((a,b)=>a.energy-b.energy)[0]?.date) || '-'}</div>
                        <div class="text-indigo-600 text-xs font-black mt-0.5">${(periodData.length?Math.min(...periodData.map(p=>p.energy)).toFixed(2):'0.00')} <span class="text-[10px] font-medium opacity-60">kWh</span></div>
                    </div>
                    <div class="bg-indigo-50/50 p-3 rounded-xl border border-indigo-100/50 group hover:bg-indigo-50 transition-colors">
                        <div class="text-indigo-600/70 text-[10px] font-bold uppercase tracking-wider mb-1">分析天数</div>
                        <div class="text-indigo-900 text-lg font-black tracking-tight">${periodData.length} <span class="text-xs font-medium opacity-60">天</span></div>
                    </div>
                </div>
            `;
            
            const timeSlotSummary = document.getElementById('timeSlotSummary');
            const existingSummary = timeSlotSummary.querySelector('.grid');
            if (existingSummary) existingSummary.remove();
            timeSlotSummary.insertAdjacentHTML('beforeend', summaryHtml);

            // 更新表格 - 添加更多列
            const tbody = document.getElementById('timeSlotComparisonTableBody');
            tbody.innerHTML = '';

            periodData.forEach((item, index) => {
                const dailyPercentage = item.totalEnergy > 0 ? ((item.energy / item.totalEnergy) * 100).toFixed(1) : 0;
                const comparison = index > 0 ? 
                    ((item.energy - periodData[0].energy) / periodData[0].energy * 100).toFixed(1) : '0.0';
                
                const row = document.createElement('tr');
                row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                row.innerHTML = `
                    <td class="border border-gray-200 p-2 text-center" style="color: ${colors[index % colors.length]}">${item.date}</td>
                    <td class="border border-gray-200 p-2 text-center">${item.energy.toFixed(2)}</td>
                    <td class="border border-gray-200 p-2 text-center">${dailyPercentage}%</td>
                    <td class="border border-gray-200 p-2 text-center">${item.hourlyAverage.toFixed(2)}</td>
                    <td class="border border-gray-200 p-2 text-center">${item.totalEnergy.toFixed(2)}</td>
                    <td class="border border-slate-200 p-2 text-center font-bold ${comparison >= 0 ? 'text-indigo-600' : 'text-slate-500'}">
                        ${comparison >= 0 ? '+' : ''}${comparison}%
                    </td>
                `;
                tbody.appendChild(row);
            });

            // 更新表格标题
            const thead = document.querySelector('#timeSlotComparisonTable thead tr');
            if (thead) {
                thead.innerHTML = `
                    <th class="border border-gray-200 p-2 text-center font-medium">日期</th>
                    <th class="border border-gray-200 p-2 text-center font-medium">时段用电量 (kWh)</th>
                    <th class="border border-gray-200 p-2 text-center font-medium">占全天比例</th>
                    <th class="border border-gray-200 p-2 text-center font-medium">时段平均小时 (kWh)</th>
                    <th class="border border-gray-200 p-2 text-center font-medium">日总用电量 (kWh)</th>
                    <th class="border border-gray-200 p-2 text-center font-medium">对比首日 (%)</th>
                `;
            }
        }

        // 添加时段统计对比按钮到界面
        setTimeout(() => {
            const controlPanel = document.querySelector('.bg-white.rounded-lg.shadow-sm.border.border-gray-200');
            if (controlPanel) {
                const timeSlotBtn = document.createElement('button');
                timeSlotBtn.innerHTML = '<i class="fa fa-clock mr-1"></i> 时段统计';
                timeSlotBtn.className = 'bg-primary text-white px-3 py-1 rounded text-sm hover:bg-primary/90 ml-2';
                timeSlotBtn.onclick = showTimeSlotComparison;
                controlPanel.appendChild(timeSlotBtn);
            }

            // 添加事件监听
            document.getElementById('closeTimeSlotComparison').addEventListener('click', closeTimeSlotComparison);
            document.getElementById('timeSlotStart').addEventListener('change', updateTimeSlotComparison);
            document.getElementById('timeSlotEnd').addEventListener('change', updateTimeSlotComparison);
        }, 1000);

        // 作者信息弹窗功能
        document.addEventListener('DOMContentLoaded', function() {
            const authorModal = document.getElementById('authorInfoModal');
            const startBtn = document.getElementById('startUsingBtn');
            const authorSection = document.getElementById('authorSection');
            
            // 更新导航可见性
            function updateNavigationVisibility(show) {
                const navWrapper = document.querySelector('.nav-wrapper');
                const navToggle = document.getElementById('navToggle');
                if (navWrapper) navWrapper.style.display = show ? 'flex' : 'none';
                if (navToggle) navToggle.style.display = show ? 'flex' : 'none';
            }
            
            // 直接显示作者信息弹窗
            updateNavigationVisibility(false);
            setTimeout(() => {
                if (authorModal) {
                    authorModal.style.display = 'flex';
                }
            }, 500);
            
            // 更新图标和背景
            const iconContainer = authorModal.querySelector('.w-24.h-24');
            const icon = authorModal.querySelector('.fa-lock');
            if (iconContainer && icon) {
                iconContainer.classList.replace('from-indigo-500', 'from-indigo-500');
                iconContainer.classList.replace('to-purple-600', 'to-indigo-600');
                iconContainer.classList.remove('rotate-3');
                icon.className = 'fa fa-user-tie text-white text-4xl';
            }
            
            // 点击开始使用按钮
            if (startBtn) {
                startBtn.addEventListener('click', function() {
                    // 标记为已访问
                    sessionStorage.setItem('hasVisited', 'true');
                    
                    // 显示导航栏
                    updateNavigationVisibility(true);
                    
                    // 展开导航栏
                    const navContainer = document.getElementById('leftNav');
                    if (navContainer && navContainer.classList.contains('collapsed')) {
                        toggleNav();
                    }
                    
                    // 隐藏弹窗动画
                    if (authorModal) {
                        authorModal.style.transition = 'opacity 0.3s ease';
                        authorModal.style.opacity = '0';
                        setTimeout(() => {
                            authorModal.style.display = 'none';
                        }, 300);
                    }
                    
                    // 显示欢迎通知
                    showNotification('欢迎使用', '欢迎使用负荷曲线分析工具！', 'success');
                });
            }
            
            // ESC键关闭弹窗（仅在已验证密码后生效）
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && authorModal && authorModal.style.display === 'flex' && authorSection && !authorSection.classList.contains('hidden')) {
                    sessionStorage.setItem('hasVisited', 'true');
                    updateNavigationVisibility(true);
                    authorModal.style.display = 'none';
                }
            });
            
            // 点击背景关闭弹窗（仅在已验证密码后生效）
            if (authorModal) {
                authorModal.addEventListener('click', function(e) {
                    if (e.target === authorModal && authorSection && !authorSection.classList.contains('hidden')) {
                        sessionStorage.setItem('hasVisited', 'true');
                        updateNavigationVisibility(true);
                        authorModal.style.display = 'none';
                    }
                });
            }
        });

        // 移除清除验证缓存的功能，因为不再需要
        // 注释掉原有的清除缓存功能
        /*
        function clearAccessCache() {
            localStorage.removeItem('invitationCode');
            location.reload();
        }

        // 添加一个隐藏的管理功能（按Ctrl+Shift+A）
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && e.shiftKey && e.key === 'A') {
                clearAccessCache();
            }
        });
        */
