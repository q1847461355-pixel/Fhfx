/**
 * 应用状态管理模块
 * 集中管理全局状态和配置
 */

const AppState = (function() {
    const state = {
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
            multiplierMode: 'single',
            ptColumn: '',
            ctColumn: '',
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
            barChartAggregation: 'daily',
            allCurvesMode: false,
            summaryMode: false,
            summarySubMode: 'all'
        },
        chart: null,
        dailyTotalBarChart: null,
        meteringPoints: [],
        memoryManagement: {
            maxDataRows: 100000,
            maxTotalMemoryMB: 200,
            currentMemoryUsageMB: 0,
            enableGarbageCollection: true
        }
    };

    const lastChartUpdateParams = {
        dates: null,
        meteringPoints: null,
        dataLength: 0,
        startTime: null,
        endTime: null,
        mode: null
    };

    const subscribers = new Map();
    let stateId = 0;

    function getState() {
        return state;
    }

    function setState(path, value) {
        const keys = path.split('.');
        let current = state;
        
        for (let i = 0; i < keys.length - 1; i++) {
            current = current[keys[i]];
        }
        
        const lastKey = keys[keys.length - 1];
        const oldValue = current[lastKey];
        current[lastKey] = value;
        
        notifySubscribers(path, value, oldValue);
    }

    function updateConfig(newConfig) {
        Object.assign(state.config, newConfig);
        notifySubscribers('config', state.config);
    }

    function updateVisualization(newVis) {
        Object.assign(state.visualization, newVis);
        notifySubscribers('visualization', state.visualization);
    }

    function subscribe(path, callback) {
        const id = ++stateId;
        if (!subscribers.has(path)) {
            subscribers.set(path, new Map());
        }
        subscribers.get(path).set(id, callback);
        
        return () => {
            subscribers.get(path)?.delete(id);
        };
    }

    function notifySubscribers(path, newValue, oldValue) {
        const pathSubscribers = subscribers.get(path);
        if (pathSubscribers) {
            pathSubscribers.forEach(callback => {
                try {
                    callback(newValue, oldValue);
                } catch (e) {
                    console.error('Subscriber error:', e);
                }
            });
        }
        
        const allSubscribers = subscribers.get('*');
        if (allSubscribers) {
            allSubscribers.forEach(callback => {
                try {
                    callback(path, newValue, oldValue);
                } catch (e) {
                    console.error('Subscriber error:', e);
                }
            });
        }
    }

    function resetState() {
        state.files = [];
        state.worksheets = [];
        state.parsedData = [];
        state.processedData = [];
        state.originalData = null;
        state.chart = null;
        state.dailyTotalBarChart = null;
        state.meteringPoints = [];
        notifySubscribers('reset', state);
    }

    function getLastChartParams() {
        return lastChartUpdateParams;
    }

    function updateLastChartParams(params) {
        Object.assign(lastChartUpdateParams, params);
    }

    function getMemoryUsage() {
        if (performance.memory) {
            return Math.round(performance.memory.usedJSHeapSize / 1024 / 1024);
        }
        return state.memoryManagement.currentMemoryUsageMB;
    }

    function checkMemoryLimit() {
        const usage = getMemoryUsage();
        if (usage > state.memoryManagement.maxTotalMemoryMB) {
            console.warn(`Memory usage (${usage}MB) exceeds limit (${state.memoryManagement.maxTotalMemoryMB}MB)`);
            return false;
        }
        return true;
    }

    return {
        getState,
        setState,
        updateConfig,
        updateVisualization,
        subscribe,
        resetState,
        getLastChartParams,
        updateLastChartParams,
        getMemoryUsage,
        checkMemoryLimit
    };
})();

window.AppState = AppState;

export default AppState;
