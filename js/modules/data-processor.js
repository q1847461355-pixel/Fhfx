/**
 * 数据处理模块
 * 负责Excel文件解析、数据转换和处理
 */

const DataProcessor = (function() {
    const DATE_FORMATS = [
        'YYYY-MM-DD',
        'YYYY/MM/DD',
        'MM-DD-YYYY',
        'MM/DD/YYYY',
        'DD-MM-YYYY',
        'DD/MM/YYYY',
        'YYYYMMDD',
        'YY-MM-DD',
        'YY/MM/DD'
    ];

    function parseDate(dateStr, format = 'YYYY-MM-DD') {
        if (!dateStr) return null;
        
        if (dateStr instanceof Date) {
            return isNaN(dateStr.getTime()) ? null : dateStr;
        }

        const str = String(dateStr).trim();
        
        const excelDate = parseFloat(str);
        if (!isNaN(excelDate) && excelDate > 0 && excelDate < 100000) {
            const date = new Date((excelDate - 25569) * 86400 * 1000);
            return isNaN(date.getTime()) ? null : date;
        }

        const patterns = {
            'YYYY-MM-DD': /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/,
            'YYYY/MM/DD': /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/,
            'MM-DD-YYYY': /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/,
            'MM/DD/YYYY': /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/,
            'DD-MM-YYYY': /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/,
            'DD/MM/YYYY': /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/,
            'YYYYMMDD': /^(\d{4})(\d{2})(\d{2})$/,
            'YY-MM-DD': /^(\d{2})[-/](\d{1,2})[-/](\d{1,2})$/,
        };

        const pattern = patterns[format];
        if (!pattern) {
            const defaultMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
            if (defaultMatch) {
                return new Date(parseInt(defaultMatch[1]), parseInt(defaultMatch[2]) - 1, parseInt(defaultMatch[3]));
            }
            return null;
        }

        const match = str.match(pattern);
        if (!match) return null;

        let year, month, day;
        
        if (format.startsWith('YYYY') || format === 'YYYYMMDD') {
            year = parseInt(match[1]);
            month = parseInt(match[2]) - 1;
            day = parseInt(match[3]);
        } else if (format.startsWith('MM')) {
            month = parseInt(match[1]) - 1;
            day = parseInt(match[2]);
            year = parseInt(match[3]);
        } else if (format.startsWith('DD')) {
            day = parseInt(match[1]);
            month = parseInt(match[2]) - 1;
            year = parseInt(match[3]);
        } else if (format.startsWith('YY')) {
            year = 2000 + parseInt(match[1]);
            month = parseInt(match[2]) - 1;
            day = parseInt(match[3]);
        }

        const date = new Date(year, month, day);
        return isNaN(date.getTime()) ? null : date;
    }

    function detectDateFormat(dates) {
        for (const format of DATE_FORMATS) {
            let validCount = 0;
            for (const dateStr of dates.slice(0, 20)) {
                if (parseDate(dateStr, format)) {
                    validCount++;
                }
            }
            if (validCount > dates.slice(0, 20).length * 0.8) {
                return format;
            }
        }
        return 'YYYY-MM-DD';
    }

    function parseNumericValue(value) {
        if (value === null || value === undefined || value === '') {
            return null;
        }
        
        const num = parseFloat(String(value).replace(/,/g, ''));
        return isNaN(num) ? null : num;
    }

    function processWideFormat(data, config) {
        const result = [];
        const dateColumn = config.dateColumn;
        const startCol = config.dataStartIndex;
        const endCol = config.dataEndIndex;
        const multiplier = config.multiplier;
        const invalidHandling = config.invalidDataHandling;

        for (const row of data) {
            const dateStr = row[dateColumn];
            const date = parseDate(dateStr, config.dateFormat);
            
            if (!date) continue;

            const hourlyData = new Array(24).fill(null);
            const timeInterval = config.timeInterval;
            const pointsPerHour = 60 / timeInterval;
            
            for (let i = startCol; i <= endCol && i < row.length; i++) {
                const pointIndex = i - startCol;
                const hour = Math.floor(pointIndex / pointsPerHour);
                
                if (hour >= 0 && hour < 24) {
                    let value = parseNumericValue(row[i]);
                    
                    if (value !== null) {
                        value *= multiplier;
                        
                        if (invalidHandling === 'ignore' && (value === 0 || value < 0)) {
                            value = null;
                        }
                    }
                    
                    if (hourlyData[hour] === null) {
                        hourlyData[hour] = 0;
                    }
                    hourlyData[hour] += value || 0;
                }
            }

            result.push({
                date: dateStr,
                dateObj: date,
                hourlyData,
                meteringPoint: row[config.meteringPointColumn] || '默认'
            });
        }

        return result;
    }

    function processLongFormat(data, config) {
        const result = new Map();
        const dateColumn = config.dateColumn;
        const timeColumn = config.timeColumn;
        const valueColumn = config.dataStartColumn;
        const multiplier = config.multiplier;

        for (const row of data) {
            const dateStr = row[dateColumn];
            const timeStr = row[timeColumn];
            const value = parseNumericValue(row[valueColumn]);

            if (!dateStr || !timeStr || value === null) continue;

            const date = parseDate(dateStr, config.dateFormat);
            if (!date) continue;

            const hourMatch = timeStr.match(/(\d{1,2})/);
            const hour = hourMatch ? parseInt(hourMatch[1]) : null;
            
            if (hour === null || hour < 0 || hour > 23) continue;

            const key = `${dateStr}_${row[config.meteringPointColumn] || '默认'}`;
            
            if (!result.has(key)) {
                result.set(key, {
                    date: dateStr,
                    dateObj: date,
                    hourlyData: new Array(24).fill(null),
                    meteringPoint: row[config.meteringPointColumn] || '默认'
                });
            }

            const entry = result.get(key);
            entry.hourlyData[hour] = value * multiplier;
        }

        return Array.from(result.values());
    }

    function fillMissingData(hourlyData, method = 'linear') {
        const result = [...hourlyData];
        const nullIndices = [];
        
        for (let i = 0; i < result.length; i++) {
            if (result[i] === null) {
                nullIndices.push(i);
            }
        }

        if (nullIndices.length === 0) return result;

        if (method === 'linear') {
            for (const idx of nullIndices) {
                let prevIdx = idx - 1;
                let nextIdx = idx + 1;
                
                while (prevIdx >= 0 && result[prevIdx] === null) prevIdx--;
                while (nextIdx < result.length && result[nextIdx] === null) nextIdx++;
                
                if (prevIdx >= 0 && nextIdx < result.length) {
                    const prevVal = result[prevIdx];
                    const nextVal = result[nextIdx];
                    const ratio = (idx - prevIdx) / (nextIdx - prevIdx);
                    result[idx] = prevVal + (nextVal - prevVal) * ratio;
                } else if (prevIdx >= 0) {
                    result[idx] = result[prevIdx];
                } else if (nextIdx < result.length) {
                    result[idx] = result[nextIdx];
                } else {
                    result[idx] = 0;
                }
            }
        } else if (method === 'zero') {
            for (const idx of nullIndices) {
                result[idx] = 0;
            }
        }

        return result;
    }

    function calculateStatistics(data) {
        if (!data || data.length === 0) {
            return null;
        }

        let totalEnergy = 0;
        let maxLoad = -Infinity;
        let minLoad = Infinity;
        let maxLoadTime = null;
        let minLoadTime = null;
        const hourlySums = new Array(24).fill(0);
        const hourlyCounts = new Array(24).fill(0);

        for (const day of data) {
            for (let h = 0; h < 24; h++) {
                const val = day.hourlyData[h];
                if (val !== null && !isNaN(val)) {
                    totalEnergy += val;
                    hourlySums[h] += val;
                    hourlyCounts[h]++;
                    
                    if (val > maxLoad) {
                        maxLoad = val;
                        maxLoadTime = h;
                    }
                    if (val < minLoad) {
                        minLoad = val;
                        minLoadTime = h;
                    }
                }
            }
        }

        const avgLoad = totalEnergy / (data.length * 24);
        const loadFactor = maxLoad > 0 ? (avgLoad / maxLoad) * 100 : 0;

        return {
            totalEnergy,
            avgLoad,
            maxLoad,
            minLoad,
            maxLoadTime,
            minLoadTime,
            loadFactor,
            hourlyAverages: hourlySums.map((sum, i) => hourlyCounts[i] > 0 ? sum / hourlyCounts[i] : 0),
            dayCount: data.length
        };
    }

    function aggregateByMonth(data) {
        const monthly = new Map();
        
        for (const day of data) {
            const date = day.dateObj || parseDate(day.date);
            if (!date) continue;
            
            const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            
            if (!monthly.has(monthKey)) {
                monthly.set(monthKey, {
                    month: monthKey,
                    totalEnergy: 0,
                    dayCount: 0,
                    maxLoad: -Infinity,
                    minLoad: Infinity
                });
            }
            
            const entry = monthly.get(monthKey);
            let dayTotal = 0;
            
            for (let h = 0; h < 24; h++) {
                const val = day.hourlyData[h];
                if (val !== null) {
                    dayTotal += val;
                    if (val > entry.maxLoad) entry.maxLoad = val;
                    if (val < entry.minLoad) entry.minLoad = val;
                }
            }
            
            entry.totalEnergy += dayTotal;
            entry.dayCount++;
        }

        return Array.from(monthly.values()).sort((a, b) => a.month.localeCompare(b.month));
    }

    return {
        parseDate,
        detectDateFormat,
        parseNumericValue,
        processWideFormat,
        processLongFormat,
        fillMissingData,
        calculateStatistics,
        aggregateByMonth,
        DATE_FORMATS
    };
})();

window.DataProcessor = DataProcessor;

export default DataProcessor;
