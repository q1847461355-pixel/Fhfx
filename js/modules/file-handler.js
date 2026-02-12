/**
 * 文件处理模块
 * 负责文件上传、解析和导出
 */

const FileHandler = (function() {
    const SUPPORTED_TYPES = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/csv'
    ];

    const MAX_FILE_SIZE = 50 * 1024 * 1024;

    async function readFile(file) {
        return new Promise((resolve, reject) => {
            if (!validateFile(file)) {
                reject(new Error('文件格式不支持或文件过大'));
                return;
            }

            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    resolve(workbook);
                } catch (error) {
                    reject(new Error('文件解析失败: ' + error.message));
                }
            };

            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    function validateFile(file) {
        if (!file) return false;

        const extension = file.name.split('.').pop().toLowerCase();
        const validExtensions = ['xlsx', 'xls', 'csv'];

        if (!validExtensions.includes(extension)) {
            return false;
        }

        if (file.size > MAX_FILE_SIZE) {
            return false;
        }

        return true;
    }

    function getWorksheets(workbook) {
        return workbook.SheetNames.map(name => ({
            name,
            data: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' })
        }));
    }

    function parseWorksheet(sheetData, options = {}) {
        const {
            headerRow = 0,
            startRow = 1,
            endRow = -1
        } = options;

        if (!sheetData || sheetData.length === 0) {
            return { headers: [], data: [] };
        }

        const headers = sheetData[headerRow] || [];
        const dataEndRow = endRow === -1 ? sheetData.length : endRow;
        const data = sheetData.slice(startRow, dataEndRow);

        return { headers, data };
    }

    async function exportToExcel(data, filename, options = {}) {
        const {
            sheetName = 'Sheet1',
            headers = null
        } = options;

        let worksheetData;
        
        if (headers) {
            worksheetData = [headers, ...data];
        } else {
            worksheetData = data;
        }

        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

        XLSX.writeFile(workbook, filename);
    }

    async function exportToCSV(data, filename, options = {}) {
        const {
            delimiter = ',',
            headers = null
        } = options;

        let csvContent;
        
        if (headers) {
            csvContent = [headers.join(delimiter)];
            data.forEach(row => {
                csvContent.push(row.join(delimiter));
            });
        } else {
            csvContent = data.map(row => row.join(delimiter));
        }

        const blob = new Blob([csvContent.join('\n')], { type: 'text/csv;charset=utf-8;' });
        saveAs(blob, filename);
    }

    async function exportToJSON(data, filename) {
        const jsonStr = JSON.stringify(data, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
        saveAs(blob, filename);
    }

    async function exportReport(reportData, format = 'xlsx') {
        const {
            title = '负荷分析报告',
            summary = {},
            dailyData = [],
            hourlyAverages = [],
            statistics = {}
        } = reportData;

        if (format === 'xlsx') {
            const workbook = XLSX.utils.book_new();

            const summarySheet = [
                ['负荷分析报告'],
                [''],
                ['报告标题', title],
                ['生成时间', new Date().toLocaleString()],
                [''],
                ['统计摘要'],
                ['总用电量', `${summary.totalEnergy?.toFixed(2) || 0} kWh`],
                ['平均负荷', `${summary.avgLoad?.toFixed(2) || 0} kW`],
                ['最大负荷', `${summary.maxLoad?.toFixed(2) || 0} kW`],
                ['最小负荷', `${summary.minLoad?.toFixed(2) || 0} kW`],
                ['负荷率', `${summary.loadFactor?.toFixed(2) || 0}%`],
                ['分析天数', summary.dayCount || 0]
            ];
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(summarySheet), '统计摘要');

            if (dailyData.length > 0) {
                const dailySheet = [
                    ['日期', ...Array.from({ length: 24 }, (_, i) => `${i}:00`), '日用电量'],
                    ...dailyData.map(day => [
                        day.date,
                        ...day.hourlyData.map(v => v?.toFixed(2) || ''),
                        day.hourlyData.reduce((sum, v) => sum + (v || 0), 0).toFixed(2)
                    ])
                ];
                XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(dailySheet), '日负荷数据');
            }

            if (hourlyAverages.length > 0) {
                const hourlySheet = [
                    ['时段', '平均负荷'],
                    ...hourlyAverages.map((val, i) => [`${i}:00-${i + 1}:00`, val.toFixed(2)])
                ];
                XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(hourlySheet), '时段平均');
            }

            XLSX.writeFile(workbook, `${title}.xlsx`);
        } else if (format === 'json') {
            await exportToJSON({
                title,
                generatedAt: new Date().toISOString(),
                summary,
                dailyData,
                hourlyAverages,
                statistics
            }, `${title}.json`);
        }
    }

    function createDragDropZone(element, callbacks = {}) {
        const {
            onDrop,
            onDragOver,
            onDragLeave,
            onError
        } = callbacks;

        element.addEventListener('dragover', (e) => {
            e.preventDefault();
            element.classList.add('drag-over');
            onDragOver?.(e);
        });

        element.addEventListener('dragleave', (e) => {
            e.preventDefault();
            element.classList.remove('drag-over');
            onDragLeave?.(e);
        });

        element.addEventListener('drop', async (e) => {
            e.preventDefault();
            element.classList.remove('drag-over');

            const files = Array.from(e.dataTransfer.files);
            const validFiles = files.filter(validateFile);

            if (validFiles.length === 0) {
                onError?.(new Error('请上传有效的 Excel 或 CSV 文件'));
                return;
            }

            onDrop?.(validFiles);
        });
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    return {
        readFile,
        validateFile,
        getWorksheets,
        parseWorksheet,
        exportToExcel,
        exportToCSV,
        exportToJSON,
        exportReport,
        createDragDropZone,
        formatFileSize,
        SUPPORTED_TYPES,
        MAX_FILE_SIZE
    };
})();

window.FileHandler = FileHandler;

export default FileHandler;
