/**
 * 数据模板配置
 * 用于生成模拟数据和自动填充配置
 */
const DATA_TEMPLATES = [
    {
        id: '96point_wide',
        name: '标准96点宽表 (15分钟间隔)',
        description: '最常见的负荷数据格式，一行代表一天，包含96列功率数据。',
        config: {
            dataStructure: 'columnToRow',
            dataType: 'instantPower',
            dateColumn: '日期',
            dataStartColumn: '00:15',
            dataEndColumn: '00:00',
            timeInterval: 15,
            useMultiplierColumn: false,
            multiplier: 1.0
        },
        generateMockData: () => {
            const headers = ['日期', '用户编号', '用户名称'];
            for (let i = 1; i <= 96; i++) {
                const hour = Math.floor(((i * 15) % (24 * 60)) / 60);
                const min = (i * 15) % 60;
                const timeStr = `${hour.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}`;
                headers.push(timeStr === '00:00' && i === 96 ? '00:00' : timeStr);
            }
            
            const data = [];
            const startDate = new Date('2024-01-01');
            for (let d = 0; d < 3; d++) {
                const currentDate = new Date(startDate);
                currentDate.setDate(startDate.getDate() + d);
                const dateStr = currentDate.toISOString().split('T')[0];
                const row = [dateStr, 'USER001', '示例用户A'];
                for (let i = 0; i < 96; i++) {
                    row.push((Math.random() * 100 + 50).toFixed(2));
                }
                data.push(row);
            }
            return { headers, data, sheetName: '96点数据示例' };
        }
    },
    {
        id: '24point_wide',
        name: '标准24点宽表 (1小时间隔)',
        description: '按小时记录的负荷数据，一行一天，包含24列功率数据。',
        config: {
            dataStructure: 'columnToRow',
            dataType: 'instantPower',
            dateColumn: '日期',
            dataStartColumn: '1时',
            dataEndColumn: '24时',
            timeInterval: 60,
            useMultiplierColumn: false,
            multiplier: 1.0
        },
        generateMockData: () => {
            const headers = ['日期', '计量点', '单位'];
            for (let i = 1; i <= 24; i++) {
                headers.push(`${i}时`);
            }
            
            const data = [];
            const startDate = new Date('2024-01-01');
            for (let d = 0; d < 3; d++) {
                const currentDate = new Date(startDate);
                currentDate.setDate(startDate.getDate() + d);
                const dateStr = currentDate.toISOString().split('T')[0];
                const row = [dateStr, 'MP001', 'kW'];
                for (let i = 0; i < 24; i++) {
                    row.push((Math.random() * 200 + 100).toFixed(2));
                }
                data.push(row);
            }
            return { headers, data, sheetName: '24点数据示例' };
        }
    },
    {
        id: 'long_table',
        name: '流水式窄表 (时间+数值)',
        description: '每行记录一个时间点的数值，适用于SCADA系统导出的原始流水账。',
        config: {
            dataStructure: 'columnToColumn',
            dataType: 'instantPower',
            dateColumn: '采集日期',
            timeColumn: '采集时间',
            dataStartColumn: '功率值(kW)',
            timeInterval: 15,
            useMultiplierColumn: true,
            multiplierColumn: '倍率'
        },
        generateMockData: () => {
            const headers = ['采集日期', '采集时间', '功率值(kW)', '倍率', '设备ID'];
            const data = [];
            const startDate = new Date('2024-01-01');
            for (let d = 0; d < 1; d++) {
                const currentDate = new Date(startDate);
                currentDate.setDate(startDate.getDate() + d);
                const dateStr = currentDate.toISOString().split('T')[0];
                for (let i = 0; i < 96; i++) {
                    const hour = Math.floor((i * 15) / 60);
                    const min = (i * 15) % 60;
                    const timeStr = `${hour.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}:00`;
                    data.push([dateStr, timeStr, (Math.random() * 50 + 20).toFixed(3), '40', 'DEV_001']);
                }
            }
            return { headers, data, sheetName: '窄表数据示例' };
        }
    },
    {
        id: 'multi_mp_wide',
        name: '多计量点96点宽表',
        description: '包含多个计量点的宽表数据，系统将自动识别并提供筛选功能。',
        config: {
            dataStructure: 'columnToRow',
            dataType: 'instantPower',
            dateColumn: '日期',
            meteringPointColumn: '计量点编号',
            dataStartColumn: '00:15',
            dataEndColumn: '00:00',
            timeInterval: 15,
            useMultiplierColumn: false,
            multiplier: 1.0
        },
        generateMockData: () => {
            const headers = ['日期', '计量点编号', '计量点名称'];
            for (let i = 1; i <= 96; i++) {
                const hour = Math.floor(((i * 15) % (24 * 60)) / 60);
                const min = (i * 15) % 60;
                const timeStr = `${hour.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}`;
                headers.push(timeStr === '00:00' && i === 96 ? '00:00' : timeStr);
            }
            
            const data = [];
            const startDate = new Date('2024-01-01');
            const mps = [
                { id: 'MP_001', name: '车间A总表' },
                { id: 'MP_002', name: '车 workshop B' },
                { id: 'MP_003', name: '办公楼' }
            ];
            
            for (const mp of mps) {
                for (let d = 0; d < 2; d++) {
                    const currentDate = new Date(startDate);
                    currentDate.setDate(startDate.getDate() + d);
                    const dateStr = currentDate.toISOString().split('T')[0];
                    const row = [dateStr, mp.id, mp.name];
                    for (let i = 0; i < 96; i++) {
                        row.push((Math.random() * 80 + 20).toFixed(2));
                    }
                    data.push(row);
                }
            }
            return { headers, data, sheetName: '多计量点示例' };
        }
    }
];
