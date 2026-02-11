        // ==================== 调试工具配置 ====================
        // 调试模式开关，设置为true可启用详细日志输出
        const DEBUG_MODE = true;

        // 页面初始化状态标志（全局共享）
        window.isPageInitializing = true;

        /**
         * 异步分片处理大数组，避免阻塞主线程
         * @param {Array} items - 要处理的数组
         * @param {Function} processFn - 处理每个元素的函数 (item, index) => void
         * @param {Object} options - 配置项 { chunkSize: 100, onProgress: (processed, total) => void }
         * @returns {Promise}
         */
        function processChunksAsync(items, processFn, options = {}) {
            return new Promise((resolve, reject) => {
                const { chunkSize = 500, onProgress } = options;
                const total = items.length;
                let index = 0;

                function nextChunk() {
                    try {
                        const end = Math.min(index + chunkSize, total);
                        for (; index < end; index++) {
                            processFn(items[index], index);
                        }

                        if (onProgress) {
                            onProgress(index, total);
                        }

                        if (index < total) {
                            // 使用 setTimeout(..., 0) 让出主线程，允许 UI 更新
                            setTimeout(nextChunk, 0);
                        } else {
                            resolve();
                        }
                    } catch (err) {
                        reject(err);
                    }
                }

                nextChunk();
            });
        }

        // 全局 Loading 控制
        function showGlobalLoading(message = '处理中...') {
            let loader = document.getElementById('globalLoadingOverlay');
            if (!loader) {
                loader = document.createElement('div');
                loader.id = 'globalLoadingOverlay';
                loader.className = 'fixed inset-0 bg-gray-900 bg-opacity-50 z-[9999] flex items-center justify-center transition-opacity duration-300';
                loader.innerHTML = `
                    <div class="bg-white rounded-xl shadow-2xl p-8 max-w-sm w-full text-center transform transition-all scale-100">
                        <div class="inline-block animate-spin rounded-full h-12 w-12 border-4 border-indigo-100 border-t-indigo-600 mb-4"></div>
                        <h3 class="text-lg font-bold text-gray-800 mb-2" id="globalLoadingTitle">处理中</h3>
                        <p class="text-gray-500 text-sm mb-4" id="globalLoadingMessage">正在分析数据...</p>
                        <div class="w-full bg-gray-200 rounded-full h-2.5 dark:bg-gray-700 overflow-hidden">
                            <div id="globalLoadingBar" class="bg-indigo-600 h-2.5 rounded-full transition-all duration-300" style="width: 0%"></div>
                        </div>
                        <p class="text-xs text-gray-400 mt-2 text-right" id="globalLoadingPercent">0%</p>
                    </div>
                `;
                document.body.appendChild(loader);
            }
            
            document.getElementById('globalLoadingTitle').textContent = message;
            document.getElementById('globalLoadingMessage').textContent = '请稍候...';
            document.getElementById('globalLoadingBar').style.width = '0%';
            document.getElementById('globalLoadingPercent').textContent = '0%';
            loader.classList.remove('hidden', 'opacity-0', 'pointer-events-none');
            return loader;
        }

        function updateGlobalLoading(percent, message) {
            const loader = document.getElementById('globalLoadingOverlay');
            if (!loader || loader.classList.contains('hidden')) return;
            
            if (message) document.getElementById('globalLoadingMessage').textContent = message;
            const safePercent = Math.min(100, Math.max(0, percent));
            document.getElementById('globalLoadingBar').style.width = `${safePercent}%`;
            document.getElementById('globalLoadingPercent').textContent = `${Math.round(safePercent)}%`;
        }

        function hideGlobalLoading() {
            const loader = document.getElementById('globalLoadingOverlay');
            if (loader) {
                loader.classList.add('opacity-0', 'pointer-events-none');
                setTimeout(() => {
                    loader.classList.add('hidden');
                }, 300);
            }
        }

        
        // 警告级别配置
        const WARNING_LEVEL = {
            LOW: 1,      // 低级别警告，仅记录
            MEDIUM: 2,   // 中级别警告，记录并显示
            HIGH: 3      // 高级别警告，记录、显示并可能中断操作
        };
        
        // 当前警告级别阈值，高于此级别的警告将被显示
        const CURRENT_WARNING_THRESHOLD = WARNING_LEVEL.MEDIUM;
        
        /**
         * 调试日志函数
         * @param {string} message - 日志消息
         * @param {string} type - 日志类型（log, info, warn, error）
         * @param {any} data - 附加数据
         */
        function debugLog(message, type = 'log', data = null) {
            if (!DEBUG_MODE) return;
            
            const timestamp = new Date().toISOString();
            const prefix = `[DEBUG ${timestamp}]`;
            
            switch(type) {
                case 'info':
                    console.info(`${prefix} ${message}`, data || '');
                    break;
                case 'warn':
                    console.warn(`${prefix} ${message}`, data || '');
                    break;
                case 'error':
                    console.error(`${prefix} ${message}`, data || '');
                    break;
                default:
                    console.log(`${prefix} ${message}`, data || '');
            }
        }
        
        /**
         * 显示警告消息
         * @param {string} message - 警告消息
         * @param {number} level - 警告级别
         * @param {boolean} showToast - 是否显示提示框
         */
        function showWarning(message, level = WARNING_LEVEL.MEDIUM, showToast = true) {
            // 记录警告日志
            debugLog(`WARNING [Level ${level}]: ${message}`, 'warn');
            
            // 如果警告级别低于当前阈值，不显示提示
            if (level < CURRENT_WARNING_THRESHOLD) return;
            
            // 显示警告提示
            if (showToast) {
                const toast = document.createElement('div');
                toast.className = `fixed top-4 right-4 p-4 rounded-2xl shadow-xl z-50 transition-all border-2 backdrop-blur-md ${
                    level >= WARNING_LEVEL.HIGH ? 'bg-white/90 border-indigo-500 text-indigo-900 shadow-indigo-100' : 
                    level >= WARNING_LEVEL.MEDIUM ? 'bg-white/90 border-indigo-400 text-indigo-800 shadow-indigo-50/50' : 
                    'bg-white/90 border-indigo-200 text-indigo-700 shadow-slate-100'
                }`;
                toast.innerHTML = `
                    <div class="flex items-center">
                        <i class="fa ${
                            level >= WARNING_LEVEL.HIGH ? 'fa-exclamation-circle' : 
                            level >= WARNING_LEVEL.MEDIUM ? 'fa-exclamation-triangle' : 
                            'fa-info-circle'
                        } mr-2"></i>
                        <span>${message}</span>
                        <button class="ml-4 text-lg" onclick="this.parentElement.parentElement.remove()" aria-label="关闭通知">×</button>
                    </div>
                `;
                document.body.appendChild(toast);
                
                // 5秒后自动移除提示
                setTimeout(() => {
                    if (toast.parentElement) {
                        toast.remove();
                    }
                }, 5000);
            }
        }
        
        /**
         * 性能监控函数
         * @param {string} label - 性能监控标签
         * @param {Function} fn - 要执行的函数
         * @returns {any} 函数执行结果
         */
        function monitorPerformance(label, fn) {
            if (!DEBUG_MODE) return fn();
            
            const startTime = performance.now();
            debugLog(`开始性能监控: ${label}`, 'info');
            
            try {
                const result = fn();
                const endTime = performance.now();
                const duration = endTime - startTime;
                
                debugLog(`性能监控完成: ${label}, 耗时: ${duration.toFixed(2)}ms`, 'info');
                
                // 如果执行时间超过阈值，显示警告
                if (duration > 1000) {
                    showWarning(`操作耗时较长: ${label} (${duration.toFixed(2)}ms)`, WARNING_LEVEL.MEDIUM);
                }
                
                return result;
            } catch (error) {
                const endTime = performance.now();
                const duration = endTime - startTime;
                debugLog(`性能监控出错: ${label}, 耗时: ${duration.toFixed(2)}ms, 错误: ${error.message}`, 'error', error);
                throw error;
            }
        }
        
        // ==================== Tailwind CSS 配置 ====================
        // Tailwind CSS 自定义配置
        if (typeof tailwind !== 'undefined') {
            tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        // UI/UX Pro Max 推荐：分析仪表盘配色 - 专业蓝色 + 琥珀色高亮
                        primary: {
                            50: '#eff6ff',
                            100: '#dbeafe',
                            200: '#bfdbfe',
                            300: '#93c5fd',
                            400: '#60a5fa',
                            500: '#3b82f6',
                            600: '#2563eb',
                            700: '#1d4ed8',
                            800: '#1e40af', // 主色：专业深蓝
                            900: '#1e3a8a',
                            DEFAULT: '#1e40af'
                        },
                        // 数据可视化配色 - 适合图表的多色方案
                        data: {
                            blue: '#1e40af',
                            cyan: '#0891b2',
                            teal: '#0d9488',
                            green: '#059669',
                            amber: '#f59e0b', // 高亮色
                            orange: '#ea580c',
                            rose: '#e11d48',
                            purple: '#7c3aed'
                        },
                        // 辅助色
                        indigo: {
                            50: '#eef2ff',
                            100: '#e0e7ff',
                            200: '#c7d2fe',
                            300: '#a5b4fc',
                            400: '#818cf8',
                            500: '#6366f1',
                            600: '#4f46e5',
                            700: '#4338ca',
                            800: '#3730a3',
                            900: '#312e81',
                            DEFAULT: '#4f46e5'
                        },
                        emerald: {
                            50: '#ecfdf5',
                            100: '#d1fae5',
                            200: '#a7f3d0',
                            300: '#6ee7b7',
                            400: '#34d399',
                            500: '#10b981',
                            600: '#059669',
                            700: '#047857',
                            800: '#065f46',
                            900: '#064e3b',
                            DEFAULT: '#10b981'
                        },
                        amber: {
                            50: '#fffbeb',
                            100: '#fef3c7',
                            200: '#fde68a',
                            300: '#fcd34d',
                            400: '#fbbf24',
                            500: '#f59e0b', // 强调色：琥珀色
                            600: '#d97706',
                            700: '#b45309',
                            800: '#92400e',
                            900: '#78350f',
                            DEFAULT: '#f59e0b'
                        },
                        slate: {
                            50: '#f8fafc',
                            100: '#f1f5f9',
                            200: '#e2e8f0',
                            300: '#cbd5e1',
                            400: '#94a3b8',
                            500: '#64748b',
                            600: '#475569',
                            700: '#334155',
                            800: '#1e293b',
                            900: '#0f172a',
                            DEFAULT: '#64748b'
                        },
                        secondary: '#6366f1',
                        accent: '#10b981',
                        neutral: '#64748b',
                        success: '#10b981',
                        warning: '#f59e0b',
                        danger: '#ef4444',
                        light: '#f8fafc',
                        dark: '#0f172a'
                    },
                    fontFamily: {
                        sans: ['Fira Sans', 'Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'Helvetica Neue', 'Arial', 'sans-serif'],
                        mono: ['Fira Code', 'JetBrains Mono', 'Consolas', 'Monaco', 'monospace']
                    },
                    fontSize: {
                        '2xs': ['0.65rem', { lineHeight: '0.85rem' }],
                        'xs': ['0.75rem', { lineHeight: '1rem' }],
                        'sm': ['0.875rem', { lineHeight: '1.25rem' }],
                        'base': ['1rem', { lineHeight: '1.5rem' }],
                        'lg': ['1.125rem', { lineHeight: '1.75rem' }],
                        'xl': ['1.25rem', { lineHeight: '1.75rem' }],
                        '2xl': ['1.5rem', { lineHeight: '2rem' }],
                        '3xl': ['1.875rem', { lineHeight: '2.25rem' }],
                        '4xl': ['2.25rem', { lineHeight: '2.5rem' }]
                    },
                    spacing: {
                        '18': '4.5rem',
                        '88': '22rem',
                        '128': '32rem'
                    },
                    borderRadius: {
                        'xl': '0.75rem',
                        '2xl': '1rem',
                        '3xl': '1.5rem',
                        '4xl': '2rem'
                    },
                    boxShadow: {
                        'soft': '0 2px 15px -3px rgba(0, 0, 0, 0.07), 0 10px 20px -2px rgba(0, 0, 0, 0.04)',
                        'medium': '0 4px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
                        'strong': '0 10px 40px -10px rgba(0, 0, 0, 0.15), 0 2px 10px -2px rgba(0, 0, 0, 0.05)',
                        'inner-soft': 'inset 0 2px 4px 0 rgba(0, 0, 0, 0.06)'
                    }
                }
            }
            };
        }
