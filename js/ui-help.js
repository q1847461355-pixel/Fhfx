        document.addEventListener('DOMContentLoaded', function() {
            const helpBtn = document.getElementById('helpBtn');
            const helpPanel = document.getElementById('helpPanel');
            const closeHelpBtn = document.getElementById('closeHelpBtn');
            const aboutBtn = document.getElementById('aboutBtn');

            // 打开使用说明面板
            if (helpBtn) {
                helpBtn.addEventListener('click', function() {
                    if (helpPanel) {
                        helpPanel.classList.remove('hidden');
                        helpPanel.classList.add('flex');
                        setTimeout(() => {
                            const content = document.getElementById('helpModalContent');
                            if (content) {
                                content.classList.remove('scale-95', 'opacity-0');
                                content.classList.add('scale-100', 'opacity-100');
                            }
                        }, 10);
                        document.body.style.overflow = 'hidden'; // 防止背景滚动
                    }
                });
            }

            // 关闭使用说明面板
            function closeHelpPanel() {
                if (helpPanel) {
                    const content = document.getElementById('helpModalContent');
                    if (content) {
                        content.classList.remove('scale-100', 'opacity-100');
                        content.classList.add('scale-95', 'opacity-0');
                    }
                    setTimeout(() => {
                        helpPanel.classList.add('hidden');
                        helpPanel.classList.remove('flex');
                        document.body.style.overflow = ''; // 恢复背景滚动
                    }, 300);
                }
            }

            if (closeHelpBtn) {
                closeHelpBtn.addEventListener('click', closeHelpPanel);
            }
            
            // 点击背景关闭面板
            if (helpPanel) {
                helpPanel.addEventListener('click', function(e) {
                    if (e.target === helpPanel) {
                        closeHelpPanel();
                    }
                });
            }

            // ESC键关闭面板
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && helpPanel && !helpPanel.classList.contains('hidden')) {
                    closeHelpPanel();
                }
            });

            // 关于按钮点击事件
            if (aboutBtn) {
                aboutBtn.addEventListener('click', function() {
                    showAboutModal();
                });
            }

            // 添加键盘快捷键
            document.addEventListener('keydown', function(e) {
                // Ctrl+Shift+H 打开帮助
                if (e.ctrlKey && e.shiftKey && e.key === 'H') {
                    e.preventDefault();
                    if (helpPanel) {
                        helpPanel.classList.remove('hidden');
                        helpPanel.classList.add('flex');
                        setTimeout(() => {
                            const content = document.getElementById('helpModalContent');
                            if (content) {
                                content.classList.remove('scale-95', 'opacity-0');
                                content.classList.add('scale-100', 'opacity-100');
                            }
                        }, 10);
                        document.body.style.overflow = 'hidden';
                    }
                }
            });

            // 添加平滑滚动效果
            if (helpPanel) {
                const helpContent = helpPanel.querySelector('.max-h-screen');
                if (helpContent) {
                    helpContent.addEventListener('wheel', function(e) {
                        e.stopPropagation();
                    });
                }
            }
        });
