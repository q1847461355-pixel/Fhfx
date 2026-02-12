/**
 * 模块加载器
 * 负责动态导入和初始化所有模块
 */

const ModuleLoader = (function() {
    const modules = {};
    const moduleStatus = {};
    const moduleDependencies = {
        'cdn-manager': [],
        'state-manager': [],
        'data-processor': [],
        'chart-renderer': ['cdn-manager'],
        'ui-components': [],
        'file-handler': ['cdn-manager']
    };

    async function loadModule(name) {
        if (moduleStatus[name] === 'loaded') {
            return modules[name];
        }

        if (moduleStatus[name] === 'loading') {
            return new Promise((resolve) => {
                const checkLoaded = setInterval(() => {
                    if (moduleStatus[name] === 'loaded') {
                        clearInterval(checkLoaded);
                        resolve(modules[name]);
                    }
                }, 50);
            });
        }

        moduleStatus[name] = 'loading';

        const deps = moduleDependencies[name] || [];
        for (const dep of deps) {
            await loadModule(dep);
        }

        try {
            const modulePath = `./modules/${name}.js`;
            const module = await import(modulePath);
            modules[name] = module.default || module;
            moduleStatus[name] = 'loaded';
            console.log(`[ModuleLoader] ${name} loaded successfully`);
            return modules[name];
        } catch (error) {
            moduleStatus[name] = 'error';
            console.error(`[ModuleLoader] Failed to load ${name}:`, error);
            throw error;
        }
    }

    async function loadModules(names) {
        const results = {};
        for (const name of names) {
            results[name] = await loadModule(name);
        }
        return results;
    }

    async function loadAll() {
        const allModules = Object.keys(moduleDependencies);
        return loadModules(allModules);
    }

    function getModule(name) {
        return modules[name];
    }

    function getStatus(name) {
        return moduleStatus[name];
    }

    function isLoaded(name) {
        return moduleStatus[name] === 'loaded';
    }

    return {
        loadModule,
        loadModules,
        loadAll,
        getModule,
        getStatus,
        isLoaded
    };
})();

window.ModuleLoader = ModuleLoader;

export default ModuleLoader;
