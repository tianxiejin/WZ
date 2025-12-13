/**
 * 轴承供应链成本核算模型系统 - 数据加载模块
 * 使用 PapaParse 解析 CSV 文件，提供统一的数据访问接口
 */

const DataLoader = {
    // 缓存已加载的数据
    cache: {},
    
    // 基础路径（自动检测）
    basePath: '',

    /**
     * 初始化基础路径
     */
    init() {
        // 获取当前页面的路径，确定CSV文件的相对位置
        const scripts = document.getElementsByTagName('script');
        for (let script of scripts) {
            if (script.src && script.src.includes('data-loader.js')) {
                // data-loader.js 在 js/ 目录下，CSV在上级目录
                const jsPath = script.src.substring(0, script.src.lastIndexOf('/'));
                this.basePath = jsPath.substring(0, jsPath.lastIndexOf('/') + 1);
                break;
            }
        }
        // 如果未找到，使用当前页面的基础路径
        if (!this.basePath) {
            this.basePath = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
        }
        console.log('DataLoader basePath:', this.basePath);
    },

    /**
     * 加载CSV文件并解析
     * @param {string} filename - CSV文件名
     * @returns {Promise<Array>} 解析后的数据数组
     */
    async loadCSV(filename) {
        // 初始化基础路径（仅一次）
        if (!this.basePath) {
            this.init();
        }
        
        // 检查缓存
        if (this.cache[filename]) {
            return this.cache[filename];
        }

        // 构建完整URL
        const fullUrl = this.basePath + filename;
        console.log('Loading CSV:', fullUrl);

        try {
            const response = await fetch(fullUrl);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            
            return new Promise((resolve, reject) => {
                Papa.parse(csvText, {
                    header: true,
                    skipEmptyLines: 'greedy',  // 更严格地跳过空行
                    dynamicTyping: true,
                    complete: (results) => {
                        // 额外过滤：移除所有字段都为空的记录
                        const cleanedData = results.data.filter(row => {
                            if (!row) return false;
                            const values = Object.values(row);
                            return values.some(v => v != null && String(v).trim() !== '');
                        });
                        this.cache[filename] = cleanedData;
                        resolve(cleanedData);
                    },
                    error: (error) => {
                        reject(error);
                    }
                });
            });
        } catch (error) {
            console.error(`加载 ${filename} 失败:`, error);
            throw error;
        }
    },

    /**
     * 加载轴承产品数据
     */
    async loadProducts() {
        return this.loadCSV('bearing_products.csv');
    },

    /**
     * 加载原材料数据
     */
    async loadMaterials() {
        return this.loadCSV('raw_materials.csv');
    },

    /**
     * 加载供应商数据
     */
    async loadSuppliers() {
        return this.loadCSV('suppliers.csv');
    },

    /**
     * 加载作业活动数据
     */
    async loadActivities() {
        return this.loadCSV('activities.csv');
    },

    /**
     * 加载生产工序数据
     */
    async loadProcesses() {
        return this.loadCSV('production_processes.csv');
    },

    /**
     * 加载生产订单数据
     */
    async loadOrders() {
        return this.loadCSV('sample_production_orders.csv');
    },

    /**
     * 加载成本分摊费率数据
     */
    async loadCostRates() {
        return this.loadCSV('cost_allocation_rates.csv');
    },

    /**
     * 加载材料消耗明细数据
     */
    async loadMaterialConsumption() {
        return this.loadCSV('material_consumption_details.csv');
    },

    /**
     * 加载工序消耗明细数据
     */
    async loadProcessConsumption() {
        return this.loadCSV('process_consumption_details.csv');
    },

    /**
     * 加载所有数据
     */
    async loadAll() {
        const [products, materials, suppliers, activities, processes, orders, costRates, materialConsumption, processConsumption] = await Promise.all([
            this.loadProducts(),
            this.loadMaterials(),
            this.loadSuppliers(),
            this.loadActivities(),
            this.loadProcesses(),
            this.loadOrders(),
            this.loadCostRates(),
            this.loadMaterialConsumption(),
            this.loadProcessConsumption()
        ]);

        return {
            products,
            materials,
            suppliers,
            activities,
            processes,
            orders,
            costRates,
            materialConsumption,
            processConsumption
        };
    },

    /**
     * 清除缓存
     */
    clearCache() {
        this.cache = {};
    }
};

/**
 * 数据处理工具函数
 */
const DataUtils = {
    /**
     * 格式化货币
     */
    formatCurrency(value, decimals = 2) {
        if (value === null || value === undefined) return '¥0.00';
        return '¥' + Number(value).toLocaleString('zh-CN', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
    },

    /**
     * 格式化百分比
     */
    formatPercent(value, decimals = 1) {
        if (value === null || value === undefined) return '0%';
        return Number(value).toFixed(decimals) + '%';
    },

    /**
     * 格式化数字
     */
    formatNumber(value, decimals = 0) {
        if (value === null || value === undefined) return '0';
        return Number(value).toLocaleString('zh-CN', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
    },

    /**
     * 计算毛利率
     */
    calculateMargin(costPrice, salesPrice) {
        if (!salesPrice || salesPrice === 0) return 0;
        return ((salesPrice - costPrice) / salesPrice) * 100;
    },

    /**
     * 按字段分组
     */
    groupBy(array, key) {
        return array.reduce((result, item) => {
            const groupKey = item[key];
            if (!result[groupKey]) {
                result[groupKey] = [];
            }
            result[groupKey].push(item);
            return result;
        }, {});
    },

    /**
     * 计算数组某字段的总和
     */
    sum(array, key) {
        return array.reduce((total, item) => total + (Number(item[key]) || 0), 0);
    },

    /**
     * 计算数组某字段的平均值
     */
    average(array, key) {
        if (array.length === 0) return 0;
        return this.sum(array, key) / array.length;
    },

    /**
     * 获取唯一值
     */
    unique(array, key) {
        return [...new Set(array.map(item => item[key]))];
    },

    /**
     * 显示加载提示
     */
    showLoading(container) {
        if (typeof container === 'string') {
            container = document.querySelector(container);
        }
        if (container) {
            container.innerHTML = '<div class="loading">数据加载中...</div>';
        }
    },

    /**
     * 显示错误提示
     */
    showError(container, message) {
        if (typeof container === 'string') {
            container = document.querySelector(container);
        }
        if (container) {
            container.innerHTML = `<div class="error">加载失败: ${message}</div>`;
        }
    },

    /**
     * 显示成功提示
     */
    showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        toast.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${type === 'success' ? '#4CAF50' : '#f44336'};
            color: white;
            padding: 15px 20px;
            border-radius: 5px;
            z-index: 1000;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            animation: fadeIn 0.3s ease;
        `;
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'fadeOut 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
};

// 添加CSS动画
const style = document.createElement('style');
style.textContent = `
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes fadeOut {
        from { opacity: 1; transform: translateY(0); }
        to { opacity: 0; transform: translateY(-10px); }
    }
    .loading {
        text-align: center;
        padding: 40px;
        color: #666;
        font-size: 16px;
    }
    .loading::before {
        content: '';
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid #1e5799;
        border-top-color: transparent;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin-right: 10px;
        vertical-align: middle;
    }
    @keyframes spin {
        to { transform: rotate(360deg); }
    }
    .error {
        text-align: center;
        padding: 40px;
        color: #e74c3c;
        font-size: 16px;
    }
`;
document.head.appendChild(style);
