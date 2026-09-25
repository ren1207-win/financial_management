// 系统配置文件
const CONFIG = {
    // 默认数据源路径
    DATA_SOURCE_PATH: 'data.xlsx',

    // 月度收支数据源路径
    MONTHLY_DATA_PATH: 'monthly_financial_data.csv',

    // 月度投资收益数据源路径
    INVEST_DATA_PATH: 'invest_income_data.csv',

    // 数据源名称
    DATA_SOURCE_NAME: 'data.xlsx',

    // 是否使用固定数据源（true: 使用固定路径，false: 允许用户上传）
    USE_FIXED_DATA_SOURCE: true,

    // 默认展示时间范围（年）
    DEFAULT_YEARS: 1,

    // 最大时间范围（年）
    MAX_YEARS: 5
};

// 导出配置
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CONFIG;
} else {
    window.CONFIG = CONFIG;
}