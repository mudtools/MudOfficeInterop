namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定查询参数的数据类型
/// </summary>
public enum XlParameterDataType
{
    /// <summary>
    /// 未知类型
    /// </summary>
    xlParamTypeUnknown = 0,

    /// <summary>
    /// 字符串
    /// </summary>
    xlParamTypeChar = 1,

    /// <summary>
    /// 数值
    /// </summary>
    xlParamTypeNumeric = 2,

    /// <summary>
    /// 小数
    /// </summary>
    xlParamTypeDecimal = 3,

    /// <summary>
    /// 整数
    /// </summary>
    xlParamTypeInteger = 4,

    /// <summary>
    /// 小整数
    /// </summary>
    xlParamTypeSmallInt = 5,

    /// <summary>
    /// 浮点数
    /// </summary>
    xlParamTypeFloat = 6,

    /// <summary>
    /// 实数
    /// </summary>
    xlParamTypeReal = 7,

    /// <summary>
    /// 双精度浮点数
    /// </summary>
    xlParamTypeDouble = 8,

    /// <summary>
    /// 可变长度字符串
    /// </summary>
    xlParamTypeVarChar = 12,

    /// <summary>
    /// 日期
    /// </summary>
    xlParamTypeDate = 9,

    /// <summary>
    /// 时间
    /// </summary>
    xlParamTypeTime = 10,

    /// <summary>
    /// 时间戳
    /// </summary>
    xlParamTypeTimestamp = 11,

    /// <summary>
    /// 长字符串
    /// </summary>
    xlParamTypeLongVarChar = -1,

    /// <summary>
    /// 二进制
    /// </summary>
    xlParamTypeBinary = -2,

    /// <summary>
    /// 可变长度二进制
    /// </summary>
    xlParamTypeVarBinary = -3,

    /// <summary>
    /// 长二进制
    /// </summary>
    xlParamTypeLongVarBinary = -4,

    /// <summary>
    /// 大整数
    /// </summary>
    xlParamTypeBigInt = -5,

    /// <summary>
    /// 微小整数
    /// </summary>
    xlParamTypeTinyInt = -6,

    /// <summary>
    /// 位
    /// </summary>
    xlParamTypeBit = -7,

    /// <summary>
    /// Unicode 字符串
    /// </summary>
    xlParamTypeWChar = -8
}