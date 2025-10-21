namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在邮件合并操作中哪个记录处于活动状态的常量
/// </summary>
public enum WdMailMergeActiveRecord
{
    /// <summary>
    /// 无活动记录
    /// </summary>
    wdNoActiveRecord = -1,

    /// <summary>
    /// 结果集中的下一条记录
    /// </summary>
    wdNextRecord = -2,

    /// <summary>
    /// 结果集中的上一条记录
    /// </summary>
    wdPreviousRecord = -3,

    /// <summary>
    /// 结果集中的第一条记录
    /// </summary>
    wdFirstRecord = -4,

    /// <summary>
    /// 结果集中的最后一条记录
    /// </summary>
    wdLastRecord = -5,

    /// <summary>
    /// 数据源中的第一条记录
    /// </summary>
    wdFirstDataSourceRecord = -6,

    /// <summary>
    /// 数据源中的最后一条记录
    /// </summary>
    wdLastDataSourceRecord = -7,

    /// <summary>
    /// 数据源中的下一条记录
    /// </summary>
    wdNextDataSourceRecord = -8,

    /// <summary>
    /// 数据源中的上一条记录
    /// </summary>
    wdPreviousDataSourceRecord = -9
}