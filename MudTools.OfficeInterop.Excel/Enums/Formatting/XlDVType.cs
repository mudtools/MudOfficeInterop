namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定Excel数据有效性验证类型
/// </summary>
public enum XlDVType
{
    /// <summary>
    /// 仅在用户更改值时进行验证
    /// </summary>
    xlValidateInputOnly,
    
    /// <summary>
    /// 整数值
    /// </summary>
    xlValidateWholeNumber,
    
    /// <summary>
    /// 小数值
    /// </summary>
    xlValidateDecimal,
    
    /// <summary>
    /// 值必须存在于指定列表中
    /// </summary>
    xlValidateList,
    
    /// <summary>
    /// 日期值
    /// </summary>
    xlValidateDate,
    
    /// <summary>
    /// 时间值
    /// </summary>
    xlValidateTime,
    
    /// <summary>
    /// 文本长度
    /// </summary>
    xlValidateTextLength,
    
    /// <summary>
    /// 使用任意公式验证数据有效性
    /// </summary>
    xlValidateCustom
}