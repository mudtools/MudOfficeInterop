namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定对Excel工作表数据执行的筛选操作类型
/// </summary>
public enum XlFilterAction
{
    /// <summary>
    /// 将筛选结果复制到工作表中的其他位置
    /// </summary>
    xlFilterCopy = 2,

    /// <summary>
    /// 在原位置隐藏不符合筛选条件的行
    /// </summary>
    xlFilterInPlace = 1
}