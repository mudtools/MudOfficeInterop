namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中一个范围的字符对象接口，继承自IExcelCharacters接口
/// </summary>
public interface IExcelRangeCharacters : IExcelCharacters
{
    /// <summary>
    /// 获取指定起始位置和长度的字符子集
    /// </summary>
    /// <param name="start">起始位置（从1开始）</param>
    /// <param name="length">要获取的字符长度</param>
    /// <returns>表示指定范围内字符的IExcelCharacters对象</returns>
    IExcelCharacters this[int start, int length] { get; }
}