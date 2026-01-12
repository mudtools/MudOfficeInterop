
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示XML数据导入到Excel操作的结果
/// </summary>
public enum XlXmlImportResult
{
    /// <summary>
    /// XML导入成功
    /// </summary>
    xlXmlImportSuccess,

    /// <summary>
    /// XML导入时部分元素被截断
    /// </summary>
    xlXmlImportElementsTruncated,

    /// <summary>
    /// XML导入时验证失败
    /// </summary>
    xlXmlImportValidationFailed
}