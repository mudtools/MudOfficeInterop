namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在打印文档时要打印的内容项
/// </summary>
public enum WdPrintOutItem
{
    /// <summary>
    /// 打印文档内容
    /// </summary>
    wdPrintDocumentContent = 0,

    /// <summary>
    /// 打印文档属性
    /// </summary>
    wdPrintProperties = 1,

    /// <summary>
    /// 打印批注
    /// </summary>
    wdPrintComments = 2,

    /// <summary>
    /// 打印标记
    /// </summary>
    wdPrintMarkup = 2,

    /// <summary>
    /// 打印样式
    /// </summary>
    wdPrintStyles = 3,

    /// <summary>
    /// 打印自动图文集词条
    /// </summary>
    wdPrintAutoTextEntries = 4,

    /// <summary>
    /// 打印快捷键分配
    /// </summary>
    wdPrintKeyAssignments = 5,

    /// <summary>
    /// 打印信封
    /// </summary>
    wdPrintEnvelope = 6,

    /// <summary>
    /// 打印带有标记的文档
    /// </summary>
    wdPrintDocumentWithMarkup = 7
}