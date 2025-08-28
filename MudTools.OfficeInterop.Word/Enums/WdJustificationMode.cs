namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在调整文本对齐时Word如何处理字符间距的枚举类型。
/// 该设置影响完全对齐段落中字符的间距调整方式。
/// </summary>
public enum WdJustificationMode
{
    /// <summary>
    /// 扩展字符间距以填充行宽
    /// </summary>
    wdJustificationModeExpand,

    /// <summary>
    /// 压缩字符间距以适应行宽
    /// </summary>
    wdJustificationModeCompress,

    /// <summary>
    /// 仅压缩假名字符（日文平假名和片假名）的间距
    /// </summary>
    wdJustificationModeCompressKana
}