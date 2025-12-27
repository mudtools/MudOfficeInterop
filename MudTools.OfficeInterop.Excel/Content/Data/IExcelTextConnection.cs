//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel文本连接接口，用于处理文本文件的数据连接
/// 继承自IDisposable接口，支持资源释放
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTextConnection : IOfficeObject<IExcelTextConnection>, IDisposable
{
    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置文本文件的起始行
    /// </summary>
    int TextFileStartRow { get; set; }

    /// <summary>
    /// 获取或设置文本文件的分隔符类型
    /// </summary>
    XlTextParsingType TextFileParseType { get; set; }

    /// <summary>
    /// 获取或设置文本文件的分隔符
    /// </summary>
    XlTextQualifier TextFileTextQualifier { get; set; }

    /// <summary>
    /// 获取或设置文本文件的固定宽度列
    /// </summary>
    object TextFileFixedColumnWidths { get; set; }

    /// <summary>
    /// 获取或设置文本文件的平台编码
    /// </summary>
    XlPlatform TextFilePlatform { get; set; }

    [ComPropertyWrap(NeedConvert = true)]
    string Connection { get; set; }

    /// <summary>
    /// 获取或设置文本文件是否包含标题行
    /// </summary>
    bool TextFileHeaderRow { get; set; }

    /// <summary>
    /// 获取或设置文本文件各列的数据类型
    /// </summary>
    object TextFileColumnDataTypes { get; set; }

    /// <summary>
    /// 获取或设置是否使用逗号作为文本文件的分隔符
    /// </summary>
    bool TextFileCommaDelimiter { get; set; }

    /// <summary>
    /// 获取或设置是否将连续的分隔符视为单个分隔符
    /// </summary>
    bool TextFileConsecutiveDelimiter { get; set; }

    /// <summary>
    /// 获取或设置文本文件的小数分隔符
    /// </summary>
    string TextFileDecimalSeparator { get; set; }

    /// <summary>
    /// 获取或设置文本文件的其他自定义分隔符
    /// </summary>
    string TextFileOtherDelimiter { get; set; }

    /// <summary>
    /// 获取或设置是否使用制表符作为文本文件的分隔符
    /// </summary>
    bool TextFileTabDelimiter { get; set; }

    /// <summary>
    /// 获取或设置是否使用分号作为文本文件的分隔符
    /// </summary>
    bool TextFileSemicolonDelimiter { get; set; }

    /// <summary>
    /// 获取或设置在刷新数据时是否提示用户
    /// </summary>
    bool TextFilePromptOnRefresh { get; set; }

    /// <summary>
    /// 获取或设置是否使用空格作为文本文件的分隔符
    /// </summary>
    bool TextFileSpaceDelimiter { get; set; }

    /// <summary>
    /// 获取或设置文本文件的千位分隔符
    /// </summary>
    string TextFileThousandsSeparator { get; set; }

    /// <summary>
    /// 获取或设置是否将末尾带负号的数字视为负数
    /// </summary>
    bool TextFileTrailingMinusNumbers { get; set; }

    /// <summary>
    /// 获取或设置文本文件的视觉布局类型
    /// </summary>
    XlTextVisualLayoutType TextFileVisualLayout { get; set; }
}