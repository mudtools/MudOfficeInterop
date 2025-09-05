//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel文本连接接口，用于处理文本文件的数据连接
/// 继承自IDisposable接口，支持资源释放
/// </summary>
public interface IExcelTextConnection : IDisposable
{
    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    IExcelWorkbookConnection Parent { get; }

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
    int[] TextFileFixedColumnWidths { get; set; }

    /// <summary>
    /// 获取或设置文本文件的平台编码
    /// </summary>
    XlPlatform TextFilePlatform { get; set; }

}