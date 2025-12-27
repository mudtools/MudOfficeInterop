//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表数据连接的接口，用于与Excel工作表中的数据连接进行交互
/// 该接口封装了COM对象，提供对Excel工作表数据连接属性的访问和操作能力
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelWorksheetDataConnection : IOfficeObject<IExcelWorksheetDataConnection>, IDisposable
{
    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据连接对象
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Connection { get; }

    /// <summary>
    /// 获取连接的命令文本
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型
    /// </summary>
    XlCmdType CommandType { get; set; }

}