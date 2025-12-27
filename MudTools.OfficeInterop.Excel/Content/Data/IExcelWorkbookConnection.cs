//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作簿连接的接口，提供对工作簿中各种数据连接的访问和操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelWorkbookConnection : IOfficeObject<IExcelWorkbookConnection>, IDisposable
{

    /// <summary>
    /// 获取父级连接集合
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 WorksheetFunction 对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置连接名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置连接描述
    /// </summary>
    string Description { get; set; }

    /// <summary>
    /// 获取连接类型
    /// </summary>
    XlConnectionType Type { get; }

    /// <summary>
    /// 获取连接的OLE DB连接对象
    /// </summary>
    IExcelOLEDBConnection OLEDBConnection { get; }

    /// <summary>
    /// 获取连接的ODBC连接对象
    /// </summary>
    IExcelODBCConnection ODBCConnection { get; }

    /// <summary>
    /// 获取连接的模型连接对象
    /// </summary>
    IExcelModelConnection ModelConnection { get; }

    /// <summary>
    /// 获取工作表数据连接对象
    /// </summary>
    IExcelWorksheetDataConnection WorksheetDataConnection { get; }

    /// <summary>
    /// 获取文本连接对象
    /// </summary>
    IExcelTextConnection TextConnection { get; }

    /// <summary>
    /// 获取数据源连接对象
    /// </summary>
    IExcelDataFeedConnection DataFeedConnection { get; }

    /// <summary>
    /// 获取连接关联的区域集合
    /// </summary>
    IExcelRanges Ranges { get; }

    /// <summary>
    /// 刷新连接
    /// </summary>
    void Refresh();

    /// <summary>
    /// 删除连接
    /// </summary>
    void Delete();
}