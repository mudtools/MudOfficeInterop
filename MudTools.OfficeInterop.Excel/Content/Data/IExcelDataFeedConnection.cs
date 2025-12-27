//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel数据源连接接口，用于定义数据连接的基本操作和属性
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDataFeedConnection : IOfficeObject<IExcelDataFeedConnection>, IDisposable
{
    /// <summary>
    /// 获取父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取连接的命令文本
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取或设置刷新时是否提示文件未找到
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取或设置是否保存密码
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 获取或设置是否始终使用连接文件
    /// </summary>
    bool AlwaysUseConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置连接字符串
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Connection { get; set; }

    /// <summary>
    /// 获取或设置是否启用刷新功能
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 获取上次刷新日期时间
    /// </summary>
    DateTime RefreshDate { get; }

    /// <summary>
    /// 获取当前是否正在刷新数据
    /// </summary>
    bool Refreshing { get; }

    /// <summary>
    /// 获取或设置刷新周期（秒）
    /// </summary>
    int RefreshPeriod { get; set; }

    /// <summary>
    /// 获取或设置服务器凭据方法
    /// </summary>
    XlCredentialsMethod ServerCredentialsMethod { get; set; }

    /// <summary>
    /// 获取或设置源数据文件路径
    /// </summary>
    string SourceDataFile { get; set; }

    /// <summary>
    /// 获取或设置源连接文件路径
    /// </summary>
    string SourceConnectionFile { get; set; }

    /// <summary>
    /// 刷新数据源连接
    /// </summary>
    void Refresh();

    /// <summary>
    /// 取消正在进行的数据刷新操作
    /// </summary>
    void CancelRefresh();

    /// <summary>
    /// 将数据连接保存为 Office 数据连接 (ODC) 文件
    /// </summary>
    /// <param name="ODCFileName">要保存的 ODC 文件名</param>
    /// <param name="Description">连接描述信息</param>
    /// <param name="Keywords">连接关键字</param>
    void SaveAsODC(string ODCFileName, string Description, string Keywords);
}
