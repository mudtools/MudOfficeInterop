//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的OLEDB连接接口，用于与外部数据源建立连接并进行数据交互
/// 该接口封装了对Excel OLEDB连接对象的操作，提供了连接管理、数据刷新等功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelOLEDBConnection : IDisposable
{

    /// <summary>
    /// 获取连接的父级工作簿连接
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取条件值对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置连接的命令文本
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string CommandText { get; set; }

    /// <summary>
    /// 获取或设置命令类型
    /// </summary>
    XlCmdType CommandType { get; set; }

    /// <summary>
    /// 获取或设置连接字符串
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Connection { get; set; }

    /// <summary>
    /// 获取ADO连接对象
    /// </summary>
    object ADOConnection { get; }

    /// <summary>
    /// 获取或设置连接是否启用背景刷新
    /// </summary>
    bool BackgroundQuery { get; set; }


    /// <summary>
    /// 获取或设置是否启用连接
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 获取或设置刷新时是否提示用户
    /// </summary>
    bool RefreshOnFileOpen { get; set; }

    /// <summary>
    /// 获取或设置是否保存密码
    /// </summary>
    bool SavePassword { get; set; }

    /// <summary>
    /// 获取或设置源数据文件
    /// </summary>
    string SourceDataFile { get; set; }

    /// <summary>
    /// 获取源工作簿连接名称
    /// </summary>
    string SourceConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置是否始终使用连接文件
    /// </summary>
    bool AlwaysUseConnectionFile { get; set; }

    /// <summary>
    /// 获取或设置是否使用本地连接
    /// </summary>
    bool UseLocalConnection { get; set; }

    /// <summary>
    /// 获取或设置最大钻取记录数
    /// </summary>
    int MaxDrillthroughRecords { get; set; }

    /// <summary>
    /// 获取是否为OLAP连接
    /// </summary>
    bool OLAP { get; }

    /// <summary>
    /// 获取连接是否已建立
    /// </summary>
    bool IsConnected { get; }

    /// <summary>
    /// 获取或设置刷新周期（秒）
    /// </summary>
    int RefreshPeriod { get; set; }

    /// <summary>
    /// 获取是否正在刷新
    /// </summary>
    bool Refreshing { get; }

    /// <summary>
    /// 获取或设置是否保持连接
    /// </summary>
    bool MaintainConnection { get; set; }

    /// <summary>
    /// 获取或设置本地连接对象
    /// </summary>
    object LocalConnection { get; set; }

    /// <summary>
    /// 获取上次刷新日期时间
    /// </summary>
    DateTime RefreshDate { get; }

    /// <summary>
    /// 获取或设置服务器凭据方法
    /// </summary>
    XlCredentialsMethod ServerCredentialsMethod { get; set; }

    /// <summary>
    /// 获取或设置可靠连接设置
    /// </summary>
    XlRobustConnect RobustConnect { get; set; }

    /// <summary>
    /// 获取或设置服务器SSO应用程序ID
    /// </summary>
    string ServerSSOApplicationID { get; set; }

    /// <summary>
    /// 获取或设置是否从服务器检索填充颜色
    /// </summary>
    bool ServerFillColor { get; set; }

    /// <summary>
    /// 获取或设置是否从服务器检索字体样式
    /// </summary>
    bool ServerFontStyle { get; set; }

    /// <summary>
    /// 获取或设置是否从服务器检索数字格式
    /// </summary>
    bool ServerNumberFormat { get; set; }

    /// <summary>
    /// 获取或设置是否从服务器检索文本颜色
    /// </summary>
    bool ServerTextColor { get; set; }

    /// <summary>
    /// 获取或设置是否以Office UI语言检索数据
    /// </summary>
    bool RetrieveInOfficeUILang { get; set; }

    /// <summary>
    /// 获取计算成员集合
    /// </summary>
    IExcelCalculatedMembers CalculatedMembers { get; }

    /// <summary>
    /// 获取或设置区域设置ID
    /// </summary>
    int LocaleID { get; set; }

    /// <summary>
    /// 建立连接
    /// </summary>
    void MakeConnection();

    /// <summary>
    /// 刷新OLEDB连接
    /// </summary>
    void Refresh();

    /// <summary>
    /// 取消刷新操作
    /// </summary>
    void CancelRefresh();

    /// <summary>
    /// 重新连接数据库
    /// </summary>
    void Reconnect();

    /// <summary>
    /// 将连接保存为ODC文件
    /// </summary>
    /// <param name="ODCFileName">ODC文件名</param>
    /// <param name="description">连接描述信息</param>
    /// <param name="keywords">连接关键字</param>
    void SaveAsODC(string ODCFileName, string? description = null, string? keywords = null);


    /// <summary>
    /// 测试OLEDB连接
    /// </summary>
    /// <returns>连接是否成功</returns>
    [IgnoreGenerator]
    bool TestConnection();

    /// <summary>
    /// 执行SQL命令
    /// </summary>
    /// <param name="sql">SQL命令</param>
    /// <returns>受影响的行数</returns>
    [IgnoreGenerator]
    int ExecuteCommand(string sql);

}
