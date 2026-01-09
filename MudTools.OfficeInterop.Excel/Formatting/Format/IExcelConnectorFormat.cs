//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 连接符（Connector）格式设置的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ConnectorFormat
/// 用于控制连接符的类型、起始/终止连接对象及连接点。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelConnectorFormat : IOfficeObject<IExcelConnectorFormat, MsExcel.ConnectorFormat>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置连接符的类型（直线、曲线、直角等）。
    /// 使用 <see cref="MsoConnectorType"/> 枚举。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoConnectorType Type { get; set; }

    /// <summary>
    /// 获取或设置连接符起始端所连接的形状对象。
    /// 设置后，连接符将自动吸附到该形状。
    /// </summary>
    IExcelShape? BeginConnectedShape { get; }

    /// <summary>
    /// 获取或设置连接符终止端所连接的形状对象。
    /// 设置后，连接符将自动吸附到该形状。
    /// </summary>
    IExcelShape? EndConnectedShape { get; }

    /// <summary>
    /// 获取连接符起始端是否已连接到某个形状。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool BeginConnected { get; }

    /// <summary>
    /// 获取连接符终止端是否已连接到某个形状。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool EndConnected { get; }

    /// <summary>
    /// 获取或设置连接符起始端所连接的形状上的连接点索引（从1开始）。
    /// 仅当 BeginConnected 为 true 时有效。
    /// </summary>
    int BeginConnectionSite { get; }

    /// <summary>
    /// 获取或设置连接符终止端所连接的形状上的连接点索引（从1开始）。
    /// 仅当 EndConnected 为 true 时有效。
    /// </summary>
    int EndConnectionSite { get; }

    /// <summary>
    /// 将连接符起始端连接到指定形状的指定连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的目标形状。</param>
    /// <param name="connectionSite">连接点索引（从1开始，通常1~n，取决于形状）。</param>
    void BeginConnect(IExcelShape connectedShape, int connectionSite);

    /// <summary>
    /// 将连接符终止端连接到指定形状的指定连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的目标形状。</param>
    /// <param name="connectionSite">连接点索引（从1开始）。</param>
    void EndConnect(IExcelShape connectedShape, int connectionSite);

    /// <summary>
    /// 断开连接符起始端的连接（不再吸附到任何形状）。
    /// </summary>
    void BeginDisconnect();

    /// <summary>
    /// 断开连接符终止端的连接（不再吸附到任何形状）。
    /// </summary>
    void EndDisconnect();
}