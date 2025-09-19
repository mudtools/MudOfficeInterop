
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 连接符（Connector）格式设置的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ConnectorFormat
/// 用于控制连接符的类型、起始/终止连接对象及连接点。
/// </summary>
public interface IExcelConnectorFormat : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置连接符的类型（直线、曲线、直角等）。
    /// 使用 <see cref="MsoConnectorType"/> 枚举。
    /// </summary>
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
    bool BeginConnected { get; }

    /// <summary>
    /// 获取连接符终止端是否已连接到某个形状。
    /// </summary>
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