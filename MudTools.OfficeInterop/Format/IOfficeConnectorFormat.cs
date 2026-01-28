//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中连接符格式的接口封装。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeConnectorFormat : IOfficeObject<IOfficeConnectorFormat, MsCore.ConnectorFormat>, IDisposable
{
    /// <summary>
    /// 获取或设置连接符的类型。
    /// </summary>
    MsoConnectorType Type { get; set; }

    /// <summary>
    /// 获取连接符的起始连接站点。
    /// </summary>
    int BeginConnectionSite { get; }

    /// <summary>
    /// 获取连接符的结束连接站点。
    /// </summary>
    int EndConnectionSite { get; }

    /// <summary>
    /// 获取连接符起始连接的形状。
    /// </summary>
    IOfficeShape? BeginConnectedShape { get; }

    /// <summary>
    /// 获取连接符结束连接的形状。
    /// </summary>
    IOfficeShape? EndConnectedShape { get; }

    /// <summary>
    /// 获取连接符起始点是否已连接。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool BeginConnected { get; }

    /// <summary>
    /// 获取连接符结束点是否已连接。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool EndConnected { get; }

    /// <summary>
    /// 将连接符的起始点连接到指定形状。
    /// </summary>
    /// <param name="shape">要连接的形状。</param>
    /// <param name="connectionSite">连接站点索引。</param>
    void BeginConnect(IOfficeShape shape, int connectionSite);

    /// <summary>
    /// 断开连接符的起始点连接。
    /// </summary>
    void BeginDisconnect();

    /// <summary>
    /// 将连接符的结束点连接到指定形状。
    /// </summary>
    /// <param name="shape">要连接的形状。</param>
    /// <param name="connectionSite">连接站点索引。</param>
    void EndConnect(IOfficeShape shape, int connectionSite);

    /// <summary>
    /// 断开连接符的结束点连接。
    /// </summary>
    void EndDisconnect();
}