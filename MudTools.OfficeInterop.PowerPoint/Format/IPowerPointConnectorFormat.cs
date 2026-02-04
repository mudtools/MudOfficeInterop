//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示连接线格式，用于管理两个形状之间的连接线。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointConnectorFormat : IOfficeObject<IPowerPointConnectorFormat, MsPowerPoint.ConnectorFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此对象的应用程序对象。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此对象的创建者标识符。
    /// </summary>
    /// <value>创建者的整数标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此连接线格式的父对象。
    /// </summary>
    /// <value>父对象。</value>
    object Parent { get; }

    /// <summary>
    /// 将连接线的起点连接到指定形状的连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的形状。</param>
    /// <param name="connectionSite">连接点在形状上的位置索引。</param>
    void BeginConnect(IPowerPointShape connectedShape, int connectionSite);

    /// <summary>
    /// 断开连接线的起点与其连接形状的链接。
    /// </summary>
    void BeginDisconnect();

    /// <summary>
    /// 将连接线的终点连接到指定形状的连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的形状。</param>
    /// <param name="connectionSite">连接点在形状上的位置索引。</param>
    void EndConnect(IPowerPointShape connectedShape, int connectionSite);

    /// <summary>
    /// 断开连接线的终点与其连接形状的链接。
    /// </summary>
    void EndDisconnect();

    /// <summary>
    /// 获取一个值，指示连接线的起点是否已连接到形状。
    /// </summary>
    /// <value>如果起点已连接则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool BeginConnected { get; }

    /// <summary>
    /// 获取连接线起点所连接的形状。
    /// </summary>
    /// <value>起点连接的形状；如果未连接则为 null。</value>
    IPowerPointShape? BeginConnectedShape { get; }

    /// <summary>
    /// 获取连接线起点所连接形状上的连接点位置索引。
    /// </summary>
    /// <value>连接点位置索引。</value>
    int BeginConnectionSite { get; }

    /// <summary>
    /// 获取一个值，指示连接线的终点是否已连接到形状。
    /// </summary>
    /// <value>如果终点已连接则为 true；否则为 false。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool EndConnected { get; }

    /// <summary>
    /// 获取连接线终点所连接的形状。
    /// </summary>
    /// <value>终点连接的形状；如果未连接则为 null。</value>
    IPowerPointShape? EndConnectedShape { get; }

    /// <summary>
    /// 获取连接线终点所连接形状上的连接点位置索引。
    /// </summary>
    /// <value>连接点位置索引。</value>
    int EndConnectionSite { get; }

    /// <summary>
    /// 获取或设置连接线的类型。
    /// </summary>
    /// <value>连接线类型的枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoConnectorType Type { get; set; }
}