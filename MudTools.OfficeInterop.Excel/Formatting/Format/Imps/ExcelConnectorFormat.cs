//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// ConnectorFormat COM 对象的封装实现类。
/// 负责管理 COM 对象生命周期，提供安全的属性访问和资源释放。
/// </summary>
internal class ExcelConnectorFormat : IExcelConnectorFormat
{
    /// <summary>
    /// 内部持有的原始 COM 对象。
    /// </summary>
    internal MsExcel.ConnectorFormat _connectorFormat;

    /// <summary>
    /// 标记对象是否已被释放。
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装类。
    /// </summary>
    /// <param name="connectorFormat">原始的 ConnectorFormat COM 对象，不可为 null。</param>
    /// <exception cref="ArgumentNullException">当传入的 connectorFormat 为 null 时抛出。</exception>
    internal ExcelConnectorFormat(MsExcel.ConnectorFormat connectorFormat)
    {
        _connectorFormat = connectorFormat ?? throw new ArgumentNullException(nameof(connectorFormat));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的受保护虚方法，支持派生类重写。
    /// </summary>
    /// <param name="disposing">是否由用户代码显式调用释放。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放托管资源：释放 COM 对象
            if (_connectorFormat != null)
            {
                Marshal.ReleaseComObject(_connectorFormat);
                _connectorFormat = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 公开的 Dispose 方法，用于显式释放资源。
    /// 调用后对象不应再被使用。
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取此对象的父对象（通常是 Shape）。
    /// </summary>
    public object Parent => _connectorFormat?.Parent;

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// 返回封装后的 <see cref="IExcelApplication"/> 接口实例。
    /// </summary>
    public IExcelApplication Application =>
        _connectorFormat?.Application != null
            ? new ExcelApplication(_connectorFormat.Application as MsExcel.Application)
            : null;

    /// <summary>
    /// 获取或设置连接符的类型（直线、曲线、直角等）。
    /// 默认值：msoConnectorMixed
    /// </summary>
    public MsoConnectorType Type
    {
        get => _connectorFormat != null
            ? _connectorFormat.Type.EnumConvert(MsoConnectorType.msoConnectorTypeMixed)
            : MsoConnectorType.msoConnectorTypeMixed;

        set
        {
            if (_connectorFormat != null)
                _connectorFormat.Type = value.EnumConvert(MsCore.MsoConnectorType.msoConnectorTypeMixed);
        }
    }

    /// <summary>
    /// 获取或设置连接符起始端所连接的形状对象。
    /// 设置后，连接符将自动吸附到该形状。
    /// 获取时返回封装后的 <see cref="IExcelShape"/>。
    /// </summary>
    public IExcelShape? BeginConnectedShape
    {
        get => _connectorFormat?.BeginConnectedShape != null
            ? new ExcelShape(_connectorFormat.BeginConnectedShape)
            : null;
    }

    /// <summary>
    /// 获取或设置连接符终止端所连接的形状对象。
    /// 设置后，连接符将自动吸附到该形状。
    /// 获取时返回封装后的 <see cref="IExcelShape"/>。
    /// </summary>
    public IExcelShape? EndConnectedShape
    {
        get => _connectorFormat?.EndConnectedShape != null
            ? new ExcelShape(_connectorFormat.EndConnectedShape)
            : null;
    }

    /// <summary>
    /// 获取连接符起始端是否已连接到某个形状。
    /// </summary>
    public bool BeginConnected =>
        _connectorFormat != null && _connectorFormat.BeginConnected.ConvertToBool();

    /// <summary>
    /// 获取连接符终止端是否已连接到某个形状。
    /// </summary>
    public bool EndConnected =>
        _connectorFormat != null && _connectorFormat.EndConnected.ConvertToBool();

    /// <summary>
    /// 获取或设置连接符起始端所连接的形状上的连接点索引（从1开始）。
    /// 仅当 BeginConnected 为 true 时有效。
    /// </summary>
    public int BeginConnectionSite
    {
        get => _connectorFormat?.BeginConnectionSite ?? 0;
    }

    /// <summary>
    /// 获取或设置连接符终止端所连接的形状上的连接点索引（从1开始）。
    /// 仅当 EndConnected 为 true 时有效。
    /// </summary>
    public int EndConnectionSite
    {
        get => _connectorFormat?.EndConnectionSite ?? 0;
    }

    /// <summary>
    /// 将连接符起始端连接到指定形状的指定连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的目标形状（封装接口）。</param>
    /// <param name="connectionSite">连接点索引（从1开始）。</param>
    public void BeginConnect(IExcelShape connectedShape, int connectionSite)
    {
        if (_connectorFormat == null || connectedShape == null) return;

        var comShape = (connectedShape as ExcelShape)?._shape;
        if (comShape != null)
            _connectorFormat.BeginConnect(comShape, connectionSite);
    }

    /// <summary>
    /// 将连接符终止端连接到指定形状的指定连接点。
    /// </summary>
    /// <param name="connectedShape">要连接的目标形状（封装接口）。</param>
    /// <param name="connectionSite">连接点索引（从1开始）。</param>
    public void EndConnect(IExcelShape connectedShape, int connectionSite)
    {
        if (_connectorFormat == null || connectedShape == null) return;

        var comShape = (connectedShape as ExcelShape)?._shape;
        if (comShape != null)
            _connectorFormat.EndConnect(comShape, connectionSite);
    }

    /// <summary>
    /// 断开连接符起始端的连接（不再吸附到任何形状）。
    /// </summary>
    public void BeginDisconnect()
    {
        _connectorFormat?.BeginDisconnect();
    }

    /// <summary>
    /// 断开连接符终止端的连接（不再吸附到任何形状）。
    /// </summary>
    public void EndDisconnect()
    {
        _connectorFormat?.EndDisconnect();
    }
}