//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.ConnectorFormat 的二次封装实现类。
/// 提供安全访问连接符格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeConnectorFormat : IOfficeConnectorFormat
{
    private MsCore.ConnectorFormat _connectorFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 ConnectorFormat 对象。
    /// </summary>
    /// <param name="connectorFormat">原始的 COM ConnectorFormat 对象。</param>
    internal OfficeConnectorFormat(MsCore.ConnectorFormat connectorFormat)
    {
        _connectorFormat = connectorFormat ?? throw new ArgumentNullException(nameof(connectorFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public MsoConnectorType Type
    {
        get => _connectorFormat?.Type != null ? (MsoConnectorType)(int)_connectorFormat?.Type : MsoConnectorType.msoConnectorTypeMixed;
        set
        {
            if (_connectorFormat != null) _connectorFormat.Type = (MsCore.MsoConnectorType)(int)value;
        }
    }

    /// <inheritdoc/>
    public int BeginConnectionSite => _connectorFormat?.BeginConnectionSite ?? 0;

    /// <inheritdoc/>
    public int EndConnectionSite => _connectorFormat?.EndConnectionSite ?? 0;

    /// <inheritdoc/>
    public IOfficeShape BeginConnectedShape
    {
        get
        {
            if (_connectorFormat?.BeginConnectedShape != null)
                return new OfficeShape(_connectorFormat.BeginConnectedShape);
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeShape EndConnectedShape
    {
        get
        {
            if (_connectorFormat?.EndConnectedShape != null)
                return new OfficeShape(_connectorFormat.EndConnectedShape);
            return null;
        }
    }

    /// <inheritdoc/>
    public bool BeginConnected => _connectorFormat?.BeginConnected == MsCore.MsoTriState.msoTrue;

    /// <inheritdoc/>
    public bool EndConnected => _connectorFormat?.EndConnected == MsCore.MsoTriState.msoTrue;
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void BeginConnect(IOfficeShape shape, int connectionSite)
    {
        if (_connectorFormat != null && shape is OfficeShape officeShape)
        {
            var comShape = officeShape._shape;
            _connectorFormat.BeginConnect(comShape, connectionSite);
        }
    }

    /// <inheritdoc/>
    public void BeginDisconnect()
    {
        _connectorFormat?.BeginDisconnect();
    }

    /// <inheritdoc/>
    public void EndConnect(IOfficeShape shape, int connectionSite)
    {
        if (_connectorFormat != null && shape is OfficeShape officeShape)
        {
            var comShape = officeShape._shape;
            _connectorFormat.EndConnect(comShape, connectionSite);
        }
    }

    /// <inheritdoc/>
    public void EndDisconnect()
    {
        _connectorFormat?.EndDisconnect();
    }

    #endregion  

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _connectorFormat != null)
        {
            Marshal.ReleaseComObject(_connectorFormat);
            _connectorFormat = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
