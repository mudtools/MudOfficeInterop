//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;


/// <summary>
/// SmartArt COM 对象的封装实现类
/// </summary>
internal class OfficeSmartArt : IOfficeSmartArt
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficeSmartArt));

    internal MsCore.SmartArt? _smartArt;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化 SmartArt 封装对象
    /// </summary>
    /// <param name="smartArt">原始 SmartArt COM 对象，不能为空</param>
    internal OfficeSmartArt(MsCore.SmartArt smartArt)
    {
        _smartArt = smartArt ?? throw new ArgumentNullException(nameof(smartArt));
        _disposedValue = false;
    }

    #region 属性实现

    public IOfficeSmartArtNodes? AllNodes
    {
        get
        {
            if (_smartArt == null) return null;

            return new OfficeSmartArtNodes(_smartArt.AllNodes);
        }
    }

    public IOfficeSmartArtNodes? Nodes
    {
        get
        {
            if (_smartArt == null) return null;
            return new OfficeSmartArtNodes(_smartArt.Nodes);
        }
    }

    public IOfficeSmartArtLayout? Layout
    {
        get
        {
            if (_smartArt == null) return null;
            return new OfficeSmartArtLayout(_smartArt.Layout);
        }
    }

    public IOfficeSmartArtQuickStyle? QuickStyle
    {
        get
        {
            if (_smartArt == null) return null;
            return new OfficeSmartArtQuickStyle(_smartArt.QuickStyle);
        }
    }

    public IOfficeSmartArtColor? Color
    {
        get
        {
            if (_smartArt == null) return null;
            return new OfficeSmartArtColor(_smartArt.Color);
        }
    }

    public bool Reverse
    {
        get
        {
            if (_smartArt == null) return false;
            return _smartArt.Reverse.ConvertToBool();
        }
        set
        {
            if (_smartArt == null) return;
            _smartArt.Reverse = value.ConvertTriState();
        }
    }
    #endregion

    public void Reset()
    {
        if (_smartArt == null)
            return;

        try
        {
            _smartArt.Reset();
        }
        catch (COMException ex)
        {
            log.Error("重置 SmartArt 失败。", ex);
        }
    }

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _smartArt != null)
        {
            Marshal.ReleaseComObject(_smartArt);
            _smartArt = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    ~OfficeSmartArt()
    {
        Dispose(false);
    }

    #endregion
}