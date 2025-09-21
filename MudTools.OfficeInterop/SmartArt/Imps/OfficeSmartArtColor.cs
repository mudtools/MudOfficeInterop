//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.SmartArtColor 的二次封装实现类。
/// 提供安全访问 SmartArt 颜色方案属性的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeSmartArtColor : IOfficeSmartArtColor
{
    private MsCore.SmartArtColor _smartArtColor;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 SmartArtColor 对象。
    /// </summary>
    /// <param name="smartArtColor">原始的 COM SmartArtColor 对象。</param>
    internal OfficeSmartArtColor(MsCore.SmartArtColor smartArtColor)
    {
        _smartArtColor = smartArtColor ?? throw new ArgumentNullException(nameof(smartArtColor));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string Name => _smartArtColor?.Name ?? string.Empty;

    /// <inheritdoc/>
    public string Description => _smartArtColor?.Description ?? string.Empty;

    /// <inheritdoc/>
    public string Category => _smartArtColor?.Category ?? string.Empty;

    /// <inheritdoc/>
    public string Id => _smartArtColor?.Id ?? string.Empty;

    /// <inheritdoc/>
    public IOfficeSmartArtColors? Parent
    {
        get
        {
            if (_smartArtColor?.Parent != null)
            {
                if (_smartArtColor.Parent is MsCore.SmartArtColors smartArtColor)
                    return new OfficeSmartArtColors(smartArtColor);
                return null;
            }
            return null;
        }
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

        if (disposing && _smartArtColor != null)
        {
            Marshal.ReleaseComObject(_smartArtColor);
            _smartArtColor = null;
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
