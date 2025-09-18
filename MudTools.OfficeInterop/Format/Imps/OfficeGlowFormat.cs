//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.GlowFormat 的二次封装实现类。
/// 提供安全访问发光格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeGlowFormat : IOfficeGlowFormat
{
    private MsCore.GlowFormat _glowFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 GlowFormat 对象。
    /// </summary>
    /// <param name="glowFormat">原始的 COM GlowFormat 对象。</param>
    internal OfficeGlowFormat(MsCore.GlowFormat glowFormat)
    {
        _glowFormat = glowFormat ?? throw new ArgumentNullException(nameof(glowFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IOfficeColorFormat Color
    {
        get
        {
            if (_glowFormat?.Color != null)
            {
                return new OfficeColorFormat(_glowFormat.Color);
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public float Radius
    {
        get => _glowFormat?.Radius ?? 0f;
        set
        {
            if (_glowFormat != null)
                _glowFormat.Radius = value;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _glowFormat?.Transparency ?? 0f;
        set
        {
            if (_glowFormat != null)
                _glowFormat.Transparency = value;
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

        if (disposing && _glowFormat != null)
        {
            Marshal.ReleaseComObject(_glowFormat);
            _glowFormat = null;
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