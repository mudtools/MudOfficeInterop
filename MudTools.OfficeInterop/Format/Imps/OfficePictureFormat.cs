//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.PictureFormat 的二次封装实现类。
/// 提供安全访问图片格式属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficePictureFormat : IOfficePictureFormat
{
    private MsCore.PictureFormat _pictureFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 PictureFormat 对象。
    /// </summary>
    /// <param name="pictureFormat">原始的 COM PictureFormat 对象。</param>
    internal OfficePictureFormat(MsCore.PictureFormat pictureFormat)
    {
        _pictureFormat = pictureFormat ?? throw new ArgumentNullException(nameof(pictureFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public float Brightness
    {
        get => _pictureFormat?.Brightness ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Brightness = value;
        }
    }

    /// <inheritdoc/>
    public float Contrast
    {
        get => _pictureFormat?.Contrast ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.Contrast = value;
        }
    }

    /// <inheritdoc/>
    public int TransparencyColor
    {
        get => _pictureFormat?.TransparencyColor ?? 0;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.TransparencyColor = value;
        }
    }

    /// <inheritdoc/>
    public float CropLeft
    {
        get => _pictureFormat?.CropLeft ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropLeft = value;
        }
    }

    /// <inheritdoc/>
    public float CropRight
    {
        get => _pictureFormat?.CropRight ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropRight = value;
        }
    }

    /// <inheritdoc/>
    public float CropTop
    {
        get => _pictureFormat?.CropTop ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropTop = value;
        }
    }

    /// <inheritdoc/>
    public float CropBottom
    {
        get => _pictureFormat?.CropBottom ?? 0f;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.CropBottom = value;
        }
    }


    /// <inheritdoc/>
    public bool TransparentBackground
    {
        get => _pictureFormat?.TransparentBackground == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_pictureFormat != null)
                _pictureFormat.TransparentBackground = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void IncrementBrightness(float increment)
    {
        _pictureFormat?.IncrementBrightness(increment);
    }

    public void IncrementContrast(float increment)
    {
        _pictureFormat?.IncrementContrast(increment);
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

        if (disposing && _pictureFormat != null)
        {
            Marshal.ReleaseComObject(_pictureFormat);
            _pictureFormat = null;
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