//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureFormat 的实现类。
/// </summary>
internal class WordPictureFormat : IWordPictureFormat
{
    private MsWord.PictureFormat _pictureFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="pictureFormat">原始 COM PictureFormat 对象。</param>
    internal WordPictureFormat(MsWord.PictureFormat pictureFormat)
    {
        _pictureFormat = pictureFormat ?? throw new ArgumentNullException(nameof(pictureFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application =>
        _pictureFormat != null ? new WordApplication(_pictureFormat.Application) : null;

    public IWordCrop? Crop =>
         _pictureFormat?.Crop != null ? new WordCrop(_pictureFormat.Crop) : null;

    /// <inheritdoc/>
    public float Brightness
    {
        get
        {
            return _pictureFormat?.Brightness ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.Brightness = Math.Max(-1.0f, Math.Min(1.0f, value));
            }
        }
    }

    /// <inheritdoc/>
    public float Contrast
    {
        get
        {
            return _pictureFormat?.Contrast ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.Contrast = Math.Max(0.0f, Math.Min(1.0f, value));

            }
        }
    }

    /// <inheritdoc/>
    public MsoPictureColorType ColorType
    {
        get => _pictureFormat?.ColorType != null ? (MsoPictureColorType)(int)_pictureFormat?.ColorType : MsoPictureColorType.msoPictureMixed;
        set
        {
            if (_pictureFormat != null) _pictureFormat.ColorType = (MsCore.MsoPictureColorType)(int)value;
        }
    }

    /// <inheritdoc/>
    public float CropLeft
    {
        get
        {

            return _pictureFormat?.CropLeft ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.CropLeft = value;

            }
        }
    }

    /// <inheritdoc/>
    public float CropRight
    {
        get
        {

            return _pictureFormat?.CropRight ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.CropRight = value;

            }
        }
    }

    /// <inheritdoc/>
    public float CropTop
    {
        get
        {

            return _pictureFormat?.CropTop ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.CropTop = value;

            }
        }
    }

    /// <inheritdoc/>
    public float CropBottom
    {
        get
        {

            return _pictureFormat?.CropBottom ?? 0f;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.CropBottom = value;

            }
        }
    }

    /// <inheritdoc/>
    public int TransparencyColor
    {
        get
        {

            return _pictureFormat?.TransparencyColor ?? 0;
        }
        set
        {
            if (_pictureFormat != null)
            {
                _pictureFormat.TransparencyColor = value;

            }
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

    /// <inheritdoc/>
    public object? Parent => _pictureFormat?.Parent;

    /// <inheritdoc/>
    public IWordSoftEdgeFormat? SoftEdge =>
        _pictureFormat?.Parent is MsWord.Shape shape && shape.SoftEdge != null
            ? new WordSoftEdgeFormat(shape.SoftEdge) : null;

    /// <inheritdoc/>
    public IWordGlowFormat? Glow =>
         _pictureFormat?.Parent is MsWord.Shape shape && shape.Glow != null
            ? new WordGlowFormat(shape.Glow) : null;

    /// <inheritdoc/>
    public IWordReflectionFormat? Reflection =>
        _pictureFormat?.Parent is MsWord.Shape shape && shape.Reflection != null
            ? new WordReflectionFormat(shape.Reflection) : null;

    /// <inheritdoc/>
    public bool IsLinked => !string.IsNullOrEmpty(Filename) && System.IO.File.Exists(Filename);

    /// <inheritdoc/>
    public string Filename => _pictureFormat?.Parent is MsWord.Shape shape
        ? shape.LinkFormat?.SourceFullName ?? string.Empty : string.Empty;

    /// <inheritdoc/>
    public long FileSize
    {
        get
        {
            if (!string.IsNullOrEmpty(Filename) && File.Exists(Filename))
            {
                return new FileInfo(Filename).Length;
            }
            return 0;
        }
    }

    /// <inheritdoc/>
    public bool HasTransparency => TransparentBackground;

    /// <inheritdoc/>
    public bool IsGrayscale => ColorType == MsoPictureColorType.msoPictureGrayscale;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void AdjustBrightness(float brightness)
    {
        if (_pictureFormat != null)
        {
            try
            {
                Brightness = Math.Max(-1.0f, Math.Min(1.0f, brightness));

            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法调整图片亮度。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void AdjustContrast(float contrast)
    {
        if (_pictureFormat != null)
        {
            try
            {
                Contrast = Math.Max(0.0f, Math.Min(1.0f, contrast));

            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法调整图片对比度。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_pictureFormat != null)
        {
            try
            {
                _pictureFormat.Brightness = 0f;
                _pictureFormat.Contrast = 0.5f;
                _pictureFormat.ColorType = MsCore.MsoPictureColorType.msoPictureAutomatic;
                _pictureFormat.CropLeft = 0f;
                _pictureFormat.CropRight = 0f;
                _pictureFormat.CropTop = 0f;
                _pictureFormat.CropBottom = 0f;

            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法重置图片格式。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void SetTransparentColor(int rgb)
    {
        if (_pictureFormat != null)
        {
            try
            {
                _pictureFormat.TransparencyColor = rgb;
                _pictureFormat.TransparentBackground = MsCore.MsoTriState.msoTrue;

            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法设置透明色。", ex);
            }
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordPictureFormat targetPicture)
    {
        if (_pictureFormat == null || targetPicture == null)
            return;

        try
        {
            targetPicture.Brightness = this.Brightness;
            targetPicture.Contrast = this.Contrast;
            targetPicture.ColorType = this.ColorType;
            targetPicture.CropLeft = this.CropLeft;
            targetPicture.CropRight = this.CropRight;
            targetPicture.CropTop = this.CropTop;
            targetPicture.CropBottom = this.CropBottom;
            targetPicture.TransparencyColor = this.TransparencyColor;
            targetPicture.TransparentBackground = this.TransparentBackground;

        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制图片格式。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Update()
    {
        if (_pictureFormat?.Parent is MsWord.Shape shape && shape.LinkFormat != null)
        {
            try
            {
                shape.LinkFormat.Update();

                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }
        return false;
    }

    /// <inheritdoc/>
    public bool BreakLink()
    {
        if (_pictureFormat?.Parent is MsWord.Shape shape && shape.LinkFormat != null)
        {
            try
            {
                shape.LinkFormat.BreakLink();

                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }
        return false;
    }

    /// <inheritdoc/>
    public bool ValidateParameters(float brightness, float contrast)
    {
        return brightness >= -1.0f && brightness <= 1.0f &&
               contrast >= 0.0f && contrast <= 1.0f;
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放父形状对象中的柔化边缘
            if (_pictureFormat?.Parent is MsWord.Shape shape && shape.SoftEdge != null)
            {
                try
                {
                    Marshal.ReleaseComObject(shape.SoftEdge);
                }
                catch
                {
                    // 忽略释放异常
                }
            }
            // 释放图片格式对象本身
            if (_pictureFormat != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_pictureFormat);
                }
                catch
                {
                    // 忽略释放异常
                }
                _pictureFormat = null;
            }
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