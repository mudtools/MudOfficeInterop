//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Crop 的实现类。
/// </summary>
internal class WordCrop : IWordCrop
{
    private MsCore.Crop _crop;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="crop">原始 COM Crop 对象。</param>
    internal WordCrop(MsCore.Crop crop)
    {
        _crop = crop ?? throw new ArgumentNullException(nameof(crop));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _crop != null ? new WordApplication(_crop.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public float ShapeLeft
    {
        get
        {

            return _crop?.ShapeLeft ?? 0f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.ShapeLeft = value;

            }
        }
    }

    /// <inheritdoc/>
    public float ShapeTop
    {
        get
        {

            return _crop?.ShapeTop ?? 0f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.ShapeTop = value;

            }
        }
    }

    /// <inheritdoc/>
    public float ShapeWidth
    {
        get
        {

            return _crop?.ShapeWidth ?? 0f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.ShapeWidth = Math.Max(0, value);

            }
        }
    }

    /// <inheritdoc/>
    public float ShapeHeight
    {
        get
        {

            return _crop?.ShapeHeight ?? 0f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.ShapeHeight = Math.Max(0, value);

            }
        }
    }

    /// <inheritdoc/>
    public float PictureWidth
    {
        get
        {

            return _crop?.PictureWidth ?? 1f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.PictureWidth = value;

            }
        }
    }

    /// <inheritdoc/>
    public float PictureHeight
    {
        get
        {

            return _crop?.PictureHeight ?? 1f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.PictureHeight = value;

            }
        }
    }

    /// <inheritdoc/>
    public float PictureOffsetX
    {
        get
        {

            return _crop?.PictureOffsetX ?? 1f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.PictureOffsetX = value;

            }
        }
    }

    /// <inheritdoc/>
    public float PictureOffsetY
    {
        get
        {

            return _crop?.PictureOffsetY ?? 1f;
        }
        set
        {
            if (_crop != null)
            {
                _crop.PictureOffsetY = value;

            }
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void SetShapeSize(float width, float height)
    {
        if (_crop != null)
        {
            _crop.ShapeWidth = Math.Max(0, width);
            _crop.ShapeHeight = Math.Max(0, height);
        }
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

        if (disposing && _crop != null)
        {
            Marshal.ReleaseComObject(_crop);
            _crop = null;
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