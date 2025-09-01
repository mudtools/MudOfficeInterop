//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.TextFrame 的二次封装实现类。
/// 提供安全访问文本框架属性和方法的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeTextFrame : IOfficeTextFrame
{
    private MsCore.TextFrame _textFrame;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 TextFrame 对象。
    /// </summary>
    /// <param name="textFrame">原始的 COM TextFrame 对象。</param>
    internal OfficeTextFrame(MsCore.TextFrame textFrame)
    {
        _textFrame = textFrame ?? throw new ArgumentNullException(nameof(textFrame));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IOfficeShape Parent
    {
        get
        {
            if (_textFrame?.Parent != null)
            {
                // 注意：这里需要根据实际的 Parent 类型进行处理
                // 通常 Parent 是 Shape 对象
                try
                {
                    var parentShape = _textFrame.Parent as MsCore.Shape;
                    if (parentShape != null)
                        return new OfficeShape(parentShape);
                }
                catch
                {
                    // 如果转换失败，返回 null
                    return null;
                }
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public float MarginLeft
    {
        get => _textFrame?.MarginLeft ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginLeft = value;
        }
    }

    /// <inheritdoc/>
    public float MarginRight
    {
        get => _textFrame?.MarginRight ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginRight = value;
        }
    }

    /// <inheritdoc/>
    public float MarginTop
    {
        get => _textFrame?.MarginTop ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginTop = value;
        }
    }

    /// <inheritdoc/>
    public float MarginBottom
    {
        get => _textFrame?.MarginBottom ?? 0f;
        set
        {
            if (_textFrame != null)
                _textFrame.MarginBottom = value;
        }
    }

    /// <inheritdoc/>
    public MsoTextOrientation Orientation
    {
        get => _textFrame?.Orientation != null ? (MsoTextOrientation)(int)_textFrame?.Orientation : MsoTextOrientation.msoTextOrientationMixed;
        set
        {
            if (_textFrame != null) _textFrame.Orientation = (MsCore.MsoTextOrientation)(int)value;
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

        if (disposing && _textFrame != null)
        {
            Marshal.ReleaseComObject(_textFrame);
            _textFrame = null;
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