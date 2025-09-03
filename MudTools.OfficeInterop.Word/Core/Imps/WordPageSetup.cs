//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档页面设置实现类
/// </summary>
internal class WordPageSetup : IWordPageSetup
{
    internal readonly MsWord.PageSetup _pageSetup;
    private bool _disposedValue;

    /// <summary>
    /// 获取或设置上边距
    /// </summary>
    public float TopMargin
    {
        get => _pageSetup.TopMargin;
        set => _pageSetup.TopMargin = value;
    }

    /// <summary>
    /// 获取或设置下边距
    /// </summary>
    public float BottomMargin
    {
        get => _pageSetup.BottomMargin;
        set => _pageSetup.BottomMargin = value;
    }

    /// <summary>
    /// 获取或设置左边距
    /// </summary>
    public float LeftMargin
    {
        get => _pageSetup.LeftMargin;
        set => _pageSetup.LeftMargin = value;
    }

    /// <summary>
    /// 获取或设置右边距
    /// </summary>
    public float RightMargin
    {
        get => _pageSetup.RightMargin;
        set => _pageSetup.RightMargin = value;
    }

    /// <summary>
    /// 获取或设置页面宽度
    /// </summary>
    public float PageWidth
    {
        get => _pageSetup.PageWidth;
        set => _pageSetup.PageWidth = value;
    }

    /// <summary>
    /// 获取或设置页面高度
    /// </summary>
    public float PageHeight
    {
        get => _pageSetup.PageHeight;
        set => _pageSetup.PageHeight = value;
    }

    /// <summary>
    /// 获取或设置页面方向（0=纵向，1=横向）
    /// </summary>
    public int Orientation
    {
        get => (int)_pageSetup.Orientation;
        set => _pageSetup.Orientation = (MsWord.WdOrientation)value;
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="pageSetup">COM PageSetup 对象</param>
    internal WordPageSetup(MsWord.PageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

