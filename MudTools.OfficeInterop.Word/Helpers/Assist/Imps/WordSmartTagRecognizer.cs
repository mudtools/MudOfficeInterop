//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordSmartTagRecognizer : IWordSmartTagRecognizer
{
    private MsWord.SmartTagRecognizer _recognizer;
    private bool _disposedValue;

    internal WordSmartTagRecognizer(MsWord.SmartTagRecognizer recognizer)
    {
        _recognizer = recognizer ?? throw new ArgumentNullException(nameof(recognizer));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    public IWordApplication? Application => _recognizer?.Application != null ? new WordApplication(_recognizer.Application) : null;

    /// <summary>
    /// 获取代表指定对象的父对象的对象 [[10]]。
    /// </summary>
    public object Parent => _recognizer?.Parent;

    /// <summary>
    /// 获取标识指定智能标记识别程序的标签 [[4]]。
    /// </summary>
    public string Caption => _recognizer?.Caption;

    /// <summary>
    /// 获取智能标记识别程序的完整名称，包括路径 [[9]]。
    /// </summary>
    public string FullName => _recognizer?.FullName;

    public string ProgID => _recognizer?.ProgID;

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值确定是否启用指定的智能标记识别程序 [[7]]。
    /// </summary>
    public bool Enabled
    {
        get => _recognizer?.Enabled ?? false;
        set
        {
            if (_recognizer != null)
                _recognizer.Enabled = value;
        }
    }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序 [[13]]。
    /// </summary>
    public int Creator => _recognizer?.Creator ?? 0;

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _recognizer != null)
        {
            Marshal.ReleaseComObject(_recognizer);
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}