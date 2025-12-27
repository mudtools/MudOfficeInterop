//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordSmartTagAction : IWordSmartTagAction
{
    private MsWord.SmartTagAction _action;
    private bool _disposedValue;

    internal WordSmartTagAction(MsWord.SmartTagAction action)
    {
        _action = action ?? throw new ArgumentNullException(nameof(action));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象 [[9]]。
    /// </summary>
    public IWordApplication? Application => _action?.Application != null ? new WordApplication(_action.Application) : null;

    /// <summary>
    /// 获取代表指定对象的父对象的对象 [[18]]。
    /// </summary>
    public object? Parent => _action?.Parent;

    /// <summary>
    /// 获取智能标记操作的名称。
    /// </summary>
    public string Name => _action?.Name;

    /// <summary>
    /// 获取代表智能文档控件类型的 WdSmartTagControlType [[15]]。
    /// </summary>
    public WdSmartTagControlType Type =>
         _action?.Type != null ? (WdSmartTagControlType)(int)_action?.Type : WdSmartTagControlType.wdControlTextbox;

    /// <summary>
    /// 获取或设置一个整数，该整数代表智能文档列表框控件中所选项的索引号 [[7]]。
    /// </summary>
    public int ListSelection
    {
        get => _action?.ListSelection ?? 0;
        set
        {
            if (_action != null)
                _action.ListSelection = value;
        }
    }

    public string TextboxText
    {
        get => _action?.TextboxText ?? string.Empty;
        set
        {
            if (_action != null)
                _action.TextboxText = value;
        }
    }

    public int RadioGroupSelection
    {
        get => _action?.RadioGroupSelection ?? 0;
        set
        {
            if (_action != null)
                _action.RadioGroupSelection = value;
        }
    }

    public bool ExpandDocumentFragment
    {
        get => _action?.ExpandDocumentFragment ?? false;
        set
        {
            if (_action != null)
                _action.ExpandDocumentFragment = value;
        }
    }

    public bool ExpandHelp
    {
        get => _action?.ExpandHelp ?? false;
        set
        {
            if (_action != null)
                _action.ExpandHelp = value;
        }
    }

    public bool PresentInPane
    {
        get => _action?.PresentInPane ?? false;
    }

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值代表智能文档复选框控件的状态。
    /// </summary>
    public bool CheckboxState
    {
        get => _action?.CheckboxState ?? false;
        set
        {
            if (_action != null)
                _action.CheckboxState = value;
        }
    }

    /// <summary>
    /// 获取或设置一个 Object 类型的值，该值代表智能文档 ActiveX 控件的值。
    /// </summary>
    public object ActiveXControl
    {
        get => _action?.ActiveXControl;
    }

    #endregion

    #region 方法实现

    /// <summary>
    /// 执行指定的智能标记操作 [[14]]。
    /// </summary>
    public void Execute()
    {
        _action?.Execute();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _action != null)
        {
            Marshal.ReleaseComObject(_action);
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