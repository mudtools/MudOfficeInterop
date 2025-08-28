namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// <see cref="IWordTemplate"/> 接口的具体实现类，封装了 Microsoft.Office.Interop.Word.Template 对象。
/// 自动管理 COM 资源的生命周期，确保安全释放。
/// </summary>
internal class WordTemplate : IWordTemplate
{
    // 内部持有的 COM Template 对象
    private MsWord.Template _template;
    // 是否已释放资源
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装一个现有的 Word Template COM 对象
    /// </summary>
    /// <param name="template">来自 Interop 的 Template 实例</param>
    internal WordTemplate(MsWord.Template template)
    {
        _template = template ?? throw new ArgumentNullException(nameof(template));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _template != null ? new WordApplication(_template.Application) : null;

    /// <summary>
    /// 获取模板的完整文件路径（如：C:\Users\...\Normal.dotm）
    /// </summary>
    public string FullName => _template?.FullName;

    /// <summary>
    /// 获取模板的文件名（如：Normal.dotm）
    /// </summary>
    public string Name => _template?.Name;

    public string? Path => _template?.Path;

    public string? NoLineBreakBefore
    {
        get => _template?.NoLineBreakBefore;
        set => _template.NoLineBreakBefore = value;
    }

    public string? NoLineBreakAfter
    {
        get => _template?.NoLineBreakAfter;
        set => _template.NoLineBreakAfter = value;
    }

    public int NoProofing
    {
        get => _template.NoProofing;
        set => _template.NoProofing = value;
    }

    public IWordAutoTextEntries? AutoTextEntries => _template != null ? new WordAutoTextEntries(_template.AutoTextEntries) : null;

    public IWordBuildingBlockEntries? BuildingBlockEntries => _template != null ? new WordBuildingBlockEntries(_template.BuildingBlockEntries) : null;

    public WdTemplateType Type => (WdTemplateType)(int)_template?.Type;

    public WdJustificationMode JustificationMode
    {
        get => (WdJustificationMode)(int)_template?.JustificationMode;
        set => _template.JustificationMode = (MsWord.WdJustificationMode)(int)value;
    }

    public WdFarEastLineBreakLevel FarEastLineBreakLevel
    {
        get => (WdFarEastLineBreakLevel)(int)_template?.FarEastLineBreakLevel;
        set => _template.FarEastLineBreakLevel = (MsWord.WdFarEastLineBreakLevel)(int)value;
    }
    public WdFarEastLineBreakLanguageID FarEastLineBreakLanguage
    {
        get => (WdFarEastLineBreakLanguageID)(int)_template?.FarEastLineBreakLanguage;
        set => _template.FarEastLineBreakLanguage = (MsWord.WdFarEastLineBreakLanguageID)(int)value;
    }


    /// <summary>
    /// 获取或设置模板的“已保存”状态。若为 false，则关闭时会提示保存。
    /// </summary>
    public bool Saved
    {
        get => _template.Saved;
        set
        {
            if (_template != null)
                _template.Saved = value;
        }
    }


    #endregion

    #region 方法实现

    /// <summary>
    /// 保存对模板的当前修改
    /// </summary>
    public void Save()
    {
        _template?.Save();
    }

    public IWordDocument OpenAsDocument()
    {
        var doc = _template?.OpenAsDocument();
        return new WordDocument(doc, Application);
    }


    #endregion

    #region IDisposable 支持

    /// <summary>
    /// 释放由当前对象使用的所有资源（托管和非托管）
    /// </summary>
    /// <param name="disposing">是否由用户代码调用（true），或由 GC 调用（false）</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _template != null)
        {
            try
            {
                // 多次调用 ReleaseComObject 直至引用计数为 0
                while (Marshal.ReleaseComObject(_template) > 0) { }
            }
            catch (InvalidComObjectException)
            {
                // COM 对象可能已被释放
            }
            catch (COMException)
            {
                // 其他 COM 错误忽略
            }
            finally
            {
                _template = null;
            }
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 执行与释放或重置非托管资源相关的应用程序定义任务。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}