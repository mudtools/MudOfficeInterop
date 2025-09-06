//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using Microsoft.Office.Interop.Word;
using MudTools.OfficeInterop.Imps;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 应用程序实现类
/// </summary>
internal partial class WordApplication : IWordApplication
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordApplication));
    private static readonly object MissingValue = System.Reflection.Missing.Value;
    private MsWord.Application _application;
    private IWordDocument _activeDocument;
    private bool _disposedValue;
    private IWordWindows _windows;
    private IWordDocuments _documents;
    private IWordSelection _selection;

    #region 应用程序基础属性

    /// <summary>
    /// 获取或设置应用程序的可见性
    /// </summary>
    public WordAppVisibility Visibility
    {
        get => _application.Visible ? WordAppVisibility.Visible : WordAppVisibility.Hidden;
        set => _application.Visible = value == WordAppVisibility.Visible;
    }

    /// <summary>
    /// 获取父对象。对于 Application 对象，通常返回 null。
    /// </summary>
    /// <inheritdoc/>
    public object Parent => _application?.Parent;

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    /// <inheritdoc/>
    public int Creator => _application?.Creator ?? 0;

    /// <summary>
    /// 获取应用程序的名称
    /// </summary>
    /// <inheritdoc/>
    public string Name
    {
        get => _application?.Name ?? string.Empty;
    }

    /// <summary>
    /// 获取应用程序的版本号
    /// </summary>
    /// <inheritdoc/>
    public string Version => _application?.Version ?? string.Empty;

    /// <summary>
    /// 获取应用程序的路径
    /// </summary>
    /// <inheritdoc/>
    public string Path => _application?.Path ?? string.Empty;

    /// <summary>
    /// 获取用于分隔文件夹名称的字符。
    /// </summary>
    /// <inheritdoc/>
    public string PathSeparator => _application?.PathSeparator ?? string.Empty;

    /// <summary>
    /// 获取或设置应用程序窗口的水平位置
    /// </summary>
    /// <inheritdoc/>
    public float Left
    {
        get => _application?.Left ?? 0;
        set { if (_application != null) _application.Left = Convert.ToInt32(value); }
    }

    /// <summary>
    /// 获取或设置应用程序窗口的垂直位置
    /// </summary>
    /// <inheritdoc/>
    public float Top
    {
        get => _application?.Top ?? 0;
        set { if (_application != null) _application.Top = Convert.ToInt32(value); }
    }

    /// <summary>
    /// 获取或设置应用程序窗口的宽度
    /// </summary>
    /// <inheritdoc/>
    public float Width
    {
        get => _application?.Width ?? 0;
        set { if (_application != null) _application.Width = Convert.ToInt32(value); }
    }

    /// <summary>
    /// 获取或设置应用程序窗口的高度
    /// </summary>
    /// <inheritdoc/>
    public float Height
    {
        get => _application?.Height ?? 0;
        set { if (_application != null) _application.Height = Convert.ToInt32(value); }
    }

    /// <summary>
    /// 获取或设置应用程序窗口的状态
    /// </summary>
    public int WindowState
    {
        get => _application?.WindowState != null ? (int)_application?.WindowState : (int)WdWindowState.wdWindowStateNormal;
        set
        {
            if (_application != null) _application.WindowState = (MsWord.WdWindowState)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置指定文档窗口或任务窗口的状态。
    /// </summary>
    /// <inheritdoc/>
    public WdWindowState WordWindowState
    {
        get => _application?.WindowState != null ? (WdWindowState)(int)_application?.WindowState : WdWindowState.wdWindowStateNormal;
        set
        {
            if (_application != null) _application.WindowState = (MsWord.WdWindowState)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序窗口的描述文字文本。
    /// </summary>
    /// <inheritdoc/>
    public string Caption
    {
        get => _application?.Caption ?? string.Empty;
        set { if (_application != null) _application.Caption = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示状态栏。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayStatusBar
    {
        get => _application?.DisplayStatusBar ?? false;
        set { if (_application != null) _application.DisplayStatusBar = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示滚动条。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayScrollBars
    {
        get => _application?.DisplayScrollBars ?? false;
        set { if (_application != null) _application.DisplayScrollBars = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在"文件"菜单上显示最近使用的文件的名称。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayRecentFiles
    {
        get => _application?.DisplayRecentFiles ?? false;
        set { if (_application != null) _application.DisplayRecentFiles = value; }
    }

    /// <summary>
    /// 获取 Word 文档窗口可设置的最大宽度（以磅为单位）。
    /// </summary>
    /// <inheritdoc/>
    public int UsableWidth => _application?.UsableWidth ?? 0;

    /// <summary>
    /// 获取 Word 文档窗口的高度设置为最大高度 (以磅为单位)。
    /// </summary>
    /// <inheritdoc/>
    public int UsableHeight => _application?.UsableHeight ?? 0;

    /// <summary>
    /// 获取或设置一个值，该值指示应用程序是否可见
    /// </summary>
    /// <inheritdoc/>
    public bool Visible
    {
        get => _application?.Visible ?? false;
        set { if (_application != null) _application.Visible = value; }
    }

    #endregion

    #region 应用程序对象属性

    /// <summary>
    /// 获取表示活动文档的 Document 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordDocument? ActiveDocument => _application?.ActiveDocument != null ? new WordDocument(_application.ActiveDocument) : null;

    /// <summary>
    /// 获取表示活动窗口的 Window 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordWindow? ActiveWindow => _application?.ActiveWindow != null ? new WordWindow(_application.ActiveWindow) : null;

    /// <summary>
    /// 获取表示所有打开的文档的 Documents 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordDocuments? Documents => _application?.Documents != null ? new WordDocuments(_application.Documents) : null;

    /// <summary>
    /// 获取表示所有可用模板的 Templates 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordTemplates? Templates => _application?.Templates != null ? new WordTemplates(_application.Templates) : null;

    /// <summary>
    /// 获取表示所有可用加载项的 AddIns 集合。
    /// </summary>
    public IWordAddIns? AddIns => _application?.AddIns != null ? new WordAddIns(_application.AddIns) : null;

    /// <summary>
    /// 获取表示 Normal 模板的 Template 对象。
    /// </summary>
    public IWordTemplate? NormalTemplate => _application?.NormalTemplate != null ? new WordTemplate(_application.NormalTemplate) : null;

    /// <summary>
    /// 获取表示所有文档窗口的 Windows 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordWindows? Windows => _application?.Windows != null ? new WordWindows(_application.Windows) : null;

    #endregion

    #region 应用程序控制方法

    /// <summary>
    /// 激活应用程序
    /// </summary>
    /// <inheritdoc/>
    public void Activate()
    {
        _application?.Activate();
    }

    /// <summary>
    /// 退出 Microsoft Word 应用程序。
    /// </summary>
    public void Quit()
    {
        _application?.Quit();
    }

    /// <summary>
    /// 退出 Microsoft Word 应用程序。
    /// </summary>
    /// <inheritdoc/>
    public void Quit(
        WdSaveOptions? saveChanges = null,
        WdOriginalFormat? originalFormat = null,
        bool? routeDocument = null)
    {
        var originalFormatObj = Type.Missing;
        if (originalFormat != null)
            originalFormatObj = (MsWord.WdOriginalFormat)(int)originalFormat;

        var saveChangesObj = Type.Missing;
        if (saveChanges != null)
            saveChangesObj = (MsWord.WdSaveOptions)(int)saveChanges;

        _application?.Quit(
           saveChangesObj,
            originalFormatObj, routeDocument.ComArgsVal());
    }

    #endregion

    #region 打印方法

    /// <summary>
    /// 打印当前文档或选定内容。
    /// </summary>
    /// <inheritdoc/>
    public void PrintOut(ref object background, ref object append, ref object range, ref object outputFileName,
                         ref object from, ref object to, ref object item, ref object copies, ref object pages,
                         ref object pageType, ref object printToFile, ref object collate, ref object fileName,
                         ref object lineEnding, ref object outputPrinterName)
    {
        _application?.PrintOut(ref background, ref append, ref range, ref outputFileName,
                               ref from, ref to, ref item, ref copies, ref pages,
                               ref pageType, ref printToFile, ref collate, ref fileName,
                               ref lineEnding, ref outputPrinterName);
    }

    /// <summary>
    /// 将文档另存为 PDF 或 XPS 格式。
    /// </summary>
    /// <inheritdoc/>
    public void ExportAsFixedFormat(string outputFileName,
        WdExportFormat exportFormat,
        bool openAfterExport = false,
        WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint,
        WdExportRange range = WdExportRange.wdExportAllDocument,
        int from = 1, int to = 1,
        WdExportItem item = WdExportItem.wdExportDocumentContent,
        bool includeDocProps = false,
        bool keepIRM = true,
        WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
        bool docStructureTags = true,
        bool bitmapMissingFonts = true,
        bool useISO19005_1 = false,
         object fixedFormatExtClassPtr = null)
    {
        _application?.ActiveDocument?.ExportAsFixedFormat(
            outputFileName, exportFormat, openAfterExport,
            optimizeFor, range, from, to, item, includeDocProps,
            keepIRM, createBookmarks, docStructureTags,
            bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr);
    }

    #endregion

    #region 宏和运行方法

    /// <summary>
    /// 运行指定的宏
    /// </summary>
    /// <param name="macroName">要运行的宏名称</param>
    /// <param name="args">传递给宏的参数</param>
    /// <returns>宏执行结果</returns>
    public object Run(string macroName, params object[] args)
    {
        if (_application == null || string.IsNullOrEmpty(macroName))
            return null;

        try
        {
            object result = _application.Run(macroName, args);
            return result;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 选择和选项属性

    /// <summary>
    /// 获取表示所选区域或插入点的 Selection 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordSelection? Selection => _application?.Selection != null ? new WordSelection(_application.Selection) : null;

    /// <summary>
    /// 获取或设置一个值，该值指示运行宏时的一些警告和消息的处理的方式。
    /// </summary>
    /// <inheritdoc/>
    public WdAlertLevel DisplayAlerts
    {
        get => _application?.DisplayAlerts != null ? (WdAlertLevel)(int)_application?.DisplayAlerts : WdAlertLevel.wdAlertsNone;
        set
        {
            if (_application != null) _application.DisplayAlerts = (MsWord.WdAlertLevel)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时显示自动完成提示。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayAutoCompleteTips
    {
        get => _application?.DisplayAutoCompleteTips ?? false;
        set { if (_application != null) _application.DisplayAutoCompleteTips = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否将批注、脚注、尾注和超链接显示为提示。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayScreenTips
    {
        get => _application?.DisplayScreenTips ?? false;
        set { if (_application != null) _application.DisplayScreenTips = value; }
    }

    /// <summary>
    /// 获取表示 Microsoft Word 中应用程序设置的 Options 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordOptions? Options =>
         _application?.Options != null ? new WordOptions(_application.Options) : null;

    /// <summary>
    /// 获取或设置一个值，该值指示 Word 处理 Ctrl+Break 用户中断的方式。
    /// </summary>
    /// <inheritdoc/>
    public WdEnableCancelKey EnableCancelKey
    {
        get => _application?.EnableCancelKey != null ? (WdEnableCancelKey)(int)_application?.EnableCancelKey : WdEnableCancelKey.wdCancelDisabled;
        set
        {
            if (_application != null) _application.EnableCancelKey = (MsWord.WdEnableCancelKey)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 在键入时是否自动检测所使用的语言。
    /// </summary>
    /// <inheritdoc/>
    public bool CheckLanguage
    {
        get => _application?.CheckLanguage ?? false;
        set { if (_application != null) _application.CheckLanguage = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否打开屏幕更新。
    /// </summary>
    /// <inheritdoc/>
    public bool ScreenUpdating
    {
        get => _application?.ScreenUpdating ?? false;
        set { if (_application != null) _application.ScreenUpdating = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否打开拼写和语法检查。
    /// </summary>
    /// <inheritdoc/>
    public bool CheckSpellingAsYouType
    {
        get => _application?.Options?.CheckSpellingAsYouType ?? false;
        set { if (_application?.Options != null) _application.Options.CheckSpellingAsYouType = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时检查语法。
    /// </summary>
    /// <inheritdoc/>
    public bool CheckGrammarAsYouType
    {
        get => _application?.Options?.CheckGrammarAsYouType ?? false;
        set { if (_application?.Options != null) _application.Options.CheckGrammarAsYouType = value; }
    }

    #endregion

    #region 语言和字典属性

    /// <summary>
    /// 获取表示"语言"对话框中列出的校对语言的 Languages 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordLanguages? Languages => _application?.Languages != null ? new WordLanguages(_application.Languages) : null;

    /// <summary>
    /// 获取表示所有可用字体名称的 FontNames 集合。
    /// </summary>
    public IWordFontNames? FontNames => _application?.FontNames != null ? new WordFontNames(_application.FontNames) : null;

    /// <summary>
    /// 获取表示所有可用纵向字体名称的 FontNames 集合。
    /// </summary>
    public IWordFontNames? PortraitFontNames => _application?.PortraitFontNames != null ? new WordFontNames(_application.PortraitFontNames) : null;

    /// <summary>
    /// 获取表示所有可用横向字体名称的 FontNames 集合。
    /// </summary>
    public IWordFontNames? LandscapeFontNames => _application?.LandscapeFontNames != null ? new WordFontNames(_application.LandscapeFontNames) : null;

    /// <summary>
    /// 获取表示活动自定义字典集合的 Dictionaries 对象。
    /// </summary>
    public IWordDictionaries? CustomDictionaries => _application?.CustomDictionaries != null ? new WordDictionaries(_application.CustomDictionaries) : null;

    /// <summary>
    /// 获取表示当前自动更正选项、条目和异常的 AutoCorrect 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordAutoCorrect? AutoCorrect =>
        _application?.AutoCorrect != null ? new WordAutoCorrect(_application.AutoCorrect) : null;

    /// <summary>
    /// 获取表示对电子邮件进行的自动更正的 AutoCorrect 对象。
    /// </summary>
    public IWordAutoCorrect? AutoCorrectEmail =>
        _application?.AutoCorrectEmail != null ? new WordAutoCorrect(_application.AutoCorrectEmail) : null;

    /// <summary>
    /// 获取表示项目符号、编号和大纲编号模板库的 ListGalleries 集合。
    /// </summary>
    public IWordListGalleries? ListGalleries =>
        _application?.ListGalleries != null ? new WordListGalleries(_application.ListGalleries) : null;

    #endregion

    #region 文件和用户属性

    /// <summary>
    /// 获取表示最近访问的文件的 RecentFiles 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordRecentFiles? RecentFiles =>
        _application?.RecentFiles != null ? new WordRecentFiles(_application.RecentFiles) : null;

    /// <summary>
    /// 获取或设置启动文件夹的完整路径（不包括最后的分隔符）。
    /// </summary>
    /// <inheritdoc/>
    public string StartupPath
    {
        get => _application?.StartupPath ?? string.Empty;
        set { if (_application != null) _application.StartupPath = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置用户的邮件地址。
    /// </summary>
    /// <inheritdoc/>
    public string UserAddress
    {
        get => _application?.UserAddress ?? string.Empty;
        set { if (_application != null) _application.UserAddress = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置用户的姓名缩写。
    /// </summary>
    /// <inheritdoc/>
    public string UserInitials
    {
        get => _application?.UserInitials ?? string.Empty;
        set { if (_application != null) _application.UserInitials = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置用户名。
    /// </summary>
    /// <inheritdoc/>
    public string UserName
    {
        get => _application?.UserName ?? string.Empty;
        set { if (_application != null) _application.UserName = value ?? string.Empty; }
    }

    #endregion

    #region 文档操作方法

    /// <summary>
    /// 创建一个空白文档
    /// </summary>
    /// <returns>新建的文档对象</returns>
    public IWordDocument BlankDocument()
    {
        try
        {
            var doc = _application.Documents.Add();
            var wordDoc = new WordDocument(doc);
            MemorizeActiveDocument(wordDoc);
            return wordDoc;
        }
        catch (Exception ex)
        {
            log.Error("Failed to create blank document.", ex);
            throw new InvalidOperationException("Failed to create blank document.", ex);
        }
    }

    /// <summary>
    /// 打开一个现有文档。
    /// </summary>
    /// <inheritdoc/>
    public IWordDocument? OpenDocument(string fileName, bool confirmConversions = true, bool readOnly = false, bool addToRecentFiles = true,
                                     string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "",
                                     string writePasswordTemplate = "", WdOpenFormat format = WdOpenFormat.wdOpenFormatAuto,
                                     MsoEncoding encoding = MsoEncoding.msoEncodingSimplifiedChineseAutoDetect, bool visible = true)
    {
        if (_application == null || string.IsNullOrWhiteSpace(fileName)) return null;

        try
        {
            var document = _application.Documents.Open(fileName, confirmConversions, readOnly, addToRecentFiles,
                                                     passwordDocument, passwordTemplate, revert,
                                                     writePasswordDocument, writePasswordTemplate, format,
                                                     (MsCore.MsoEncoding)(int)encoding, visible);
            return document != null ? new WordDocument(document) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to open document '{fileName}': {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// 新建一个文档。
    /// </summary>
    /// <inheritdoc/>
    public IWordDocument? NewDocument(object template, object newTemplate)
    {
        if (_application?.Documents == null) return null;

        try
        {
            var document = _application.Documents.Add(ref template, ref newTemplate);
            return document != null ? new WordDocument(document) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to create new document: {ex.Message}", ex);
            return null;
        }
    }

    /// <summary>
    /// 保护文档。
    /// </summary>
    /// <inheritdoc/>
    public void Protect(MsWord.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock)
    {
        _application?.ActiveDocument?.Protect(type, ref noReset, ref password, ref useIRM, ref enforceStyleLock);
    }

    /// <summary>
    /// 取消保护文档。
    /// </summary>
    /// <inheritdoc/>
    public void Unprotect(object password)
    {
        _application?.ActiveDocument?.Unprotect(ref password);
    }

    /// <summary>
    /// 保存所有打开的文档。
    /// </summary>
    /// <inheritdoc/>
    public void SaveAll()
    {
        _application?.Documents?.Save(MissingValue, MissingValue);
    }

    #endregion

    #region 查找和替换方法

    /// <summary>
    /// 执行查找操作。
    /// </summary>
    /// <inheritdoc/>
    public bool FindText(string findText)
    {
        if (_application?.Selection?.Find == null || string.IsNullOrWhiteSpace(findText)) return false;

        var find = _application.Selection.Find;
        find.ClearFormatting();
        find.Text = findText;
        return find.Execute();
    }

    /// <summary>
    /// 替换文本。
    /// </summary>
    /// <inheritdoc/>
    public int ReplaceText(string findText, string replaceWith, MsWord.WdReplace replace)
    {
        if (_application?.Selection?.Find == null) return 0;

        var find = _application.Selection.Find;
        find.ClearFormatting();
        find.Text = findText ?? string.Empty;
        find.Replacement.ClearFormatting();
        find.Replacement.Text = replaceWith ?? string.Empty;

        // 执行替换所有操作
        if (replace == MsWord.WdReplace.wdReplaceAll)
        {
            int count = 0;
            while (find.Execute(
                findText,
                 MissingValue, MissingValue, MissingValue, MissingValue,
                MissingValue, MissingValue, MissingValue, MissingValue,
                replaceWith, replace, MissingValue, MissingValue,
                MissingValue, MissingValue))
            {
                count++;
            }
            return count;
        }
        else
        {
            // 执行单次替换或查找
            return find.Execute(
                findText,
                MissingValue, MissingValue, MissingValue, MissingValue,
                MissingValue, MissingValue, MissingValue, MissingValue,
                replaceWith, replace, MissingValue, MissingValue,
                MissingValue, MissingValue) ? 1 : 0;
        }
    }

    #endregion

    #region 国际化方法和属性

    /// <summary>
    /// 获取有关当前国家/地区和国际设置的信息。
    /// </summary>
    /// <inheritdoc/>
    public object GetInternational(WdInternationalIndex index)
    {
        return _application?.International[(MsWord.WdInternationalIndex)(int)index];
    }

    /// <summary>
    /// 获取一个值，该值指示引用对象的指定变量是否有效。
    /// </summary>
    /// <inheritdoc/>
    public bool IsObjectValid(object obj)
    {
        return _application?.IsObjectValid[obj] ?? false;
    }

    #endregion

    #region 文件转换和任务属性

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 如何处理调用需要尚未安装的功能的方法和属性。
    /// </summary>
    /// <inheritdoc/>
    public MsoFeatureInstall FeatureInstall
    {
        get => _application?.FeatureInstall != null ? (MsoFeatureInstall)(int)_application.FeatureInstall : MsoFeatureInstall.msoFeatureInstallNone;
        set
        {
            if (_application != null) _application.FeatureInstall = (MsCore.MsoFeatureInstall)(int)value;
        }
    }

    /// <summary>
    /// 获取表示所有可用文件转换器的 FileConverters 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordFileConverters? FileConverters => _application?.FileConverters != null ? new WordFileConverters(_application.FileConverters) : null;

    /// <summary>
    /// 获取表示所有正在运行的应用程序的 Tasks 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordTasks? Tasks => _application?.Tasks != null ? new WordTasks(_application.Tasks) : null;

    /// <summary>
    /// 获取表示所有内置对话框的 Dialogs 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordDialogs? Dialogs => _application?.Dialogs != null ? new WordDialogs(_application.Dialogs) : null;

    /// <summary>
    /// 获取表示所有自定义键绑定的 KeyBindings 集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordKeyBindings? KeyBindings => _application?.KeyBindings != null ? new WordKeyBindings(_application.KeyBindings) : null;

    /// <summary>
    /// 获取表示所有已加载的 COM 加载项的 COMAddIns 集合。
    /// </summary>
    /// <inheritdoc/>
    public object? COMAddIns => _application?.COMAddIns;

    /// <summary>
    /// 获取表示命令栏的 CommandBars 对象
    /// </summary>
    /// <inheritdoc/>
    public IOfficeCommandBars? CommandBars => _application?.CommandBars != null ? new OfficeCommandBars(_application.CommandBars) : null;

    #endregion

    #region 邮件相关属性

    /// <summary>
    /// 获取表示电子邮件创作的全局首选项的 EmailOptions 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordEmailOptions? EmailOptions => _application?.EmailOptions != null ? new WordEmailOptions(_application.EmailOptions) : null;

    /// <summary>
    /// 获取或设置用于电子邮件的模板。
    /// </summary>
    /// <inheritdoc/>
    public string EmailTemplate
    {
        get => _application?.EmailTemplate ?? string.Empty;
        set { if (_application != null) _application.EmailTemplate = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取表示邮件标签的 MailingLabel 对象。
    /// </summary>
    /// <inheritdoc/>
    public IWordMailingLabel? MailingLabel => _application?.MailingLabel != null ? new WordMailingLabel(_application.MailingLabel) : null;

    /// <summary>
    /// 获取表示活动电子邮件的 MailMessage 对象。
    /// </summary>
    public IWordMailMessage? MailMessage => _application?.MailMessage != null ? new WordMailMessage(_application.MailMessage) : null;

    /// <summary>
    /// 获取邮件系统的类型。
    /// </summary>
    public WdMailSystem MailSystem => _application?.MailSystem != null ? (WdMailSystem)(int)_application.MailSystem : WdMailSystem.wdNoMailSystem;

    /// <summary>
    /// 获取一个值，该值指示是否安装了 MAPI。
    /// </summary>
    public bool MAPIAvailable => _application?.MAPIAvailable ?? false;

    /// <summary>
    /// 获取一个值，该值指示插入点是否位于电子邮件标头字段中。
    /// </summary>
    public bool FocusInMailHeader => _application?.FocusInMailHeader ?? false;

    /// <summary>
    /// 获取或设置一个值，该值指示是否在全屏模式下打开附件。
    /// </summary>
    /// <inheritdoc/>
    public bool OpenAttachmentsInFullScreen
    {
        get => _application?.OpenAttachmentsInFullScreen ?? false;
        set { if (_application != null) _application.OpenAttachmentsInFullScreen = value; }
    }

    #endregion

    #region 安全和限制属性

    /// <summary>
    /// 获取或设置自动化安全级别。
    /// </summary>
    /// <inheritdoc/>
    public MsoAutomationSecurity AutomationSecurity
    {
        get => _application?.AutomationSecurity != null ? (MsoAutomationSecurity)(int)_application?.AutomationSecurity : MsoAutomationSecurity.msoAutomationSecurityByUI;
        set
        {
            if (_application != null) _application.AutomationSecurity = (MsCore.MsoAutomationSecurity)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置文件验证方式。
    /// </summary>
    /// <inheritdoc/>
    public MsoFileValidationMode FileValidation
    {
        get => _application?.FileValidation != null ? (MsoFileValidationMode)(int)_application?.FileValidation : MsoFileValidationMode.msoFileValidationDefault;
        set
        {
            if (_application != null) _application.FileValidation = (MsCore.MsoFileValidationMode)(int)value;
        }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否限制链接样式。
    /// </summary>
    /// <inheritdoc/>
    public bool RestrictLinkedStyles
    {
        get => _application?.RestrictLinkedStyles ?? false;
        set { if (_application != null) _application.RestrictLinkedStyles = value; }
    }

    #endregion

    #region 系统信息属性

    /// <summary>
    /// 获取一个值，该值指示是否安装了数学协处理器。
    /// </summary>
    public bool MathCoprocessorAvailable => _application?.MathCoprocessorAvailable ?? false;

    /// <summary>
    /// 获取一个值，该值指示是否有可用于系统的鼠标。
    /// </summary>
    public bool MouseAvailable => _application?.MouseAvailable ?? false;

    /// <summary>
    /// 获取 NUM LOCK 键的状态。
    /// </summary>
    public bool NumLock => _application?.NumLock ?? false;

    /// <summary>
    /// 获取 CAPS LOCK 键的状态。
    /// </summary>
    public bool CapsLock => _application?.CapsLock ?? false;

    /// <summary>
    /// 获取 Word 应用程序的内部版本号。
    /// </summary>
    public string Build => _application?.Build ?? string.Empty;

    /// <summary>
    /// 获取一个值，该值指示文档或应用程序是否由用户创建或打开。
    /// </summary>
    /// <inheritdoc/>
    public bool UserControl
    {
        get => _application?.UserControl ?? false;
    }

    #endregion

    #region 文件对话框方法

    /// <summary>
    /// 创建文件对话框
    /// </summary>
    /// <param name="fileDialogType">文件对话框类型</param>
    /// <returns>文件对话框对象</returns>
    public IOfficeFileDialog CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        var dialog = _application?.FileDialog[(MsCore.MsoFileDialogType)(int)fileDialogType];
        return dialog != null ? new OfficeFileDialog(dialog) : null;
    }

    /// <summary>
    /// 获取文件对话框。
    /// </summary>
    /// <inheritdoc/>
    public IOfficeFileDialog? FileDialog(MsoFileDialogType fileDialogType)
    {
        return CreateFileDialog(fileDialogType);
    }

    #endregion

    #region 智能标记属性

    /// <summary>
    /// 获取智能标记识别器集合。
    /// </summary>
    /// <inheritdoc/>
    public IWordSmartTagRecognizers? SmartTagRecognizers => _application?.SmartTagRecognizers != null ? new WordSmartTagRecognizers(_application?.SmartTagRecognizers) : null;

    /// <summary>
    /// 获取智能标记类型集合。
    /// </summary>
    public IWordSmartTagTypes? SmartTagTypes => _application?.SmartTagTypes != null ? new WordSmartTagTypes(_application.SmartTagTypes) : null;

    #endregion

    #region 其他属性和方法

    /// <summary>
    /// 获取一个值，该值指示是否支持任意 XML。
    /// </summary>
    /// <inheritdoc/>
    public bool ArbitraryXMLSupportAvailable =>
        _application?.ArbitraryXMLSupportAvailable ?? false;

    /// <summary>
    /// 获取表示 Microsoft Office 帮助查看器的 IAssistance 对象。
    /// </summary>
    public IOfficeAssistance? Assistance =>
        _application?.Assistant != null ? new OfficeAssistance(_application.Assistant) : null;

    /// <summary>
    /// 获取表示在将表格和图片等项目插入文档中时自动添加的标题的 AutoCaptions 集合。
    /// </summary>
    public IWordAutoCaptions? AutoCaptions =>
        _application?.AutoCaptions != null ? new WordAutoCaptions(_application.AutoCaptions) : null;

    /// <summary>
    /// 获取后台打印队列中打印作业的编号。
    /// </summary>
    public int BackgroundPrintingStatus => _application?.BackgroundPrintingStatus ?? 0;

    /// <summary>
    /// 获取排队在后台保存的文件数。
    /// </summary>
    public int BackgroundSavingStatus => _application?.BackgroundSavingStatus ?? 0;

    /// <summary>
    /// 获取或设置一个值，该值指示是否可以使用 Word 打开 HTML 文件。
    /// </summary>
    /// <inheritdoc/>
    public string BrowseExtraFileTypes
    {
        get => _application?.BrowseExtraFileTypes ?? string.Empty;
        set { if (_application != null) _application.BrowseExtraFileTypes = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取表示垂直滚动条上的"选择浏览对象"工具的 Browser 对象。
    /// </summary>
    public IWordBrowser? Browser =>
         _application?.Browser != null ? new WordBrowser(_application.Browser) : null;

    /// <summary>
    /// 获取 Word 应用程序的内部版本号。
    /// </summary>
    public string BuildFull => _application?.BuildFull ?? string.Empty;

    /// <summary>
    /// 获取或设置一个值，该值指示在比较和合并文档时是否默认使用"法律黑线"选项。
    /// </summary>
    /// <inheritdoc/>
    public bool DefaultLegalBlackline
    {
        get => _application?.DefaultLegalBlackline ?? false;
        set { if (_application != null) _application.DefaultLegalBlackline = value; }
    }

    /// <summary>
    /// 获取或设置在"另存为"对话框中的"另存为类型"框中显示的默认格式。
    /// </summary>
    /// <inheritdoc/>
    public string DefaultSaveFormat
    {
        get => _application?.DefaultSaveFormat ?? string.Empty;
        set { if (_application != null) _application.DefaultSaveFormat = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。
    /// </summary>
    /// <inheritdoc/>
    public string DefaultTableSeparator
    {
        get => _application?.DefaultTableSeparator ?? string.Empty;
        set { if (_application != null) _application.DefaultTableSeparator = value ?? string.Empty; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示文档信息面板。
    /// </summary>
    /// <inheritdoc/>
    public bool DisplayDocumentInformationPanel
    {
        get => _application?.DisplayDocumentInformationPanel ?? false;
        set { if (_application != null) _application.DisplayDocumentInformationPanel = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在全屏模式下打开附件。
    /// </summary>
    /// <inheritdoc/>
    public bool DontResetInsertionPointProperties
    {
        get => _application?.DontResetInsertionPointProperties ?? false;
        set { if (_application != null) _application.DontResetInsertionPointProperties = value; }
    }

    /// <summary>
    /// 获取表示所有活动的自定义转换字典的 HangulHanjaConversionDictionaries 集合。
    /// </summary>
    public IWordHangulHanjaConversionDictionaries? HangulHanjaDictionaries =>
        _application?.HangulHanjaDictionaries != null ? new WordHangulHanjaConversionDictionaries(_application.HangulHanjaDictionaries) : null;

    /// <summary>
    /// 获取或设置一个值，该值指示是否在受保护的视图中打开文件。
    /// </summary>
    public bool IsSandboxed => _application?.IsSandboxed ?? false;

    /// <summary>
    /// 获取表示所选 Microsoft Word 用户界面的语言设置。
    /// </summary>
    public MsoLanguageID Language =>
        _application?.Language != null ? (MsoLanguageID)(int)_application.Language : MsoLanguageID.msoLanguageIDSimplifiedChinese;

    /// <summary>
    /// 获取表示语言设置的 LanguageSettings 对象
    /// </summary>
    public IOfficeLanguageSettings? LanguageSettings =>
        _application?.LanguageSettings != null ? new OfficeLanguageSettings(_application.LanguageSettings) : null;

    /// <summary>
    /// 获取表示在其中存储包含正在运行过程的模块的模板或文档的 Template 或 Document 对象。
    /// </summary>
    public object MacroContainer => _application?.MacroContainer;

    /// <summary>
    /// 获取表示公式的自动更正条目的 OMathAutoCorrect 对象。
    /// </summary>
    public IWordOMathAutoCorrect? OMathAutoCorrect =>
         _application?.OMathAutoCorrect != null ? new WordOMathAutoCorrect(_application.OMathAutoCorrect) : null;

    /// <summary>
    /// 获取一个 PickerDialog 对象，该对象提供在对话框中选择人员或数据的功能。
    /// </summary>
    public IOfficePickerDialog? PickerDialog =>
        _application?.PickerDialog != null ? new OfficePickerDialog(_application.PickerDialog) : null;

    /// <summary>
    /// 获取一个值，该值指示打印预览是否为当前视图。
    /// </summary>
    /// <inheritdoc/>
    public bool PrintPreview
    {
        get => _application?.PrintPreview ?? false;
        set { if (_application != null) _application.PrintPreview = value; }
    }

    /// <summary>
    /// 获取表示所有受保护的视图窗口的 ProtectedViewWindows 集合。
    /// </summary>
    public IWordProtectedViewWindows? ProtectedViewWindows =>
        _application?.ProtectedViewWindows != null ? new WordProtectedViewWindows(_application.ProtectedViewWindows) : null;

    /// <summary>
    /// 获取或设置一个值，该值指示启动 Microsoft Word 时是否显示任务窗格。
    /// </summary>
    /// <inheritdoc/>
    public bool ShowStartupDialog
    {
        get => _application?.ShowStartupDialog ?? false;
        set { if (_application != null) _application.ShowStartupDialog = value; }
    }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示样式预览。
    /// </summary>
    /// <inheritdoc/>
    public bool ShowStylePreviews
    {
        get => _application?.ShowStylePreviews ?? false;
        set { if (_application != null) _application.ShowStylePreviews = value; }
    }

    /// <summary>
    /// 获取一个值，该值指示应用程序是否处于特殊模式（例如 CopyText 模式或 MoveText 模式）。
    /// </summary>
    public bool SpecialMode => _application?.SpecialMode ?? false;

    /// <summary>
    /// 获取或设置状态栏中显示的文本。
    /// </summary>
    /// <inheritdoc/>
    public void SetStatusBar(string text)
    {
        if (_application != null) _application.StatusBar = text ?? string.Empty;
    }

    /// <summary>
    /// 获取表示 Microsoft Word 中最常执行的任务的 TaskPanes 集合。
    /// </summary>
    public IWordTaskPanes? TaskPanes =>
        _application?.TaskPanes != null ? new WordTaskPanes(_application.TaskPanes) : null;

    /// <summary>
    /// 获取一个 UndoRecord 对象，该对象提供撤消堆栈中的自定义入口点。
    /// </summary>
    public IWordUndoRecord? UndoRecord =>
        _application?.UndoRecord != null ? new WordUndoRecord(_application.UndoRecord) : null;

    /// <summary>
    /// 获取自动化对象 (Word.Basic) ，其中包括 Microsoft Word 6.0 版和 Windows 95 Word 中提供的所有 WordBasic 语句和函数的方法。
    /// </summary>
    public object WordBasic => _application?.WordBasic;

    /// <summary>
    /// 获取表示当前在应用程序中加载的一组颜色样式的 SmartArtColors 对象。
    /// </summary>
    public IOfficeSmartArtColors? SmartArtColors =>
        _application?.SmartArtColors != null ? new OfficeSmartArtColors(_application.SmartArtColors) : null;

    /// <summary>
    /// 获取表示当前在应用程序中加载的 SmartArt 布局集的 SmartArtLayouts 对象。
    /// </summary>
    public IOfficeSmartArtLayouts? SmartArtLayouts =>
         _application?.SmartArtLayouts != null ? new OfficeSmartArtLayouts(_application.SmartArtLayouts) : null;

    /// <summary>
    /// 获取表示应用程序中当前加载的 SmartArt 样式集的 SmartArtQuickStyles 对象。
    /// </summary>
    public IOfficeSmartArtQuickStyles? SmartArtQuickStyles =>
         _application?.SmartArtQuickStyles != null ? new OfficeSmartArtQuickStyles(_application.SmartArtQuickStyles) : null;

    /// <summary>
    /// 获取表示活动加密会话的对象。
    /// </summary>
    public int ActiveEncryptionSession =>
         _application?.ActiveEncryptionSession != null ? _application.ActiveEncryptionSession : 0;

    /// <summary>
    /// 获取或设置一个值，该值指示图表数据点是否被跟踪。
    /// </summary>
    /// <inheritdoc/>
    public bool ChartDataPointTrack
    {
        get => _application?.ChartDataPointTrack ?? false;
        set { if (_application != null) _application.ChartDataPointTrack = value; }
    }

    /// <summary>
    /// 获取一个 FileSearch 对象，该对象可用于使用绝对路径或相对路径搜索文件。
    /// </summary>
    public IOfficeFileSearch? FileSearch =>
        _application?.FileSearch != null ? new OfficeFileSearch(_application.FileSearch) : null;

    #endregion

    #region 系统信息方法

    /// <summary>
    /// 获取系统信息
    /// </summary>
    /// <returns>系统信息对象</returns>
    public IWordSystemInfo GetSystemInfo()
    {
        try
        {
            return new WordSystemInfo
            {
                OSVersion = Environment.OSVersion.ToString(),
                TotalMemory = Environment.WorkingSet,
                AvailableMemory = 0, // .NET 中没有直接获取可用内存的方法
                ProcessorCount = Environment.ProcessorCount,
                SystemUpTime = DateTime.Now - TimeSpan.FromTicks(Environment.TickCount)
            };
        }
        catch (Exception ex)
        {
            log.Error("Failed to get system information.", ex);
            throw new InvalidOperationException("Failed to get system information.", ex);
        }
    }

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 WordApplication 实例
    /// </summary>
    internal WordApplication()
    {
        _application = new MsWord.Application();
        _applicationEvent = _application;
        InitializeApp();
        ConnectEvents();
    }

    /// <summary>
    /// 使用现有的 Word 应用程序实例初始化 WordApplication
    /// </summary>
    /// <param name="application">现有的 Word 应用程序实例</param>
    internal WordApplication(MsWord.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _applicationEvent = application;
        InitializeApp();
        ConnectEvents();
    }

    #endregion

    #region 私有方法

    /// <summary>
    /// 初始化应用程序
    /// </summary>
    private void InitializeApp()
    {
        _application.DisplayAlerts = MsWord.WdAlertLevel.wdAlertsMessageBox;
        _disposedValue = false;
        _activeDocument = null;
    }

    /// <summary>
    /// 记住活动文档
    /// </summary>
    /// <param name="document">活动文档</param>
    private void MemorizeActiveDocument(IWordDocument document)
    {
        _activeDocument = document;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放相关对象
            _selection?.Dispose();
            _documents?.Dispose();
            _windows?.Dispose();
            DisconnectEvents();
            if (_application != null)
            {
                try
                {
                    _application.DisplayAlerts = MsWord.WdAlertLevel.wdAlertsAll;

                    if (Visibility == WordAppVisibility.Hidden)
                    {
                        object saveChanges = MsWord.WdSaveOptions.wdDoNotSaveChanges;
                        object originalFormat = MissingValue;
                        object routeDocument = MissingValue;
                        _application.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                    }
                    else
                    {
                        _application.Visible = true;
                    }

                    try { while (Marshal.ReleaseComObject(_application) > 0) ; } catch { }
                }
                catch
                {
                    // 忽略释放失败的情况
                }
            }

            GC.Collect();
        }

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

    #endregion
}