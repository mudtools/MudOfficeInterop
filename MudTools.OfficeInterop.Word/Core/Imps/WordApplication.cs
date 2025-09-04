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


    public WordAppVisibility Visibility
    {
        get => _application.Visible ? WordAppVisibility.Visible : WordAppVisibility.Hidden;
        set => _application.Visible = value == WordAppVisibility.Visible;
    }


    #region 基本属性实现 (Basic Properties Implementation)
    /// <inheritdoc/>
    public object Parent => _application?.Parent;

    /// <inheritdoc/>
    public int Creator => _application?.Creator ?? 0;

    /// <inheritdoc/>
    public string ActivePrinter
    {
        get => _application?.ActivePrinter ?? string.Empty;
        set { if (_application != null) _application.ActivePrinter = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public IWordDocument? ActiveDocument => _application?.ActiveDocument != null ? new WordDocument(_application.ActiveDocument) : null;

    /// <inheritdoc/>
    public IWordWindow? ActiveWindow => _application?.ActiveWindow != null ? new WordWindow(_application.ActiveWindow) : null;

    /// <inheritdoc/>
    public IWordDocuments? Documents => _application?.Documents != null ? new WordDocuments(_application.Documents) : null;

    /// <inheritdoc/>
    public IWordTemplates? Templates => _application?.Templates != null ? new WordTemplates(_application.Templates) : null;

    public IWordAddIns? AddIns => _application?.AddIns != null ? new WordAddIns(_application.AddIns) : null;

    public IWordTemplate? NormalTemplate => _application?.NormalTemplate != null ? new WordTemplate(_application.NormalTemplate) : null;

    /// <inheritdoc/>
    public string Name
    {
        get => _application?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public string Version => _application?.Version ?? string.Empty;

    /// <inheritdoc/>
    public string Path => _application?.Path ?? string.Empty;

    /// <inheritdoc/>
    public string PathSeparator => _application?.PathSeparator ?? string.Empty;

    #endregion

    #region 窗口和显示属性实现 (Window & Display Properties Implementation)

    /// <inheritdoc/>
    public float Left
    {
        get => _application?.Left ?? 0;
        set { if (_application != null) _application.Left = Convert.ToInt32(value); }
    }

    /// <inheritdoc/>
    public float Top
    {
        get => _application?.Top ?? 0;
        set { if (_application != null) _application.Top = Convert.ToInt32(value); }
    }

    /// <inheritdoc/>
    public float Width
    {
        get => _application?.Width ?? 0;
        set { if (_application != null) _application.Width = Convert.ToInt32(value); }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _application?.Height ?? 0;
        set { if (_application != null) _application.Height = Convert.ToInt32(value); }
    }

    /// <inheritdoc/>
    public WdWindowState WordWindowState
    {
        get => _application?.WindowState != null ? (WdWindowState)(int)_application?.WindowState : WdWindowState.wdWindowStateNormal;
        set
        {
            if (_application != null) _application.WindowState = (MsWord.WdWindowState)(int)value;
        }
    }

    /// <inheritdoc/>
    public string Caption
    {
        get => _application?.Caption ?? string.Empty;
        set { if (_application != null) _application.Caption = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public bool DisplayStatusBar
    {
        get => _application?.DisplayStatusBar ?? false;
        set { if (_application != null) _application.DisplayStatusBar = value; }
    }

    /// <inheritdoc/>
    public bool DisplayScrollBars
    {
        get => _application?.DisplayScrollBars ?? false;
        set { if (_application != null) _application.DisplayScrollBars = value; }
    }

    /// <inheritdoc/>
    public bool DisplayRecentFiles
    {
        get => _application?.DisplayRecentFiles ?? false;
        set { if (_application != null) _application.DisplayRecentFiles = value; }
    }

    /// <inheritdoc/>
    public int UsableWidth => _application?.UsableWidth ?? 0;

    /// <inheritdoc/>
    public int UsableHeight => _application?.UsableHeight ?? 0;

    /// <inheritdoc/>
    public IWordWindows? Windows => _application?.Windows != null ? new WordWindows(_application.Windows) : null;

    #endregion

    #region 基本方法实现 (Basic Methods Implementation)

    /// <inheritdoc/>
    public void Quit(ref object saveChanges, ref object originalFormat, ref object routeDocument)
    {
        _application?.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
    }

    /// <inheritdoc/>
    public void Activate()
    {
        _application?.Activate();
    }

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

    #endregion

    #region 选择和查找属性实现 (Selection & Find Properties Implementation)

    /// <inheritdoc/>
    public IWordSelection? Selection => _application?.Selection != null ? new WordSelection(_application.Selection) : null;

    /// <inheritdoc/>
    public WdAlertLevel DisplayAlerts
    {
        get => _application?.DisplayAlerts != null ? (WdAlertLevel)(int)_application?.DisplayAlerts : WdAlertLevel.wdAlertsNone;
        set
        {
            if (_application != null) _application.DisplayAlerts = (MsWord.WdAlertLevel)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool DisplayAutoCompleteTips
    {
        get => _application?.DisplayAutoCompleteTips ?? false;
        set { if (_application != null) _application.DisplayAutoCompleteTips = value; }
    }

    /// <inheritdoc/>
    public bool DisplayScreenTips
    {
        get => _application?.DisplayScreenTips ?? false;
        set { if (_application != null) _application.DisplayScreenTips = value; }
    }

    #endregion

    #region 选项和设置属性实现 (Options & Settings Properties Implementation)

    /// <inheritdoc/>
    public IWordOptions? Options =>
         _application?.Options != null ? new WordOptions(_application.Options) : null;

    /// <inheritdoc/>
    public WdEnableCancelKey EnableCancelKey
    {
        get => _application?.EnableCancelKey != null ? (WdEnableCancelKey)(int)_application?.EnableCancelKey : WdEnableCancelKey.wdCancelDisabled;
        set
        {
            if (_application != null) _application.EnableCancelKey = (MsWord.WdEnableCancelKey)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool CheckLanguage
    {
        get => _application?.CheckLanguage ?? false;
        set { if (_application != null) _application.CheckLanguage = value; }
    }

    /// <inheritdoc/>
    public bool ScreenUpdating
    {
        get => _application?.ScreenUpdating ?? false;
        set { if (_application != null) _application.ScreenUpdating = value; }
    }

    /// <inheritdoc/>
    public bool CheckSpellingAsYouType
    {
        get => _application?.Options?.CheckSpellingAsYouType ?? false;
        set { if (_application?.Options != null) _application.Options.CheckSpellingAsYouType = value; }
    }

    /// <inheritdoc/>
    public bool CheckGrammarAsYouType
    {
        get => _application?.Options?.CheckGrammarAsYouType ?? false;
        set { if (_application?.Options != null) _application.Options.CheckGrammarAsYouType = value; }
    }

    #endregion

    #region 语言和字典属性实现 (Language & Dictionary Properties Implementation)

    /// <inheritdoc/>
    public IWordLanguages? Languages => _application?.Languages != null ? new WordLanguages(_application.Languages) : null;
    public IWordFontNames? FontNames => _application?.FontNames != null ? new WordFontNames(_application.FontNames) : null;
    public IWordFontNames? PortraitFontNames => _application?.PortraitFontNames != null ? new WordFontNames(_application.PortraitFontNames) : null;
    public IWordFontNames? LandscapeFontNames => _application?.LandscapeFontNames != null ? new WordFontNames(_application.LandscapeFontNames) : null;
    public IWordDictionaries? CustomDictionaries => _application?.CustomDictionaries != null ? new WordDictionaries(_application.CustomDictionaries) : null;

    #endregion

    #region 自动更正和列表属性实现 (AutoCorrect & Lists Properties Implementation)

    /// <inheritdoc/>
    public IWordAutoCorrect? AutoCorrect => _application?.AutoCorrect != null ? new WordAutoCorrect(_application.AutoCorrect) : null;
    public IWordAutoCorrect? AutoCorrectEmail => _application?.AutoCorrectEmail != null ? new WordAutoCorrect(_application.AutoCorrectEmail) : null;
    public IWordListGalleries? ListGalleries => _application?.ListGalleries != null ? new WordListGalleries(_application.ListGalleries) : null;

    #endregion

    #region 文件和模板属性实现 (File & Template Properties Implementation)

    /// <inheritdoc/>
    public IWordRecentFiles RecentFiles => _application?.RecentFiles != null ? new WordRecentFiles(_application.RecentFiles) : null;

    /// <inheritdoc/>
    public string StartupPath
    {
        get => _application?.StartupPath ?? string.Empty;
        set { if (_application != null) _application.StartupPath = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string UserAddress
    {
        get => _application?.UserAddress ?? string.Empty;
        set { if (_application != null) _application.UserAddress = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string UserInitials
    {
        get => _application?.UserInitials ?? string.Empty;
        set { if (_application != null) _application.UserInitials = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string UserName
    {
        get => _application?.UserName ?? string.Empty;
        set { if (_application != null) _application.UserName = value ?? string.Empty; }
    }

    #endregion

    #region 更多方法实现 (More Methods Implementation)

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

    /// <inheritdoc/>
    public bool FindText(string findText)
    {
        if (_application?.Selection?.Find == null || string.IsNullOrWhiteSpace(findText)) return false;

        var find = _application.Selection.Find;
        find.ClearFormatting();
        find.Text = findText;
        return find.Execute();
    }

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




    /// <inheritdoc/>
    public object GetInternational(WdInternationalIndex index)
    {
        return _application?.International[(MsWord.WdInternationalIndex)(int)index];
    }

    #endregion

    #region 自动化和 COM 属性实现 (Automation & COM Properties Implementation)

    /// <inheritdoc/>
    public MsoFeatureInstall FeatureInstall
    {
        get => _application?.FeatureInstall != null ? (MsoFeatureInstall)(int)_application?.FeatureInstall : MsoFeatureInstall.msoFeatureInstallNone;
        set
        {
            if (_application != null) _application.FeatureInstall = (MsCore.MsoFeatureInstall)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool IsObjectValid(object obj)
    {
        return _application?.IsObjectValid[obj] ?? false;
    }

    /// <inheritdoc/>
    public IWordFileConverters? FileConverters => _application?.FileConverters != null ? new WordFileConverters(_application.FileConverters) : null;
    public IWordTasks? Tasks => _application?.Tasks != null ? new WordTasks(_application.Tasks) : null;
    public IWordDialogs? Dialogs => _application?.Dialogs;
    public IWordKeyBindings? KeyBindings => _application?.KeyBindings != null ? new WordKeyBindings(_application.KeyBindings) : null;
    public object? COMAddIns => _application?.COMAddIns;
    public IOfficeCommandBars? CommandBars => _application?.CommandBars != null ? new OfficeCommandBars(_application?.CommandBars) : null;

    #endregion

    #region 邮件相关属性实现 (Mail Properties Implementation)

    /// <inheritdoc/>
    public IWordEmailOptions EmailOptions => _application?.EmailOptions;

    /// <inheritdoc/>
    public string EmailTemplate
    {
        get => _application?.EmailTemplate ?? string.Empty;
        set { if (_application != null) _application.EmailTemplate = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public IWordMailingLabel MailingLabel => _application?.MailingLabel;
    public IWordMailMessage MailMessage => _application?.MailMessage;
    public IWordWdMailSystem MailSystem => _application?.MailSystem ?? MsWord.WdMailSystem.wdNoMailSystem;
    public bool MAPIAvailable => _application?.MAPIAvailable ?? false;
    public bool FocusInMailHeader => _application?.FocusInMailHeader ?? false;

    /// <inheritdoc/>
    public bool OpenAttachmentsInFullScreen
    {
        get => _application?.OpenAttachmentsInFullScreen ?? false;
        set { if (_application != null) _application.OpenAttachmentsInFullScreen = value; }
    }

    #endregion

    #region 安全性相关属性实现 (Security Properties Implementation)

    /// <inheritdoc/>
    public MsoAutomationSecurity AutomationSecurity
    {
        get => _application?.AutomationSecurity != null ? (MsoAutomationSecurity)(int)_application?.AutomationSecurity : MsoAutomationSecurity.msoAutomationSecurityByUI.msoFileValidationDefault;
        set
        {
            if (_application != null) _application.AutomationSecurity = (MsCore.MsoAutomationSecurity)(int)value;
        }
    }

    /// <inheritdoc/>
    public MsoFileValidationMode FileValidation
    {
        get => _application?.FileValidation != null ? (MsoFileValidationMode)(int)_application?.FileValidation : MsoFileValidationMode.msoFileValidationDefault;
        set
        {
            if (_application != null) _application.FileValidation = (MsCore.MsoFileValidationMode)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool RestrictLinkedStyles
    {
        get => _application?.RestrictLinkedStyles ?? false;
        set { if (_application != null) _application.RestrictLinkedStyles = value; }
    }

    #endregion

    #region 系统和环境属性实现 (System & Environment Properties Implementation)

    /// <inheritdoc/>
    public IWordSystem System => _application?.System;
    public bool MathCoprocessorAvailable => _application?.MathCoprocessorAvailable ?? false;
    public bool MouseAvailable => _application?.MouseAvailable ?? false;
    public bool NumLock => _application?.NumLock ?? false;
    public bool CapsLock => _application?.CapsLock ?? false;
    public string Build => _application?.Build ?? string.Empty;

    /// <inheritdoc/>
    public bool UserControl
    {
        get => _application?.UserControl ?? false;
    }

    #endregion

    #region 更多方法实现 (More Methods Implementation)

    /// <inheritdoc/>
    public void Protect(MsWord.WdProtectionType type, object noReset, object password, object useIRM, object enforceStyleLock)
    {
        _application?.ActiveDocument?.Protect(type, ref noReset, ref password, ref useIRM, ref enforceStyleLock);
    }

    /// <inheritdoc/>
    public void Unprotect(object password)
    {
        _application?.ActiveDocument?.Unprotect(ref password);
    }

    /// <inheritdoc/>
    public void SaveAll()
    {
        _application?.Documents?.Save(MissingValue, MissingValue);
    }


    /// <inheritdoc/>
    public IWordSynonymInfo SynonymInfo(string word, object languageID)
    {
        return _application?.SynonymInfo[word, ref languageID];
    }

    /// <inheritdoc/>
    public IOfficeFileDialog FileDialog(MsoFileDialogType fileDialogType)
    {
        var dialog = _application?.FileDialog[(MsCore.MsoFileDialogType)(int)fileDialogType];
        return dialog != null ? new OfficeFileDialog(dialog) : null;
    }

    /// <inheritdoc/>
    public IWordSmartTagRecognizers SmartTagRecognizers => _application?.SmartTagRecognizers;
    public IWordSmartTagTypes SmartTagTypes => _application?.SmartTagTypes;

    #endregion

    #region 剩余属性实现 (Remaining Properties Implementation)

    /// <inheritdoc/>
    public bool ArbitraryXMLSupportAvailable => _application?.ArbitraryXMLSupportAvailable ?? false;
    public object Assistance => _application?.Assistance;
    public MsWord.AutoCaptions AutoCaptions => _application?.AutoCaptions;
    public int BackgroundPrintingStatus => _application?.BackgroundPrintingStatus ?? 0;
    public int BackgroundSavingStatus => _application?.BackgroundSavingStatus ?? 0;
    public MsWord.Bibliography Bibliography => _application?.Bibliography;

    /// <inheritdoc/>
    public string BrowseExtraFileTypes
    {
        get => _application?.BrowseExtraFileTypes ?? string.Empty;
        set { if (_application != null) _application.BrowseExtraFileTypes = value ?? string.Empty; }
    }

    public MsWord.Browser Browser => _application?.Browser;
    public string BuildFull => _application?.BuildFull ?? string.Empty;

    /// <inheritdoc/>
    public bool DefaultLegalBlackline
    {
        get => _application?.DefaultLegalBlackline ?? false;
        set { if (_application != null) _application.DefaultLegalBlackline = value; }
    }

    /// <inheritdoc/>
    public string DefaultSaveFormat
    {
        get => _application?.DefaultSaveFormat ?? string.Empty;
        set { if (_application != null) _application.DefaultSaveFormat = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string DefaultTableSeparator
    {
        get => _application?.DefaultTableSeparator ?? string.Empty;
        set { if (_application != null) _application.DefaultTableSeparator = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public bool DisplayDocumentInformationPanel
    {
        get => _application?.DisplayDocumentInformationPanel ?? false;
        set { if (_application != null) _application.DisplayDocumentInformationPanel = value; }
    }

    /// <inheritdoc/>
    public bool DontResetInsertionPointProperties
    {
        get => _application?.DontResetInsertionPointProperties ?? false;
        set { if (_application != null) _application.DontResetInsertionPointProperties = value; }
    }

    public MsWord.HangulHanjaConversionDictionaries HangulHanjaDictionaries => _application?.HangulHanjaDictionaries;
    public bool IsSandboxed => _application?.IsSandboxed ?? false;
    public MsoLanguageID Language => _application?.Language != null ? (MsoLanguageID)(int)_application?.Language : MsoLanguageID.msoLanguageIDSimplifiedChinese;
    public IOfficeLanguageSettings? LanguageSettings => _application?.LanguageSettings != null ? new OfficeLanguageSettings(_application?.LanguageSettings) : null;
    public object MacroContainer => _application?.MacroContainer;
    public MsWord.OMathAutoCorrect OMathAutoCorrect => _application?.OMathAutoCorrect;
    public object PickerDialog => _application?.PickerDialog;

    /// <inheritdoc/>
    public bool PrintPreview
    {
        get => _application?.PrintPreview ?? false;
        set { if (_application != null) _application.PrintPreview = value; }
    }

    public MsWord.ProtectedViewWindows ProtectedViewWindows => _application?.ProtectedViewWindows;

    /// <inheritdoc/>
    public bool ShowStartupDialog
    {
        get => _application?.ShowStartupDialog ?? false;
        set { if (_application != null) _application.ShowStartupDialog = value; }
    }

    /// <inheritdoc/>
    public bool ShowStylePreviews
    {
        get => _application?.ShowStylePreviews ?? false;
        set { if (_application != null) _application.ShowStylePreviews = value; }
    }

    public bool SpecialMode => _application?.SpecialMode ?? false;

    /// <inheritdoc/>
    public string StatusBar
    {
        get => _application?.StatusBar ?? string.Empty;
        set { if (_application != null) _application.StatusBar = value ?? string.Empty; }
    }

    public IWordTaskPanes TaskPanes => _application?.TaskPanes;
    public object UndoRecord => _application?.UndoRecord;

    /// <inheritdoc/>
    public bool Visible
    {
        get => _application?.Visible ?? false;
        set { if (_application != null) _application.Visible = value; }
    }

    public object WordBasic => _application?.WordBasic;
    public MsWord.XMLNamespaces XMLNamespaces => _application?.XMLNamespaces;
    public object SmartArtColors => _application?.SmartArtColors;
    public object SmartArtLayouts => _application?.SmartArtLayouts;
    public object SmartArtQuickStyles => _application?.SmartArtQuickStyles;
    public object ActiveEncryptionSession => _application?.ActiveEncryptionSession;

    /// <inheritdoc/>
    public bool ChartDataPointTrack
    {
        get => _application?.ChartDataPointTrack ?? false;
        set { if (_application != null) _application.ChartDataPointTrack = value; }
    }

    public IWordFileSearch FileSearch => _application?.FileSearch;

    #endregion

    #region 剩余方法实现 (Remaining Methods Implementation)


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


    internal WordApplication()
    {
        _application = new MsWord.Application();
        _applicationEvent = _application;
        InitializeApp();
        ConnectEvents();
    }

    internal WordApplication(MsWord.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _applicationEvent = application;
        InitializeApp();
        ConnectEvents();
    }


    private void InitializeApp()
    {
        _application.DisplayAlerts = MsWord.WdAlertLevel.wdAlertsMessageBox;
        _disposedValue = false;
        _activeDocument = null;
    }


    private void MemorizeActiveDocument(IWordDocument document)
    {
        _activeDocument = document;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放相关对象
            _selection?.Dispose();
            _documents?.Dispose();
            _windows?.Dispose();
            _wordTemplate?.Dispose();
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

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
