//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定要显示的对话框。
/// </summary>
public enum XlBuiltInDialog
{
    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOpen = 1,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOpenLinks = 2,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSaveAs = 5,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFileDelete = 6,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPageSetup = 7,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPrint = 8,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPrinterSetup = 9,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogArrangeAll = 12,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWindowSize = 13,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWindowMove = 14,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogRun = 17,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSetPrintTitles = 23,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFont = 26,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDisplay = 27,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogProtectDocument = 28,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCalculation = 32,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogExtract = 35,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDataDelete = 36,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSort = 39,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDataSeries = 40,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogTable = 41,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatNumber = 42,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAlignment = 43,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogStyle = 44,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogBorder = 45,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCellProtection = 46,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogColumnWidth = 47,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogClear = 52,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPasteSpecial = 53,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEditDelete = 54,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsert = 55,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPasteNames = 58,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDefineName = 61,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCreateNames = 62,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormulaGoto = 63,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormulaFind = 64,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryArea = 67,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryBar = 68,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryColumn = 69,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryLine = 70,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryPie = 71,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryScatter = 72,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCombination = 73,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGridlines = 76,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAxes = 78,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAttachText = 80,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPatterns = 84,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMainChart = 85,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOverlay = 86,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScale = 87,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatLegend = 88,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatText = 89,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogParse = 91,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogUnhide = 94,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkspace = 95,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogActivate = 103,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCopyPicture = 108,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDeleteName = 110,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDeleteFormat = 111,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogNew = 119,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogRowHeight = 127,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatMove = 128,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatSize = 129,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormulaReplace = 130,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSelectSpecial = 132,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogApplyNames = 133,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogReplaceFont = 134,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSplit = 137,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOutline = 142,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSaveWorkbook = 145,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCopyChart = 147,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatFont = 150,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogNote = 154,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSetUpdateStatus = 159,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogColorPalette = 161,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChangeLink = 166,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAppMove = 170,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAppSize = 171,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMainChartType = 185,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOverlayChartType = 186,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOpenMail = 188,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSendMail = 189,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogStandardFont = 190,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogConsolidate = 191,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSortSpecial = 192,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dArea = 193,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dColumn = 194,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dLine = 195,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dPie = 196,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogView3d = 197,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGoalSeek = 198,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkgroup = 199,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFillGroup = 200,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogUpdateLink = 201,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPromote = 202,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDemote = 203,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogShowDetail = 204,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogObjectProperties = 207,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSaveNewObject = 208,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogApplyStyle = 212,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAssignToObject = 213,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogObjectProtection = 214,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCreatePublisher = 217,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSubscribeTo = 218,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogShowToolbar = 220,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPrintPreview = 222,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEditColor = 223,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatMain = 225,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatOverlay = 226,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEditSeries = 228,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDefineStyle = 229,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryRadar = 249,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEditionOptions = 251,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogZoom = 256,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsertObject = 259,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSize = 261,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMove = 262,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatAuto = 269,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dBar = 272,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGallery3dSurface = 273,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCustomizeToolbar = 276,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookAdd = 281,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookMove = 282,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookCopy = 283,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookOptions = 284,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSaveWorkspace = 285,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartWizard = 288,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAssignToTool = 293,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPlacement = 300,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFillWorkgroup = 301,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookNew = 302,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScenarioCells = 305,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScenarioAdd = 307,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScenarioEdit = 308,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScenarioSummary = 311,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotTableWizard = 312,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotFieldProperties = 313,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsCalculation = 318,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsEdit = 319,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsView = 320,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAddinManager = 321,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMenuEditor = 322,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAttachToolbars = 323,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsChart = 325,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogVbaInsertFile = 328,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogVbaProcedureDefinition = 330,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogRoutingSlip = 336,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMailLogon = 339,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsertPicture = 342,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryDoughnut = 344,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartTrend = 350,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookInsert = 354,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsTransition = 355,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsGeneral = 356,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFilterAdvanced = 370,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMailNextLetter = 378,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDataLabel = 379,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsertTitle = 380,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFontProperties = 381,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMacroOptions = 382,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookUnhide = 384,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookName = 386,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogGalleryCustom = 388,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAddChartAutoformat = 390,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartAddData = 392,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogTabOrder = 394,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSubtotalCreate = 398,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookTabSplit = 415,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWorkbookProtect = 417,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScrollbarProperties = 420,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotShowPages = 421,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogTextToColumns = 422,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatCharttype = 423,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotFieldGroup = 433,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotFieldUngroup = 434,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCheckboxProperties = 435,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogLabelProperties = 436,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogListboxProperties = 437,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEditboxProperties = 438,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOpenText = 441,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPushbuttonProperties = 445,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFilter = 447,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFunctionWizard = 450,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSaveCopyAs = 456,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsListsAdd = 458,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesAxes = 460,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesX = 461,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesY = 462,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogErrorbarX = 463,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogErrorbarY = 464,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFormatChart = 465,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesOrder = 466,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMailEditMailer = 470,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogStandardWidth = 472,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogScenarioMerge = 473,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogProperties = 474,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSummaryInfo = 474,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFindFile = 475,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogActiveCellFont = 476,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogVbaMakeAddin = 478,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogFileSharing = 481,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogAutoCorrect = 485,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCustomViews = 493,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsertNameLabel = 496,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesShape = 504,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartOptionsDataLabels = 505,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartOptionsDataTable = 506,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSetBackgroundPicture = 509,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDataValidation = 525,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartType = 526,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartLocation = 527,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    _xlDialogPhonetic = 538,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartSourceData = 540,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    _xlDialogChartSourceData = 541,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSeriesOptions = 557,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotTableOptions = 567,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotSolveOrder = 568,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotCalculatedField = 570,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotCalculatedItem = 572,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogConditionalFormatting = 583,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogInsertHyperlink = 596,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogProtectSharing = 620,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogOptionsME = 647,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPublishAsWebPage = 653,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPhonetic = 656,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogNewWebQuery = 667,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogImportTextFile = 666,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogExternalDataProperties = 530,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsGeneral = 683,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsFiles = 684,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsPictures = 685,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsEncoding = 686,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsFonts = 687,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPivotClientServerSet = 689,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPropertyFields = 754,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogSearch = 731,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogEvaluateFormula = 709,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogDataLabelMultiple = 723,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogChartOptionsDataLabelMultiple = 724,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogErrorChecking = 732,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogWebOptionsBrowsers = 773,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogCreateList = 796,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogPermission = 832,

    /// <summary>
    /// 显示常量名称对应的对话框。
    /// </summary>
    xlDialogMyPermission = 834,

    /// <summary>
    /// 文档检查器对话框。
    /// </summary>
    xlDialogDocumentInspector = 862,

    /// <summary>
    /// 名称管理器对话框。
    /// </summary>
    xlDialogNameManager = 977,

    /// <summary>
    /// 新建名称对话框。
    /// </summary>
    xlDialogNewName = 978,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSparklineInsertLine = 1133,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSparklineInsertColumn = 1134,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSparklineInsertWinLoss = 1135,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSlicerSettings = 1179,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSlicerCreation = 1182,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSlicerPivotTableConnections = 1184,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogPivotTableSlicerConnections = 1183,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogPivotTableWhatIfAnalysisSettings = 1153,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSetManager = 1109,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSetMDXEditor = 1208,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSetTupleEditorOnRows = 1107,

    /// <summary>
    /// 内部使用。
    /// </summary>
    xlDialogSetTupleEditorOnColumns = 1108,

    /// <summary>
    /// 关系管理对话框。
    /// </summary>
    xlDialogManageRelationships = 1271,

    /// <summary>
    /// 创建关系对话框。
    /// </summary>
    xlDialogCreateRelationship = 1272,

    /// <summary>
    /// 推荐的数据透视表对话框。
    /// </summary>
    xlDialogRecommendedPivotTables = 1258
}