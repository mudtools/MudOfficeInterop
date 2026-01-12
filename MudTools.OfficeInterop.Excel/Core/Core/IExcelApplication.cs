//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Vbe;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel应用程序接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel", ComClassName = "_Application", NoneConstructor = true, NoneDisposed = true)]
public interface IExcelApplication : IOfficeObject<IExcelApplication, MsExcel._Application>, IDisposable, IOfficeApplication
{
    /// <summary>
    /// 获取一个表示指定对象创建者的 Application 对象（可用于 OLE 自动化对象以返回该对象的应用程序）。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]

    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]

    IExcelApplication? Parent { get; }

    /// <summary>
    /// 获取一个表示活动窗口（最顶层的窗口）或指定窗口中的活动单元格的 Range 对象。如果窗口未显示工作表，则此属性失败。
    /// </summary>
    IExcelRange? ActiveCell { get; }

    /// <summary>
    /// 获取一个表示活动图表（嵌入式图表或图表工作表）的 Chart 对象。当嵌入式图表被选中或激活时，它被视为活动图表。当没有活动图表时，此属性返回 null。
    /// </summary>
    IExcelChart? ActiveChart { get; }

    /// <summary>
    /// 获取或设置活动打印机的名称。
    /// </summary>
    string ActivePrinter { get; set; }

    /// <summary>
    /// 获取一个表示活动工作簿或指定窗口或工作簿中的活动工作表（最顶层的工作表）的对象。如果没有活动工作表，则返回 null。
    /// </summary>
    object ActiveSheet { get; }

    /// <summary>
    /// 获取一个表示活动窗口（最顶层的窗口）的 Window 对象。如果没有打开的窗口，则返回 null。
    /// </summary>
    IExcelWindow? ActiveWindow { get; }

    /// <summary>
    /// 获取一个表示活动窗口（最顶层的窗口）中的工作簿的 Workbook 对象。如果没有打开的窗口，或者信息窗口或剪贴板窗口是活动窗口，则返回 null。
    /// </summary>
    IExcelWorkbook? ActiveWorkbook { get; }

    /// <summary>
    /// 获取一个表示“加载项”对话框（“工具”菜单）中列出的所有加载项的 AddIns 集合。
    /// </summary>
    IExcelAddIns? AddIns { get; }

    /// <summary>
    /// 计算所有打开的工作簿。
    /// </summary>
    void Calculate();

    /// <summary>
    /// 获取一个表示活动工作表中所有单元格的 Range 对象。如果活动文档不是工作表，则此属性失败。
    /// </summary>
    IExcelRange? Cells { get; }

    /// <summary>
    /// 获取一个表示活动工作簿中所有图表工作表的 Sheets 集合。
    /// </summary>
    IExcelSheets? Charts { get; }

    /// <summary>
    /// 获取一个表示活动工作表中所有列的 Range 对象。如果活动文档不是工作表，则此属性失败。
    /// </summary>
    IExcelRange? Columns { get; }

    /// <summary>
    /// 获取 Microsoft Excel 接收的最后一个 DDE 确认消息中包含的应用程序特定的 DDE 返回代码。
    /// </summary>
    int DDEAppReturnCode { get; }

    /// <summary>
    /// 通过指定的 DDE 通道在另一个应用程序中运行命令或执行其他操作。
    /// </summary>
    /// <param name="channel">通道号，由 DDEInitiate 方法返回。</param>
    /// <param name="executeString">接收应用程序中定义的消息字符串。</param>
    void DDEExecute(int channel, string executeString);

    /// <summary>
    /// 打开到应用程序的 DDE 通道。
    /// </summary>
    /// <param name="app">应用程序名称。</param>
    /// <param name="topic">描述要打开通道的应用程序中的内容——通常是该应用程序的文档。</param>
    /// <returns>通道号。</returns>
    int? DDEInitiate(string app, string topic);

    /// <summary>
    /// 将数据发送到应用程序。
    /// </summary>
    /// <param name="channel">通道号，由 DDEInitiate 方法返回。</param>
    /// <param name="item">要向其发送数据的项。</param>
    /// <param name="data">要发送到应用程序的数据。</param>
    void DDEPoke(int channel, object item, object data);

    /// <summary>
    /// 从指定的应用程序请求信息。此方法始终返回一个数组。
    /// </summary>
    /// <param name="channel">通道号，由 DDEInitiate 方法返回。</param>
    /// <param name="item">要请求的项。</param>
    /// <returns>请求的信息。</returns>
    object? DDERequest(int channel, string item);

    /// <summary>
    /// 关闭到另一个应用程序的通道。
    /// </summary>
    /// <param name="channel">通道号，由 DDEInitiate 方法返回。</param>
    void DDETerminate(int channel);

    /// <summary>
    /// 将 Microsoft Excel 名称转换为对象或值。
    /// </summary>
    /// <param name="name">对象的名称，使用 Microsoft Excel 的命名约定。</param>
    /// <returns>转换后的对象或值。</returns>
    object? Evaluate(object name);

    /// <summary>
    /// 运行 Microsoft Excel 4.0 宏函数，然后返回该函数的结果。返回类型取决于函数。
    /// </summary>
    /// <param name="executeString">不带等号的 Microsoft Excel 4.0 宏语言函数。所有引用必须以 R1C1 字符串形式给出。如果字符串包含嵌入的双引号，则必须将其加倍。</param>
    /// <returns>宏函数的执行结果。</returns>
    object? ExecuteExcel4Macro(string executeString);

    /// <summary>
    /// 返回一个表示两个或多个区域矩形交集的 Range 对象。
    /// </summary>
    /// <param name="range1">第一个区域。</param>
    /// <param name="range2">第二个区域。</param>
    /// <param name="range3">第三个区域（可选）。</param>
    /// <param name="range4">第四个区域（可选）。</param>
    /// <param name="range5">第五个区域（可选）。</param>
    /// <param name="range6">第六个区域（可选）。</param>
    /// <param name="range7">第七个区域（可选）。</param>
    /// <param name="range8">第八个区域（可选）。</param>
    /// <param name="range9">第九个区域（可选）。</param>
    /// <param name="range10">第十个区域（可选）。</param>
    /// <param name="range11">第十一个区域（可选）。</param>
    /// <param name="range12">第十二个区域（可选）。</param>
    /// <param name="range13">第十三个区域（可选）。</param>
    /// <param name="range14">第十四个区域（可选）。</param>
    /// <param name="range15">第十五个区域（可选）。</param>
    /// <param name="range16">第十六个区域（可选）。</param>
    /// <param name="range17">第十七个区域（可选）。</param>
    /// <param name="range18">第十八个区域（可选）。</param>
    /// <param name="range19">第十九个区域（可选）。</param>
    /// <param name="range20">第二十个区域（可选）。</param>
    /// <param name="range21">第二十一个区域（可选）。</param>
    /// <param name="range22">第二十二个区域（可选）。</param>
    /// <param name="range23">第二十三个区域（可选）。</param>
    /// <param name="range24">第二十四个区域（可选）。</param>
    /// <param name="range25">第二十五个区域（可选）。</param>
    /// <param name="range26">第二十六个区域（可选）。</param>
    /// <param name="range27">第二十七个区域（可选）。</param>
    /// <param name="range28">第二十八个区域（可选）。</param>
    /// <param name="range29">第二十九个区域（可选）。</param>
    /// <param name="range30">第三十个区域（可选）。</param>
    /// <returns>表示区域交集的 Range 对象。</returns>
    IExcelRange? Intersect(IExcelRange range1, IExcelRange range2, IExcelRange? range3 = null,
                            IExcelRange? range4 = null, IExcelRange? range5 = null, IExcelRange? range6 = null,
                            IExcelRange? range7 = null, IExcelRange? range8 = null, IExcelRange? range9 = null,
                            IExcelRange? range10 = null, IExcelRange? range11 = null, IExcelRange? range12 = null,
                            IExcelRange? range13 = null, IExcelRange? range14 = null, IExcelRange? range15 = null,
                            IExcelRange? range16 = null, IExcelRange? range17 = null, IExcelRange? range18 = null,
                            IExcelRange? range19 = null, IExcelRange? range20 = null, IExcelRange? range21 = null,
                            IExcelRange? range22 = null, IExcelRange? range23 = null, IExcelRange? range24 = null,
                            IExcelRange? range25 = null, IExcelRange? range26 = null, IExcelRange? range27 = null,
                            IExcelRange? range28 = null, IExcelRange? range29 = null, IExcelRange? range30 = null);

    /// <summary>
    /// 获取一个表示活动工作簿中所有名称的 Names 集合。
    /// </summary>
    IExcelNames? Names { get; }

    /// <summary>
    /// 获取一个表示活动工作表中所有行的 Range 对象。如果活动文档不是工作表，则此属性失败。
    /// </summary>
    IExcelRange? Rows { get; }

    /// <summary>
    /// 运行宏或调用函数。
    /// </summary>
    /// <param name="macro">要运行的宏。可以是包含宏名称的字符串、指示函数位置的 Range 对象，或注册的 DLL（XLL）函数的注册 ID。</param>
    /// <param name="arg1">传递给函数的参数（可选）。</param>
    /// <param name="arg2">传递给函数的参数（可选）。</param>
    /// <param name="arg3">传递给函数的参数（可选）。</param>
    /// <param name="arg4">传递给函数的参数（可选）。</param>
    /// <param name="arg5">传递给函数的参数（可选）。</param>
    /// <param name="arg6">传递给函数的参数（可选）。</param>
    /// <param name="arg7">传递给函数的参数（可选）。</param>
    /// <param name="arg8">传递给函数的参数（可选）。</param>
    /// <param name="arg9">传递给函数的参数（可选）。</param>
    /// <param name="arg10">传递给函数的参数（可选）。</param>
    /// <param name="arg11">传递给函数的参数（可选）。</param>
    /// <param name="arg12">传递给函数的参数（可选）。</param>
    /// <param name="arg13">传递给函数的参数（可选）。</param>
    /// <param name="arg14">传递给函数的参数（可选）。</param>
    /// <param name="arg15">传递给函数的参数（可选）。</param>
    /// <param name="arg16">传递给函数的参数（可选）。</param>
    /// <param name="arg17">传递给函数的参数（可选）。</param>
    /// <param name="arg18">传递给函数的参数（可选）。</param>
    /// <param name="arg19">传递给函数的参数（可选）。</param>
    /// <param name="arg20">传递给函数的参数（可选）。</param>
    /// <param name="arg21">传递给函数的参数（可选）。</param>
    /// <param name="arg22">传递给函数的参数（可选）。</param>
    /// <param name="arg23">传递给函数的参数（可选）。</param>
    /// <param name="arg24">传递给函数的参数（可选）。</param>
    /// <param name="arg25">传递给函数的参数（可选）。</param>
    /// <param name="arg26">传递给函数的参数（可选）。</param>
    /// <param name="arg27">传递给函数的参数（可选）。</param>
    /// <param name="arg28">传递给函数的参数（可选）。</param>
    /// <param name="arg29">传递给函数的参数（可选）。</param>
    /// <param name="arg30">传递给函数的参数（可选）。</param>
    /// <returns>宏或函数的执行结果。</returns>
    object Run(object? macro = null, object? arg1 = null, object? arg2 = null, object? arg3 = null,
                object? arg4 = null, object? arg5 = null, object? arg6 = null, object? arg7 = null,
                object? arg8 = null, object? arg9 = null, object? arg10 = null, object? arg11 = null,
                object? arg12 = null, object? arg13 = null, object? arg14 = null, object? arg15 = null,
                object? arg16 = null, object? arg17 = null, object? arg18 = null, object? arg19 = null,
                object? arg20 = null, object? arg21 = null, object? arg22 = null, object? arg23 = null,
                object? arg24 = null, object? arg25 = null, object? arg26 = null, object? arg27 = null,
                object? arg28 = null, object? arg29 = null, object? arg30 = null);

    /// <summary>
    /// 获取活动窗口中的选定对象。
    /// </summary>
    object Selection { get; }

    /// <summary>
    /// 向活动应用程序发送击键。
    /// </summary>
    /// <param name="keys">要作为文本发送到应用程序的键或键组合。</param>
    /// <param name="wait">如果为 true，则 Microsoft Excel 在将控制权返回给宏之前等待击键被处理。如果为 false（或省略），则在不等待击键被处理的情况下继续运行宏。</param>
    void SendKeys(string keys, bool? wait = null);

    /// <summary>
    /// 获取一个表示活动工作簿中所有工作表的 Sheets 集合。
    /// </summary>
    IExcelSheets? Sheets { get; }

    /// <summary>
    /// 获取一个表示当前宏代码正在运行的工作簿的 Workbook 对象。
    /// </summary>
    IExcelWorkbook? ThisWorkbook { get; }

    /// <summary>
    /// 返回两个或多个区域的并集。
    /// </summary>
    /// <param name="range1">第一个区域。</param>
    /// <param name="range2">第二个区域。</param>
    /// <param name="range3">第三个区域（可选）。</param>
    /// <param name="range4">第四个区域（可选）。</param>
    /// <param name="range5">第五个区域（可选）。</param>
    /// <param name="range6">第六个区域（可选）。</param>
    /// <param name="range7">第七个区域（可选）。</param>
    /// <param name="range8">第八个区域（可选）。</param>
    /// <param name="range9">第九个区域（可选）。</param>
    /// <param name="range10">第十个区域（可选）。</param>
    /// <param name="range11">第十一个区域（可选）。</param>
    /// <param name="range12">第十二个区域（可选）。</param>
    /// <param name="range13">第十三个区域（可选）。</param>
    /// <param name="range14">第十四个区域（可选）。</param>
    /// <param name="range15">第十五个区域（可选）。</param>
    /// <param name="range16">第十六个区域（可选）。</param>
    /// <param name="range17">第十七个区域（可选）。</param>
    /// <param name="range18">第十八个区域（可选）。</param>
    /// <param name="range19">第十九个区域（可选）。</param>
    /// <param name="range20">第二十个区域（可选）。</param>
    /// <param name="range21">第二十一个区域（可选）。</param>
    /// <param name="range22">第二十二个区域（可选）。</param>
    /// <param name="range23">第二十三个区域（可选）。</param>
    /// <param name="range24">第二十四个区域（可选）。</param>
    /// <param name="range25">第二十五个区域（可选）。</param>
    /// <param name="range26">第二十六个区域（可选）。</param>
    /// <param name="range27">第二十七个区域（可选）。</param>
    /// <param name="range28">第二十八个区域（可选）。</param>
    /// <param name="range29">第二十九个区域（可选）。</param>
    /// <param name="range30">第三十个区域（可选）。</param>
    /// <returns>表示区域并集的 Range 对象。</returns>
    IExcelRange? Union(IExcelRange range1, IExcelRange range2, IExcelRange? range3 = null,
                        IExcelRange? range4 = null, IExcelRange? range5 = null, IExcelRange? range6 = null,
                        IExcelRange? range7 = null, IExcelRange? range8 = null, IExcelRange? range9 = null,
                        IExcelRange? range10 = null, IExcelRange? range11 = null, IExcelRange? range12 = null,
                        IExcelRange? range13 = null, IExcelRange? range14 = null, IExcelRange? range15 = null,
                        IExcelRange? range16 = null, IExcelRange? range17 = null, IExcelRange? range18 = null,
                        IExcelRange? range19 = null, IExcelRange? range20 = null, IExcelRange? range21 = null,
                        IExcelRange? range22 = null, IExcelRange? range23 = null, IExcelRange? range24 = null,
                        IExcelRange? range25 = null, IExcelRange? range26 = null, IExcelRange? range27 = null,
                        IExcelRange? range28 = null, IExcelRange? range29 = null, IExcelRange? range30 = null);

    /// <summary>
    /// 获取一个表示所有工作簿中所有窗口的 Windows 集合。
    /// </summary>
    IExcelWindows? Windows { get; }

    /// <summary>
    /// 获取一个表示所有打开的工作簿的 Workbooks 集合。
    /// </summary>
    IExcelWorkbooks? Workbooks { get; }

    /// <summary>
    /// 获取 WorksheetFunction 对象。
    /// </summary>
    IExcelWorksheetFunction? WorksheetFunction { get; }

    /// <summary>
    /// 获取一个表示活动工作簿中所有工作表的 Sheets 集合。
    /// </summary>
    IExcelSheets? Worksheets { get; }

    /// <summary>
    /// 获取一个表示指定工作簿中所有 Microsoft Excel 4.0 国际宏工作表的 Sheets 集合。
    /// </summary>
    IExcelSheets? Excel4IntlMacroSheets { get; }

    /// <summary>
    /// 获取一个表示指定工作簿中所有 Microsoft Excel 4.0 宏工作表的 Sheets 集合。
    /// </summary>
    IExcelSheets? Excel4MacroSheets { get; }

    /// <summary>
    /// 激活一个 Microsoft 应用程序。如果应用程序已在运行，则此方法激活正在运行的应用程序。如果应用程序未运行，则此方法启动应用程序的新实例。
    /// </summary>
    /// <param name="index">指定要激活的 Microsoft 应用程序的 XlMSApplication 常量。</param>
    void ActivateMicrosoftApp(XlMSApplication index);

    /// <summary>
    /// 将自定义图表自动格式添加到可用的图表自动格式列表中。
    /// </summary>
    /// <param name="chart">包含在应用新图表自动格式时将应用的格式的图表。</param>
    /// <param name="name">自动格式的名称。</param>
    /// <param name="description">自定义自动格式的描述（可选）。</param>
    void AddChartAutoFormat(IExcelChart chart, string name, string? description = null);

    /// <summary>
    /// 为自定义自动填充和/或自定义排序添加自定义列表。
    /// </summary>
    /// <param name="listArray">指定源数据，可以是字符串数组或 Range 对象。</param>
    /// <param name="byRow">仅当 listArray 是 Range 对象时使用。如果为 true，则从区域中的每一行创建自定义列表。如果为 false，则从区域中的每一列创建自定义列表。如果省略此参数，并且区域中的行数多于列数（或行数和列数相等），则 Microsoft Excel 从区域中的每一列创建自定义列表。如果省略此参数，并且区域中的列数多于行数，则 Microsoft Excel 从区域中的每一行创建自定义列表。</param>
    void AddCustomList(string[] listArray, bool? byRow = null);

    /// <summary>
    /// 为自定义自动填充和/或自定义排序添加自定义列表。
    /// </summary>
    /// <param name="range">指定源数据，可以是字符串数组或 Range 对象。</param>
    /// <param name="byRow">仅当 listArray 是 Range 对象时使用。如果为 true，则从区域中的每一行创建自定义列表。如果为 false，则从区域中的每一列创建自定义列表。如果省略此参数，并且区域中的行数多于列数（或行数和列数相等），则 Microsoft Excel 从区域中的每一列创建自定义列表。如果省略此参数，并且区域中的列数多于行数，则 Microsoft Excel 从区域中的每一行创建自定义列表。</param>
    void AddCustomList(IExcelRange range, bool? byRow = null);

    /// <summary>
    /// 获取或设置一个值，指示在拖放编辑操作期间覆盖非空单元格之前，Microsoft Excel 是否显示消息。
    /// </summary>
    bool AlertBeforeOverwriting { get; set; }

    /// <summary>
    /// 获取或设置备用启动文件夹的名称。
    /// </summary>
    string AltStartupPath { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示打开包含链接的文件时，Microsoft Excel 是否询问用户是否更新链接。如果为 false，则链接会自动更新，不显示对话框。
    /// </summary>
    bool AskToUpdateLinks { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用插入和删除动画。
    /// </summary>
    bool EnableAnimations { get; set; }

    /// <summary>
    /// 获取一个表示 Microsoft Excel 自动更正属性的 AutoCorrect 对象。
    /// </summary>
    IExcelAutoCorrect? AutoCorrect { get; }


    /// <summary>
    /// 获取或设置一个值，指示在将工作簿保存到磁盘之前是否计算工作簿（如果 Calculation 属性设置为 xlManual）。即使更改 Calculation 属性，此属性也会保留。
    /// </summary>
    bool CalculateBeforeSave { get; set; }

    /// <summary>
    /// 获取或设置计算模式。
    /// </summary>
    XlCalculation Calculation { get; set; }

    /// <summary>
    /// 获取有关如何调用 Visual Basic 的信息。
    /// </summary>
    /// <param name="index">数组的索引。仅当属性返回数组时才使用此参数。</param>
    /// <returns>调用者信息。</returns>
    [MethodIndex]
    object? Caller(object? index = null);

    /// <summary>
    /// 获取应用程序窗口中显示的标题。如果不设置名称或将名称设置为空，则此属性返回“Microsoft Excel”。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用单元格的拖放操作。
    /// </summary>
    bool CellDragAndDrop { get; set; }

    /// <summary>
    /// 将厘米测量值转换为点（一点等于 0.035 厘米）。
    /// </summary>
    /// <param name="centimeters">要转换为点的厘米值。</param>
    /// <returns>转换后的点数。</returns>
    double? CentimetersToPoints(double centimeters);

    /// <summary>
    /// 检查单个单词的拼写。如果单词在其中一个词典中找到，则返回 true；如果未找到，则返回 false。
    /// </summary>
    /// <param name="word">要检查的单词。</param>
    /// <param name="customDictionary">一个字符串，指示如果在主词典中找不到单词时要检查的自定义词典的文件名。如果省略此参数，则使用当前指定的词典。</param>
    /// <param name="ignoreUppercase">如果为 true，则 Microsoft Excel 忽略全部大写的单词。如果为 false，则 Microsoft Excel 检查全部大写的单词。如果省略此参数，则使用当前设置。</param>
    /// <returns>如果单词拼写正确，则为 true；否则为 false。</returns>
    bool? CheckSpelling(string word, string? customDictionary = null, bool? ignoreUppercase = null);

    /// <summary>
    /// 获取剪贴板上当前格式的数组。要确定剪贴板上是否有特定格式，请将数组中的每个元素与备注部分中列出的适当常量进行比较。
    /// </summary>
    /// <param name="index">要返回的数组元素。如果省略此参数，则属性返回剪贴板上当前所有格式的整个数组。</param>
    /// <returns>剪贴板格式数组。</returns>
    [MethodIndex]
    object? ClipboardFormats(object? index = null);

    /// <summary>
    /// 获取或设置一个值，指示是否可以显示 Microsoft Office 剪贴板。
    /// </summary>
    bool DisplayClipboardWindow { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示手写识别是否仅限于数字和标点符号。
    /// </summary>
    bool ConstrainNumeric { get; set; }

    /// <summary>
    /// 在公式中的单元格引用之间进行转换，在 A1 和 R1C1 引用样式之间、相对引用和绝对引用之间，或同时进行两者。
    /// </summary>
    /// <param name="formula">包含要转换的公式的字符串。这必须是有效的公式，并且必须以等号开头。</param>
    /// <param name="fromReferenceStyle">公式的引用样式。</param>
    /// <param name="toReferenceStyle">要返回的引用样式。如果省略此参数，则引用样式不会更改；公式保持 FromReferenceStyle 指定的样式。</param>
    /// <param name="toAbsolute">指定转换后的引用类型。如果省略此参数，则引用类型不会更改。</param>
    /// <param name="relativeTo">一个 Range 对象，包含一个单元格。相对引用与此单元格相关。</param>
    /// <returns>转换后的公式。</returns>
    object? ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle,
                           object? toReferenceStyle = null, object? toAbsolute = null, IExcelRange? relativeTo = null);

    /// <summary>
    /// 获取或设置一个值，指示对象是否与单元格一起被剪切、复制、提取和排序。
    /// </summary>
    bool CopyObjectsWithCells { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 中鼠标指针的外观。
    /// </summary>
    XlMousePointer Cursor { get; set; }

    /// <summary>
    /// 获取定义的（包括内置列表）自定义列表的数量。
    /// </summary>
    int CustomListCount { get; }

    /// <summary>
    /// 获取或设置剪切或复制模式的状态。
    /// </summary>
    XlCutCopyMode CutCopyMode { get; set; }

    /// <summary>
    /// 获取或设置数据输入模式。在数据输入模式下，您只能在当前选定范围中未锁定的单元格中输入数据。
    /// </summary>
    int DataEntryMode { get; set; }

    /// <summary>
    /// 删除可用图表自动格式列表中的自定义图表自动格式。
    /// </summary>
    /// <param name="name">要删除的自定义自动格式的名称。</param>
    void DeleteChartAutoFormat(string name);

    /// <summary>
    /// 删除自定义列表。
    /// </summary>
    /// <param name="listNum">自定义列表编号。此数字必须大于或等于 5（Microsoft Excel 有四个无法删除的内置自定义列表）。</param>
    void DeleteCustomList(int listNum);

    /// <summary>
    /// 获取一个表示所有内置对话框的 Dialogs 集合。
    /// </summary>
    IExcelDialogs? Dialogs { get; }

    /// <summary>
    /// 获取或设置一个值，指示宏运行时 Microsoft Excel 是否显示某些警报和消息。
    /// </summary>
    bool DisplayAlerts { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示编辑栏。
    /// </summary>
    bool DisplayFormulaBar { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否处于全屏模式。
    /// </summary>
    bool DisplayFullScreen { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示包含注释的单元格是否显示单元格提示并包含注释指示器（其右上角的小点）。
    /// </summary>
    bool DisplayNoteIndicator { get; set; }

    /// <summary>
    /// 获取或设置单元格显示注释和指示器的方式。
    /// </summary>
    XlCommentDisplayMode DisplayCommentIndicator { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示最近使用的文件列表。
    /// </summary>
    bool DisplayRecentFiles { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示所有工作簿的滚动条。
    /// </summary>
    bool DisplayScrollBars { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示状态栏。
    /// </summary>
    bool DisplayStatusBar { get; set; }

    /// <summary>
    /// 等效于双击活动单元格。
    /// </summary>
    void DoubleClick();

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否允许在单元格中直接编辑。
    /// </summary>
    bool EditDirectlyInCell { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用自动完成功能。
    /// </summary>
    bool EnableAutoComplete { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 如何处理用户对正在运行的过程的中断（通过 CTRL+BREAK、ESC 或 COMMAND+PERIOD）。
    /// </summary>
    XlEnableCancelKey EnableCancelKey { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为 Microsoft Office 启用声音。
    /// </summary>
    bool EnableSound { get; set; }

    /// <summary>
    /// 获取有关已安装文件转换器的信息。如果未安装转换器，则返回 null。
    /// </summary>
    /// <param name="index1">转换器的长名称，包括 Windows 中的文件类型搜索字符串，例如“Lotus 1-2-3 Files (*.wk*)”。</param>
    /// <param name="index2">转换器 DLL 或代码资源的路径。</param>
    /// <returns>文件转换器信息。</returns>
    [MethodIndex]
    object? FileConverters(string? index1 = null, string? index2 = null);

    /// <summary>
    /// 获取用于文件搜索的 FileSearch 对象。此属性仅在 Microsoft Windows 中可用。
    /// </summary>
    IOfficeFileSearch? FileSearch { get; }

    /// <summary>
    /// 获取或设置一个值，指示在将 Application.FixedDecimal 属性设置为 true 后输入的所有数据是否都将使用由 FixedDecimalPlaces 属性设置的固定小数位数进行格式化。
    /// </summary>
    bool FixedDecimal { get; set; }

    /// <summary>
    /// 获取或设置当 FixedDecimal 属性设置为 true 时使用的固定小数位数。
    /// </summary>
    int FixedDecimalPlaces { get; set; }

    /// <summary>
    /// 返回自定义列表（字符串数组）。
    /// </summary>
    /// <param name="listNum">列表编号。</param>
    /// <returns>自定义列表内容。</returns>
    object? GetCustomListContents(int listNum);

    /// <summary>
    /// 返回字符串数组的自定义列表编号。您可以使用此方法来匹配内置列表和自定义定义的列表。
    /// </summary>
    /// <param name="listArray">字符串数组。</param>
    /// <returns>自定义列表编号。</returns>
    int? GetCustomListNum(string[] listArray);

    /// <summary>
    /// 显示标准“打开”对话框并从用户处获取文件名，而无需实际打开任何文件。
    /// </summary>
    /// <param name="fileFilter">指定文件筛选条件的字符串。此字符串由文件筛选字符串对组成，后跟 MS-DOS 通配符文件筛选规范，每个部分和每个对用逗号分隔。每个单独的对列在“文件类型”下拉列表框中。</param>
    /// <param name="filterIndex">默认文件筛选条件的索引号，从 1 到 FileFilter 中指定的筛选器数量。如果省略此参数或大于存在的筛选器数量，则使用第一个文件筛选器。</param>
    /// <param name="title">对话框的标题。如果省略此参数，则标题为“打开”。</param>
    /// <param name="buttonText">仅限 Macintosh。</param>
    /// <param name="multiSelect">如果为 true，则允许选择多个文件名。如果为 false，则仅允许选择一个文件名。默认值为 false。</param>
    /// <returns>选择的文件名。</returns>
    [ValueConvert]
    string? GetOpenFilename(string? fileFilter = null, int? filterIndex = null, string? title = null,
                            string? buttonText = null, bool? multiSelect = null);

    /// <summary>
    /// 显示标准“另存为”对话框并从用户处获取文件名，而无需实际保存任何文件。
    /// </summary>
    /// <param name="initialFilename">建议的文件名。如果省略此参数，则 Microsoft Excel 使用活动工作簿的名称。</param>
    /// <param name="fileFilter">指定文件筛选条件的字符串。此字符串由文件筛选字符串对组成，后跟 MS-DOS 通配符文件筛选规范，每个部分和每个对用逗号分隔。每个单独的对列在“文件类型”下拉列表框中。</param>
    /// <param name="filterIndex">默认文件筛选条件的索引号，从 1 到 FileFilter 中指定的筛选器数量。如果省略此参数或大于存在的筛选器数量，则使用第一个文件筛选器。</param>
    /// <param name="title">对话框的标题。如果省略此参数，则使用默认标题。</param>
    /// <param name="buttonText">仅限 Macintosh。</param>
    /// <returns>保存的文件名。</returns>
    [ValueConvert]
    string? GetSaveAsFilename(string? initialFilename = null, string? fileFilter = null,
                             int? filterIndex = null, string? title = null, string? buttonText = null);

    /// <summary>
    /// 选择任何工作簿中的任何区域或 Visual Basic 过程，并在工作簿尚未处于活动状态时将其激活。
    /// </summary>
    /// <param name="reference">目标。可以是 Range 对象、包含 R1C1 样式表示法的单元格引用的字符串或包含 Visual Basic 过程名称的字符串。如果省略此参数，则目标是您上次使用 Goto 方法选择的最后一个区域。</param>
    /// <param name="scroll">如果为 true，则滚动窗口以使区域的左上角出现在窗口的左上角。如果为 false，则不滚动窗口。默认为 false。</param>
    void Goto(string? reference = null, bool? scroll = null);

    /// <summary>
    /// 选择任何工作簿中的任何区域或 Visual Basic 过程，并在工作簿尚未处于活动状态时将其激活。
    /// </summary>
    /// <param name="reference">目标。可以是 Range 对象、包含 R1C1 样式表示法的单元格引用的字符串或包含 Visual Basic 过程名称的字符串。如果省略此参数，则目标是您上次使用 Goto 方法选择的最后一个区域。</param>
    /// <param name="scroll">如果为 true，则滚动窗口以使区域的左上角出现在窗口的左上角。如果为 false，则不滚动窗口。默认为 false。</param>
    void Goto(IExcelRange? reference = null, bool? scroll = null);

    /// <summary>
    /// 显示帮助主题。
    /// </summary>
    /// <param name="helpFile">要显示的联机帮助文件的名称。如果未指定此参数，则使用 Microsoft Excel 帮助。</param>
    /// <param name="helpContextID">帮助主题的上下文 ID 号。如果未指定此参数，则显示“帮助主题”对话框。</param>
    void Help(string? helpFile = null, string? helpContextID = null);

    /// <summary>
    /// 获取或设置一个值，指示是否忽略远程 DDE 请求。
    /// </summary>
    bool IgnoreRemoteRequests { get; set; }

    /// <summary>
    /// 将英寸测量值转换为点。
    /// </summary>
    /// <param name="inches">要转换为点的英寸值。</param>
    /// <returns>转换后的点数。</returns>
    double? InchesToPoints(double inches);

    /// <summary>
    /// 显示用户输入对话框。返回在对话框中输入的信息。
    /// </summary>
    /// <param name="prompt">要在对话框中显示的消息。可以是字符串、数字、日期或布尔值。</param>
    /// <param name="title">输入框的标题。如果省略此参数，则默认标题为“输入”。</param>
    /// <param name="defaultValue">指定对话框首次显示时将出现在文本框中的值。如果省略此参数，则文本框为空。此值可以是 Range 对象。</param>
    /// <param name="left">对话框相对于屏幕左上角的 x 位置（以磅为单位）。</param>
    /// <param name="top">对话框相对于屏幕左上角的 y 位置（以磅为单位）。</param>
    /// <param name="helpFile">此输入框的帮助文件名。如果 HelpFile 和 HelpContextID 参数同时存在，则对话框中会出现一个“帮助”按钮。</param>
    /// <param name="helpContextID">HelpFile 中帮助主题的上下文 ID 号。</param>
    /// <param name="type">指定返回数据类型。如果省略此参数，则对话框返回文本。</param>
    /// <returns>输入框的返回值。</returns>
    object? InputBox(string prompt, string? title = null, object? defaultValue = null,
                    int? left = null, int? top = null, string? helpFile = null,
                    string? helpContextID = null, int? type = null);

    /// <summary>
    /// 显示用户输入对话框。返回在对话框中输入的信息。
    /// </summary>
    /// <param name="prompt">要在对话框中显示的消息。可以是字符串、数字、日期或布尔值。</param>
    /// <param name="title">输入框的标题。如果省略此参数，则默认标题为“输入”。</param>
    /// <param name="defaultValue">指定对话框首次显示时将出现在文本框中的值。如果省略此参数，则文本框为空。此值可以是 Range 对象。</param>
    /// <param name="left">对话框相对于屏幕左上角的 x 位置（以磅为单位）。</param>
    /// <param name="top">对话框相对于屏幕左上角的 y 位置（以磅为单位）。</param>
    /// <param name="helpFile">此输入框的帮助文件名。如果 HelpFile 和 HelpContextID 参数同时存在，则对话框中会出现一个“帮助”按钮。</param>
    /// <param name="helpContextID">HelpFile 中帮助主题的上下文 ID 号。</param>
    /// <param name="type">指定返回数据类型。如果省略此参数，则对话框返回文本。</param>
    /// <returns>输入框的返回值。</returns>
    object? InputBox(string prompt, string? title = null, IExcelRange? defaultValue = null,
                    int? left = null, int? top = null, string? helpFile = null,
                    string? helpContextID = null, int? type = null);

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否处于交互模式；此属性通常为 true。如果将此属性设置为 false，Microsoft Excel 将阻止来自键盘和鼠标的所有输入（除了显示由您的代码的对话框的输入）。阻止用户输入将防止用户在代码移动或激活 Microsoft Excel 对象时干扰代码。
    /// </summary>
    bool Interactive { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否将使用迭代来解决循环引用。
    /// </summary>
    bool Iteration { get; set; }

    /// <summary>
    /// 获取 Library 文件夹的路径，但不包括最终分隔符。
    /// </summary>
    string LibraryPath { get; }

    /// <summary>
    /// 对应于“宏选项”对话框中的选项。您还可以使用此方法在内置或新类别中显示用户定义函数（UDF）。
    /// </summary>
    /// <param name="macro">宏名称或用户定义函数（UDF）的名称。</param>
    /// <param name="description">宏描述。</param>
    /// <param name="hasMenu">此参数被忽略。</param>
    /// <param name="menuText">此参数被忽略。</param>
    /// <param name="hasShortcutKey">如果为 true，则为宏分配快捷键（还必须指定 ShortcutKey）。如果此参数为 false，则不分配快捷键给宏。如果宏已有快捷键，则将此参数设置为 false 将删除该快捷键。默认值为 false。</param>
    /// <param name="shortcutKey">如果 HasShortcutKey 为 true，则为必需；否则忽略。快捷键。</param>
    /// <param name="category">指定现有宏函数类别（例如财务、日期和时间或用户定义）的整数。还可以为自定义类别指定字符串。</param>
    /// <param name="statusBar">宏的状态栏文本。</param>
    /// <param name="helpContextID">分配给宏的帮助主题的上下文 ID 的整数。</param>
    /// <param name="helpFile">包含由 HelpContextId 定义的帮助主题的帮助文件的名称。</param>
    void MacroOptions(string? macro = null, string? description = null, bool? hasMenu = null,
                    string? menuText = null, bool? hasShortcutKey = null, string? shortcutKey = null,
                    object? category = null, string? statusBar = null, string? helpContextID = null, string? helpFile = null);

    /// <summary>
    /// 关闭由 Microsoft Excel 建立的 MAPI 邮件会话。
    /// </summary>
    void MailLogoff();

    /// <summary>
    /// 登录到 MAPI 邮件或 Microsoft Exchange 并建立邮件会话。如果 Microsoft Mail 尚未运行，则必须使用此方法建立邮件会话，然后才能使用邮件或文档路由功能。
    /// </summary>
    /// <param name="name">邮件帐户名或 Microsoft Exchange 配置文件名称。如果省略此参数，则使用默认邮件帐户名。</param>
    /// <param name="password">邮件帐户密码。在 Microsoft Exchange 中忽略此参数。</param>
    /// <param name="downloadNewMail">如果为 true，则立即下载新邮件。</param>
    void MailLogon(string? name = null, string? password = null, bool? downloadNewMail = null);

    /// <summary>
    /// 获取作为十六进制字符串的 MAPI 邮件会话号（如果有活动会话），如果没有会话，则返回 null。
    /// </summary>
    object MailSession { get; }

    /// <summary>
    /// 获取主机上安装的邮件系统。
    /// </summary>
    XlMailSystem MailSystem { get; }

    /// <summary>
    /// 获取一个值，指示数学协处理器是否可用。
    /// </summary>
    bool MathCoprocessorAvailable { get; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 解决循环引用时每次迭代之间的最大变化量。
    /// </summary>
    double MaxChange { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 可用于解决循环引用的最大迭代次数。
    /// </summary>
    int MaxIterations { get; set; }

    /// <summary>
    /// 获取 Microsoft Excel 仍可使用的内存量（以字节为单位）。
    /// </summary>
    int MemoryFree { get; }

    /// <summary>
    /// 获取一个值，指示鼠标是否可用。
    /// </summary>
    bool MouseAvailable { get; }

    /// <summary>
    /// 获取或设置一个值，指示按下 ENTER（RETURN）键后活动单元格是否立即移动。
    /// </summary>
    bool MoveAfterReturn { get; set; }

    /// <summary>
    /// 获取或设置用户按下 ENTER 时活动单元格移动的方向。
    /// </summary>
    XlDirection MoveAfterReturnDirection { get; set; }

    /// <summary>
    /// 获取表示最近使用的文件列表的 RecentFiles 集合。
    /// </summary>
    IExcelRecentFiles? RecentFiles { get; }


    /// <summary>
    /// 获取存储模板的网络路径。如果网络路径不存在，则此属性返回空字符串。
    /// </summary>
    string NetworkTemplatesPath { get; }

    /// <summary>
    /// 获取一个包含由最近查询表或数据透视表操作生成的所有 ODBC 错误的 ODBCErrors 集合。
    /// </summary>
    IExcelODBCErrors? ODBCErrors { get; }

    /// <summary>
    /// 获取或设置 ODBC 查询超时时间（以秒为单位）。默认值为 45 秒。
    /// </summary>
    int ODBCTimeout { get; set; }

    /// <summary>
    /// 获取当前操作系统的名称和版本号，例如“Windows (32-bit) 4.00”或“Macintosh 7.00”。
    /// </summary>
    string OperatingSystem { get; }

    /// <summary>
    /// 获取注册的组织名称。
    /// </summary>
    string OrganizationName { get; }

    /// <summary>
    /// 获取路径分隔符字符（“\”）。
    /// </summary>
    string PathSeparator { get; }

    /// <summary>
    /// 获取最后四个选定区域或名称的数组。数组中的每个元素都是一个 Range 对象。
    /// </summary>
    /// <param name="index">先前区域或名称的索引号（从 1 到 4）。</param>
    /// <returns>先前选定的区域或名称。</returns>
    [MethodIndex]
    object? PreviousSelections(int? index = null);

    /// <summary>
    /// 获取或设置一个值，指示数据透视表是否使用结构化选择。
    /// </summary>
    bool PivotTableSelection { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示首次保存文件时，Microsoft Excel 是否询问摘要信息。
    /// </summary>
    bool PromptForSummaryInfo { get; set; }


    /// <summary>
    /// 如果宏记录器打开，则记录代码。
    /// </summary>
    /// <param name="basicCode">一个字符串，指定如果宏记录器正在记录到 Visual Basic 模块中时将记录的 Visual Basic 代码。该字符串将记录在一行上。如果字符串包含回车符（ASCII 字符 10，或代码中的 Chr$(10)），则它将记录在多行上。</param>
    /// <param name="xlmCode">此参数被忽略。</param>
    void RecordMacro(string? basicCode = null, object? xlmCode = null);

    /// <summary>
    /// 获取一个值，指示宏是否使用相对引用记录；如果记录是绝对的，则为 false。
    /// </summary>
    bool RecordRelative { get; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 显示单元格引用以及行和列标题的方式，可以是 A1 或 R1C1 引用样式。
    /// </summary>
    XlReferenceStyle ReferenceStyle { get; set; }

    /// <summary>
    /// 获取使用 REGISTER 或 REGISTER.ID 宏函数注册的动态链接库（DLL）或代码资源中的函数的信息。
    /// </summary>
    /// <param name="index1">DLL 或代码资源的名称。</param>
    /// <param name="index2">函数名称。</param>
    /// <returns>注册函数信息。</returns>
    [MethodIndex]
    object? RegisteredFunctions(string? index1 = null, string? index2 = null);

    /// <summary>
    /// 加载 XLL 代码资源并自动注册资源中包含的函数和命令。
    /// </summary>
    /// <param name="filename">指定要加载的 XLL 的名称。</param>
    /// <returns>如果加载成功，则为 true；否则为 false。</returns>
    bool? RegisterXLL(string filename);

    /// <summary>
    /// 重复最后一个用户界面操作。
    /// </summary>
    void Repeat();

    /// <summary>
    /// 重置传递名单，以便可以使用相同的名单（相同的收件人列表和传递信息）启动新的传递。传递必须完成后才能使用此方法。在其他时间使用此方法会导致错误。
    /// </summary>
    void ResetTipWizard();

    /// <summary>
    /// 保存对指定工作簿的更改。
    /// </summary>
    /// <param name="filename">要保存的文件名。</param>
    void Save(string? filename = null);

    /// <summary>
    /// 保存当前工作区。
    /// </summary>
    /// <param name="filename">保存的文件名。</param>
    void SaveWorkspace(string? filename = null);

    /// <summary>
    /// 获取或设置一个值，指示是否打开屏幕更新。
    /// </summary>
    bool ScreenUpdating { get; set; }

    /// <summary>
    /// 指定 Microsoft Excel 在创建新图表时将使用的图表模板的名称。
    /// </summary>
    /// <param name="formatName">指定自定义自动格式的名称。此名称可以是命名自定义自动格式的字符串，也可以是用于指定内置图表模板的特殊常量 xlBuiltIn。</param>
    /// <param name="gallery">指定图库的名称。</param>
    void SetDefaultChart(string? formatName = null, string? gallery = null);

    /// <summary>
    /// 获取或设置 Microsoft Excel 自动插入到新工作簿中的工作表数量。
    /// </summary>
    int SheetsInNewWorkbook { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示图表是否显示图表提示名称。默认值为 true。
    /// </summary>
    bool ShowChartTipNames { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示图表是否显示图表提示值。默认值为 true。
    /// </summary>
    bool ShowChartTipValues { get; set; }

    /// <summary>
    /// 获取或设置标准字体的名称。
    /// </summary>
    string StandardFont { get; set; }

    /// <summary>
    /// 获取或设置标准字体大小（以磅为单位）。
    /// </summary>
    double StandardFontSize { get; set; }

    /// <summary>
    /// 获取启动文件夹的完整路径，不包括最终分隔符。
    /// </summary>
    string StartupPath { get; }

    /// <summary>
    /// 获取或设置状态栏中的文本。
    /// </summary>
    object StatusBar { get; set; }

    /// <summary>
    /// 获取存储模板的本地路径。
    /// </summary>
    string TemplatesPath { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否打开工具提示。
    /// </summary>
    bool ShowToolTips { get; set; }

    /// <summary>
    /// 获取或设置保存文件的默认格式。
    /// </summary>
    XlFileFormat DefaultSaveFormat { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 菜单或帮助键，通常为“/”。
    /// </summary>
    string TransitionMenuKey { get; set; }

    /// <summary>
    /// 获取或设置按下 Microsoft Excel 菜单键时采取的操作。
    /// </summary>
    int TransitionMenuKeyAction { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示过渡导航键是否处于活动状态。
    /// </summary>
    bool TransitionNavigKeys { get; set; }

    /// <summary>
    /// 取消最后一个用户界面操作。
    /// </summary>
    void Undo();

    /// <summary>
    /// 获取窗口在应用程序窗口区域中可以占用的最大高度（以磅为单位）。
    /// </summary>
    double UsableHeight { get; }

    /// <summary>
    /// 获取窗口在应用程序窗口区域中可以占用的最大宽度（以磅为单位）。
    /// </summary>
    double UsableWidth { get; }

    /// <summary>
    /// 获取或设置一个值，指示应用程序是否对用户可见或是否由用户创建或启动。如果使用 CreateObject 或 GetObject 函数以编程方式创建或启动应用程序并且应用程序是隐藏的，则为 false。
    /// </summary>
    bool UserControl { get; set; }

    /// <summary>
    /// 获取或设置当前用户的名称。
    /// </summary>
    string UserName { get; set; }

    /// <summary>
    /// 获取“Microsoft Excel”。
    /// </summary>
    string Value { get; }

    /// <summary>
    /// 获取一个表示 Visual Basic 编辑器的 VBE 对象。
    /// </summary>
    IVbeApplication? VBE { get; }

    /// <summary>
    /// 将用户定义函数标记为易失性。每当工作表中任何单元格发生计算时，都必须重新计算易失性函数。仅当输入变量更改时，才会重新计算非易失性函数。如果此方法不在用于计算工作表单元格的用户定义函数内部，则无效。
    /// </summary>
    /// <param name="volatilee">如果为 true，则将函数标记为易失性。如果为 false，则将函数标记为非易失性。默认值为 true。</param>
    void Volatile(bool? volatilee = null);

    /// <summary>
    /// 获取一个值，指示计算机是否正在 Microsoft Windows for Pen Computing 下运行。
    /// </summary>
    bool WindowsForPens { get; }

    /// <summary>
    /// 获取或设置窗口的状态。
    /// </summary>
    XlWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 显示新窗口和工作表的默认方向。
    /// </summary>
    int DefaultSheetDirection { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是使用视觉光标还是逻辑光标。
    /// </summary>
    int CursorMovement { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否为从右到左的语言显示控制字符。
    /// </summary>
    bool ControlCharacters { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为指定的对象启用事件。
    /// </summary>
    bool EnableEvents { get; set; }

    /// <summary>
    /// 暂停正在运行的宏，直到指定的时间。如果指定的时间已到，则返回 true。
    /// </summary>
    /// <param name="time">希望宏恢复的时间，采用 Microsoft Excel 日期格式。</param>
    /// <returns>如果时间已到，则为 true；否则为 false。</returns>
    bool? Wait(object time);

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否自动将格式和公式扩展到添加到列表的新数据。
    /// </summary>
    bool ExtendList { get; set; }

    /// <summary>
    /// 获取 OLEDBErrors 集合，该集合表示最近 OLE DB 查询返回的错误信息。
    /// </summary>
    IExcelOLEDBErrors? OLEDBErrors { get; }

    /// <summary>
    /// 返回指定文本字符串的日语拼音文本。仅当您为 Microsoft Office 选择或安装了日语语言支持时，此方法才可用。
    /// </summary>
    /// <param name="text">要转换为拼音文本的文本。如果省略此参数，则返回先前指定的文本的下一个可能的拼音文本字符串（如果有）。如果没有更多可能的拼音文本字符串，则返回空字符串。</param>
    /// <returns>拼音文本。</returns>
    string? GetPhonetic(string? text = null);

    /// <summary>
    /// 获取 Microsoft Excel 的 COMAddIns 集合，该集合表示当前安装的 COM 加载项。
    /// </summary>
    IOfficeCOMAddIns? COMAddIns { get; }

    /// <summary>
    /// 获取 Microsoft Excel 的全局唯一标识符（GUID）。
    /// </summary>
    string ProductCode { get; }

    /// <summary>
    /// 获取用户计算机上安装 COM 加载项的位置的路径。
    /// </summary>
    string UserLibraryPath { get; }

    /// <summary>
    /// 获取或设置一个值，指示格式化为百分比的单元格中的条目是否在输入后不自动乘以 100。
    /// </summary>
    bool AutoPercentEntry { get; set; }

    /// <summary>
    /// 强制对所有打开工作簿中的数据进行全面计算。
    /// </summary>
    void CalculateFull();

    /// <summary>
    /// 显示“打开”对话框。
    /// </summary>
    /// <returns>如果成功打开文件，则为 true；否则为 false。</returns>
    bool? FindFile();

    /// <summary>
    /// 获取一个数字，其最右边的四位数字是次要计算引擎版本号，而其他（左侧）数字是 Microsoft Excel 的主要版本。
    /// </summary>
    int CalculationVersion { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否为每个打开的工作簿显示单独的 Windows 任务栏按钮。默认值为 true。
    /// </summary>
    bool ShowWindowsInTaskbar { get; set; }

    /// <summary>
    /// 获取或设置一个值（常量），该值指定 Microsoft Excel 如何处理对需要尚未安装的功能的方法和属性的调用。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoFeatureInstall FeatureInstall { get; set; }

    /// <summary>
    /// 获取一个值，指示 Microsoft Excel 应用程序是否准备就绪。
    /// </summary>
    bool Ready { get; }

    /// <summary>
    /// 获取或设置要查找的单元格格式类型的搜索条件。
    /// </summary>
    IExcelCellFormat? FindFormat { get; set; }

    /// <summary>
    /// 设置要在替换单元格格式时使用的替换条件。然后，替换条件将用于 Range 对象的 Replace 方法的后续调用中。
    /// </summary>
    IExcelCellFormat? ReplaceFormat { get; set; }

    /// <summary>
    /// 获取一个指示应用程序对于在 Microsoft Excel 中执行的任何计算的计算状态的 XlCalculationState 常量。
    /// </summary>
    XlCalculationState CalculationState { get; }

    /// <summary>
    /// 获取或设置一个指定在执行计算时可以中断 Microsoft Excel 的键的 XlCalculationInterruptKey 常量。
    /// </summary>
    XlCalculationInterruptKey CalculationInterruptKey { get; set; }

    /// <summary>
    /// 获取一个表示在工作表重新计算时被跟踪的区域的 Watches 对象。
    /// </summary>
    IExcelWatches? Watches { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否可以显示函数工具提示。
    /// </summary>
    bool DisplayFunctionToolTips { get; set; }

    /// <summary>
    /// 获取或设置一个表示 Microsoft Excel 在以编程方式打开文件时使用的安全模式的 MsoAutomationSecurity 常量。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutomationSecurity AutomationSecurity { get; set; }

    /// <summary>
    /// 获取一个表示文件对话框实例的 FileDialog 对象。
    /// </summary>
    /// <param name="fileDialogType">文件对话框的类型。</param>
    /// <returns>文件对话框对象。</returns>
    [MethodIndex]
    IOfficeFileDialog? FileDialog([ComNamespace("MsCore")] MsoFileDialogType fileDialogType);

    /// <summary>
    /// 对所有打开的工作簿，强制进行全面计算并重建依赖项。
    /// </summary>
    void CalculateFullRebuild();

    /// <summary>
    /// 获取或设置一个值，指示是否可以显示“粘贴选项”按钮。
    /// </summary>
    bool DisplayPasteOptions { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否应显示“插入选项”按钮。
    /// </summary>
    bool DisplayInsertOptions { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否可以获取数据透视表报告数据。
    /// </summary>
    bool GenerateGetPivotData { get; set; }

    /// <summary>
    /// 获取一个 AutoRecover 对象，该对象在定时间隔内备份所有文件格式。
    /// </summary>
    IExcelAutoRecover? AutoRecover { get; }

    /// <summary>
    /// 获取 Microsoft Excel 窗口的顶级窗口句柄的整数指示符。
    /// </summary>
    int Hwnd { get; }

    /// <summary>
    /// 获取调用 Microsoft Excel 的实例的实例句柄。
    /// </summary>
    int Hinstance { get; }

    /// <summary>
    /// 停止 Microsoft Excel 应用程序中的重新计算。
    /// </summary>
    /// <param name="keepAbort">允许对区域执行重新计算。</param>
    void CheckAbort(bool? keepAbort = null);

    /// <summary>
    /// 获取一个表示应用程序的错误检查选项的 ErrorCheckingOptions 对象。
    /// </summary>
    IExcelErrorCheckingOptions? ErrorCheckingOptions { get; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否在您键入时自动将超链接格式化为您键入的格式。如果为 false，则 Excel 不会在您键入时自动将超链接格式化为您键入的格式。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceHyperlinks { get; set; }

    /// <summary>
    /// 获取应用程序的 SmartTagRecognizers 集合。
    /// </summary>
    IExcelSmartTagRecognizers? SmartTagRecognizers { get; }

    /// <summary>
    /// 获取一个 NewFile 对象。
    /// </summary>
    IOfficeNewFile? NewWorkbook { get; }

    /// <summary>
    /// 获取一个表示应用程序拼写选项的 SpellingOptions 对象。
    /// </summary>
    IExcelSpellingOptions? SpellingOptions { get; }

    /// <summary>
    /// 获取一个 Speech 对象。
    /// </summary>
    IExcelSpeech? Speech { get; }

    /// <summary>
    /// 获取或设置一个值，指示为另一个国家/地区的标准纸张大小格式化的文档（例如 A4）是否自动调整，以便在您所在国家/地区的标准纸张大小（例如 Letter）上正确打印。
    /// </summary>
    bool MapPaperSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为 Microsoft Excel 应用程序显示“新建工作簿”任务窗格。
    /// </summary>
    bool ShowStartupDialog { get; set; }

    /// <summary>
    /// 获取或设置用作小数分隔符的字符。
    /// </summary>
    string DecimalSeparator { get; set; }

    /// <summary>
    /// 获取或设置用作千位分隔符的字符。
    /// </summary>
    string ThousandsSeparator { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用 Microsoft Excel 的系统分隔符。
    /// </summary>
    bool UseSystemSeparators { get; set; }

    /// <summary>
    /// 获取调用用户定义函数的单元格作为 Range 对象。
    /// </summary>
    IExcelRange? ThisCell { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示“文档操作”任务窗格。
    /// </summary>
    bool DisplayDocumentActionTaskPane { get; set; }

    /// <summary>
    /// 打开“XML 源”任务窗格并显示由 XmlMap 参数指定的 XML 映射。
    /// </summary>
    /// <param name="xmlMap">要在任务窗格中显示的 XML 映射。</param>
    void DisplayXMLSourcePane(object? xmlMap = null);

    /// <summary>
    /// 获取一个布尔值，指示 Microsoft Excel 中是否提供 XML 功能。
    /// </summary>
    bool ArbitraryXMLSupportAvailable { get; }

    /// <summary>
    /// 获取或设置应用程序中使用的度量单位。
    /// </summary>
    int MeasurementUnit { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示用户选择文本时是否显示迷你工具栏。
    /// </summary>
    bool ShowSelectionFloaties { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示用户在工作簿窗口中右键单击时是否显示迷你工具栏。
    /// </summary>
    bool ShowMenuFloaties { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在功能区中显示“开发工具”选项卡。
    /// </summary>
    bool ShowDevTools { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用支持预览的库时是否显示库预览。
    /// </summary>
    bool EnableLivePreview { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示文档属性面板。
    /// </summary>
    bool DisplayDocumentInformationPanel { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用 ClearType 在菜单、功能区、对话框文本中显示字体。
    /// </summary>
    bool AlwaysUseClearType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示开发人员尝试使用现有函数名称创建新函数时是否发出警报。
    /// </summary>
    bool WarnOnFunctionNameConflict { get; set; }

    /// <summary>
    /// 允许用户以行指定编辑栏的高度。
    /// </summary>
    int FormulaBarHeight { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在构建单元格公式时是否显示相关函数和已定义名称的列表。
    /// </summary>
    bool DisplayFormulaAutoComplete { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是使用传统表示法方法还是新的结构化引用表示法方法来引用公式中的表格。
    /// </summary>
    XlGenerateTableRefs GenerateTableRefs { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Office 帮助查看器的 IAssistance 对象。
    /// </summary>
    IOfficeAssistance? Assistance { get; }

    /// <summary>
    /// 运行所有挂起的到 OLEDB 和 OLAP 数据源的查询。
    /// </summary>
    void CalculateUntilAsyncQueriesDone();

    /// <summary>
    /// 获取或设置一个值，指示当用户尝试执行影响比 Office 中心 UI 中指定的更多单元格的操作时，是否显示警报消息。
    /// </summary>
    bool EnableLargeOperationAlert { get; set; }

    /// <summary>
    /// 获取或设置超出该值将触发警报的操作所需的最大单元格数。
    /// </summary>
    int LargeOperationCellThousandCount { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当工作表由 VBA 代码计算时，是否执行对 OLAP 数据源的异步查询。
    /// </summary>
    bool DeferAsyncQueries { get; set; }

    /// <summary>
    /// 获取一个 MultiThreadedCalculation 对象，该对象控制 Excel 2007 中新增的多线程重新计算设置。
    /// </summary>
    IExcelMultiThreadedCalculation? MultiThreadedCalculation { get; }

    /// <summary>
    /// 返回指定 URL 的站点上运行的 SharePoint Foundation 实例的版本号。
    /// </summary>
    /// <param name="url">要检查的站点的 URL。</param>
    /// <returns>SharePoint 版本号。</returns>
    int? SharePointVersion(string url);

    /// <summary>
    /// 获取活动加密会话。
    /// </summary>
    int ActiveEncryptionSession { get; }

    /// <summary>
    /// 获取或设置一个值，指示 Excel 是否使用高质量模式打印图形。
    /// </summary>
    bool HighQualityModeForGraphics { get; set; }

    /// <summary>
    /// 获取一个 FileExportConverters 集合，该集合表示可用于 Microsoft Excel 保存文件的所有文件转换器。
    /// </summary>
    IExcelFileExportConverters? FileExportConverters { get; }

    /// <summary>
    /// 获取应用程序中当前加载的 SmartArt 布局集。
    /// </summary>
    IOfficeSmartArtLayouts? SmartArtLayouts { get; }

    /// <summary>
    /// 获取应用程序中当前加载的 SmartArt 快速样式集。
    /// </summary>
    IOfficeSmartArtQuickStyles? SmartArtQuickStyles { get; }

    /// <summary>
    /// 获取应用程序中当前加载的颜色样式集。
    /// </summary>
    IOfficeSmartArtColors? SmartArtColors { get; }

    /// <summary>
    /// 获取一个 AddIns2 对象集合，该集合表示 Microsoft Excel 中当前可用或打开的所有加载项，无论它们是否安装。
    /// </summary>
    IExcelAddIns2? AddIns2 { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否打开与打印机的通信。
    /// </summary>
    bool PrintCommunication { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Excel 是否允许在计算群集上运行 XLL 加载项中的用户定义函数。
    /// </summary>
    bool UseClusterConnector { get; set; }

    /// <summary>
    /// 获取或设置用于运行 XLL 加载项中的用户定义函数的高性能计算（HPC）群集连接器的名称。
    /// </summary>
    string ClusterConnector { get; set; }

    /// <summary>
    /// 获取一个 ProtectedViewWindows 集合，该集合表示应用程序中打开的所有“受保护的视图”窗口。
    /// </summary>
    IExcelProtectedViewWindows? ProtectedViewWindows { get; }

    /// <summary>
    /// 获取一个表示活动“受保护的视图”窗口（最顶层的窗口）的 ProtectedViewWindow 对象。
    /// </summary>
    IExcelProtectedViewWindow? ActiveProtectedViewWindow { get; }

    /// <summary>
    /// 获取一个值，指示指定的工作簿是否在“受保护的视图”窗口中打开。
    /// </summary>
    bool IsSandboxed { get; }

    /// <summary>
    /// 获取或设置一个值，指示 Excel 是否使用 ISO 8601 格式保存日期和时间值。
    /// </summary>
    bool SaveISO8601Dates { get; set; }

    /// <summary>
    /// 获取由指定 _Application 对象表示的 Microsoft Excel 2010 实例的句柄。
    /// </summary>
    object HinstancePtr { get; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 在打开文件之前如何验证文件。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoFileValidationMode FileValidation { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Excel 如何验证数据透视表报告的数据缓存的内容。
    /// </summary>
    XlFileValidationPivotMode FileValidationPivot { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示快速分析工具。
    /// </summary>
    bool ShowQuickAnalysis { get; set; }

    /// <summary>
    /// 获取 QuickAnalysis 对象。
    /// </summary>
    IExcelQuickAnalysis? QuickAnalysis { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用快速填充功能。
    /// </summary>
    bool FlashFill { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用宏动画。
    /// </summary>
    bool EnableMacroAnimations { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用图表数据点跟踪。
    /// </summary>
    bool ChartDataPointTrack { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用快速填充模式。
    /// </summary>
    bool FlashFillMode { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否合并实例。
    /// </summary>
    bool MergeInstances { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否启用检查文件扩展名。
    /// </summary>
    bool EnableCheckFileExtensions { get; set; }

    /// <summary>
    /// 创建一个工作簿
    /// </summary>
    /// <param name="templatePath">模板文件</param>
    /// <returns></returns>
    [IgnoreGenerator]
    IExcelWorkbook CreateFrom(string templatePath);

    /// <summary>
    /// 创建一个空白工作簿
    /// </summary>
    /// <returns></returns>
    [IgnoreGenerator]
    IExcelWorkbook BlankWorkbook();

    /// <summary>
    /// 打开一个工作簿
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <returns></returns>
    [IgnoreGenerator]
    IExcelWorkbook Open(string filePath);

    /// <summary>
    /// 打开工作簿
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="updateLinks">是否更新链接</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="format">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="ignoreReadOnlyRecommended">是否忽略只读建议</param>
    /// <param name="origin">文本来源</param>
    /// <param name="delimiter">文本分隔符</param>
    /// <param name="editable">是否可编辑</param>
    /// <param name="notify">是否通知</param>
    /// <param name="converter">格式转换器</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <returns>打开的工作簿对象</returns>
    [IgnoreGenerator]
    IExcelWorkbook? OpenWorkbook(string filename, int updateLinks = 0, bool readOnly = false,
                                     int format = 1, string password = "", string writeResPassword = "",
                                     bool ignoreReadOnlyRecommended = false, int origin = 0,
                                     string delimiter = ",", bool editable = true, bool notify = false,
                                     int converter = 0, bool addToMru = true);

    #region 事件

    /// <summary>
    /// 当新建工作簿时触发
    /// </summary>
    event WorkbookNewEventHandler WorkbookNew;

    /// <summary>
    /// 当工作簿打开时触发
    /// </summary>
    event WorkbookOpenEventHandler WorkbookOpen;

    /// <summary>
    /// 当工作簿被激活时触发
    /// </summary>
    event WorkbookActivateEventHandler WorkbookActivate;

    /// <summary>
    /// 当工作簿被取消激活时触发
    /// </summary>
    event WorkbookDeactivateEventHandler WorkbookDeactivate;

    /// <summary>
    /// 当工作簿即将关闭时触发
    /// </summary>
    event WorkbookBeforeCloseEventHandler WorkbookBeforeClose;

    /// <summary>
    /// 当工作簿即将保存时触发
    /// </summary>
    event WorkbookBeforeSaveEventHandler WorkbookBeforeSave;

    /// <summary>
    /// 当工作表内容发生改变时触发
    /// </summary>
    event SheetChangeEventHandler SheetChange;

    /// <summary>
    /// 当工作表被激活时触发
    /// </summary>
    event SheetActivateEventHandler SheetActivate;

    /// <summary>
    /// 当工作表被取消激活时触发
    /// </summary>
    event SheetDeactivateEventHandler SheetDeactivate;

    /// <summary>
    /// 当工作表选择区域发生改变时触发
    /// </summary>
    event SheetSelectionChangeEventHandler SheetSelectionChange;

    /// <summary>
    /// 当工作表被双击前触发
    /// </summary>
    event SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;

    /// <summary>
    /// 当工作表被右键单击前触发
    /// </summary>
    event SheetBeforeRightClickEventHandler SheetBeforeRightClick;

    /// <summary>
    /// 当工作表计算完成时触发
    /// </summary>
    event SheetCalculateEventHandler SheetCalculate;

    /// <summary>
    /// 当Excel应用程序窗口大小改变时触发
    /// </summary>
    event WindowResizeEventHandler WindowResize;

    /// <summary>
    /// 当Excel应用程序窗口失去焦点或变为非活动状态时触发
    /// </summary>
    event WindowDeactivateEventHandler WindowDeactivate;

    /// <summary>
    /// 当Excel应用程序窗口获得焦点或变为活动状态时触发
    /// </summary>
    event WindowActivateEventHandler WindowActivate;
    #endregion
}