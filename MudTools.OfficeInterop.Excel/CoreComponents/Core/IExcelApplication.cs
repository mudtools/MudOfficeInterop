//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel
{
    /// <summary>
    /// Excel应用程序接口
    /// </summary>
    public interface IExcelApplication : IOfficeApplication
    {
        #region 基础属性

        /// <summary>
        /// 获取或设置单元格拖放功能是否启用
        /// </summary>
        public bool CellDragAndDrop { get; set; }

        /// <summary>
        /// 获取或设置是否启用实时预览
        /// </summary>
        public bool EnableLivePreview { get; set; }

        /// <summary>
        /// 获取或设置是否显示浮动工具栏
        /// </summary>
        public bool ShowSelectionFloaties { get; set; }

        /// <summary>
        /// 获取或设置是否显示开发工具选项卡
        /// </summary>
        public bool ShowDevTools { get; set; }

        /// <summary>
        /// 获取或设置是否忽略远程请求
        /// </summary>
        bool IgnoreRemoteRequests { get; set; }

        /// <summary>
        /// 获取Excel应用程序窗口的句柄
        /// </summary>
        int? Hwnd { get; }

        /// <summary>
        /// 获取用于查找操作的单元格格式
        /// </summary>
        IExcelCellFormat FindFormat { get; }

        /// <summary>
        /// 获取或设置批注指示器的显示模式
        /// </summary>
        XlCommentDisplayMode DisplayCommentIndicator { get; set; }

        /// <summary>
        /// 获取或设置鼠标指针的类型
        /// </summary>
        XlMousePointer Cursor { get; set; }

        /// <summary>
        /// 获取或设置应用程序是否处于用户控制状态
        /// 对应 Application.UserControl 属性
        /// </summary>
        bool UserControl { get; set; }

        /// <summary>
        /// 获取或设置应用程序是否显示警告信息
        /// 对应 Application.DisplayAlerts 属性
        /// </summary>
        bool DisplayAlerts { get; set; }

        /// <summary>
        /// 获取或设置应用程序是否启用事件
        /// 对应 Application.EnableEvents 属性
        /// </summary>
        bool EnableEvents { get; set; }

        /// <summary>
        /// 获取或设置应用程序是否启用动画
        /// 对应 Application.EnableAnimations 属性
        /// </summary>
        bool EnableAnimations { get; set; }

        /// <summary>
        /// 获取或设置取消键的行为方式
        /// </summary>
        XlEnableCancelKey EnableCancelKey { get; set; }

        /// <summary>
        /// 获取应用程序的当前用户
        /// 对应 Application.UserName 属性
        /// </summary>
        string UserName { get; set; }

        /// <summary>
        /// 获取应用程序的组织名称
        /// 对应 Application.OrganizationName 属性
        /// </summary>
        string OrganizationName { get; }


        /// <summary>
        /// 获取或设置应用程序窗口的标题栏文本。
        /// </summary>
        string Caption { get; set; }

        /// <summary>
        /// 获取或设置新建工作簿中包含的工作表数量。
        /// </summary>
        int SheetsInNewWorkbook { get; set; }

        /// <summary>
        /// 获取或设置应用程序的默认字体名称。
        /// </summary>
        string StandardFont { get; set; }

        /// <summary>
        /// 获取或设置应用程序的默认字体大小。
        /// </summary>
        double StandardFontSize { get; set; }

        /// <summary>
        /// 获取或设置是否启用迭代计算。
        /// </summary>
        bool Iteration { get; set; }

        /// <summary>
        /// 获取或设置迭代计算时的最大误差。
        /// </summary>
        double MaxChange { get; set; }

        /// <summary>
        /// 获取或设置迭代计算时的最大迭代次数。
        /// </summary>
        int MaxIterations { get; set; }

        /// <summary>
        /// 获取或设置是否在状态栏显示计算过程。
        /// </summary>
        bool DisplayInsertOptions { get; set; }

        /// <summary>
        /// 获取或设置是否显示最近使用的文件列表。
        /// </summary>
        bool DisplayRecentFiles { get; set; }


        /// <summary>
        /// 获取或设置是否允许在单元格内直接编辑。
        /// </summary>
        bool EditDirectlyInCell { get; set; }

        /// <summary>
        /// 获取或设置是否启用自动完成功能。
        /// </summary>
        bool EnableAutoComplete { get; set; }

        /// <summary>
        /// 获取或设置是否显示图表工具提示中的系列名称。
        /// </summary>
        bool ShowChartTipNames { get; set; }

        /// <summary>
        /// 获取或设置是否显示图表工具提示中的数值。
        /// </summary>
        bool ShowChartTipValues { get; set; }

        /// <summary>
        /// 获取或设置是否显示屏幕提示。
        /// </summary>
        bool ShowToolTips { get; set; }

        /// <summary>
        /// 获取或设置打开文件时是否提示更新链接。
        /// </summary>
        bool AskToUpdateLinks { get; set; }

        /// <summary>
        /// 获取或设置覆盖文件时是否显示警告。
        /// </summary>
        bool AlertBeforeOverwriting { get; set; }

        /// <summary>
        /// 获取或设置过渡菜单键。
        /// </summary>
        string TransitionMenuKey { get; set; }

        /// <summary>
        /// 获取或设置过渡菜单键的操作。
        /// </summary>
        int TransitionMenuKeyAction { get; set; }

        /// <summary>
        /// 获取或设置按 Enter 键后是否移动选定区域。
        /// </summary>
        bool MoveAfterReturn { get; set; }

        /// <summary>
        /// 获取或设置按 Enter 键后移动选定区域的方向。
        /// </summary>
        XlDirection MoveAfterReturnDirection { get; set; }

        /// <summary>
        /// 获取应用程序窗口的可用高度（像素）。
        /// </summary>
        double UsableHeight { get; }

        /// <summary>
        /// 获取应用程序窗口的可用宽度（像素）。
        /// </summary>
        double UsableWidth { get; }

        /// <summary>
        /// 获取或设置状态栏显示的文本。
        /// </summary>
        string StatusBar { get; set; }

        #endregion

        #region 工作簿管理

        /// <summary>
        /// 获取应用程序中的工作簿集合
        /// 对应 Application.Workbooks 属性
        /// </summary>
        IExcelWorkbooks Workbooks { get; }

        /// <summary>
        /// 获取当前活动的工作簿
        /// 对应 Application.ActiveWorkbook 属性
        /// </summary>
        IExcelWorkbook? ActiveWorkbook { get; }

        /// <summary>
        /// 获取当前活动的窗口
        /// 对应 Application.ActiveWindow 属性
        /// </summary>
        IExcelWindow? ActiveWindow { get; }

        /// <summary>
        /// 获取工作表函数对象
        /// 对应 Application.WorksheetFunction 属性
        /// </summary>
        IExcelWorksheetFunction? WorksheetFunction { get; }

        /// <summary>
        /// 获取应用程序的ThisWorkbook
        /// 对应 Application.ThisWorkbook 属性
        /// </summary>
        IExcelWorkbook? ThisWorkbook { get; }

        /// <summary>
        /// 获取工作簿的数量
        /// </summary>
        int WorkbooksCount { get; }

        /// <summary>
        /// 获取指定索引的工作簿
        /// </summary>
        /// <param name="index">工作簿索引</param>
        /// <returns>工作簿对象</returns>
        IExcelWorkbook? GetWorkbook(int index);

        /// <summary>
        /// 获取指定名称的工作簿
        /// </summary>
        /// <param name="name">工作簿名称</param>
        /// <returns>工作簿对象</returns>
        IExcelWorkbook? GetWorkbook(string name);

        #endregion

        #region 工作表管理

        /// <summary>
        /// 获取当前活动的工作表
        /// 对应 Application.ActiveSheet 属性
        /// </summary>
        IExcelCommonSheet? ActiveSheet { get; }

        /// <summary>
        /// 获取当前活动的工作表
        /// 对应 Application.ActiveSheet 属性
        /// </summary>
        IExcelWorksheet? ActiveSheetWrap { get; }

        /// <summary>
        /// 获取当前活动的单元格区域
        /// 对应 Application.ActiveCell 属性
        /// </summary>
        IExcelRange? ActiveCell { get; }

        /// <summary>
        /// 表示用户当前在Excel界面中选中的对象，并且在绝大多数情况下，它可能返回的是一个Range、Chart、ChartObject、Shape、PivotTable。
        /// </summary>
        object? Selection { get; }

        /// <summary>
        /// 表示用户当前在Excel界面中选中的对象，并且在绝大多数情况下，它可能返回的是一个Range、Chart、ChartObject、Shape、PivotTable。
        /// </summary>
        /// <typeparam name="T">IExcelRange、IExcelChart、IExcelShape等类型</typeparam>
        /// <returns></returns>
        T? SelectionWrap<T>() where T : IDisposable;

        /// <summary>
        /// 表示用户当前在Excel界面中选中Range对象，用于操作工作表中的单元格或单元格区域。
        /// </summary>
        IExcelRange? SelectionRange { get; }

        /// <summary>
        /// 获取当前活动行集合
        /// </summary>
        IExcelRows ActiveRows { get; }

        /// <summary>
        /// 获取当前活动列集合
        /// </summary>
        IExcelColumns ActiveColumns { get; }

        /// <summary>
        /// 获取当前活动工作表的所有列集合
        /// 对应 Application.Columns 属性
        /// </summary>
        IExcelRange Columns { get; }

        /// <summary>
        /// 获取当前活动工作表的所有行集合
        /// 对应 Application.Rows 属性
        /// </summary>
        IExcelRange Rows { get; }

        /// <summary>
        /// 获取当前活动工作表的所有单元格集合
        /// 对应 Application.Cells 属性
        /// </summary>
        IExcelRange Cells { get; }

        /// <summary>
        /// 获取工作表集合
        /// 对应 Application.Sheets 属性
        /// </summary>
        IExcelSheets Sheets { get; }

        /// <summary>
        /// 获取工作表集合
        /// 对应 Application.Worksheets 属性
        /// </summary>
        IExcelSheets Worksheets { get; }

        /// <summary>
        /// 获取最近使用的文件列表
        /// 对应 Application.RecentFiles 属性
        /// </summary>
        IExcelRecentFiles RecentFiles { get; }

        /// <summary>
        /// 获取窗口集合
        /// 对应 Application.Windows 属性
        /// </summary>
        IExcelWindows Windows { get; }

        /// <summary>
        /// 获取受保护视图窗口集合
        /// </summary>
        IExcelProtectedViewWindows ProtectedViewWindows { get; }

        /// <summary>
        /// 获取活动的受保护视图窗口
        /// </summary>
        IExcelProtectedViewWindow? ActiveProtectedViewWindow { get; }

        /// <summary>
        /// 获取当前的插件集合。
        /// </summary>
        IExcelAddIns? AddIns { get; }
        #endregion

        #region 计算设置

        /// <summary>
        /// 获取或设置计算模式
        /// 对应 Application.Calculation 属性
        /// </summary>
        XlCalculation Calculation { get; set; }

        /// <summary>
        /// 获取或设置是否自动重算
        /// 对应 Application.CalculateBeforeSave 属性
        /// </summary>
        bool CalculateBeforeSave { get; set; }

        /// <summary>
        /// 获取或设置是否启用多线程计算
        /// 对应 Application.MultiThreadedCalculation 属性
        /// </summary>
        bool MultiThreadedCalculation { get; set; }

        /// <summary>
        /// 手动计算所有打开的工作簿
        /// 对应 Application.CalculateFull 方法
        /// </summary>
        void Calculate();

        /// <summary>
        /// 重新计算所有打开的工作簿
        /// 对应 Application.CalculateFullRebuild 方法
        /// </summary>
        void CalculateFull();

        /// <summary>
        /// 计算指定工作表
        /// </summary>
        /// <param name="worksheet">要计算的工作表</param>
        void CalculateWorksheet(IExcelWorksheet worksheet);


        /// <summary>
        /// 获取多个区域的交集区域
        /// </summary>
        /// <param name="ranges">要计算交集的区域集合</param>
        /// <returns>交集区域（无交集时返回 null）</returns>
        /// <exception cref="ArgumentNullException">输入区域为空时抛出</exception>
        /// <exception cref="ArgumentException">区域数量不足时抛出</exception>
        IExcelRange Intersect(params IExcelRange[] ranges);

        /// <summary>
        /// 获取多个区域的并集区域
        /// </summary>
        /// <param name="ranges">要合并的区域集合</param>
        /// <returns>并集区域（无效输入时返回 null）</returns>
        /// <exception cref="ArgumentNullException">输入区域为空时抛出</exception>
        /// <exception cref="ArgumentException">区域数量不足时抛出</exception>
        IExcelRange Union(params IExcelRange[] ranges);

        /// <summary>
        /// 检查多个区域是否相邻（可形成连续区域）
        /// </summary>
        /// <param name="ranges">要检查的区域集合</param>
        /// <returns>如果区域相邻返回true，否则返回false</returns>
        bool AreContiguous(params IExcelRange[] ranges);

        /// <summary>
        /// 合并区域并应用格式
        /// </summary>
        /// <param name="format">要应用的单元格格式</param>
        /// <param name="ranges">要合并和格式化的区域集合</param>
        void FormatUnionRange(IExcelCellFormat format, params IExcelRange[] ranges);

        /// <summary>
        /// 将单元格格式应用到指定区域
        /// </summary>
        /// <param name="range">目标区域</param>
        /// <param name="format">单元格格式配置</param>
        /// <param name="applyToSubAreas">是否应用到不连续的子区域</param>
        void ApplyCellFormat(IExcelRange range, IExcelCellFormat format, bool applyToSubAreas = false);

        /// <summary>
        /// 计算 Excel 公式
        /// </summary>
        /// <param name="formula">Excel 公式字符串</param>
        /// <returns>计算结果</returns>
        object Evaluate(string formula);

        /// <summary>
        /// 计算 Excel 公式（带参数）
        /// </summary>
        /// <param name="formula">包含占位符的公式模板</param>
        /// <param name="args">公式参数</param>
        /// <returns>计算结果</returns>
        object Evaluate(string formula, params object[] args);

        /// <summary>
        /// 强类型计算结果（数值类型）
        /// </summary>
        /// <remarks>
        /// <code>
        /// double sum = evaluator.EvaluateToNumber("=SUM(1, 2, 3)");
        /// </code>
        /// </remarks>
        double EvaluateToNumber(string formula);

        /// <summary>
        /// 强类型计算结果（布尔类型）
        /// </summary>
        /// <remarks>
        /// <code>
        ///  bool isTrue = evaluator.EvaluateToBool("=AND(TRUE, 1=1)");
        /// </code>
        /// </remarks>
        bool EvaluateToBool(string formula);

        /// <summary>
        /// 强类型计算结果（日期类型）
        /// </summary>
        /// <remarks>
        /// <code>
        ///  DateTime date = evaluator.EvaluateToDateTime("=DATE(2023, 12, 31)");
        /// </code>
        /// </remarks>
        DateTime EvaluateToDateTime(string formula);

        /// <summary>
        /// 强类型计算结果（字符串类型）
        /// </summary>
        /// <remarks>
        /// <code>
        ///  string name = "John";
        ///  int age = 30;
        ///  string result = evaluator.EvaluateToString( "=\"{0} is {1} years old\"", name, age);
        /// </code>
        /// </remarks>
        string EvaluateToString(string formula);

        /// <summary>
        /// 计算并返回二维数组结果（用于范围计算结果）
        /// </summary>
        /// <remarks>
        /// <code>
        ///  object[,] array = evaluator.EvaluateToArray("={1,2,3;4,5,6}");
        /// </code>
        /// </remarks>
        object[,] EvaluateToArray(string formula);
        #endregion

        #region 屏幕和显示

        /// <summary>
        /// 获取或设置是否显示滚动条
        /// 对应 Application.DisplayScrollBars 属性
        /// </summary>
        bool DisplayScrollBars { get; set; }

        /// <summary>
        /// 获取或设置是否以全屏模式显示
        /// 对应 Application.DisplayFullScreen 属性
        /// </summary>
        bool DisplayFullScreen { get; set; }

        /// <summary>
        /// 获取或设置是否在任务栏中显示窗口
        /// 对应 Application.ShowWindowsInTaskbar 属性
        /// </summary>
        bool ShowWindowsInTaskbar { get; set; }

        /// <summary>
        /// 获取或设置是否显示公式栏
        /// 对应 Application.DisplayFormulaBar 属性
        /// </summary>
        bool DisplayFormulaBar { get; set; }

        /// <summary>
        /// 获取或设置屏幕更新是否启用
        /// 对应 Application.ScreenUpdating 属性
        /// </summary>
        bool ScreenUpdating { get; set; }

        /// <summary>
        /// 获取或设置是否显示状态栏
        /// 对应 Application.DisplayStatusBar 属性
        /// </summary>
        bool DisplayStatusBar { get; set; }

        #endregion

        #region 文件操作

        /// <summary>
        /// 创建一个新的空白工作簿
        /// </summary>
        /// <returns>新建的空白工作簿</returns>
        IExcelWorkbook BlankWorkbook();

        /// <summary>
        /// 打开工作簿
        /// 对应 Application.Workbooks.Open 方法
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
        IExcelWorkbook OpenWorkbook(string filename, int updateLinks = 0, bool readOnly = false,
                                   int format = 1, string password = "", string writeResPassword = "",
                                   bool ignoreReadOnlyRecommended = false, int origin = 0,
                                   string delimiter = ",", bool editable = true, bool notify = false,
                                   int converter = 0, bool addToMru = true);

        /// <summary>
        /// 新建工作簿
        /// 对应 Application.Workbooks.Add 方法
        /// </summary>
        /// <param name="template">模板文件路径</param>
        /// <returns>新建的工作簿对象</returns>
        IExcelWorkbook NewWorkbook(string template = "");

        /// <summary>
        /// 从模板创建工作簿
        /// </summary>
        /// <param name="templatePath">模板文件路径</param>
        /// <returns>基于模板创建的工作簿</returns>
        IExcelWorkbook CreateFrom(string templatePath);

        /// <summary>
        /// 打开已有的 Excel 文件
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>打开的工作簿对象</returns>
        IExcelWorkbook Open(string filePath);

        /// <summary>
        /// 保存所有工作簿
        /// 对应 Application.Save 方法
        /// </summary>
        void SaveAll();

        /// <summary>
        /// 关闭所有工作簿
        /// 对应 Application.Workbooks.Close 方法
        /// </summary>
        /// <param name="saveChanges">是否保存更改</param>
        void CloseAllWorkbooks(bool saveChanges = true);

        #endregion

        #region 宏和自动化
        /// <summary>
        /// 在 Excel 的撤销列表中注册一个自定义的撤销操作。
        /// 当用户选择此操作时，Excel 将尝试运行指定的宏。
        /// </summary>
        /// <param name="undoText">显示在 Excel 撤销列表中的文本。</param>
        /// <param name="macroProcedureName">
        /// 当用户选择撤销时要执行的宏的完整名称。
        /// 通常格式为 "WorkbookName.xlam!MacroName"。
        /// </param>
        /// <exception cref="System.ArgumentNullException">
        /// 如果 undoText 或 macroProcedureName 为 null 或空。
        /// </exception>
        /// <exception cref="System.InvalidOperationException">
        /// 如果内部的 Application 对象为 null。
        /// </exception>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// 如果与 Excel 的交互失败（例如，宏名称格式错误），可能会抛出 COM 异常。
        /// </exception>
        /// <remarks>
        /// 调用此方法后，必须确保指定的 macroProcedureName 对应的宏存在于 Excel 会话中，
        /// 并且该宏实现了相应的撤销逻辑。
        /// </remarks>
        void OnUndo(string undoText, string macroProcedureName);
        /// <summary>
        /// 执行Excel 4.0宏函数
        /// 对应 Application.ExecuteExcel4Macro 方法
        /// </summary>
        /// <param name="macro">宏函数</param>
        /// <returns>执行结果</returns>
        object ExecuteExcel4Macro(string macro);
        #endregion

        #region 系统信息

        /// <summary>
        /// 获取操作系统版本
        /// 对应 Application.OperatingSystem 属性
        /// </summary>
        string OperatingSystem { get; }

        /// <summary>
        /// 获取系统内存信息
        /// 对应 Application.MemoryFree 属性
        /// </summary>
        int MemoryFree { get; }

        /// <summary>
        /// 获取总内存信息
        /// 对应 Application.MemoryTotal 属性
        /// </summary>
        int MemoryTotal { get; }

        /// <summary>
        /// 获取启动时间
        /// </summary>
        DateTime StartupTime { get; }

        /// <summary>
        /// 获取运行时间（秒）
        /// </summary>
        double RunTime { get; }

        #endregion

        #region 国际化设置
        /// <summary>
        /// 获取与应用程序国际设置相关的信息。
        /// </summary>
        /// <param name="index">指定要获取的国际设置项。</param>
        /// <returns>与指定索引对应的国际设置值。</returns>
        string? International(XlApplicationInternational index);
        /// <summary>
        /// 获取或设置测量单位
        /// 对应 Application.MeasurementUnit 属性
        /// </summary>
        int MeasurementUnit { get; set; }

        /// <summary>
        /// 获取或设置默认文件路径
        /// 对应 Application.DefaultFilePath 属性
        /// </summary>
        string DefaultFilePath { get; set; }

        /// <summary>
        /// 获取或设置模板路径
        /// 对应 Application.TemplatesPath 属性
        /// </summary>
        string TemplatesPath { get; }

        #endregion

        #region 剪贴板操作

        /// <summary>
        /// 复制对象到剪贴板
        /// 对应 Application.CutCopyMode 属性
        /// </summary>
        XlCutCopyMode CutCopyMode { get; set; }

        /// <summary>
        /// 清除剪贴板
        /// </summary>
        void ClearClipboard();

        /// <summary>
        /// 获取剪贴板内容类型
        /// </summary>
        string ClipboardContentType { get; }

        #endregion

        #region 对话框和用户界面

        /// <summary>
        /// 显示"打开"对话框获取文件名
        /// </summary>
        /// <param name="filter">文件过滤器（如："Excel Files (*.xlsx), *.xlsx|All Files (*.*), *.*"）</param>
        /// <param name="filterIndex">默认过滤器索引</param>
        /// <param name="title">对话框标题</param>
        /// <param name="buttonText">按钮文本（仅Mac）</param>
        /// <param name="multiSelect">是否允许多选</param>
        /// <returns>选择的文件路径（多选时为数组），取消时为null</returns>
        public IList<string> GetOpenFilenames(
            string filter = "所有文件 (*.*)|*.*",
            int filterIndex = 1,
            string title = "打开文件",
            string buttonText = "打开",
            bool multiSelect = true);

        /// <summary>
        /// 显示文件打开对话框
        /// 对应 Application.GetOpenFilename 方法
        /// </summary>
        /// <param name="fileFilter">文件过滤器</param>
        /// <param name="title">对话框标题</param>
        /// <returns>选择的文件路径</returns>
        object GetOpenFilename(string fileFilter = "所有文件 (*.*)|*.*",
                                 string title = "打开文件");

        /// <summary>
        /// 显示"另存为"对话框获取文件名
        /// </summary>
        /// <param name="initialFilename">初始文件名</param>
        /// <param name="filter">文件过滤器</param>
        /// <param name="filterIndex">默认过滤器索引</param>
        /// <param name="title">对话框标题</param>
        /// <param name="buttonText">按钮文本（仅Mac）</param>
        /// <returns>选择的文件路径，取消时为null</returns>
        string GetSaveAsFilename(
           string initialFilename = "",
           string filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*",
           int filterIndex = 1,
           string title = "另存为",
           string buttonText = "保存");

        /// <summary>
        /// 显示文件保存对话框
        /// 对应 Application.GetSaveAsFilename 方法
        /// </summary>
        /// <param name="initialFilename">初始文件名</param>
        /// <param name="fileFilter">文件过滤器</param>
        /// <param name="title">对话框标题</param>
        /// <returns>保存的文件路径</returns>
        string GetSaveFilename(string initialFilename = "",
                                 string fileFilter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*",
                                 string title = "另存为");

        /// <summary>
        /// 获取自定义序列的序号
        /// </summary>
        /// <param name="listItems">序列项目</param>
        /// <returns>序列序号（未找到返回-1）</returns>
        int GetCustomListNum(IList<string> listItems);

        /// <summary>
        /// 获取自定义序列的内容
        /// </summary>
        /// <param name="listNum">序列序号</param>
        /// <returns>序列内容数组</returns>
        IList<string> GetCustomListContents(int listNum);

        /// <summary>
        /// 显示一个提示用户输入信息的对话框，并返回 object 类型的结果。
        /// 注意：如果用户点击取消，通常会抛出 COMException (HResult = 0x800A03EC)。
        /// </summary>
        /// <param name="prompt">显示在对话框中的消息。</param>
        /// <param name="title">对话框标题栏的文本。</param>
        /// <param name="defaultValue">文本框中的默认值。</param>
        /// <param name="left">对话框距屏幕左边的距离。</param>
        /// <param name="top">对话框距屏幕上边的距离。</param>
        /// <param name="helpFile">帮助文件的名称。</param>
        /// <param name="helpContextID">帮助文件中帮助主题的上下文编号。</param>
        /// <param name="type">
        /// 指定在对话框中返回的数据类型。默认值为 2 (文本)。
        /// 可以是以下值的和：
        /// 0 - 公式, 1 - 数字, 2 - 文本, 4 - 逻辑值, 8 - 单元格引用, 16 - 错误值, 64 - 数组。
        /// 如果类型为 8 (单元格引用)，则返回一个 Range 对象。
        /// </param>
        /// <returns>用户输入的值。类型取决于 type 参数。如果用户取消，则抛出 COMException。</returns>
        object InputBox(
           string prompt,
           object? title = null,
           object? defaultValue = null,
           object? left = null,
           object? top = null,
           object? helpFile = null,
           object? helpContextID = null,
           int? type = 2);

        /// <summary>
        /// 显示一个提示用户输入信息的对话框。
        /// 这是对 Microsoft.Office.Interop.Excel.Application.InputBox 方法的封装。
        /// </summary>
        /// <param name="prompt">显示在对话框中的消息 (最多 255 个字符)。</param>
        /// <param name="title">对话框标题栏的文本。如果省略，则使用应用程序名称。</param>
        /// <param name="defaultValue">文本框中的默认值。如果省略，则文本框为空。</param>
        /// <param name="left">对话框距屏幕左边的距离（以点为单位）。如果省略，则 Excel 设置对话框位置。</param>
        /// <param name="top">对话框距屏幕上边的距离（以点为单位）。如果省略，则 Excel 设置对话框位置。</param>
        /// <param name="helpFile">帮助文件的名称。如果省略，则不显示帮助。</param>
        /// <param name="helpContextID">帮助文件中帮助主题的上下文编号。如果省略，则不显示帮助。</param>
        /// <param name="type">
        /// 指定在对话框中返回的数据类型。默认值为 2 (文本)。
        /// 可以是以下值的和：
        /// 0 - 公式, 1 - 数字, 2 - 文本, 4 - 逻辑值, 8 - 单元格引用, 16 - 错误值, 64 - 数组。
        /// 如果类型为 8 (单元格引用)，则返回一个 Range 对象。
        /// </param>
        /// <returns>
        /// 一个 InputBoxResult 对象，指示用户是点击了确定还是取消，以及返回的值。
        /// 如果用户点击取消，ResultType 为 Cancel，Value 为 null。
        /// 如果发生错误，ResultType 为 Error，Value 可能包含错误信息。
        /// 如果用户点击确定，ResultType 为 Ok，Value 包含用户输入的值（类型取决于 type 参数）。
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// 如果内部的 _application 对象为 null。
        /// </exception>
        InputBoxResult ShowInputBox(
           string prompt,
           string? title = null,
           object? defaultValue = null,
           object? left = null,
           object? top = null,
           object? helpFile = null,
           object? helpContextID = null,
           object? type = null);
        /// <summary>
        /// 显示一个提示用户输入文本的对话框。
        /// </summary>
        /// <param name="prompt">显示在对话框中的消息。</param>
        /// <param name="title">对话框标题栏的文本。</param>
        /// <param name="defaultValue">文本框中的默认文本。</param>
        /// <returns>包含用户输入文本的 InputBoxResult。</returns>
        InputBoxResult<string?> ShowInputBoxText(string prompt, string? title = null, string? defaultValue = null);
        /// <summary>
        /// 显示一个提示用户输入数字的对话框。
        /// </summary>
        /// <param name="prompt">显示在对话框中的消息。</param>
        /// <param name="title">对话框标题栏的文本。</param>
        /// <param name="defaultValue">文本框中的默认数字。</param>
        /// <returns>包含用户输入数字的 InputBoxResult。</returns>
        InputBoxResult<double?> ShowInputBoxNumber(string prompt, string? title = null, double? defaultValue = null);
        /// <summary>
        /// 显示一个提示用户选择单元格引用的对话框。
        /// </summary>
        /// <param name="prompt">显示在对话框中的消息。</param>
        /// <param name="title">对话框标题栏的文本。</param>
        /// <returns>包含用户选择的 Range 对象的 InputBoxResult。</returns>
        InputBoxResult<IExcelRange> ShowInputBoxRangeSelection(string prompt, string? title = null);

        #endregion

        #region 打印设置

        /// <summary>
        /// 获取或设置是否使用系统对话框打印
        /// 对应 Application.UseSystemSeparators 属性
        /// </summary>
        bool UseSystemSeparators { get; set; }

        /// <summary>
        /// 获取或设置默认打印机
        /// </summary>
        string DefaultPrinter { get; set; }

        /// <summary>
        /// 获取打印机列表
        /// </summary>
        /// <returns>打印机名称数组</returns>
        string[] GetPrinterList();

        /// <summary>
        /// 打印预览所有工作簿
        /// </summary>
        void PrintPreviewAll();

        #endregion

        #region 错误处理

        /// <summary>
        /// 获取或设置错误检查选项
        /// 对应 Application.ErrorCheckingOptions 属性
        /// </summary>
        IExcelErrorCheckingOptions ErrorCheckingOptions { get; }

        #endregion

        #region 性能监控

        /// <summary>
        /// 获取性能统计信息
        /// </summary>
        /// <returns>性能统计对象</returns>
        ApplicationPerformance GetPerformanceStats();

        /// <summary>
        /// 重置性能统计
        /// </summary>
        void ResetPerformanceStats();

        /// <summary>
        /// 获取内存使用情况
        /// </summary>
        /// <returns>内存使用信息</returns>
        MemoryInfo GetMemoryInfo();

        /// <summary>
        /// 获取CPU使用率
        /// </summary>
        /// <returns>CPU使用率百分比</returns>
        double GetCPUUsage();

        #endregion

        #region 操作方法
        /// <summary>
        /// Range函数包装
        /// </summary>
        /// <param name="cell1">区域的第一个单元格</param>
        /// <param name="cell2">区域的第二个单元格</param>
        /// <returns>Excel区域对象</returns>
        IExcelRange? Range(object? cell1, object? cell2 = null);

        /// <summary>
        /// 选定指定的区域或对象。
        /// </summary>
        /// <param name="reference">要选定的区域或对象（可以是 Range, Sheet 名称等）。</param>
        /// <param name="scroll">是否滚动到选定区域。</param>
        void Goto(object reference, bool scroll = true);

        /// <summary>
        /// 将公式从一种引用样式转换为另一种。
        /// </summary>
        /// <param name="formula">要转换的公式。</param>
        /// <param name="fromReferenceStyle">源引用样式。</param>
        /// <param name="toReferenceStyle">目标引用样式。</param>
        /// <param name="toAbsolute">如何转换引用（绝对、相对等）。</param>
        /// <param name="relativeTo">相对引用的基准单元格。</param>
        /// <returns>转换后的公式。</returns>
        string ConvertFormula(string formula,
            XlReferenceStyle fromReferenceStyle,
            XlReferenceStyle toReferenceStyle,
            int toAbsolute = 1,
            object? relativeTo = null);

        /// <summary>
        /// 检查指定文本的拼写。
        /// </summary>
        /// <param name="text">要检查的文本。</param>
        /// <param name="customDictionary">自定义词典的名称。</param>
        /// <param name="ignoreUpper">是否忽略全大写单词。</param>
        /// <returns>如果拼写正确返回 True，否则返回 False。</returns>
        bool CheckSpelling(string text, object? customDictionary = null, object? ignoreUpper = null);


        /// <summary>
        /// 为指定的键或键组合指定过程（宏）。
        /// </summary>
        /// <param name="key">键或键组合（例如 "^c" 代表 Ctrl+C）。</param>
        /// <param name="procedure">要运行的过程名称（宏名）。</param>
        void OnKey(string key, string procedure = "");

        /// <summary>
        /// 最小化应用程序
        /// </summary>
        void Minimize();

        /// <summary>
        /// 最大化应用程序
        /// </summary>
        void Maximize();

        /// <summary>
        /// 恢复应用程序
        /// </summary>
        void Restore();

        /// <summary>
        /// 发送按键到应用程序
        /// 对应 Application.SendKeys 方法
        /// </summary>
        /// <param name="keys">按键字符串</param>
        /// <param name="wait">是否等待</param>
        void SendKeys(string keys, bool wait = true);

        /// <summary>
        /// 等待指定时间
        /// 对应 Application.Wait 方法
        /// </summary>
        /// <param name="time">等待到的时间</param>
        void Wait(DateTime time);

        /// <summary>
        /// 延迟指定毫秒数
        /// </summary>
        /// <param name="milliseconds">毫秒数</param>
        void Delay(int milliseconds);
        #endregion

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
}