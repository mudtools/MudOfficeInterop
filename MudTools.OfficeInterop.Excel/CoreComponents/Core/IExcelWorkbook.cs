//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Vbe;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Workbook 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Workbook 的安全访问和操作
/// </summary>
public interface IExcelWorkbook : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取一个值，该值指示工作簿是否受密码保护
    /// </summary>
    /// <value>
    /// 如果工作簿受密码保护，则为 true；否则为 false
    /// </value>
    bool HasPassword { get; }

    /// <summary>
    /// 获取工作簿的名称
    /// 对应 Workbook.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取工作簿的完整路径
    /// 对应 Workbook.FullName 属性
    /// </summary>
    string FullName { get; }

    /// <summary>
    /// 获取工作簿的路径
    /// 对应 Workbook.Path 属性
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取工作簿的多用户编辑状态
    /// 对应 Workbook.MultiUserEditing 属性
    /// </summary>
    bool MultiUserEditing { get; }

    /// <summary>
    /// 获取工作簿是否已保存
    /// 对应 Workbook.Saved 属性
    /// </summary>
    bool Saved { get; set; }

    /// <summary>
    /// 获取工作簿的结构保护状态
    /// 对应 Workbook.ProtectStructure 属性
    /// </summary>
    bool ProtectStructure { get; }

    /// <summary>
    /// 获取或设置工作簿中图形对象的显示方式
    /// 对应 Workbook.DisplayDrawingObjects 属性
    /// </summary>
    XlDisplayDrawingObjects DisplayDrawingObjects { get; set; }

    /// <summary>
    /// 获取工作簿是否受保护
    /// 对应 Workbook.ProtectStructure 属性
    /// </summary>
    bool IsProtected { get; }

    /// <summary>
    /// 获取工作簿的只读状态
    /// 对应 Workbook.ReadOnly 属性
    /// </summary>
    bool ReadOnly { get; }

    /// <summary>
    /// 获取工作簿中的外部数据连接集合
    /// 对应 Workbook.Connections 属性
    /// </summary>
    IExcelConnections? Connections { get; }

    /// <summary>
    /// 获取工作簿的VB工程。
    /// 对应 Workbook.VBProject 属性
    /// </summary>
    IVbeVBProject? VBProject { get; }

    /// <summary>
    /// 获取与工作簿关联的窗口集合
    /// </summary>
    IExcelWindows? Windows { get; }

    /// <summary>
    /// 获取工作簿的修改时间
    /// </summary>
    DateTime ModifiedTime { get; }

    /// <summary>
    /// 获取工作簿的创建时间
    /// </summary>
    DateTime CreatedTime { get; }

    /// <summary>
    /// 获取工作簿的文件大小（字节）
    /// </summary>
    long FileSize { get; }

    /// <summary>
    /// 获取工作簿所在的父对象（通常是Application）
    /// 对应 Workbook.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取工作簿所在的Application对象
    /// 对应 Workbook.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取工作簿的编码名称
    /// 对应 Workbook.CodeName 属性
    /// </summary>
    string CodeName { get; }

    #endregion

    #region 工作表管理 
    /// <summary>
    /// 获取工作簿中的所有工作表集合
    /// </summary>
    IExcelSheets Worksheets { get; }

    /// <summary>
    /// 获取工作簿中的所有工作表集合（包括图表工作表等）
    /// 对应 Workbook.Sheets 属性
    /// </summary>
    IExcelSheets Sheets { get; }

    /// <summary>
    /// 获取工作簿中的工作表数量
    /// </summary>
    int WorksheetCount { get; }

    /// <summary>
    /// 获取指定索引的工作表
    /// </summary>
    /// <param name="index">工作表索引</param>
    /// <returns>工作表对象</returns>
    IExcelWorksheet? GetWorksheet(int index);

    /// <summary>
    /// 获取指定名称的工作表
    /// </summary>
    /// <param name="name">工作表名称</param>
    /// <returns>工作表对象</returns>
    IExcelWorksheet? GetWorksheet(string name);

    /// <summary>
    /// 添加新的工作表
    /// </summary>
    /// <param name="before">添加到指定工作表之前</param>
    /// <param name="after">添加到指定工作表之后</param>
    /// <param name="count">添加的工作表数量</param>
    /// <param name="type">工作表类型</param>
    /// <returns>新创建的工作表对象</returns>
    IExcelWorksheet? AddWorksheet(IExcelWorksheet before = null, IExcelWorksheet after = null,
                                int count = 1, int type = 0);

    /// <summary>
    /// 删除工作表
    /// </summary>
    /// <param name="worksheet">要删除的工作表</param>
    void DeleteWorksheet(IExcelWorksheet worksheet);

    /// <summary>
    /// 获取活动工作表
    /// </summary>
    /// <returns>活动工作表对象</returns>
    IExcelWorksheet ActiveSheet { get; }

    #endregion

    #region 保护和安全

    /// <summary>
    /// 保护工作簿结构和窗口
    /// 对应 Workbook.Protect 方法
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="structure">是否保护结构</param>
    /// <param name="windows">是否保护窗口</param>
    void Protect(string password = "", bool structure = true, bool windows = false);

    /// <summary>
    /// 取消保护工作簿
    /// 对应 Workbook.Unprotect 方法
    /// </summary>
    /// <param name="password">保护密码</param>
    void Unprotect(string password = "");

    /// <summary>
    /// 保护工作簿中的所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    void ProtectAllWorksheets(string password = "");

    /// <summary>
    /// 取消保护工作簿中的所有工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    void UnprotectAllWorksheets(string password = "");

    #endregion

    #region 操作方法

    void ExportAsFixedFormat(
      XlFixedFormatType Type,
      string Filename,
      object? Quality = null,
      object? IncludeDocProperties = null,
      object? IgnorePrintAreas = null,
      object? From = null,
      object? To = null,
      object? OpenAfterPublish = null,
      object? FixedFormatExtClassPtr = null);

    /// <summary>
    /// 保存工作簿
    /// 对应 Workbook.Save 方法
    /// </summary>
    void Save();

    /// <summary>
    /// 另存为工作簿
    /// 对应 Workbook.SaveAs 方法
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="readOnlyRecommended">是否建议只读</param>
    /// <param name="createBackup">是否创建备份</param>
    /// <param name="accessMode">访问模式</param>
    /// <param name="conflictResolution">冲突解决方式</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <param name="local">是否本地格式</param>
    void SaveAs(string filename, XlFileFormat fileFormat = XlFileFormat.xlWorkbookDefault, string? password = null,
                      string? writeResPassword = null, bool? readOnlyRecommended = false, bool? createBackup = false,
                      XlSaveAsAccessMode accessMode = XlSaveAsAccessMode.xlNoChange,
                      XlSaveConflictResolution? conflictResolution = XlSaveConflictResolution.xlLocalSessionChanges,
                      bool? addToMru = true, bool? local = false);

    /// <summary>
    /// 关闭工作簿
    /// 对应 Workbook.Close 方法
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    /// <param name="filename">文件路径</param>
    /// <param name="routeWorkbook">是否发送路由</param>
    void Close(bool saveChanges = true, string filename = "", bool routeWorkbook = false);

    /// <summary>
    /// 激活工作簿
    /// 对应 Workbook.Activate 方法
    /// </summary>
    void Activate();

    /// <summary>
    /// 选择工作簿
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制工作簿
    /// 对应 Workbook.FollowHyperlink 方法（简化实现）
    /// </summary>
    /// <param name="before">复制到指定工作簿之前</param>
    /// <param name="after">复制到指定工作簿之后</param>
    void Copy(IExcelWorkbook before = null, IExcelWorkbook after = null);

    /// <summary>
    /// 打印工作簿
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOut(bool preview = false);

    /// <summary>
    /// 发送工作簿
    /// </summary>
    /// <param name="recipients">收件人</param>
    /// <param name="subject">主题</param>
    /// <param name="returnReceipt">是否要求回执</param>
    void SendMail(string recipients, string subject = "", bool returnReceipt = false);

    #endregion

    #region 高级功能
    /// <summary>
    /// 计算工作簿中的所有公式
    /// 对应 Workbook.CalculateFull 方法
    /// </summary>
    void CalculateAll();

    /// <summary>
    /// 刷新工作簿
    /// </summary>
    void RefreshAll();

    /// <summary>
    /// 应用自动筛选到所有工作表
    /// </summary>
    void AutoFilterAll();

    /// <summary>
    /// 清除工作簿中的所有内容
    /// </summary>
    void ClearAll();


    /// <summary>
    /// 获取工作簿的名称集合
    /// </summary>
    IExcelNames Names { get; }

    /// <summary>
    /// 获取工作簿的样式集合
    /// </summary>
    IExcelStyles Styles { get; }

    /// <summary>
    /// 获取工作簿的图表集合
    /// </summary>
    IExcelSheets Charts { get; }

    /// <summary>
    /// 获取工作簿的透视表缓存集合
    /// </summary>
    IExcelPivotCaches PivotCaches();

    #endregion

    #region 属性设置
    /// <summary>
    /// 获取或设置是否显示滚动条
    /// </summary>
    bool DisplayScrollBars { get; set; }

    /// <summary>
    /// 获取或设置是否显示公式栏
    /// </summary>
    bool DisplayFormulaBar { get; set; }

    /// <summary>
    /// 获取或设置窗口状态
    /// </summary>
    XlWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置窗口高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置窗口宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置窗口左边距
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置窗口顶边距
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置是否启用事件
    /// </summary>
    bool EnableEvents { get; set; }

    /// <summary>
    /// 获取或设置是否启用多线程计算
    /// </summary>
    bool MultiThreadedCalculation { get; set; }

    /// <summary>
    /// 获取或设置计算模式
    /// </summary>
    XlCalculation Calculation { get; set; }

    #endregion

    #region 事件
    /// <summary>
    /// 当工作表数据透视表更改时发生
    /// </summary>
    event WorkBookSheetPivotTableChangeEventHandler WorkBookSheetPivotTableChange;

    /// <summary>
    /// 当窗口大小调整时发生
    /// </summary>
    event WindowResizeEventHandler WindowResize;

    /// <summary>
    /// 当窗口停用时发生
    /// </summary>
    event WindowDeactivateEventHandler WindowDeActivate;

    /// <summary>
    /// 当窗口激活时发生
    /// </summary>
    event WindowActivateEventHandler WindowActivate;

    /// <summary>
    /// 当工作表计算完成时发生
    /// </summary>
    event SheetCalculateEventHandler Calculate;

    /// <summary>
    /// 当工作表即将删除时发生
    /// </summary>
    event SheetBeforeDeleteEventHandler SheetBeforeDelete;

    /// <summary>
    /// 当工作表停用时发生
    /// </summary>
    event SheetDeactivateEventHandler SheetDeactivate;

    /// <summary>
    /// 当行集完成之前发生
    /// </summary>
    event WorkBookBeforeRowsetCompleteEventHandler OnBeforeRowsetComplete;

    /// <summary>
    /// 当数据透视表打开连接时发生
    /// </summary>
    event WorkBookPivotTableOpenConnectionEventHandler PivotTableOpenConnection;

    /// <summary>
    /// 当数据透视表关闭连接时发生
    /// </summary>
    event WorkBookPivotTableCloseConnectionEventHandler PivotTableCloseConnection;

    /// <summary>
    /// 当工作簿打开时发生
    /// </summary>
    event WorkbookOpenEventHandler Open;

    /// <summary>
    /// 当工作簿中新建工作表时发生
    /// </summary>
    event WorkBookNewSheetEventHandler NewSheet;

    /// <summary>
    /// 当工作簿停用时发生
    /// </summary>
    event DeactivateEventHandler Deactivate;

    /// <summary>
    /// 当工作簿打印之前发生
    /// </summary>
    event WorkBookBeforePrintEventHandler BeforePrint;

    /// <summary>
    /// 当工作簿中新建图表时发生
    /// </summary>
    event WorkBookNewChartEventHandler NewChart;

    /// <summary>
    /// 当工作簿关闭之前发生
    /// </summary>
    event WorkbookBeforeCloseEventHandler BeforeClose;

    /// <summary>
    /// 当工作表选择区域更改时发生
    /// </summary>
    event SheetSelectionChangeEventHandler SheetSelectionChange;

    /// <summary>
    /// 当工作簿保存之后发生
    /// </summary>
    event WorkBookAfterSaveEventHandler AfterSave;

    /// <summary>
    /// 当工作簿激活时发生
    /// </summary>
    event WorkbookActivateEventHandler WorkbookActivate;

    /// <summary>
    /// 当工作表内容更改时发生
    /// </summary>
    event WorkBookSheetChangeEventHandler SheetChange;

    /// <summary>
    /// 当工作簿同步事件发生时触发
    /// </summary>
    event WorkBookSyncEventHandler Sync;
    #endregion
}
