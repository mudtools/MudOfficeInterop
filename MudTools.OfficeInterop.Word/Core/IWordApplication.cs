//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word应用程序接口
/// </summary>
public interface IWordApplication : IOfficeApplication
{
    #region 基础属性   
    /// <summary>
    /// 获取文档数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取当前活动文档
    /// </summary>
    IWordDocument ActiveDocument { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }


    /// <summary>
    /// 获取或设置 Word 是否显示警告对话框
    /// </summary>
    WdAlertLevel DisplayAlerts { get; set; }

    /// <summary>
    /// 获取或设置屏幕更新是否启用
    /// </summary>
    bool ScreenUpdating { get; set; }


    /// <summary>
    /// 获取或设置当前用户名
    /// </summary>
    string UserName { get; set; }

    /// <summary>
    /// 获取应用程序的默认语言设置
    /// </summary>
    MsoLanguageID Language { get; }

    /// <summary>
    /// 获取应用程序支持的语言集合
    /// </summary>
    IWordLanguages Languages { get; }

    /// <summary>
    /// 获取应用程序的版本号
    /// </summary>
    string Version { get; }

    /// <summary>
    /// 获取应用程序的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取应用程序的安装路径
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取应用程序的构建号
    /// </summary>
    string Build { get; }

    /// <summary>
    /// 获取或设置应用程序是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取语言设置
    /// </summary>
    IOfficeLanguageSettings LanguageSettings { get; }

    /// <summary>
    /// 获取或设置窗口状态
    /// </summary>
    int WindowState { get; set; }

    /// <summary>
    /// 获取或设置应用程序的高度
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置应用程序的宽度
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置应用程序的左边距
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置应用程序的顶边距
    /// </summary>
    float Top { get; set; }

    #endregion

    #region 核心对象

    /// <summary>
    /// 获取活动选择区域
    /// </summary>
    IWordSelection Selection { get; }

    /// <summary>
    /// 获取文档集合
    /// </summary>
    IWordDocuments Documents { get; }

    /// <summary>
    /// 获取窗口集合
    /// </summary>
    IWordWindows Windows { get; }

    /// <summary>
    /// 获取窗口数量
    /// </summary>
    int WindowCount { get; }

    /// <summary>
    /// 获取活动窗口
    /// </summary>
    IWordWindow? ActiveWindow { get; }

    IWordTemplate? NormalTemplate { get; }
    #endregion

    #region 文档操作

    /// <summary>
    /// 根据索引获取文档（从1开始）
    /// </summary>
    IWordDocument this[int index] { get; }

    /// <summary>
    /// 根据索引获取文档并激活（从1开始）
    /// </summary>
    IWordDocument Select(int index);

    /// <summary>
    /// 根据索引获取窗口
    /// </summary>
    IWordWindow GetWindow(int index);

    /// <summary>
    /// 创建新窗口
    /// </summary>
    IWordWindow NewWindow();

    /// <summary>
    /// 创建空白文档
    /// </summary>
    IWordDocument BlankDocument();

    /// <summary>
    /// 从模板创建文档
    /// </summary>
    /// <param name="templatePath">模板路径</param>
    /// <returns>文档对象</returns>
    IWordDocument CreateFrom(string templatePath);

    /// <summary>
    /// 打开现有文档
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="readOnly">是否只读打开</param>
    /// <param name="password">密码（可选）</param>
    /// <returns>文档对象</returns>
    IWordDocument Open(string filePath, bool readOnly = false, string password = null);

    #endregion

    #region 应用程序操作
    /// <summary>
    /// 运行宏
    /// </summary>
    /// <param name="macroName">宏名称</param>
    void RunMacro(string macroName);

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
    #endregion

    #region 文件和系统操作

    /// <summary>
    /// 获取最近使用的文档列表
    /// </summary>
    /// <param name="count">返回文档数量</param>
    /// <returns>文档路径列表</returns>
    IEnumerable<string> GetRecentFiles(int count = 10);

    /// <summary>
    /// 添加最近使用的文档
    /// </summary>
    /// <param name="filePath">文件路径</param>
    void AddToRecentFiles(string filePath);

    /// <summary>
    /// 设置 Word 选项
    /// </summary>
    /// <param name="optionName">选项名称</param>
    /// <param name="value">选项值</param>
    void SetOption(string optionName, object value);

    /// <summary>
    /// 获取 Word 选项值
    /// </summary>
    /// <param name="optionName">选项名称</param>
    /// <returns>选项值</returns>
    object GetOption(string optionName);

    /// <summary>
    /// 获取 Word 安装路径
    /// </summary>
    /// <returns>安装路径</returns>
    string GetInstallPath();

    /// <summary>
    /// 获取 Word 产品名称
    /// </summary>
    /// <returns>产品名称</returns>
    string GetProductName();

    /// <summary>
    /// 检查 Word 是否正在打印
    /// </summary>
    /// <returns>是否正在打印</returns>
    bool IsPrinting();

    /// <summary>
    /// 取消所有打印作业
    /// </summary>
    void CancelPrintJobs();

    /// <summary>
    /// 获取系统信息
    /// </summary>
    /// <returns>系统信息</returns>
    IWordSystemInfo GetSystemInfo();

    /// <summary>
    /// 刷新应用程序显示
    /// </summary>
    void Refresh();

    #endregion

    #region 事件

    /// <summary>
    /// 当文档打开时触发
    /// </summary>
    event DocumentOpenEventHandler DocumentOpen;

    /// <summary>
    /// 当文档关闭前触发
    /// </summary>
    event DocumentBeforeCloseEventHandler DocumentBeforeClose;

    /// <summary>
    /// 当文档保存前触发
    /// </summary>
    event DocumentBeforeSaveEventHandler DocumentBeforeSave;

    /// <summary>
    /// 当窗口激活时触发
    /// </summary>
    event WindowActivateEventHandler WindowActivate;

    /// <summary>
    /// 当窗口失活时触发
    /// </summary>
    event WindowDeactivateEventHandler WindowDeactivate;

    /// <summary>
    /// 当文档同步时触发
    /// </summary>
    event DocumentSyncEventHandler DocumentSync;

    /// <summary>
    /// 当文档变化时触发
    /// </summary>
    event DocumentChangeEventHandler DocumentChange;

    /// <summary>
    /// 当邮件合并数据源打开时触发
    /// </summary>
    event MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoad;

    /// <summary>
    /// 当邮件合并数据源验证时触发
    /// </summary>
    event MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidate;


    /// <summary>
    /// 当窗口选择改变时触发
    /// </summary>
    event WindowSelectionChangeEventHandler WindowSelectionChange;

    /// <summary>
    /// 当窗口大小改变时触发
    /// </summary>
    event WindowSizeEventHandler WindowSize;

    #endregion
}
