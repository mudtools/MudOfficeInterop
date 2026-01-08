//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// Office应用程序接口，定义了所有Office应用程序的公共功能
/// </summary>
public interface IOfficeApplication : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取应用程序的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取应用程序的版本号
    /// </summary>
    string Version { get; }

    /// <summary>
    /// 获取应用程序的完整路径
    /// </summary>
    string Path { get; }

    /// <summary>
    /// 获取或设置是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取应用程序的构建版本信息
    /// </summary>
    string? Build { get; }
    #endregion

    #region 窗口属性

    /// <summary>
    /// 获取或设置窗口状态（正常、最小化、最大化）
    /// </summary>
    [IgnoreGenerator]
    int WindowStateValue { get; set; }

    /// <summary>
    /// 获取或设置应用程序窗口的左边距
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取或设置应用程序窗口的顶边距
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置应用程序窗口的宽度
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取或设置应用程序窗口的高度
    /// </summary>
    int Height { get; set; }

    #endregion

    #region 核心对象

    /// <summary>
    /// 获取语言设置对象
    /// </summary>
    IOfficeLanguageSettings? LanguageSettings { get; }

    /// <summary>
    /// 获取应用程序的命令栏集合
    /// </summary>
    IOfficeCommandBars? CommandBars { get; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 使应用程序窗口获得焦点
    /// </summary>
    void Activate();

    /// <summary>
    /// 退出应用程序
    /// </summary>
    void Quit();

    #endregion

    #region 宏和自动化

    /// <summary>
    /// 运行指定的宏
    /// </summary>
    /// <param name="macroName">宏名称</param>
    /// <param name="args">宏参数</param>
    /// <returns>宏的返回值</returns>
    object Run(string macroName, params object[] args);

    #endregion

    #region UI 和交互      

    /// <summary>
    /// 显示一个文件对话框
    /// </summary>
    /// <param name="fileDialogType">对话框类型</param>
    /// <returns>文件对话框对象</returns>
    [IgnoreGenerator]
    IOfficeFileDialog CreateFileDialog(MsoFileDialogType fileDialogType);

    #endregion
}