//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 包含有关计算机系统的信息。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSystem : IOfficeObject<IWordSystem, MsWord.System>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前操作系统的名称（例如："Windows" 或 "Windows NT"）。
    /// </summary>
    string OperatingSystem { get; }

    /// <summary>
    /// 获取系统使用的处理器类型（例如：i486）。
    /// </summary>
    /// <remarks>
    /// 此属性已标记为过时，建议使用其他方法获取处理器信息。
    /// </remarks>
    string ProcessorType { get; }

    /// <summary>
    /// 获取操作系统的版本号。
    /// </summary>
    string Version { get; }

    /// <summary>
    /// 获取当前驱动器的可用磁盘空间（以字节为单位）。
    /// </summary>
    int FreeDiskSpace { get; }

    /// <summary>
    /// 获取系统软件指定的语言。
    /// </summary>
    string LanguageDesignation { get; }

    /// <summary>
    /// 获取水平显示分辨率（以像素为单位）。
    /// </summary>
    int HorizontalResolution { get; }

    /// <summary>
    /// 获取垂直屏幕分辨率（以像素为单位）。
    /// </summary>
    int VerticalResolution { get; }

    /// <summary>
    /// 获取或设置Windows注册表中指定子键下的条目值。
    /// 注册表路径：HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word。
    /// </summary>
    /// <param name="section">注册表子键名称，位于"HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word"子键下。</param>
    /// <param name="key">指定子键中的条目名称（例如，Options子键中的"BackgroundPrint"）。</param>
    /// <returns>注册表条目的字符串值。</returns>
    [MethodIndex]
    string? ProfileString(string section, string key);

    /// <summary>
    /// 获取或设置配置文件或Windows注册表中的字符串值。
    /// </summary>
    /// <param name="fileName">配置文件名。如果未指定路径，则假定为Windows文件夹。</param>
    /// <param name="section">
    /// 包含键的配置文件的节名。
    /// 在Windows配置文件中，节名出现在相关键之前的方括号中（不要在section参数中包含方括号）。
    /// 如果要从Windows注册表返回条目值，section应为子键的完整路径，包括子树。
    /// </param>
    /// <param name="key">
    /// 要检索的键设置或注册表条目值。
    /// 在Windows配置文件中，键名后跟等号(=)和设置。
    /// 如果要从Windows注册表返回条目值，key应为section指定的子键中的条目名称。
    /// </param>
    /// <returns>配置文件或注册表条目的字符串值。</returns>
    [MethodIndex]
    string? PrivateProfileString(string fileName, string section, string key);

    /// <summary>
    /// 获取一个值，指示系统是否安装了数学协处理器。
    /// </summary>
    bool MathCoprocessorInstalled { get; }

    /// <summary>
    /// 获取或设置指针的状态（形状）。
    /// 可以是以下WdCursorType常量之一：
    /// wdCursorIBeam、wdCursorNormal、wdCursorNorthwestArrow 或 wdCursorWait。
    /// </summary>
    WdCursorType Cursor { get; set; }

    /// <summary>
    /// 如果Microsoft系统信息应用程序未运行则启动它，如果已运行则切换到它。
    /// </summary>
    void MSInfo();

    /// <summary>
    /// 建立到网络驱动器的连接。
    /// </summary>
    /// <param name="path">网络驱动器的路径（例如："\\Project\Info"）。</param>
    /// <param name="drive">
    /// 对应于要分配给网络驱动器的字母的数字。
    /// 0（零）对应于第一个可用的驱动器字母，1对应于第二个可用的驱动器字母，依此类推。
    /// 如果省略此参数，则使用下一个可用的字母。
    /// </param>
    /// <param name="password">如果网络驱动器受密码保护，则提供密码。</param>
    void Connect(string path, int? drive = null, string? password = null);
}
