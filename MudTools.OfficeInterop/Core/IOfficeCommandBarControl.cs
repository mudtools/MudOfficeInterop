//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
public interface IOfficeCommandBarControl : IDisposable
{
    /// <summary>
    /// 获取控件的索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置控件ID
    /// </summary>
    int Id { get; }

    object Control { get; }

    /// <summary>
    /// 获取或设置控件类型
    /// </summary>
    MsoControlType Type { get; }

    /// <summary>
    /// 获取或设置控件标签
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置控件是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置控件是否启用
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取或设置控件的标签
    /// </summary>
    string Tag { get; set; }

    /// <summary>
    /// 获取或设置控件的提示文本
    /// </summary>
    string TooltipText { get; set; }

    /// <summary>
    /// 获取或设置控件的帮助文件路径
    /// </summary>
    string HelpFile { get; set; }

    /// <summary>
    /// 获取或设置控件的帮助上下文ID
    /// </summary>
    int HelpContextId { get; set; }

    /// <summary>
    /// 获取控件的左侧位置
    /// </summary>
    int Left { get; }

    /// <summary>
    /// 获取控件的顶部位置
    /// </summary>
    int Top { get; }

    /// <summary>
    /// 获取控件的宽度
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取控件的高度
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 获取或设置控件的参数
    /// </summary>
    string Parameter { get; set; }

    /// <summary>
    /// 获取父级控件集合
    /// </summary>
    IOfficeCommandBar Parent { get; }

    /// <summary>
    /// 激活控件
    /// </summary>
    void Execute();

    /// <summary>
    /// 删除控件
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制控件
    /// </summary>
    /// <param name="bar">目标命令栏</param>
    /// <param name="before">插入位置</param>
    /// <returns>复制的控件</returns>
    IOfficeCommandBarControl Copy(IOfficeCommandBar bar = null, IOfficeCommandBar before = null);

    /// <summary>
    /// 移动控件
    /// </summary>
    /// <param name="bar">目标命令栏</param>
    /// <param name="before">插入位置</param>
    void Move(IOfficeCommandBar bar = null, IOfficeCommandBar before = null);
}