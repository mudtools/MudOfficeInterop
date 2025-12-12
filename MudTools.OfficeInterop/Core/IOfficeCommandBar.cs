//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;


public interface IOfficeCommandBar : IDisposable
{
    /// <summary>
    /// 获取命令栏的索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置命令栏名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置命令栏名称（本地化）
    /// </summary>
    string NameLocal { get; set; }

    /// <summary>
    /// 获取或设置命令栏是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置命令栏位置
    /// </summary>
    MsoBarPosition Position { get; set; }

    /// <summary>
    /// 获取或设置命令栏是否为内置命令栏
    /// </summary>
    bool BuiltIn { get; }

    /// <summary>
    /// 获取命令栏的高度
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 获取命令栏的宽度
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取命令栏的左侧位置
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取命令栏的顶部位置
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置命令栏是否启用
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取命令栏控件集合（伪代码）
    /// </summary>
    IOfficeCommandBarControls Controls { get; }

    /// <summary>
    /// 获取命令栏的ID
    /// </summary>
    int Id { get; }

    /// <summary>
    /// 删除命令栏
    /// </summary>
    void Delete();

    /// <summary>
    /// 重置命令栏到默认状态
    /// </summary>
    void Reset();

    /// <summary>
    /// 显示命令栏上下文菜单
    /// </summary>
    /// <param name="x">X坐标</param>
    /// <param name="y">Y坐标</param>
    void ShowPopup(int x = 0, int y = 0);

    /// <summary>
    /// 查找控件
    /// </summary>
    /// <param name="type">控件类型</param>
    /// <param name="id">控件ID</param>
    /// <param name="tag">标签</param>
    /// <param name="visible">是否可见</param>
    /// <returns>找到的控件对象</returns>
    IOfficeCommandBarControl? FindControl(MsoControlType type = MsoControlType.msoControlButton,
                                       object? id = null, object? tag = null, object? visible = null);
}