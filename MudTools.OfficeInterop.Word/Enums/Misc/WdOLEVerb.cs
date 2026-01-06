//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 OLE 对象应执行的动词相关操作
/// </summary>
public enum WdOLEVerb
{
    /// <summary>
    /// 执行用户双击对象时调用的主要动词操作
    /// </summary>
    wdOLEVerbPrimary = 0,

    /// <summary>
    /// 向用户显示对象以供编辑或查看。用于显示新插入的对象以进行初始编辑
    /// </summary>
    wdOLEVerbShow = -1,

    /// <summary>
    /// 在单独的窗口中打开对象
    /// </summary>
    wdOLEVerbOpen = -2,

    /// <summary>
    /// 从视图中移除对象的用户界面
    /// </summary>
    wdOLEVerbHide = -3,

    /// <summary>
    /// 就地激活对象并显示对象所需的任何用户界面工具（例如菜单或工具栏）
    /// </summary>
    wdOLEVerbUIActivate = -4,

    /// <summary>
    /// 运行对象并安装其窗口，但不安装任何用户界面工具
    /// </summary>
    wdOLEVerbInPlaceActivate = -5,

    /// <summary>
    /// 强制对象丢弃可能维护的任何撤销状态；但对象仍保持激活状态
    /// </summary>
    wdOLEVerbDiscardUndoState = -6
}