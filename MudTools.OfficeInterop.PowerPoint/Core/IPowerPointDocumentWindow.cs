//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint DocumentWindow 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.PowerPoint.DocumentWindow 的安全访问和操作
/// </summary>
public interface IPowerPointDocumentWindow : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取PowerPoint应用程序窗口的句柄
    /// </summary>
    int? Hwnd { get; }
    /// <summary>
    /// 获取 Window 对象的父对象（通常是 Application）
    /// 对应 DocumentWindow.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 Window 对象所在的 Application 对象
    /// 对应 DocumentWindow.Application 属性
    /// </summary>
    IPowerPointApplication Application { get; }

    /// <summary>
    /// 获取窗口关联的演示文稿
    /// 对应 DocumentWindow.Presentation 属性
    /// </summary>
    IPowerPointPresentation Presentation { get; }

    /// <summary>
    /// 获取或设置窗口的标题栏文本
    /// 对应 DocumentWindow.Caption 属性
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取窗口是否是活动窗口
    /// 对应 DocumentWindow.Active 属性
    /// </summary>
    bool IsActive { get; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置窗口的左边距
    /// 对应 DocumentWindow.Left 属性
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置窗口的顶边距
    /// 对应 DocumentWindow.Top 属性
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置窗口的宽度
    /// 对应 DocumentWindow.Width 属性
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置窗口的高度
    /// 对应 DocumentWindow.Height 属性
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置窗口状态（正常、最小化、最大化）
    /// 对应 DocumentWindow.WindowState 属性
    /// </summary>
    PpWindowState WindowState { get; set; } // 使用 int 代表 Pp.PpWindowState
    #endregion

    #region 核心对象
    /// <summary>
    /// 获取窗口当前的选择对象
    /// 对应 DocumentWindow.Selection 属性
    /// </summary>
    IPowerPointSelection Selection { get; }

    /// <summary>
    /// 获取窗口当前的视图对象
    /// 对应 DocumentWindow.View 属性
    /// </summary>
    IPowerPointView View { get; }

    /// <summary>
    /// 获取窗口当前视图中活动的幻灯片
    /// </summary>
    IPowerPointSlide ActiveSlide { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 激活此窗口，使其成为活动窗口
    /// 对应 DocumentWindow.Activate 方法
    /// </summary>
    void Activate();

    /// <summary>
    /// 关闭窗口
    /// 对应 DocumentWindow.Close 方法
    /// </summary>
    void Close();
    #endregion

    #region 视图操作
    /// <summary>
    /// 切换到普通视图
    /// </summary>
    void ViewNormal();

    /// <summary>
    /// 切换到幻灯片浏览视图
    /// </summary>
    void ViewSlideSorter();

    /// <summary>
    /// 切换到幻灯片放映视图
    /// </summary>
    void ViewSlideShow();

    /// <summary>
    /// 切换到备注页视图
    /// </summary>
    void ViewNotesPage();
    #endregion

}
