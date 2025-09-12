//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 幻灯片接口
/// </summary>
public interface IPowerPointSlide : IDisposable
{
    /// <summary>
    /// 获取幻灯片名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取幻灯片索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取幻灯片布局
    /// </summary>
    PpSlideLayout Layout { get; set; }

    /// <summary>
    /// 获取幻灯片标题
    /// </summary>
    string Title { get; }

    /// <summary>
    /// 获取父对象（通常是 Slides 集合）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取幻灯片的形状集合
    /// </summary>
    IPowerPointShapes? Shapes { get; }

    /// <summary>
    /// 获取幻灯片的页眉页脚
    /// </summary>
    IPowerPointHeadersFooters? HeadersFooters { get; }

    /// <summary>
    /// 获取幻灯片的背景
    /// </summary>
    IPowerPointBackground? Background { get; }

    /// <summary>
    /// 获取幻灯片的母版
    /// </summary>
    IPowerPointMaster? Master { get; }

    /// <summary>
    /// 获取幻灯片的幻灯片显示
    /// </summary>
    IPowerPointSlideShowTransition? SlideShowTransition { get; }

    /// <summary>
    /// 获取幻灯片的动画设置
    /// </summary>
    IPowerPointTimeLine? TimeLine { get; }

    /// <summary>
    /// 获取幻灯片的超链接集合
    /// </summary>
    IEnumerable<IPowerPointHyperlink> Hyperlinks { get; }

    /// <summary>
    /// 获取幻灯片的标签集合
    /// </summary>
    IPowerPointTags Tags { get; }

    /// <summary>
    /// 获取幻灯片的自定义布局
    /// </summary>
    IPowerPointCustomLayout CustomLayout { get; set; }

    /// <summary>
    /// 获取幻灯片的幻灯片ID
    /// </summary>
    int SlideID { get; }

    /// <summary>
    /// 获取幻灯片编号
    /// </summary>
    int SlideNumber { get; }

    /// <summary>
    /// 激活幻灯片
    /// </summary>
    void Select();

    /// <summary>
    /// 复制幻灯片
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切幻灯片
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    void Delete();

    /// <summary>
    /// 移动幻灯片到指定位置
    /// </summary>
    /// <param name="toPosition">目标位置</param>
    void MoveTo(int toPosition);

    /// <summary>
    /// 应用设计模板
    /// </summary>
    /// <param name="designName">设计模板名称</param>
    void ApplyDesign(string designName);

    /// <summary>
    /// 应用主题
    /// </summary>
    /// <param name="themeName">主题名称</param>
    void ApplyTheme(string themeName);

    /// <summary>
    /// 导出幻灯片为图片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="filterName">图片格式</param>
    /// <param name="scaleWidth">宽度缩放</param>
    /// <param name="scaleHeight">高度缩放</param>
    void Export(string fileName, string filterName = "PNG", int scaleWidth = 0, int scaleHeight = 0);

    /// <summary>
    /// 获取幻灯片的缩略图
    /// </summary>
    /// <returns>缩略图数据</returns>
    byte[] GetThumbnail();

    /// <summary>
    /// 获取幻灯片的所有文本内容
    /// </summary>
    /// <returns>文本内容列表</returns>
    IEnumerable<string> GetAllText();


    /// <summary>
    /// 获取指定占位符
    /// </summary>
    /// <param name="placeholderIndex">占位符索引</param>
    /// <returns>形状对象</returns>
    IPowerPointShape GetPlaceholder(int placeholderIndex);

    /// <summary>
    /// 获取所有占位符
    /// </summary>
    /// <returns>形状对象列表</returns>
    IEnumerable<IPowerPointShape> GetPlaceholders();

    /// <summary>
    /// 刷新幻灯片显示
    /// </summary>
    void Refresh();
}
