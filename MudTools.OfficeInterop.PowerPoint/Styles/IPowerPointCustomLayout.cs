//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 自定义布局接口
/// </summary>
public interface IPowerPointCustomLayout : IDisposable
{
    /// <summary>
    /// 获取或设置布局名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状集合
    /// </summary>
    IPowerPointShapes Shapes { get; }

    /// <summary>
    /// 获取页眉页脚
    /// </summary>
    IPowerPointHeadersFooters HeadersFooters { get; }

    /// <summary>
    /// 获取背景
    /// </summary>
    IPowerPointBackground Background { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置是否包含主题
    /// </summary>
    bool FollowMasterBackground { get; set; }

    /// <summary>
    /// 获取布局索引
    /// </summary>
    int Index { get; }


    /// <summary>
    /// 应用到幻灯片
    /// </summary>
    /// <param name="slide">目标幻灯片</param>
    void ApplyTo(IPowerPointSlide slide);

    /// <summary>
    /// 复制布局
    /// </summary>
    /// <returns>复制的布局</returns>
    IPowerPointCustomLayout Duplicate();

    /// <summary>
    /// 删除布局
    /// </summary>
    void Delete();

    /// <summary>
    /// 重置布局
    /// </summary>
    void Reset();

    /// <summary>
    /// 刷新布局显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 获取布局的缩略图
    /// </summary>
    /// <returns>缩略图数据</returns>
    byte[] GetThumbnail();

    /// <summary>
    /// 导出布局为图片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="filterName">图片格式</param>
    void Export(string fileName, string filterName = "PNG");

    /// <summary>
    /// 获取布局信息
    /// </summary>
    /// <returns>布局信息字符串</returns>
    string GetLayoutInfo();
}