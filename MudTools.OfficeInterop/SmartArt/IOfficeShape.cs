//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;
/// <summary>
/// 表示 Office 中的一个形状对象（如文本框、图片等）的接口封装。
/// </summary>
public interface IOfficeShape : IDisposable
{
    /// <summary>
    /// 获取形状的唯一标识符
    /// </summary>
    int Id { get; }

    /// <summary>
    /// 获取或设置形状的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状的类型
    /// </summary>
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取形状的标题
    /// </summary>
    string Title { get; }

    /// <summary>
    /// 获取或设置形状的替代文本
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取或设置形状是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状的左侧位置（以磅为单位）
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置形状的顶部位置（以磅为单位）
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度（以磅为单位）
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置形状的高度（以磅为单位）
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取形状的 Z 顺序位置
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 删除形状
    /// </summary>
    void Delete();

    /// <summary>
    /// 将形状置于 Z 顺序的前面
    /// </summary>
    void ZOrder(MsoZOrderCmd ZOrderCmd);


    void Apply();

    /// <summary>
    /// 调整形状的大小
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    void Resize(float width, float height);

    /// <summary>
    /// 将形状复制到剪贴板
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切形状到剪贴板
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制形状并返回新形状的引用
    /// </summary>
    /// <returns>新形状</returns>
    IOfficeShape Duplicate();
}