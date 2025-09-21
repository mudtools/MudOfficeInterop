//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示一组 Office 形状对象的集合接口。
/// </summary>
public interface IOfficeShapes : IEnumerable<IOfficeShape>, IDisposable
{
    /// <summary>
    /// 获取集合中的形状数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取形状
    /// </summary>
    /// <param name="index">形状的索引（从1开始）</param>
    /// <returns>指定索引处的形状</returns>
    IOfficeShape? this[int index] { get; }

    /// <summary>
    /// 通过名称获取形状
    /// </summary>
    /// <param name="name">形状的名称</param>
    /// <returns>具有指定名称的形状</returns>
    IOfficeShape? this[string name] { get; }

    /// <summary>
    /// 向集合中添加新形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的形状</returns>
    IOfficeShape? AddShape(MsoAutoShapeType type, float left, float top, float width, float height);

    /// <summary>
    /// 向集合中添加文本框
    /// </summary>
    /// <param name="left">左侧位置</param>
    /// <param name="top">顶部位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="orientation"></param>
    /// <returns>新添加的文本框形状</returns>
    IOfficeShape? AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 从集合中删除所有形状
    /// </summary>
    void DeleteAll();

    /// <summary>
    /// 根据名称选择形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>选定的形状</returns>
    IOfficeShape? SelectByName(string name);

}