//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中形状集合的封装接口
/// </summary>
public interface IWordShapes : IEnumerable<IWordShape>, IDisposable
{
    /// <summary>
    /// 获取形状的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的形状对象
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    IWordShape this[int index] { get; }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    IWordShape this[string name] { get; }

    IWordShape AddOLEObject(ref object ClassType,
        ref object FileName,
        ref object LinkToFile,
        ref object DisplayAsIcon,
        ref object IconFileName,
        ref object IconIndex,
        ref object IconLabel,
        ref object Left,
        ref object Top,
        ref object Width,
        ref object Height,
        ref object Anchor);

    /// <summary>
    /// 添加文本框形状
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">上边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IWordShape AddTextbox(MsoTextOrientation orientation, float left, float top, float width, float height);

    /// <summary>
    /// 添加图片形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">上边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IWordShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument,
                         double left, double top, double width, double height);

    /// <summary>
    /// 根据索引删除形状
    /// </summary>
    /// <param name="index">要删除的形状索引</param>
    void Delete(int index);

    /// <summary>
    /// 根据名称删除形状
    /// </summary>
    /// <param name="name">要删除的形状名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除所有形状
    /// </summary>
    void DeleteAll();

    /// <summary>
    /// 查找指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象，如果未找到则返回null</returns>
    IWordShape FindByName(string name);
}