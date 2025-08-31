//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 中内嵌形状集合的封装接口
/// </summary>
public interface IWordInlineShapes : IEnumerable<IWordInlineShape>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取内嵌形状的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的内嵌形状对象
    /// </summary>
    /// <param name="index">内嵌形状索引（从1开始）</param>
    /// <returns>内嵌形状对象</returns>
    IWordInlineShape this[int index] { get; }

    /// <summary>
    /// 添加图片内嵌形状
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <returns>新添加的内嵌形状对象</returns>
    IWordInlineShape AddPicture(string fileName, bool linkToFile = false, bool saveWithDocument = true);

    /// <summary>
    /// 添加OLE对象内嵌形状
    /// </summary>
    /// <param name="classType">OLE对象类类型</param>
    /// <param name="fileName">文件名</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="displayAsIcon">是否以图标显示</param>
    /// <param name="iconFileName">图标文件路径</param>
    /// <param name="iconIndex">图标索引</param>
    /// <param name="iconLabel">图标标签</param>
    /// <returns>新创建的内嵌形状对象</returns>
    IWordInlineShape AddOLEObject(string classType = null, string fileName = null, bool linkToFile = false,
                                 bool displayAsIcon = false, string iconFileName = null, int iconIndex = 0,
                                 string iconLabel = null);

    /// <summary>
    /// 添加水平线内嵌形状
    /// </summary>
    /// <param name="fileName">水平线文件路径</param>
    /// <param name="Range">是否链接到文件</param>
    /// <returns>新创建的内嵌形状对象</returns>
    IWordInlineShape AddHorizontalLine(string fileName = null, object Range = null);


    /// <summary>
    /// 添加图表内嵌形状
    /// </summary>
    /// <param name="style">图表样式</param>
    /// <returns>新创建的内嵌形状对象</returns>
    IWordInlineShape AddChart(MsoChartType style = MsoChartType.xlArea);

    /// <summary>
    /// 根据索引删除内嵌形状
    /// </summary>
    /// <param name="index">要删除的内嵌形状索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除所有内嵌形状
    /// </summary>
    void DeleteAll();

    /// <summary>
    /// 查找指定类型的内嵌形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <returns>符合条件的内嵌形状集合</returns>
    IEnumerable<IWordInlineShape> FindByType(int type);

    /// <summary>
    /// 获取集合的父对象（伪代码）
    /// </summary>
    object Parent { get; }
}