//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Pictures 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Pictures 的安全访问和操作
/// </summary>
public interface IExcelPictures : IEnumerable<IExcelPicture>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取图片集合中的图片数量
    /// 对应 Pictures.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的图片对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">图片索引（从1开始）</param>
    /// <returns>图片对象</returns>
    IExcelPicture? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的图片对象
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>图片对象</returns>
    IExcelPicture? this[string name] { get; }

    /// <summary>
    /// 获取图片集合所在的父对象（通常是工作表）
    /// 对应 Pictures.Parent 属性
    /// </summary>
    object Parent { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向工作表添加图片
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="converter"></param>
    /// <returns>新创建的图片对象</returns>
    IExcelPicture? Add(string filename, object converter);

    /// <summary>
    /// 批量添加图片
    /// </summary>
    /// <param name="picturesData">图片数据数组</param>
    /// <returns>成功添加的图片数量</returns>
    int AddRange(PictureData[] picturesData);

    /// <summary>
    /// 从字节数组添加图片
    /// </summary>
    /// <param name="imageBytes">图片字节数组</param>
    /// <param name="imageFormat">图片格式</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图片对象</returns>
    IExcelPicture AddFromBytes(byte[] imageBytes, string imageFormat = "png",
                              double left = 0, double top = 0, double width = -1, double height = -1);

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图片
    /// </summary>
    /// <param name="name">图片名称</param>
    /// <returns>匹配的图片数组</returns>
    IExcelPicture[] FindByName(string name);

    /// <summary>
    /// 根据位置查找图片
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图片数组</returns>
    IExcelPicture[] FindByPosition(double left, double top, double tolerance = 10);

    /// <summary>
    /// 根据大小查找图片
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图片数组</returns>
    IExcelPicture[] FindBySize(double width, double height, double tolerance = 10);

    /// <summary>
    /// 获取指定区域内的所有图片
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的图片数组</returns>
    IExcelPicture[] GetPicturesInRange(IExcelRange range);

    /// <summary>
    /// 获取可见的图片
    /// </summary>
    /// <returns>可见图片数组</returns>
    IExcelPicture[] GetVisiblePictures();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图片
    /// 对应 Pictures.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的图片
    /// </summary>
    /// <param name="index">要删除的图片索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的图片
    /// </summary>
    /// <param name="picture">要删除的图片对象</param>
    void Delete(IExcelPicture picture);

    /// <summary>
    /// 批量删除图片
    /// </summary>
    /// <param name="indices">要删除的图片索引数组</param>
    void DeleteRange(int[] indices);

    #endregion

    #region 排列和布局

    /// <summary>
    /// 对齐选中的图片
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    void Align(int alignment);

    /// <summary>
    /// 分布选中的图片
    /// </summary>
    /// <param name="distribution">分布方式</param>
    void Distribute(int distribution);

    /// <summary>
    /// 统一选中图片的大小
    /// </summary>
    /// <param name="useWidth">是否使用宽度作为标准</param>
    void SizeToSame(bool useWidth = true);

    /// <summary>
    /// 按指定行列排列图片
    /// </summary>
    /// <param name="rows">行数</param>
    /// <param name="columns">列数</param>
    /// <param name="horizontalSpacing">水平间距</param>
    /// <param name="verticalSpacing">垂直间距</param>
    void ArrangeInGrid(int rows, int columns, double horizontalSpacing = 10, double verticalSpacing = 10);

    #endregion

    #region 导出和导入

    /// <summary>
    /// 获取所有图片的信息
    /// </summary>
    /// <returns>图片信息数组</returns>
    PictureInfo[] GetAllPictureInfo();

    #endregion

    #region 统计和分析

    /// <summary>
    /// 获取图片统计信息
    /// </summary>
    /// <returns>图片统计信息对象</returns>
    PictureStatistics GetStatistics();


    /// <summary>
    /// 获取图片大小分布
    /// </summary>
    /// <returns>大小分布信息</returns>
    PictureSizeDistribution GetSizeDistribution();

    #endregion
}

/// <summary>
/// 图片数据结构
/// </summary>
public class PictureData
{
    /// <summary>
    /// 图片文件路径
    /// </summary>
    public string Filename { get; set; }

    /// <summary>
    /// 左边距
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// 顶边距
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// 宽度
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// 是否链接到文件
    /// </summary>
    public bool LinkToFile { get; set; }

    /// <summary>
    /// 是否与文档一起保存
    /// </summary>
    public bool SaveWithDocument { get; set; }
}

/// <summary>
/// 图片信息结构
/// </summary>
public class PictureInfo
{
    /// <summary>
    /// 图片索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 图片名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 左边距
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// 顶边距
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// 宽度
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// 高度
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// 是否可见
    /// </summary>
    public bool Visible { get; set; }

    /// <summary>
    /// 图片格式
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// 文件路径
    /// </summary>
    public string Filename { get; set; }

    /// <summary>
    /// 是否链接
    /// </summary>
    public bool IsLinked { get; set; }

    /// <summary>
    /// 原始宽度
    /// </summary>
    public double OriginalWidth { get; set; }

    /// <summary>
    /// 原始高度
    /// </summary>
    public double OriginalHeight { get; set; }
}

/// <summary>
/// 图片统计信息结构
/// </summary>
public class PictureStatistics
{
    /// <summary>
    /// 总图片数
    /// </summary>
    public int TotalCount { get; set; }

    /// <summary>
    /// 可见图片数
    /// </summary>
    public int VisibleCount { get; set; }

    /// <summary>
    /// 隐藏图片数
    /// </summary>
    public int HiddenCount { get; set; }

    /// <summary>
    /// 链接图片数
    /// </summary>
    public int LinkedCount { get; set; }

    /// <summary>
    /// 嵌入图片数
    /// </summary>
    public int EmbeddedCount { get; set; }

    /// <summary>
    /// 平均宽度
    /// </summary>
    public double AverageWidth { get; set; }

    /// <summary>
    /// 平均高度
    /// </summary>
    public double AverageHeight { get; set; }

    /// <summary>
    /// 最大宽度
    /// </summary>
    public double MaxWidth { get; set; }

    /// <summary>
    /// 最大高度
    /// </summary>
    public double MaxHeight { get; set; }
}


/// <summary>
/// 大小分布信息结构
/// </summary>
public class PictureSizeDistribution
{
    /// <summary>
    /// 小图片数量（小于100x100像素）
    /// </summary>
    public int SmallPictures { get; set; }

    /// <summary>
    /// 中等图片数量（100x100到500x500像素）
    /// </summary>
    public int MediumPictures { get; set; }

    /// <summary>
    /// 大图片数量（500x500到1000x1000像素）
    /// </summary>
    public int LargePictures { get; set; }

    /// <summary>
    /// 超大图片数量（大于1000x1000像素）
    /// </summary>
    public int ExtraLargePictures { get; set; }
}