//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 形状集合接口
/// </summary>
public interface IPowerPointShapes : IEnumerable<IPowerPointShape>, IDisposable
{
    /// <summary>
    /// 获取形状数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 根据索引获取形状
    /// </summary>
    IPowerPointShape this[int index] { get; }

    /// <summary>
    /// 根据名称获取形状
    /// </summary>
    IPowerPointShape this[string name] { get; }


    /// <summary>
    /// 添加形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的形状</returns>
    IPowerPointShape AddShape(MsoAutoShapeType type, double left, double top, double width, double height);

    /// <summary>
    /// 添加文本框
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的文本框</returns>
    IPowerPointShape AddTextbox(MsoTextOrientation orientation, double left, double top, double width, double height);

    /// <summary>
    /// 添加图片
    /// </summary>
    /// <param name="fileName">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的图片形状</returns>
    IPowerPointShape AddPicture(string fileName, bool linkToFile, bool saveWithDocument, double left, double top, double width, double height);

    /// <summary>
    /// 添加图表
    /// </summary>
    /// <param name="type">图表类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的图表形状</returns>
    IPowerPointShape AddChart(MsoChartType type, double left, double top, double width, double height);

    /// <summary>
    /// 添加表格
    /// </summary>
    /// <param name="numRows">行数</param>
    /// <param name="numColumns">列数</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的表格形状</returns>
    IPowerPointShape AddTable(int numRows, int numColumns, double left, double top, double width, double height);

    /// <summary>
    /// 添加智能图形
    /// </summary>
    /// <param name="smartArtType">智能图形类型</param>
    /// <param name="left">左边缘位置</param>
    /// <param name="top">上边缘位置</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新添加的智能图形形状</returns>
    IPowerPointShape AddSmartArt(object smartArtType, double left, double top, double width, double height);

    IPowerPointShape AddOLEObject(
        float Left = 0f, float Top = 0f,
        float Width = -1f, float Height = -1f,
        string ClassName = "", string FileName = "", bool DisplayAsIcon = false,
        string IconFileName = "", int IconIndex = 0,
        string IconLabel = "", bool Link = false);

    /// <summary>
    /// 获取形状范围
    /// </summary>
    /// <param name="index">索引或名称数组</param>
    /// <returns>形状范围对象</returns>
    IPowerPointShapeRange Range(object index);

    /// <summary>
    /// 根据条件查找形状
    /// </summary>
    /// <param name="predicate">查找条件</param>
    /// <returns>符合条件的形状列表</returns>
    IEnumerable<IPowerPointShape> Find(Func<IPowerPointShape, bool> predicate);

    /// <summary>
    /// 按类型查找形状
    /// </summary>
    /// <param name="shapeType">形状类型</param>
    /// <returns>指定类型的形状列表</returns>
    IEnumerable<IPowerPointShape> FindByType(MsoShapeType shapeType);

    /// <summary>
    /// 按名称查找形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的形状列表</returns>
    IEnumerable<IPowerPointShape> FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 删除所有形状
    /// </summary>
    void Delete();

    /// <summary>
    /// 删除指定索引的形状
    /// </summary>
    /// <param name="index">形状索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的形状
    /// </summary>
    /// <param name="name">形状名称</param>
    void Delete(string name);

    /// <summary>
    /// 选择所有形状
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void SelectAll(bool replace = true);

    /// <summary>
    /// 取消选择所有形状
    /// </summary>
    void DeselectAll();

    /// <summary>
    /// 对齐所有形状
    /// </summary>
    /// <param name="alignCmd">对齐命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片对齐</param>
    void AlignAll(int alignCmd, bool relativeToSlide = false);

    /// <summary>
    /// 分布所有形状
    /// </summary>
    /// <param name="distributeCmd">分布命令</param>
    /// <param name="relativeToSlide">是否相对于幻灯片分布</param>
    void DistributeAll(int distributeCmd, bool relativeToSlide = false);

    /// <summary>
    /// 组合所有形状
    /// </summary>
    /// <returns>组合后的形状</returns>
    IPowerPointShape GroupAll();

    /// <summary>
    /// 获取占位符
    /// </summary>
    /// <param name="index">占位符索引</param>
    /// <returns>占位符形状</returns>
    IPowerPointShape Placeholders(int index);

    /// <summary>
    /// 获取主标题占位符
    /// </summary>
    /// <returns>主标题占位符</returns>
    IPowerPointShape Title { get; }

    /// <summary>
    /// 获取所有占位符
    /// </summary>
    IEnumerable<IPowerPointShape> GetAllPlaceholders();

    /// <summary>
    /// 按Z轴顺序排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的形状列表</returns>
    IEnumerable<IPowerPointShape> OrderByZOrder(bool ascending = true);

    /// <summary>
    /// 按名称排序
    /// </summary>
    /// <param name="ascending">是否升序排列</param>
    /// <returns>排序后的形状列表</returns>
    IEnumerable<IPowerPointShape> OrderByName(bool ascending = true);
}
