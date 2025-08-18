//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ShapeRange 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ShapeRange 的安全访问和操作
/// ShapeRange 代表一个或多个形状对象的集合
/// </summary>
public interface IExcelShapeRange : IEnumerable<IExcelShape>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取形状区域中的形状数量
    /// 对应 ShapeRange.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的形状对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    IExcelShape? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    IExcelShape? this[string name] { get; }

    /// <summary>
    /// 获取形状区域的名称
    /// 对应 ShapeRange.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状区域所在的父对象（通常是工作表）
    /// 对应 ShapeRange.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取形状区域的ID
    /// 对应 ShapeRange.ID 属性
    /// </summary>
    int ID { get; }

    #endregion

    #region 位置和大小属性

    /// <summary>
    /// 获取形状区域的左边距
    /// 对应 ShapeRange.Left 属性
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状区域的顶边距
    /// 对应 ShapeRange.Top 属性
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状区域的宽度
    /// 对应 ShapeRange.Width 属性
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状区域的高度
    /// 对应 ShapeRange.Height 属性
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取形状区域的旋转角度
    /// 对应 ShapeRange.Rotation 属性
    /// </summary>
    float Rotation { get; set; }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状区域是否可见
    /// 对应 ShapeRange.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    #endregion

    #region 格式设置

    /// <summary>
    /// 获取形状区域的填充格式对象
    /// 对应 ShapeRange.Fill 属性
    /// </summary>
    IExcelFillFormat Fill { get; }

    /// <summary>
    /// 获取形状区域的线条格式对象
    /// 对应 ShapeRange.Line 属性
    /// </summary>
    IExcelLineFormat Line { get; }

    /// <summary>
    /// 获取形状区域的文本框架对象
    /// 对应 ShapeRange.TextFrame 属性
    /// </summary>
    IExcelTextFrame TextFrame { get; }

    /// <summary>
    /// 获取形状区域的阴影格式对象
    /// 对应 ShapeRange.Shadow 属性
    /// </summary>
    IExcelShadowFormat Shadow { get; }

    /// <summary>
    /// 获取形状区域的三维格式对象
    /// 对应 ShapeRange.ThreeD 属性
    /// </summary>
    IExcelThreeDFormat ThreeD { get; }

    #endregion

    #region 文本属性

    /// <summary>
    /// 获取或设置形状区域中的文本内容
    /// 对应 ShapeRange.TextFrame.Characters.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置形状区域中文本的自动调整大小
    /// 对应 ShapeRange.TextFrame.AutoSize 属性
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取或设置形状区域中文本的水平对齐方式
    /// 对应 ShapeRange.TextFrame.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置形状区域中文本的垂直对齐方式
    /// 对应 ShapeRange.TextFrame.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向形状区域添加新的形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的形状对象</returns>
    IExcelShape? AddShape(MsoAutoShapeType type, double left, double top, double width, double height);

    /// <summary>
    /// 向形状区域添加文本框
    /// </summary>
    /// <param name="orientation">文本方向</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的文本框对象</returns>
    IExcelShape? AddTextbox(MsoTextOrientation orientation, double left, double top, double width, double height);

    /// <summary>
    /// 向形状区域添加线条
    /// </summary>
    /// <param name="x1">起点X坐标</param>
    /// <param name="y1">起点Y坐标</param>
    /// <param name="x2">终点X坐标</param>
    /// <param name="y2">终点Y坐标</param>
    /// <returns>新创建的线条对象</returns>
    IExcelShape AddLine(double x1, double y1, double x2, double y2);

    /// <summary>
    /// 向形状区域添加图片
    /// </summary>
    /// <param name="filename">图片文件路径</param>
    /// <param name="linkToFile">是否链接到文件</param>
    /// <param name="saveWithDocument">是否与文档一起保存</param>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图片对象</returns>
    IExcelShape AddPicture(string filename, bool linkToFile, bool saveWithDocument,
                          double left, double top, double width, double height);

    #endregion

    #region 选择和操作

    /// <summary>
    /// 选择形状区域
    /// 对应 ShapeRange.Select 方法
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除形状区域中的所有形状
    /// 对应 ShapeRange.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 应用自动调整选项
    /// 对应 ShapeRange.Apply 方法
    /// </summary>
    void Apply();

    /// <summary>
    /// 复制形状区域的格式
    /// 对应 ShapeRange.PickUp 方法
    /// </summary>
    void PickUp();

    #endregion

    #region 变换操作

    /// <summary>
    /// 调整形状区域大小
    /// </summary>
    /// <param name="widthScale">宽度缩放比例</param>
    /// <param name="heightScale">高度缩放比例</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小</param>
    void Scale(double widthScale, double heightScale, bool relativeToOriginalSize = false);

    /// <summary>
    /// 移动形状区域
    /// </summary>
    /// <param name="leftIncrement">左边距增量</param>
    /// <param name="topIncrement">顶边距增量</param>
    void Move(double leftIncrement, double topIncrement);

    /// <summary>
    /// 旋转形状区域
    /// </summary>
    /// <param name="rotationIncrement">旋转角度增量（度）</param>
    void Rotate(double rotationIncrement);

    #endregion

    #region 排列和布局

    /// <summary>
    /// 将形状区域置于最前面
    /// 对应 ShapeRange.ZOrder 方法
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将形状区域置于最后面
    /// 对应 ShapeRange.ZOrder 方法
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 将形状区域向前移动一层
    /// </summary>
    void BringForward();

    /// <summary>
    /// 将形状区域向后移动一层
    /// </summary>
    void SendBackward();

    /// <summary>
    /// 对齐形状区域中的形状
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="relativeTo">相对对象</param>
    void Align(MsoAlignCmd alignment, bool relativeTo = false);

    /// <summary>
    /// 分布形状区域中的形状
    /// </summary>
    /// <param name="distribution">分布方式</param>
    void Distribute(MsoDistributeCmd distribution);

    /// <summary>
    /// 统一形状区域中形状的大小
    /// </summary>
    /// <param name="useWidth">是否使用宽度作为标准</param>
    void SizeToSame(bool useWidth = true);

    #endregion

    #region 组合操作

    /// <summary>
    /// 组合形状区域中的所有形状
    /// 对应 ShapeRange.Group 方法
    /// </summary>
    /// <returns>组合后的形状对象</returns>
    IExcelShape Group();

    /// <summary>
    /// 取消组合形状区域中的组合形状
    /// 对应 ShapeRange.Ungroup 方法
    /// </summary>
    /// <returns>取消组合后的形状区域</returns>
    IExcelShapeRange Ungroup();

    /// <summary>
    /// 获取形状区域中的所有子形状
    /// </summary>
    /// <returns>子形状数组</returns>
    IExcelShape[] GetChildShapes();

    /// <summary>
    /// 获取形状区域中的顶级形状
    /// </summary>
    /// <returns>顶级形状数组</returns>
    IExcelShape[] GetTopLevelShapes();

    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据类型筛选形状
    /// </summary>
    /// <param name="type">形状类型</param>
    /// <returns>匹配的形状数组</returns>
    IExcelShape[] FilterByType(MsoShapeType type);

    /// <summary>
    /// 根据名称筛选形状
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的形状数组</returns>
    IExcelShape[] FilterByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据位置筛选形状
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的形状数组</returns>
    IExcelShape[] FilterByPosition(double left, double top, double tolerance = 10);

    /// <summary>
    /// 根据大小筛选形状
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的形状数组</returns>
    IExcelShape[] FilterBySize(double width, double height, double tolerance = 10);

    /// <summary>
    /// 获取可见的形状
    /// </summary>
    /// <returns>可见形状数组</returns>
    IExcelShape[] GetVisibleShapes();

    /// <summary>
    /// 获取隐藏的形状
    /// </summary>
    /// <returns>隐藏形状数组</returns>
    IExcelShape[] GetHiddenShapes();

    #endregion

    #region 层次结构
    /// <summary>
    /// 获取形状区域所在的图表对象（如果是图表）
    /// 对应 ShapeRange.Chart 属性
    /// </summary>
    IExcelChart Chart { get; }
    #endregion   
}