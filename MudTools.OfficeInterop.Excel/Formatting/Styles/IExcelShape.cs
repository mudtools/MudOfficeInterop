//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Shape 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Shape 的安全访问和操作
/// </summary>
public interface IExcelShape : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取形状的 OLE 格式设置属性
    /// 对应 Shape.OLEFormat 属性，提供对嵌入的 OLE 对象格式设置的访问
    /// </summary>
    IExcelOLEFormat OLEFormat { get; }

    /// <summary>
    /// 获取组合形状中单个子形状的集合
    /// 对应 Shape.GroupItems 属性，仅当形状为组合形状时可用
    /// </summary>
    IExcelGroupShapes GroupItems { get; }

    /// <summary>
    /// 获取或设置形状的名称
    /// 对应 Shape.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状的类型
    /// 对应 Shape.Type 属性
    /// </summary>
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取形状的ID
    /// 对应 Shape.ID 属性
    int ID { get; }

    /// <summary>
    /// 获取形状的父对象
    /// 对应 Shape.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置形状的定位方式
    /// 对应 Shape.Placement 属性
    /// </summary>
    XlPlacement Placement { get; set; }


    bool LockAspectRatio { get; set; }

    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置形状的左边距（以磅为单位）
    /// 对应 Shape.Left 属性
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置形状的顶边距（以磅为单位）
    /// 对应 Shape.Top 属性
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度（以磅为单位）
    /// 对应 Shape.Width 属性
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置形状的高度（以磅为单位）
    /// 对应 Shape.Height 属性
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置形状的旋转角度（以度为单位）
    /// 对应 Shape.Rotation 属性
    /// </summary>
    float Rotation { get; set; }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状是否可见
    /// 对应 Shape.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状是否锁定
    /// 对应 Shape.Locked 属性
    /// </summary>
    bool Locked { get; set; }
    #endregion

    #region 格式设置

    /// <summary>
    /// 获取形状的填充格式对象
    /// 对应 Shape.Fill 属性
    /// </summary>
    IExcelFillFormat Fill { get; }

    /// <summary>
    /// 获取形状的线条格式对象
    /// 对应 Shape.Line 属性
    /// </summary>
    IExcelLineFormat Line { get; }

    /// <summary>
    /// 获取形状的文本框架对象
    /// 对应 Shape.TextFrame 属性
    /// </summary>
    IExcelTextFrame TextFrame { get; }

    /// <summary>
    /// 获取形状的阴影格式对象
    /// 对应 Shape.Shadow 属性
    /// </summary>
    IExcelShadowFormat Shadow { get; }

    /// <summary>
    /// 获取形状的三维格式对象
    /// 对应 Shape.ThreeD 属性
    /// </summary>
    IExcelThreeDFormat ThreeD { get; }

    #endregion

    #region 文本属性

    /// <summary>
    /// 获取或设置形状中的文本内容
    /// 对应 Shape.TextFrame.Characters.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的自动调整大小
    /// 对应 Shape.TextFrame.AutoSize 属性
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的水平对齐方式
    /// 对应 Shape.TextFrame.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的垂直对齐方式
    /// 对应 Shape.TextFrame.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 选择形状
    /// 对应 Shape.Select 方法
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制形状
    /// 对应 Shape.Copy 方法
    /// </summary>
    void Copy();

    /// <summary>
    /// 复制形状
    /// 对应 Shape.Copy 方法
    /// </summary>
    void CopyPicture(XlPictureAppearance? Appearance, XlCopyPictureFormat? Format);

    /// <summary>
    /// 剪切形状
    /// 对应 Shape.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除形状
    /// 对应 Shape.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 高度缩放
    /// </summary>
    /// <param name="Factor"></param>
    /// <param name="RelativeToOriginalSize">是否相对于原始大小</param>
    /// <param name="Scale">缩放比例</param>
    void ScaleHeight(float Factor, bool RelativeToOriginalSize, double Scale);

    /// <summary>
    /// 宽度缩放
    /// </summary>
    /// <param name="Factor"></param>
    /// <param name="RelativeToOriginalSize">是否相对于原始大小</param>
    /// <param name="Scale">缩放比例</param>
    void ScaleWidth(float Factor, bool RelativeToOriginalSize, double Scale);

    /// <summary>
    /// 调整形状大小
    /// 对应 Shape.ScaleWidth 和 Shape.ScaleHeight 方法
    /// </summary>
    /// <param name="widthScale">宽度缩放比例</param>
    /// <param name="heightScale">高度缩放比例</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小</param>
    void Scale(double widthScale, double heightScale, bool relativeToOriginalSize = false);

    /// <summary>
    /// 移动形状
    /// 对应 Shape.IncrementLeft 和 Shape.IncrementTop 方法
    /// </summary>
    /// <param name="leftIncrement">左边距增量</param>
    /// <param name="topIncrement">顶边距增量</param>
    void Move(double leftIncrement, double topIncrement);

    /// <summary>
    /// 旋转形状
    /// 对应 Shape.IncrementRotation 方法
    /// </summary>
    /// <param name="rotationIncrement">旋转角度增量（度）</param>
    void Rotate(double rotationIncrement);

    /// <summary>
    /// 将形状置于最前面
    /// 对应 Shape.ZOrder 方法
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将形状置于最后面
    /// 对应 Shape.ZOrder 方法
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 取消组合形状
    /// 对应 Shape.Ungroup 方法
    /// </summary>
    /// <returns>取消组合后的形状集合</returns>
    IExcelShapes Ungroup();

    /// <summary>
    /// 应用自动调整选项
    /// 对应 Shape.Apply 方法
    /// </summary>
    void Apply();

    /// <summary>
    /// 复制形状的格式
    /// 对应 Shape.PickUp 方法
    /// </summary>
    void PickUp();

    #endregion

    #region 层次结构

    /// <summary>
    /// 获取形状所在的区域对象（如果适用）
    /// 对应 Shape.TopLeftCell 属性
    /// </summary>
    IExcelRange TopLeftCell { get; }

    /// <summary>
    /// 获取形状所在的区域对象（如果适用）
    /// 对应 Shape.BottomRightCell 属性
    /// </summary>
    IExcelRange BottomRightCell { get; }

    /// <summary>
    /// 获取形状所在的图表对象（如果是图表）
    /// 对应 Shape.Chart 属性
    /// </summary>
    IExcelChart Chart { get; }

    #endregion
}