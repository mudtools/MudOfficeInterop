//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Picture 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Picture 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPicture : IOfficeObject<IExcelPicture, MsExcel.Picture>, IDisposable
{

    /// <summary>
    /// 获取图例所在的 Application 对象
    /// 对应 Picture .Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #region 基础属性

    /// <summary>
    /// 获取或设置图片的名称
    /// 对应 Picture.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取图片的索引位置
    /// 对应 Picture.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置图片是否可见
    /// 对应 Picture.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取图片所在的父对象
    /// 对应 Picture.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图片的底层形状对象
    /// 对应 Picture.Shape 属性
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取图片的边框格式
    /// 对应 Picture.Border 属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取图片的内部填充格式
    /// 对应 Picture.Interior 属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取图片左上角所在的单元格
    /// 对应 Picture.TopLeftCell 属性
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取图片右下角所在的单元格
    /// 对应 Picture.BottomRightCell 属性
    /// </summary>
    IExcelRange? BottomRightCell { get; }

    /// <summary>
    /// 获取或设置图片是否在打印时可见
    /// 对应 Picture.PrintObject 属性
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 获取或设置图片是否被锁定
    /// 对应 Picture.Locked 属性
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置图片是否启用
    /// 对应 Picture.Enabled 属性
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取图片的 Z 轴顺序位置
    /// 对应 Picture.ZOrder 属性
    /// </summary>
    int ZOrder { get; }

    #endregion

    #region 位置和大小

    /// <summary>
    /// 获取或设置图片的左边距
    /// 对应 Picture.Left 属性
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置图片的顶边距
    /// 对应 Picture.Top 属性
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置图片的宽度
    /// 对应 Picture.Width 属性
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置图片的高度
    /// 对应 Picture.Height 属性
    /// </summary>
    double Height { get; set; }

    #endregion

    #region 图片属性
    /// <summary>
    /// 获取或设置图片的公式
    /// 对应 Picture.Formula 属性
    /// </summary>
    string? Formula { get; set; }
    /// <summary>
    /// 获取或设置图片是否显示阴影效果
    /// 对应 Picture.Shadow 属性
    /// </summary>
    bool? Shadow { get; set; }

    #endregion

    #region 操作方法
    /// <summary>
    /// 将图片置于底层
    /// 对应 Picture.SendToBack 方法
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 将图片置于顶层
    /// 对应 Picture.BringToFront 方法
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 选择图片
    /// 对应 Picture.Select 方法
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除图片
    /// 对应 Picture.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制图片到剪贴板
    /// 对应 Picture.Copy 方法
    /// </summary>
    /// <returns>返回复制的图片对象，如果复制失败则返回 null</returns>
    [ValueConvert]
    IExcelPicture? Copy();

    /// <summary>
    /// 创建图片的副本
    /// 对应 Picture.Duplicate 方法
    /// </summary>
    /// <returns>返回复制的图片对象，如果复制失败则返回 null</returns>
    [ValueConvert]
    IExcelPicture? Duplicate();

    /// <summary>
    /// 将图片复制为指定格式的图片对象
    /// 对应 Picture.CopyPicture 方法
    /// </summary>
    /// <param name="Appearance">指定图片外观类型，如屏幕显示样式或打印机样式</param>
    /// <param name="Format">指定复制的图片格式，如位图或图片格式</param>
    /// <returns>返回复制的图片对象，如果复制失败则返回 null</returns>
    [ValueConvert]
    IExcelPicture? CopyPicture(
       XlPictureAppearance Appearance = XlPictureAppearance.xlPrinter,
       XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);

    /// <summary>
    /// 剪切图片到剪贴板并删除原图片
    /// 对应 Picture.Cut 方法
    /// </summary>
    /// <returns>返回剪切的图片对象，如果剪切失败则返回 null</returns>
    [ValueConvert]
    IExcelPicture? Cut();

    #endregion
}