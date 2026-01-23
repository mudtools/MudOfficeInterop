//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Line (边框线条) 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.LineFormat 的安全访问和操作
/// 用于设置形状或图表元素的边框线条
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelLine : IOfficeObject<IExcelLine, MsExcel.Line>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取线条所在的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取线条对象所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 线条属性
    /// <summary>
    /// 获取或设置线条的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置打印时是否包含该对象
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 获取或设置对象是否被锁定
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置线条的宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置线条的高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取线条在集合中的索引值
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取线条的 Z 轴顺序
    /// </summary>
    int ZOrder { get; }

    /// <summary>
    /// 获取或设置线条距其顶部边界的距离
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置线条距其左边界的距离
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置箭头头部样式
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlArrowHeadStyle ArrowHeadStyle { get; set; }

    /// <summary>
    /// 获取或设置箭头头部宽度
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlArrowHeadWidth ArrowHeadWidth { get; set; }

    /// <summary>
    /// 获取或设置箭头头部长度
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float ArrowHeadLength { get; set; }

    /// <summary>
    /// 获取线条左上角单元格
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取线条右下角单元格
    /// </summary>
    IExcelRange? BottomRightCell { get; }

    /// <summary>
    /// 获取线条的形状范围对象
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }

    /// <summary>
    /// 获取线条的边框对象
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取或设置线条是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置线条是否启用
    /// </summary>
    bool Enabled { get; set; }
    #endregion

    #region 方法
    /// <summary>
    /// 将对象放到最前面
    /// </summary>
    /// <returns>操作结果</returns>
    object? BringToFront();

    /// <summary>
    /// 将对象放到最后面
    /// </summary>
    /// <returns>操作结果</returns>
    object? SendToBack();

    /// <summary>
    /// 剪切对象
    /// </summary>
    /// <returns>操作结果</returns>
    object? Cut();

    /// <summary>
    /// 复制对象
    /// </summary>
    /// <returns>操作结果</returns>
    object? Copy();

    /// <summary>
    /// 删除对象
    /// </summary>
    /// <returns>操作结果</returns>
    object? Delete();

    /// <summary>
    /// 复制对象
    /// </summary>
    /// <returns>复制的对象</returns>
    object? Duplicate();

    /// <summary>
    /// 复制对象图片
    /// </summary>
    /// <param name="appearance">图片外观样式</param>
    /// <param name="format">图片格式</param>
    /// <returns>操作结果</returns>
    object? CopyPicture(XlPictureAppearance appearance, XlCopyPictureFormat format);

    /// <summary>
    /// 选择对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    /// <returns>操作结果</returns>
    object? Select(bool replace = true);
    #endregion
}