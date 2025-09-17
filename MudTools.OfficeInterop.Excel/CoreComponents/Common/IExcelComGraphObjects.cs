//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 公共的图形对象接口。
/// </summary>
public interface IExcelComGraphObjects : IDisposable
{
    /// <summary>
    /// 获取或设置颜色的应用程序版本
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取控件的边框属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取控件的内部属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取控件的形状区域属性
    /// </summary>
    IExcelShapeRange? ShapeRange { get; }


    /// <summary>
    /// 获取图表对象集合中的图表数量
    /// 对应 ChartObjects.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置控件是否被锁定
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置控件的宽度
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置控件的高度
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 获取或设置控件上边缘到工作表上边缘的距离
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置控件左边缘到工作表左边缘的距离
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置控件是否显示阴影效果
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置打印工作表时是否打印该控件
    /// </summary>
    bool PrintObject { get; set; }

    /// <summary>
    /// 获取或设置控件是否可见
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 选择所有图表对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    void SelectAll(bool replace = true);

    /// <summary>
    /// 复制图表对象到剪贴板
    /// </summary>
    /// <returns>返回复制操作的结果对象</returns>
    object Copy();

    /// <summary>
    /// 剪切图表对象到剪贴板
    /// </summary>
    /// <returns>返回剪切操作的结果对象</returns>
    object Cut();

    /// <summary>
    /// 删除图表对象
    /// </summary>
    void Delete();
}
