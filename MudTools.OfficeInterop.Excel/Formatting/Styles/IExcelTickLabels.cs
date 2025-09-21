//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel TickLabels 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.TickLabels 的安全访问和操作
/// </summary>
public interface IExcelTickLabels : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取坐标轴刻度标签的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取刻度标签的父对象 (通常是 Axis)
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取刻度标签所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取刻度标签的字体对象 
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat Format { get; }

    /// <summary>
    /// 获取或设置是否自动缩放字体
    /// </summary>
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置刻度标签的数字格式
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置刻度标签的数字格式是否为关联格式
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取或设置刻度标签的本地化数字格式
    /// </summary>
    string? NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置刻度标签的方向
    /// </summary>
    int Orientation { get; set; } // 使用 int 代表 XlTickLabelOrientation


    /// <summary>
    /// 获取或设置刻度标签的阅读顺序
    /// </summary>
    int ReadingOrder { get; set; } // 使用 int 代表 XlReadingOrder

    /// <summary>
    /// 获取或设置刻度标签的偏移量 (0-1000)
    /// </summary>
    int Offset { get; set; }

    /// <summary>
    /// 获取或设置多级标签的层级
    /// </summary>
    bool MultiLevel { get; set; }

    #endregion

    #region 操作方法
    /// <summary>
    /// 选择刻度标签
    /// </summary>
    void Select();

    /// <summary>
    /// 删除刻度标签 (通常意味着重置为默认)
    /// </summary>
    void Delete();

    #endregion
}
