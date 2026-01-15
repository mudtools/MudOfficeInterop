//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ErrorBars 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ErrorBars 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelErrorBars : IOfficeObject<IExcelErrorBars, MsExcel.ErrorBars>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取误差线对象的名称
    /// 对应 ErrorBars.Name 属性
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取误差线对象的父对象 (通常是 Series 或 Chart)
    /// 对应 ErrorBars.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取误差线对象所在的 Application 对象
    /// 对应 ErrorBars.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取误差线的边框对象
    /// 对应 ErrorBars.Border 属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置误差线的结束样式
    /// 对应 ErrorBars.EndStyle 属性
    /// </summary>
    XlEndStyleCap EndStyle { get; set; }
    #endregion   

    #region 操作方法
    /// <summary>
    /// 选择误差线对象
    /// 对应 ErrorBars.Select 方法
    /// </summary>
    void Select();

    /// <summary>
    /// 删除误差线 (通常意味着隐藏误差线，即设置 Series.HasErrorBars = false)
    /// 对应 ErrorBars.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除误差线格式
    /// 对应 ErrorBars.ClearFormats 方法
    /// </summary>
    void ClearFormats();
    #endregion   
}
