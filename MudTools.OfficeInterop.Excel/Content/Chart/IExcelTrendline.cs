//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Trendline 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Trendline 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelTrendline : IOfficeObject<IExcelTrendline, MsExcel.Trendline>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置趋势线的名称
    /// 对应 Trendline.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取趋势线在集合中的索引
    /// 对应 Trendline.Index 属性
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取趋势线的父对象 (通常是 Trendlines 集合)
    /// 对应 Trendline.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取趋势线所在的 Application 对象
    /// 对应 Trendline.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置趋势线的类型
    /// 对应 Trendline.Type 属性
    /// </summary>
    XlTrendlineType Type { get; set; }

    /// <summary>
    /// 获取或设置趋势线的阶数 (多项式趋势线)
    /// 对应 Trendline.Order 属性
    /// </summary>
    int Order { get; set; }

    /// <summary>
    /// 获取或设置趋势线的周期 (移动平均趋势线)
    /// 对应 Trendline.Period 属性
    /// </summary>
    int Period { get; set; }

    /// <summary>
    /// 获取或设置向前预测周期数
    /// 对应 Trendline.Forward 属性
    /// </summary>
    int Forward { get; set; }

    /// <summary>
    /// 获取或设置向后预测周期数
    /// 对应 Trendline.Backward 属性
    /// </summary>
    int Backward { get; set; }

    /// <summary>
    /// 获取或设置趋势线与 Y 轴的交点
    /// 对应 Trendline.Intercept 属性
    /// </summary>
    double Intercept { get; set; }

    /// <summary>
    /// 获取或设置是否显示公式
    /// 对应 Trendline.DisplayEquation 属性
    /// </summary>
    bool DisplayEquation { get; set; }

    /// <summary>
    /// 获取或设置是否显示 R 平方值
    /// 对应 Trendline.DisplayRSquared 属性
    /// </summary>
    bool DisplayRSquared { get; set; }

    /// <summary>
    /// 获取或设置向前预测周期数（双精度浮点型）
    /// 对应 Trendline.Forward2 属性
    /// </summary>
    double Forward2 { get; set; }

    /// <summary>
    /// 获取或设置向后预测周期数（双精度浮点型）
    /// 对应 Trendline.Backward2 属性
    /// </summary>
    double Backward2 { get; set; }

    /// <summary>
    /// 获取或设置趋势线名称是否为自动分配
    /// 对应 Trendline.NameIsAuto 属性
    /// </summary>
    bool NameIsAuto { get; set; }

    /// <summary>
    /// 获取或设置趋势线与 Y 轴的交点是否为自动计算
    /// 对应 Trendline.InterceptIsAuto 属性
    /// </summary>
    bool InterceptIsAuto { get; set; }

    #endregion

    #region 格式设置

    /// <summary>
    /// 获取趋势线的数据标签对象
    /// 对应 Trendline.DataLabel 属性
    /// </summary>
    IExcelDataLabel? DataLabel { get; }

    /// <summary>
    /// 获取趋势线的边框对象
    /// 对应 Trendline.Border 属性 (如果适用)
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取趋势线的格式对象 (伪代码占位符，用于访问线条等格式)
    /// 对应 Trendline.Format 属性
    /// </summary>
    IExcelChartFormat? Format { get; }
    #endregion


    #region 操作方法
    /// <summary>
    /// 选择趋势线对象
    /// 对应 Trendline.Select 方法
    /// </summary>
    void Select();

    /// <summary>
    /// 删除趋势线
    /// 对应 Trendline.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 清除趋势线格式
    /// 对应 Trendline.ClearFormats 方法
    /// </summary>
    void ClearFormats();
    #endregion    
}
