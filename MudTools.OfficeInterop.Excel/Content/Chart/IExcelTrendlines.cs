//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Trendlines 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Trendlines 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelTrendlines : IEnumerable<IExcelTrendline>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取趋势线集合中的趋势线数量
    /// 对应 Trendlines.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的趋势线对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">趋势线索引（从1开始）</param>
    /// <returns>趋势线对象</returns>
    IExcelTrendline? this[int index] { get; }

    /// <summary>
    /// 获取趋势线集合所在的父对象 (通常是 Series)
    /// 对应 Trendlines.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取趋势线集合所在的 Application 对象
    /// 对应 Trendlines.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的趋势线
    /// 对应 Trendlines.Add 方法
    /// </summary>
    /// <param name="type">趋势线类型</param>
    /// <param name="order">趋势线阶数 (多项式)</param>
    /// <param name="period">趋势线周期 (移动平均)</param>
    /// <param name="forward">向前预测周期数</param>
    /// <param name="backward">向后预测周期数</param>
    /// <param name="intercept">趋势线与 Y 轴的交点</param>
    /// <param name="displayEquation">是否显示公式</param>
    /// <param name="displayRSquared">是否显示 R 平方值</param>
    /// <param name="name">趋势线名称</param>
    /// <returns>新创建的趋势线对象</returns>
    IExcelTrendline Add(XlTrendlineType type = XlTrendlineType.xlLinear, int? order = null, int? period = null, double? forward = null,
                       double? backward = null, double? intercept = null, bool? displayEquation = false,
                       bool? displayRSquared = false, string? name = "");
    #endregion
}

