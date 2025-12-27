//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Areas 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Areas 的安全访问和操作
/// </summary>
public interface IExcelAreas : IEnumerable<IExcelRange>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取区域集合中的区域数量
    /// 对应 Areas.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的区域对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">区域索引（从1开始）</param>
    /// <returns>区域对象</returns>
    IExcelRange this[int index] { get; }

    /// <summary>
    /// 获取区域集合所在的父对象（通常是 Range）
    /// 对应 Areas.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取区域集合所在的Application对象
    /// 对应 Areas.Application 属性
    /// </summary>
    IExcelApplication? Application { get; }
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据地址查找区域 (占位符，因为 Areas 通常是连续的，查找意义不大)
    /// </summary>
    /// <param name="address">区域地址</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的区域数组</returns>
    IExcelRange[] FindByAddress(string address, bool matchCase = false);

    /// <summary>
    /// 根据大小查找区域 (占位符)
    /// </summary>
    /// <param name="rowCount">行数</param>
    /// <param name="columnCount">列数</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的区域数组</returns>
    IExcelRange[] FindBySize(int rowCount, int columnCount, int tolerance = 0);

    /// <summary>
    /// 获取最大的区域
    /// </summary>
    /// <returns>最大的区域对象</returns>
    IExcelRange GetLargestArea();

    /// <summary>
    /// 获取最小的区域
    /// </summary>
    /// <returns>最小的区域对象</returns>
    IExcelRange GetSmallestArea();

    /// <summary>
    /// 获取可见的区域 (占位符，通常由父 Range 决定)
    /// </summary>
    /// <returns>可见区域数组</returns>
    IExcelRange[] GetVisibleAreas();

    /// <summary>
    /// 获取隐藏的区域 (占位符，通常由父 Range 决定)
    /// </summary>
    /// <returns>隐藏区域数组</returns>
    IExcelRange[] GetHiddenAreas();
    #endregion

    #region 操作方法  
    /// <summary>
    /// 删除指定索引的区域 (通常通过父 Range 操作，或删除整个父 Range)
    /// </summary>
    /// <param name="index">要删除的区域索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的区域对象 (通常通过父 Range 操作)
    /// </summary>
    /// <param name="area">要删除的区域对象</param>
    void Delete(IExcelRange area);

    /// <summary>
    /// 批量删除区域 (通常通过父 Range 操作)
    /// </summary>
    /// <param name="indices">要删除的区域索引数组</param>
    void DeleteRange(int[] indices);

    #endregion
}