//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotItems 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotItems 的安全访问和操作
/// </summary>
public interface IExcelPivotItems : IEnumerable<IExcelPivotItem>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表项目集合中的项目数量
    /// 对应 PivotItems.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的数据透视表项目对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">项目索引（从1开始）</param>
    /// <returns>数据透视表项目对象</returns>
    IExcelPivotItem this[int index] { get; }

    /// <summary>
    /// 获取指定名称的数据透视表项目对象
    /// </summary>
    /// <param name="name">项目名称</param>
    /// <returns>数据透视表项目对象</returns>
    IExcelPivotItem this[string name] { get; }

    /// <summary>
    /// 获取项目集合所在的父对象（通常是 PivotField）
    /// 对应 PivotItems.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取项目集合所在的Application对象
    /// 对应 PivotItems.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找项目
    /// </summary>
    /// <param name="name">项目名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的项目数组</returns>
    IExcelPivotItem[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据可见性查找项目
    /// </summary>
    /// <param name="isVisible">是否可见</param>
    /// <returns>匹配的项目数组</returns>
    IExcelPivotItem[] FindByVisibility(bool isVisible);

    /// <summary>
    /// 获取可见的项目
    /// </summary>
    /// <returns>可见项目数组</returns>
    IExcelPivotItem[] GetVisibleItems();

    /// <summary>
    /// 获取隐藏的项目
    /// </summary>
    /// <returns>隐藏项目数组</returns>
    IExcelPivotItem[] GetHiddenItems();
    #endregion

    #region 操作方法
    /// <summary>
    /// 隐藏集合中的所有项目
    /// </summary>
    void HideAll();

    /// <summary>
    /// 显示集合中的所有项目
    /// </summary>
    void ShowAll();

    /// <summary>
    /// 删除指定索引的项目 (通常意味着隐藏)
    /// </summary>
    /// <param name="index">要隐藏的项目索引</param>
    void Hide(int index);

    /// <summary>
    /// 删除指定名称的项目 (通常意味着隐藏)
    /// </summary>
    /// <param name="name">要隐藏的项目名称</param>
    void Hide(string name);

    /// <summary>
    /// 删除指定的项目对象 (通常意味着隐藏)
    /// </summary>
    /// <param name="item">要隐藏的项目对象</param>
    void Hide(IExcelPivotItem item);

    /// <summary>
    /// 批量隐藏项目
    /// </summary>
    /// <param name="indices">要隐藏的项目索引数组</param>
    void HideRange(int[] indices);
    #endregion
}
