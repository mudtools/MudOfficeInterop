//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel PivotFields 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.PivotFields 的安全访问和操作
/// </summary>
public interface IExcelPivotFields : IEnumerable<IExcelPivotField>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取数据透视表字段集合中的字段数量
    /// 对应 PivotFields.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的数据透视表字段对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">字段索引（从1开始）</param>
    /// <returns>数据透视表字段对象</returns>
    IExcelPivotField this[int index] { get; }

    /// <summary>
    /// 获取指定名称的数据透视表字段对象
    /// </summary>
    /// <param name="name">字段名称</param>
    /// <returns>数据透视表字段对象</returns>
    IExcelPivotField this[string name] { get; }

    /// <summary>
    /// 获取字段集合所在的父对象（通常是 PivotTable）
    /// 对应 PivotFields.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取字段集合所在的Application对象
    /// 对应 PivotFields.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 查找和筛选
    /// <summary>
    /// 根据名称查找字段
    /// </summary>
    /// <param name="name">字段名称</param>
    /// <param name="matchCase">是否区分大小写</param>
    /// <returns>匹配的字段数组</returns>
    IExcelPivotField[] FindByName(string name, bool matchCase = false);

    /// <summary>
    /// 根据字段方向查找字段
    /// </summary>
    /// <param name="orientation">字段方向</param>
    /// <returns>匹配的字段数组</returns>
    IExcelPivotField[] FindByOrientation(XlPivotFieldOrientation orientation);

    /// <summary>
    /// 根据字段位置查找字段
    /// </summary>
    /// <param name="position">字段位置</param>
    /// <returns>匹配的字段数组</returns>
    IExcelPivotField[] FindByPosition(int position);

    /// <summary>
    /// 获取已计算的字段
    /// </summary>
    /// <returns>已计算字段数组</returns>
    IExcelPivotField[] GetCalculatedFields();
    #endregion

    #region 操作方法 
    /// <summary>
    /// 删除指定索引的字段
    /// </summary>
    /// <param name="index">要删除的字段索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定名称的字段
    /// </summary>
    /// <param name="name">要删除的字段名称</param>
    void Delete(string name);

    /// <summary>
    /// 删除指定的字段对象
    /// </summary>
    /// <param name="field">要删除的字段对象</param>
    void Delete(IExcelPivotField field);

    /// <summary>
    /// 批量删除字段
    /// </summary>
    /// <param name="indices">要删除的字段索引数组</param>
    void DeleteRange(int[] indices);
    #endregion
}
