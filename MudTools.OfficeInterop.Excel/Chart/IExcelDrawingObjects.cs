//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


public interface IExcelDrawingObjects : IDisposable, IEnumerable<IExcelDrawing>
{
    /// <summary>
    /// 获取绘图对象集合中的对象数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取绘图对象（索引从1开始）
    /// </summary>
    /// <param name="index">绘图对象索引</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing this[int index] { get; }

    /// <summary>
    /// 根据名称获取绘图对象
    /// </summary>
    /// <param name="name">绘图对象名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing this[string name] { get; }

    /// <summary>
    /// 根据索引或名称获取绘图对象
    /// </summary>
    /// <param name="index">绘图对象索引或名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing GetItem(object index);

    /// <summary>
    /// 根据名称查找绘图对象
    /// </summary>
    /// <param name="name">对象名称</param>
    /// <returns>绘图对象</returns>
    IExcelDrawing FindByName(string name);

    /// <summary>
    /// 删除指定名称的绘图对象
    /// </summary>
    /// <param name="name">对象名称</param>
    void Remove(string name);

    /// <summary>
    /// 清除所有绘图对象
    /// </summary>
    void Clear();

    /// <summary>
    /// 选择所有绘图对象
    /// </summary>
    void SelectAll();

    /// <summary>
    /// 获取可见的绘图对象
    /// </summary>
    IEnumerable<IExcelDrawing> VisibleItems { get; }

    /// <summary>
    /// 获取锁定的绘图对象
    /// </summary>
    IEnumerable<IExcelDrawing> LockedItems { get; }
}