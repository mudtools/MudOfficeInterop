//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Styles 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Styles 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelStyles : IEnumerable<IExcelStyle?>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取样式集合所在的父对象（通常是工作簿）
    /// 对应 Styles.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取样式集合所在的Application对象
    /// 对应 Styles.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取样式集合中的样式数量
    /// 对应 Styles.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的样式对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">样式索引（从1开始）</param>
    /// <returns>样式对象</returns>
    IExcelStyle? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的样式对象
    /// </summary>
    /// <param name="name">样式名称</param>
    /// <returns>样式对象</returns>
    IExcelStyle? this[string name] { get; }


    #endregion

    #region 创建和添加

    /// <summary>
    /// 创建一个新样式，并将其添加到当前工作簿可用的样式列表中。
    /// </summary>
    /// <param name="name">必需。新样式的名称。</param>
    /// <param name="basedOn">可选。一个Range对象，引用用作新样式基础的单元格。如果省略此参数，则新创建的样式基于Normal样式。</param>
    /// <returns>新创建的Style对象。</returns>
    IExcelStyle Add(string name, IExcelRange? basedOn = null);

    /// <summary>
    /// 将另一个工作簿中的样式合并到集合中
    /// </summary>
    /// <param name="workbook">目标工作簿对象</param>
    void Merge(IExcelWorkbook workbook);
    #endregion
}

