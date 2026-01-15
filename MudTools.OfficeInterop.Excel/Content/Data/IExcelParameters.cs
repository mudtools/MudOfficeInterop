//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示指定查询表的Parameter对象集合。每个Parameter对象表示一个查询参数。每个查询表都包含一个Parameters集合，但除非查询表使用参数查询，否则该集合为空。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelParameters : IEnumerable<IExcelParameter?>, IOfficeObject<IExcelParameters, MsExcel.Parameters>, IDisposable
{
    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取表示Excel应用程序的Application对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取集合中参数的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据名称或索引号从集合中获取单个Parameter对象。
    /// </summary>
    /// <param name="index">必需。对象的名称或索引号。</param>
    /// <returns>指定名称或索引的Parameter对象。</returns>
    IExcelParameter? this[int index] { get; }

    /// <summary>
    /// 根据名称或索引号从集合中获取单个Parameter对象。
    /// </summary>
    /// <param name="name">必需。对象的名称或索引号。</param>
    /// <returns>指定名称或索引的Parameter对象。</returns>
    IExcelParameter? this[string name] { get; }

    /// <summary>
    /// 删除整个参数集合。
    /// </summary>
    void Delete();

    /// <summary>
    /// 创建一个新的查询参数。
    /// </summary>
    /// <param name="name">必需。指定参数的名称。参数名称应与SQL语句中的参数子句匹配。</param>
    /// <param name="dataType">可选。参数的数据类型。可以是任何XlParameterDataType常量。</param>
    /// <returns>新创建的Parameter对象。</returns>
    IExcelParameter Add(string name, XlParameterDataType? dataType = null);
}