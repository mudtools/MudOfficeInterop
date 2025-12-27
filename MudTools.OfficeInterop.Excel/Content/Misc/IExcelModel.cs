//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作表页面的页眉和页脚设置接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelModel : IOfficeObject<IExcelModel>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象 
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }


    /// <summary>
    /// 获取数据模型中的表集合。
    /// </summary>
    IExcelModelTables? ModelTables { get; }

    /// <summary>
    /// 获取数据模型中的关系集合。
    /// </summary>
    IExcelModelRelationships? ModelRelationships { get; }

    /// <summary>
    /// 刷新数据模型。
    /// </summary>
    void Refresh();

    /// <summary>
    /// 向数据模型添加与数据源的连接。
    /// </summary>
    /// <param name="connectionToDataSource">要添加到数据模型的数据源连接。</param>
    /// <returns>添加到数据模型的工作簿连接对象。</returns>
    IExcelWorkbookConnection? AddConnection(IExcelWorkbookConnection connectionToDataSource);

    /// <summary>
    /// 为指定的模型表创建模型工作簿连接。
    /// </summary>
    /// <param name="modelTable">要为其创建连接的模型表。</param>
    /// <returns>创建的模型工作簿连接对象。</returns>
    IExcelWorkbookConnection? CreateModelWorkbookConnection(object modelTable);

    /// <summary>
    /// 获取数据模型连接。
    /// </summary>
    IExcelWorkbookConnection? DataModelConnection { get; }

    /// <summary>
    /// 获取数据模型的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 初始化数据模型。
    /// </summary>
    void Initialize();
}