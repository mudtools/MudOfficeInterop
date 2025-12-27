//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel工作簿中的连接集合接口
/// 继承自IDisposable和IEnumerable[IExcelWorkbookConnection]接口，提供对Excel工作簿连接的管理功能
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelConnections : IOfficeObject<IExcelConnections>, IDisposable, IEnumerable<IExcelWorkbookConnection>
{

    /// <summary>
    /// 获取父级工作簿
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前对象所在的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取连接集合中的连接数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取连接（索引从1开始）
    /// </summary>
    /// <param name="index">连接索引</param>
    /// <returns>工作簿连接对象</returns>
    IExcelWorkbookConnection? this[int index] { get; }

    /// <summary>
    /// 根据名称获取连接
    /// </summary>
    /// <param name="name">连接名称</param>
    /// <returns>工作簿连接对象</returns>
    IExcelWorkbookConnection? this[string name] { get; }


    /// <summary>
    /// 添加新的工作簿连接
    /// </summary>
    /// <param name="name">连接名称</param>
    /// <param name="description">连接描述</param>
    /// <param name="connectionString">连接字符串</param>
    /// <param name="commandText">命令文本</param>
    /// <param name="lCmdType">命令类型</param>
    /// <returns>新创建的工作簿连接对象</returns>
    IExcelWorkbookConnection? Add(string name, string description, string connectionString,
                                string? commandText = null, XlCmdType lCmdType = XlCmdType.xlCmdSql);

    /// <summary>
    /// 添加新的工作簿连接（增强版）
    /// </summary>
    /// <param name="name">连接名称</param>
    /// <param name="description">连接描述</param>
    /// <param name="connectionString">连接字符串</param>
    /// <param name="commandText">命令文本</param>
    /// <param name="lCmdtype">命令类型</param>
    /// <param name="createModelConnection">是否创建模型连接</param>
    /// <param name="importRelationships">是否导入关系</param>
    /// <returns>新创建的工作簿连接对象</returns>
    IExcelWorkbookConnection? Add2(string name, string description,
        string connectionString, string? commandText = null,
        XlCmdType lCmdtype = XlCmdType.xlCmdSql, bool? createModelConnection = null,
        bool? importRelationships = null);


    /// <summary>
    /// 从文件添加新的工作簿连接
    /// </summary>
    /// <param name="filename">文件名</param>
    /// <returns>新创建的工作簿连接对象</returns>
    IExcelWorkbookConnection? AddFromFile(string filename);

    /// <summary>
    /// 从文件添加新的工作簿连接（增强版）
    /// </summary>
    /// <param name="filename">文件名</param>
    /// <param name="createModelConnection">是否创建模型连接</param>
    /// <param name="importRelationships">是否导入关系</param>
    /// <returns>新创建的工作簿连接对象</returns>
    IExcelWorkbookConnection? AddFromFile2(string filename, bool? createModelConnection, bool? importRelationships);
}
