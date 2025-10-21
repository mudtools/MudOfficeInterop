//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示邮件合并操作所使用的数据源的二次封装接口。
/// 此接口提供了对数据源的连接信息、字段、记录以及记录导航的访问。
/// </summary>
public interface IWordMailMergeDataSource : IDisposable
{
    /// <summary>
    /// 获取此数据源所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数据源的父对象（通常是 <see cref="IWordMailMerge"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据源中第一个记录的索引号。
    /// </summary>
    int FirstRecord { get; }

    /// <summary>
    /// 获取数据源中最后一个记录的索引号。
    /// </summary>
    int LastRecord { get; }

    /// <summary>
    /// 获取数据源中的记录总数。
    /// </summary>
    int RecordCount { get; }

    /// <summary>
    /// 获取数据源的表名。
    /// </summary>
    string TableName { get; }

    /// <summary>
    /// 获取或设置用于从数据源检索数据的查询字符串。
    /// </summary>
    string? QueryString { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否包含数据源中的记录。
    /// </summary>
    bool Included { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示数据源中的地址是否无效。
    /// </summary>
    bool InvalidAddress { get; set; }

    /// <summary>
    /// 获取数据源的标头源名称。
    /// </summary>
    string HeaderSourceName { get; }

    /// <summary>
    /// 获取数据源的标头源类型。
    /// </summary>
    WdMailMergeDataSource HeaderSourceType { get; }

    /// <summary>
    /// 获取数据源的类型。
    /// </summary>
    WdMailMergeDataSource Type { get; }

    /// <summary>
    /// 获取数据源中活动（当前）记录的索引号。
    /// </summary>
    WdMailMergeActiveRecord ActiveRecord { get; set; }

    /// <summary>
    /// 获取数据源的连接字符串。
    /// 该字符串描述了如何连接到外部数据源（例如，ODBC 或 OLE DB 连接）。
    /// </summary>
    string? ConnectString { get; }

    /// <summary>
    /// 获取数据源的完整路径和文件名。
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// 获取数据源中的数据字段集合。
    /// 数据字段包含当前记录的实际数据值。
    /// </summary>
    IWordMailMergeDataFields? DataFields { get; }

    /// <summary>
    /// 获取映射的数据字段集合。
    /// 映射字段用于将邮件合并域名称映射到数据源中的实际字段名称。
    /// </summary>
    IWordMappedDataFields? MappedDataFields { get; }

    /// <summary>
    /// 获取数据源中的字段集合。
    /// 每个字段代表数据源中的一列。
    /// </summary>
    IWordMailMergeFieldNames? FieldNames { get; }

    /// <summary>
    /// 设置所有记录的包含标志。
    /// </summary>
    /// <param name="Included">如果为 true，则包含所有记录；否则排除所有记录。</param>
    void SetAllIncludedFlags(bool Included);

    /// <summary>
    /// 设置所有记录的错误标志。
    /// </summary>
    /// <param name="invalid">如果为 true，则将所有记录标记为无效。</param>
    /// <param name="invalidComment">与无效状态关联的注释文本。</param>
    void SetAllErrorFlags(bool invalid, string invalidComment);

    /// <summary>
    /// 关闭数据源并释放相关资源。
    /// </summary>
    void Close();

    /// <summary>
    /// 在数据源中查找包含指定文本的记录。
    /// </summary>
    /// <param name="fieldName">要在其中搜索的字段名称。</param>
    /// <param name="text">要搜索的文本。</param>
    /// <returns>如果找到匹配的记录，则为 true；否则为 false。</returns>
    /// <exception cref="ArgumentNullException">当 <paramref name="fieldName"/> 或 <paramref name="text"/> 为 null 或空时抛出。</exception>
    bool FindRecord(string text, string? fieldName = null);
}