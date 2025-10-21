//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档的邮件合并功能的二次封装接口。
/// 此接口提供了对邮件合并操作的全面控制，包括数据源管理、执行合并和状态查询 [[2]]。
/// </summary>
public interface IWordMailMerge : IDisposable
{
    /// <summary>
    /// 获取此邮件合并对象所属的 Word 应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此邮件合并对象的父对象（通常是 <see cref="IWordDocument"/>）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个值，该值指示文档是否为邮件合并主文档。
    /// </summary>
    bool MainDocumentType { get; }

    /// <summary>
    /// 获取或设置邮件合并操作的目标。
    /// 例如，可以是新文档、打印机或电子邮件 [[11]]。
    /// </summary>
    WdMailMergeDestination Destination { get; set; }

    /// <summary>
    /// 获取邮件合并操作的当前状态。
    /// 状态可以是正常文档、仅主文档、主文档和数据源等多种情况之一。
    /// </summary>
    WdMailMergeState State { get; }

    bool HighlightMergeFields { get; set; }

    WdMailMergeMailFormat MailFormat { get; set; }

    /// <summary>
    /// 获取或设置邮件合并域代码的显示方式。
    /// 该值控制在文档中是显示域代码还是显示域结果。
    /// </summary>
    int ViewMailMergeFieldCodes { get; set; }

    /// <summary>
    /// 获取或设置是否在邮件合并过程中抑制空白行。
    /// 当设置为true时，如果合并后的行为空，则该行不会出现在结果文档中。
    /// </summary>
    bool SuppressBlankLines { get; set; }

    /// <summary>
    /// 获取或设置邮件合并结果是否作为附件发送。
    /// 当设置为true时，合并结果将作为附件发送；否则，内容将直接包含在邮件正文中。
    /// </summary>
    bool MailAsAttachment { get; set; }

    /// <summary>
    /// 获取或设置包含电子邮件地址的数据源字段名称。
    /// 此字段用于指定在邮件合并过程中使用哪个字段作为收件人邮箱地址。
    /// </summary>
    string MailAddressFieldName { get; set; }

    /// <summary>
    /// 获取或设置邮件合并生成的电子邮件的主题。
    /// 此属性指定通过邮件合并发送的电子邮件的标题。
    /// </summary>
    string MailSubject { get; set; }

    /// <summary>
    /// 获取文档中的邮件合并域集合。
    /// 可用于访问和操作文档中的各个邮件合并域。
    /// </summary>
    IWordMailMergeFields? Fields { get; }

    /// <summary>
    /// 获取邮件合并的数据源对象。
    /// 该对象提供了对连接字符串、记录和字段的访问。
    /// </summary>
    IWordMailMergeDataSource? DataSource { get; }

    /// <summary>
    /// 检查邮件合并设置的有效性并报告任何错误。
    /// </summary>
    void Check();

    /// <summary>
    /// 打开数据源编辑器以修改邮件合并的数据源内容。
    /// </summary>
    void EditDataSource();

    /// <summary>
    /// 打开头文件编辑器以修改邮件合并的页眉信息源内容。
    /// </summary>
    void EditHeaderSource();

    /// <summary>
    /// 打开主文档编辑器以修改邮件合并的主文档内容。
    /// </summary>
    void EditMainDocument();

    /// <summary>
    /// 将指定类型的数据源设置为通讯录地址簿进行邮件合并操作。
    /// </summary>
    /// <param name="Type">要使用的地址簿类型。</param>
    void UseAddressBook(string Type);


    /// <summary>
    /// 执行邮件合并操作。
    /// 根据 <see cref="Destination"/> 属性的设置，结果可能是新文档、打印输出或电子邮件。
    /// </summary>
    /// <param name="pause">如果为 true，Word 在遇到错误时会暂停并显示疑难解答对话框；如果为 false，错误将记录在一个新文档中 [[32]]。</param>
    void Execute(bool pause = false);

    /// <summary>
    /// 将外部数据源附加到邮件合并主文档。
    /// 支持多种数据源类型，如 Access 数据库、Excel 工作簿、文本文件等 [[11]]。
    /// </summary>
    /// <param name="dataSourcePath">数据源文件的完整路径。</param>
    /// <param name="confirmConversions">如果为 true，在打开文档时会提示用户进行格式转换。</param>
    /// <param name="readOnly">如果为 true，以只读方式打开数据源。</param>
    /// <param name="linkToSource">如果为 true，保持与数据源的链接；如果为 false，则嵌入数据。</param>
    /// <param name="connection">可选的连接字符串，用于指定如何连接到数据源（例如，ODBC 连接）。</param>
    /// <param name="sqlStatement">可选的 SQL 语句，用于从数据源中筛选记录。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="dataSourcePath"/> 为 null 或空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当打开数据源失败时抛出。</exception>
    void OpenDataSource(
        string dataSourcePath,
        bool confirmConversions = false,
        bool readOnly = true,
        bool linkToSource = true,
        string? connection = null,
        string? sqlStatement = null);

    /// <summary>
    /// 创建一个新的 Word 文档作为邮件合并的数据源。
    /// 新文档将包含一个表格，用于存储邮件合并数据 [[17]]。
    /// </summary>
    /// <param name="fileName">新数据源文档的完整路径和文件名。</param>
    /// <param name="headerSource">一个包含字段名（以制表符分隔）的字符串。</param>
    /// <param name="data">一个包含记录数据的二维数组，每行代表一条记录。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="fileName"/> 或 <paramref name="headerSource"/> 为 null 或空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当创建数据源失败时抛出。</exception>
    void CreateDataSource(string fileName, string headerSource, object[,] data);
}