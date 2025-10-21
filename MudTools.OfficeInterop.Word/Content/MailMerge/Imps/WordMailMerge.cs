//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMerge : IWordMailMerge
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMerge));
    private MsWord.MailMerge? _mailMerge;
    private bool _disposedValue;

    internal WordMailMerge(MsWord.MailMerge mailMerge)
    {
        _mailMerge = mailMerge ?? throw new ArgumentNullException(nameof(mailMerge));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _mailMerge != null ? new WordApplication(_mailMerge.Application) : null;

    public object? Parent => _mailMerge?.Parent;


    public WdMailMergeMainDocType MainDocumentType
    {
        get => _mailMerge?.MainDocumentType.EnumConvert(WdMailMergeMainDocType.wdNotAMergeDocument) ?? WdMailMergeMainDocType.wdNotAMergeDocument;
        set
        {
            if (_mailMerge != null)
                _mailMerge.MainDocumentType = value.EnumConvert(MsWord.WdMailMergeMainDocType.wdNotAMergeDocument);
        }
    }

    public WdMailMergeDestination Destination
    {
        get => _mailMerge?.Destination.EnumConvert(WdMailMergeDestination.wdSendToNewDocument) ?? WdMailMergeDestination.wdSendToNewDocument;
        set
        {
            if (_mailMerge != null)
                _mailMerge.Destination = value.EnumConvert(MsWord.WdMailMergeDestination.wdSendToNewDocument);
        }
    }

    public int ViewMailMergeFieldCodes
    {
        get => _mailMerge?.ViewMailMergeFieldCodes ?? 0;
        set
        {
            if (_mailMerge != null)
                _mailMerge.ViewMailMergeFieldCodes = value;
        }
    }

    public bool SuppressBlankLines
    {
        get => _mailMerge?.SuppressBlankLines ?? false;
        set
        {
            if (_mailMerge != null)
                _mailMerge.SuppressBlankLines = value;
        }
    }

    public bool MailAsAttachment
    {
        get => _mailMerge?.MailAsAttachment ?? false;
        set
        {
            if (_mailMerge != null)
                _mailMerge.MailAsAttachment = value;
        }
    }

    public string MailAddressFieldName
    {
        get => _mailMerge?.MailAddressFieldName ?? string.Empty;
        set
        {
            if (_mailMerge != null)
                _mailMerge.MailAddressFieldName = value;
        }
    }

    public bool HighlightMergeFields
    {
        get => _mailMerge?.HighlightMergeFields ?? false;
        set
        {
            if (_mailMerge != null)
                _mailMerge.HighlightMergeFields = value;
        }
    }

    public string MailSubject
    {
        get => _mailMerge?.MailSubject ?? string.Empty;
        set
        {
            if (_mailMerge != null)
                _mailMerge.MailSubject = value;
        }
    }

    public WdMailMergeState State
    {
        get => _mailMerge?.State.EnumConvert(WdMailMergeState.wdNormalDocument) ?? WdMailMergeState.wdNormalDocument;
    }

    public WdMailMergeMailFormat MailFormat
    {
        get => _mailMerge?.MailFormat.EnumConvert(WdMailMergeMailFormat.wdMailFormatHTML) ?? WdMailMergeMailFormat.wdMailFormatHTML;
        set
        {
            if (_mailMerge != null)
                _mailMerge.MailFormat = value.EnumConvert(MsWord.WdMailMergeMailFormat.wdMailFormatHTML);
        }
    }


    public IWordMailMergeDataSource? DataSource => _mailMerge?.DataSource != null ? new WordMailMergeDataSource(_mailMerge.DataSource) : null;

    public IWordMailMergeFields? Fields => _mailMerge?.Fields != null ? new WordMailMergeFields(_mailMerge.Fields) : null;

    #endregion

    #region 方法实现

    public void Execute(bool pause = false)
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            object pauseObj = pause;
            _mailMerge.Execute(ref pauseObj);
        }
        catch (Exception ex)
        {
            log.Error("执行邮件合并操作失败。", ex);
            throw new InvalidOperationException("执行邮件合并操作失败。", ex);
        }
    }

    public void OpenDataSource(string name, WdOpenFormat? format = null, bool? confirmConversions = null,
        bool? readOnly = null, bool? linkToSource = null, bool? addToRecentFiles = null, string? passwordDocument = null,
        string? passwordTemplate = null, bool? revert = null, string? writePasswordDocument = null,
        string? writePasswordTemplate = null, string? connection = null, string? sqlStatement = null,
        string? sqlStatement1 = null, bool? openExclusive = null, WdMergeSubType? subType = null)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.OpenDataSource(
                Name: name,
                Format: format.ComArgsConvert(e => e.EnumConvert(MsWord.WdOpenFormat.wdOpenFormatAuto)),
                ConfirmConversions: confirmConversions.ComArgsVal(),
                ReadOnly: readOnly.ComArgsVal(),
                LinkToSource: linkToSource.ComArgsVal(),
                AddToRecentFiles: addToRecentFiles.ComArgsVal(),
                PasswordDocument: passwordDocument.ComArgsVal(),
                PasswordTemplate: passwordTemplate.ComArgsVal(),
                Revert: revert.ComArgsVal(),
                WritePasswordDocument: writePasswordDocument.ComArgsVal(),
                WritePasswordTemplate: writePasswordTemplate.ComArgsVal(),
                Connection: connection.ComArgsVal(),
                SQLStatement: sqlStatement.ComArgsVal(),
                SQLStatement1: sqlStatement1.ComArgsVal(),
                OpenExclusive: openExclusive.ComArgsVal(),
                SubType: subType.ComArgsConvert(e => e.EnumConvert(MsWord.WdMergeSubType.wdMergeSubTypeOther))
            );
        }
        catch (Exception ex)
        {
            log.Error("打开数据源失败。", ex);
            throw new InvalidOperationException("打开数据源失败。", ex);
        }
    }


    public void OpenDataSource(
        string dataSourcePath,
        bool confirmConversions = false,
        bool readOnly = true,
        bool linkToSource = true,
        string? connection = null,
        string? sqlStatement = null)
    {
        if (string.IsNullOrEmpty(dataSourcePath))
            throw new ArgumentNullException(nameof(dataSourcePath));

        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.OpenDataSource(
                Name: dataSourcePath,
                ConfirmConversions: confirmConversions,
                ReadOnly: readOnly,
                LinkToSource: linkToSource,
                Connection: connection ?? string.Empty,
                SQLStatement: sqlStatement ?? string.Empty
            );
        }
        catch (Exception ex)
        {
            log.Error($"打开邮件合并数据源 '{dataSourcePath}' 失败。", ex);
            throw new InvalidOperationException($"打开邮件合并数据源 '{dataSourcePath}' 失败。", ex);
        }
    }

    public void CreateDataSource(string fileName, string headerSource, object[,] data)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentNullException(nameof(fileName));
        if (string.IsNullOrEmpty(headerSource))
            throw new ArgumentNullException(nameof(headerSource));

        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            // 将二维数组数据转换为制表符分隔的字符串
            var records = new System.Text.StringBuilder();
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    records.Append(data[i, j]?.ToString() ?? string.Empty);
                    if (j < cols - 1) records.Append('\t');
                }
                records.AppendLine();
            }

            object name = fileName;
            object headerRecord = headerSource;
            object dataTable = records.ToString();
            _mailMerge.CreateDataSource(
                Name: name,
                HeaderRecord: headerRecord,
                MSQuery: dataTable
            );
        }
        catch (Exception ex)
        {
            log.Error($"创建邮件合并数据源 '{fileName}' 失败。", ex);
            throw new InvalidOperationException($"创建邮件合并数据源 '{fileName}' 失败。", ex);
        }
    }

    public void Check()
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.Check();
        }
        catch (Exception ex)
        {
            log.Error("检查邮件合并数据源失败。", ex);
            throw new InvalidOperationException("检查邮件合并数据源失败。", ex);
        }
    }

    public void EditDataSource()
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.EditDataSource();
        }
        catch (Exception ex)
        {
            log.Error("编辑邮件合并数据源失败。", ex);
            throw new InvalidOperationException("编辑邮件合并数据源失败。", ex);
        }
    }

    public void EditHeaderSource()
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.EditHeaderSource();
        }
        catch (Exception ex)
        {
            log.Error("编辑邮件合并数据源失败。", ex);
            throw new InvalidOperationException("编辑邮件合并数据源失败。", ex);
        }
    }

    public void EditMainDocument()
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.EditMainDocument();
        }
        catch (Exception ex)
        {
            log.Error("编辑邮件合并数据源失败。", ex);
            throw new InvalidOperationException("编辑邮件合并数据源失败。", ex);
        }
    }

    public void UseAddressBook(string Type)
    {
        if (_mailMerge == null)
            throw new InvalidOperationException("邮件合并对象不可用。");

        try
        {
            _mailMerge.UseAddressBook(Type);
        }
        catch (Exception ex)
        {
            log.Error($"使用通讯录 '{Type}' 失败。", ex);
            throw new InvalidOperationException($"使用通讯录 '{Type}' 失败。", ex);
        }
    }


    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mailMerge != null)
        {
            Marshal.ReleaseComObject(_mailMerge);
            _mailMerge = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}