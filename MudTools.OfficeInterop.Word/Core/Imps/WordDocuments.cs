//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档集合实现类
/// </summary>
internal class WordDocuments : IWordDocuments
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordApplication));
    private readonly MsWord.Documents _documents;
    private bool _disposedValue;

    public IWordApplication Application => _documents != null ? new WordApplication(_documents.Application) : null;

    public int Count => _documents.Count;

    public object Parent => _documents.Parent;


    internal WordDocuments(MsWord.Documents documents)
    {
        _documents = documents ?? throw new ArgumentNullException(nameof(documents));
        _disposedValue = false;
    }

    public IWordDocument this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                var document = _documents[index];
                return new WordDocument(document);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get document at index {index}.", ex);
            }
        }
    }

    public IWordDocument this[string name]
    {
        get
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Document name cannot be null or empty.", nameof(name));

            try
            {
                var document = _documents[name];
                return new WordDocument(document);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get document with name '{name}'.", ex);
            }
        }
    }

    public IWordDocument Add(string? template = null)
    {
        try
        {
            MsWord.Document doc;
            if (string.IsNullOrEmpty(template))
            {
                doc = _documents.Add();
            }
            else
            {
                doc = _documents.Add(template);
            }
            return new WordDocument(doc);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add document.", ex);
        }
    }

    public IWordDocument Open(string fileName, bool readOnly = false, string? password = null)
    {
        if (!File.Exists(fileName))
            throw new FileNotFoundException("File not found.", fileName);

        try
        {
            object fileNameObj = fileName;
            object confirmConversionsObj = missing;
            object readOnlyObj = readOnly;
            object addToRecentFilesObj = missing;
            object passwordDocumentObj = string.IsNullOrEmpty(password) ? missing : (object)password;
            object passwordTemplateObj = missing;
            object revertObj = missing;
            object writePasswordDocumentObj = missing;
            object writePasswordTemplateObj = missing;
            object formatObj = missing;
            object encodingObj = missing;
            object visibleObj = missing;
            object openAndRepairObj = missing;
            object documentDirectionObj = missing;
            object noEncodingDialogObj = missing;
            object xMLTransformObj = missing;

            var doc = _documents.Open(
                ref fileNameObj,
                ref confirmConversionsObj,
                ref readOnlyObj,
                ref addToRecentFilesObj,
                ref passwordDocumentObj,
                ref passwordTemplateObj,
                ref revertObj,
                ref writePasswordDocumentObj,
                ref writePasswordTemplateObj,
                ref formatObj,
                ref encodingObj,
                ref visibleObj,
                ref openAndRepairObj,
                ref documentDirectionObj,
                ref noEncodingDialogObj,
                ref xMLTransformObj);

            return new WordDocument(doc);
        }
        catch (COMException ex)
        {
            log.Error($"Failed to open document '{fileName}': {ex.Message}", ex);
            return null;
        }
    }

    /// <inheritdoc/>
    public IWordDocument? OpenDocument(string fileName, bool confirmConversions = true, bool readOnly = false, bool addToRecentFiles = true,
                                     string passwordDocument = "", string passwordTemplate = "", bool revert = true, string writePasswordDocument = "",
                                     string writePasswordTemplate = "", WdOpenFormat format = WdOpenFormat.wdOpenFormatAuto,
                                     MsoEncoding encoding = MsoEncoding.msoEncodingSimplifiedChineseAutoDetect, bool visible = true)
    {
        if (_documents == null || string.IsNullOrWhiteSpace(fileName)) return null;

        try
        {
            var document = _documents.Open(fileName, confirmConversions, readOnly, addToRecentFiles,
                                                     passwordDocument, passwordTemplate, revert,
                                                     writePasswordDocument, writePasswordTemplate, format,
                                                     (MsCore.MsoEncoding)(int)encoding, visible);
            return document != null ? new WordDocument(document) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to open document '{fileName}': {ex.Message}", ex);
            return null;
        }
    }

    public IWordDocument GetActiveDocument()
    {
        try
        {
            var activeDoc = (_documents.Parent as MsWord.Application)?.ActiveDocument;
            return activeDoc != null ? new WordDocument(activeDoc) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get active document.", ex);
        }
    }

    public void Close(WdSaveOptions saveChanges = WdSaveOptions.wdSaveChanges,
                     WdOriginalFormat originalFormat = WdOriginalFormat.wdWordDocument,
                     bool? routeDocument = null)
    {
        try
        {
            _documents.Close((MsWord.WdSaveOptions)(int)saveChanges,
            (MsWord.WdOriginalFormat)(int)originalFormat,
            routeDocument.ComArgsVal());
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close all documents.", ex);
        }
    }

    public void Close(bool saveChanges = true)
    {
        try
        {
            object saveOption = saveChanges ? MsWord.WdSaveOptions.wdSaveChanges : MsWord.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = missing;
            object routeDocument = missing;

            _documents.Close(ref saveOption, ref originalFormat, ref routeDocument);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close all documents.", ex);
        }
    }

    public void SaveAll()
    {
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var doc = this[i];
                    doc.Save();
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save all documents.", ex);
        }
    }

    public IEnumerator<IWordDocument> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private static readonly object missing = System.Reflection.Missing.Value;

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}