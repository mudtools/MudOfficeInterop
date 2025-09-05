//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 FileTypes 的二次封装实现类。
/// 提供安全访问文件类型集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeFileTypes : IOfficeFileTypes
{
    private MsCore.FileTypes _fileTypes;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 FileTypes 对象。
    /// </summary>
    /// <param name="fileTypes">原始的 COM FileTypes 对象。</param>
    internal OfficeFileTypes(MsCore.FileTypes fileTypes)
    {
        _fileTypes = fileTypes ?? throw new ArgumentNullException(nameof(fileTypes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _fileTypes?.Count ?? 0;

    /// <inheritdoc/>
    public MsoFileType this[int index]
    {
        get
        {
            if (_fileTypes == null || index < 1 || index > Count)
                return (MsoFileType)MsCore.MsoFileType.msoFileTypeAllFiles;

            try
            {
                return (MsoFileType)(int)_fileTypes[index];
            }
            catch
            {
                return MsoFileType.msoFileTypeAllFiles;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Add(MsoFileType fileType)
    {
        _fileTypes?.Add((MsCore.MsoFileType)(int)fileType);
    }

    /// <inheritdoc/>
    public void Remove(MsoFileType fileType)
    {
        var index = IndexOf(fileType);
        if (index > 0)
        {
            _fileTypes?.Remove(index);
        }
    }

    /// <inheritdoc/>
    public bool Contains(MsoFileType fileType)
    {
        if (_fileTypes == null)
            return false;

        for (int i = 1; i <= Count; i++)
        {
            if (_fileTypes[i] == (MsCore.MsoFileType)(int)fileType)
                return true;
        }
        return false;
    }

    /// <inheritdoc/>
    public void Clear()
    {
        while (Count > 0)
        {
            _fileTypes?.Remove(1);
        }
    }

    /// <inheritdoc/>
    public int IndexOf(MsoFileType fileType)
    {
        if (_fileTypes == null)
            return 0;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                if (_fileTypes[i] == (MsCore.MsoFileType)(int)fileType)
                    return i;
            }
            return 0;
        }
        catch
        {
            return 0;
        }
    }


    /// <inheritdoc/>
    public void RemoveAt(int index)
    {
        if (_fileTypes == null || index < 1 || index > Count)
            return;

        _fileTypes?.Remove(index);
    }

    /// <inheritdoc/>
    public MsoFileType[] GetAllFileTypes()
    {
        if (_fileTypes == null)
            return [];

        try
        {
            var array = new MsoFileType[Count];
            for (int i = 1; i <= Count; i++)
            {
                array[i - 1] = (MsoFileType)(int)_fileTypes[i];
            }
            return array;
        }
        catch
        {
            return [];
        }
    }

    /// <inheritdoc/>
    public void SetDefault()
    {
        // 清除现有类型并设置默认类型
        Clear();

        // 添加常用的默认文件类型
        Add(MsoFileType.msoFileTypeAllFiles);
        Add(MsoFileType.msoFileTypeWordDocuments);
        Add(MsoFileType.msoFileTypeExcelWorkbooks);
        Add(MsoFileType.msoFileTypeOfficeFiles);
        Add(MsoFileType.msoFileTypePowerPointPresentations);
        Add(MsoFileType.msoFileTypeTemplates);
        Add(MsoFileType.msoFileTypeWordDocuments);
    }

    /// <inheritdoc/>
    public string GetDescription(MsoFileType fileType)
    {
        return fileType switch
        {
            MsoFileType.msoFileTypeAllFiles => "所有文件",
            MsoFileType.msoFileTypeBinders => "活页夹文件",
            MsoFileType.msoFileTypeCalendarItem => "日历项目",
            MsoFileType.msoFileTypeContactItem => "联系人项目",
            MsoFileType.msoFileTypeDatabases => "数据库文件",
            MsoFileType.msoFileTypeDataConnectionFiles => "数据连接文件",
            MsoFileType.msoFileTypeExcelWorkbooks => "Excel 工作簿",
            MsoFileType.msoFileTypeJournalItem => "日记项目",
            MsoFileType.msoFileTypeMailItem => "邮件项目",
            MsoFileType.msoFileTypeNoteItem => "笔记项目",
            MsoFileType.msoFileTypeOfficeFiles => "Office 文件",
            MsoFileType.msoFileTypeOutlookItems => "Outlook 项目",
            MsoFileType.msoFileTypePhotoDrawFiles => "PhotoDraw 文件",
            MsoFileType.msoFileTypePowerPointPresentations => "PowerPoint 演示文稿",
            MsoFileType.msoFileTypeProjectFiles => "Project 文件",
            MsoFileType.msoFileTypePublisherFiles => "Publisher 文件",
            MsoFileType.msoFileTypeTaskItem => "任务项目",
            MsoFileType.msoFileTypeTemplates => "模板文件",
            MsoFileType.msoFileTypeVisioFiles => "Visio 文件",
            MsoFileType.msoFileTypeWebPages => "网页文件",
            MsoFileType.msoFileTypeWordDocuments => "Word 文档",
            MsoFileType.msoFileTypeDocumentImagingFiles => "文档影像文件类型",
            MsoFileType.msoFileTypeDesignerFiles => "Designer文件类型",
            _ => "未知文件类型",
        };
    }

    /// <inheritdoc/>
    public string GetExtension(MsoFileType fileType)
    {
        return fileType switch
        {
            MsoFileType.msoFileTypeWordDocuments => ".doc,.docx",
            MsoFileType.msoFileTypeExcelWorkbooks => ".xls,.xlsx",
            MsoFileType.msoFileTypePowerPointPresentations => ".ppt,.pptx",
            MsoFileType.msoFileTypePublisherFiles => ".pub",
            MsoFileType.msoFileTypeVisioFiles => ".vsd,.vsdx",
            MsoFileType.msoFileTypeProjectFiles => ".mpp",
            MsoFileType.msoFileTypeTemplates => ".dot,.dotx,.xlt,.xltm,.pot,.potx",
            MsoFileType.msoFileTypeWebPages => ".htm,.html",
            _ => "*.*",
        };
    }

    #endregion

    #region IEnumerable<MsoFileType> 实现

    /// <inheritdoc/>
    public IEnumerator<MsoFileType> GetEnumerator()
    {
        if (_fileTypes == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            yield return (MsoFileType)(int)_fileTypes[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fileTypes != null)
        {
            try
            {
                Marshal.ReleaseComObject(_fileTypes);
            }
            catch
            {
                // 忽略释放异常
            }
            _fileTypes = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}