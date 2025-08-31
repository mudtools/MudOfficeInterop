//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Fields 的实现类。
/// </summary>
internal class WordFields : IWordFields
{
    private MsWord.Fields _fields;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="fields">原始 COM Fields 对象。</param>
    internal WordFields(MsWord.Fields fields)
    {
        _fields = fields ?? throw new ArgumentNullException(nameof(fields));
        _disposedValue = false;
    }

    #region 属性实现
    /// <inheritdoc/>
    public IWordApplication? Application => _fields != null ? new WordApplication(_fields.Application) : null;

    /// <inheritdoc/>
    public object Parent => _fields?.Parent;

    /// <inheritdoc/>
    public int Count => _fields?.Count ?? 0;

    /// <inheritdoc/>
    public IWordField this[int index]
    {
        get
        {
            if (_fields == null || index < 1 || index > Count)
                return null;

            try
            {
                var field = _fields[index];
                return field != null ? new WordField(field) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IWordField this[IWordRange range]
    {
        get
        {
            if (_fields == null || range == null)
                return null;

            try
            {
                var wordRange = (range as WordRange)?._range;
                if (wordRange != null)
                {
                    // 查找与范围相交的域
                    for (int i = 1; i <= Count; i++)
                    {
                        var field = _fields[i];
                        if (field != null &&
                            field.Code.Start <= wordRange.End &&
                            field.Code.End >= wordRange.Start)
                        {
                            return new WordField(field);
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public object Parent => _fields?.Parent;

    /// <inheritdoc/>
    public IWordField FirstField => this[1];

    /// <inheritdoc/>
    public IWordField LastField => this[Count];

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public int Update()
    {
        if (_fields == null)
            return 0;

        int updatedCount = 0;
        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && field.Update())
            {
                updatedCount++;
            }
        }
        return updatedCount;
    }

    /// <inheritdoc/>
    public void Unlink()
    {
        if (_fields == null)
            return;

        _fields.Unlink();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        if (_fields == null)
            return;

        // 从后往前删除，避免索引变化问题
        for (int i = Count; i >= 1; i--)
        {
            _fields[i]?.Delete();
        }
    }

    /// <inheritdoc/>
    public List<IWordField> GetFieldsByType(WdFieldType fieldType)
    {
        var fields = new List<IWordField>();

        if (_fields == null)
            return fields;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && field.Type == (MsWord.WdFieldType)(int)fieldType)
            {
                fields.Add(new WordField(field));
            }
        }

        return fields;
    }

    /// <inheritdoc/>
    public bool ContainsType(WdFieldType fieldType)
    {
        if (_fields == null)
            return false;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && field.Type == (MsWord.WdFieldType)(int)fieldType)
            {
                return true;
            }
        }

        return false;
    }

    /// <inheritdoc/>
    public int GetCountByType(WdFieldType fieldType)
    {
        if (_fields == null)
            return 0;

        int count = 0;
        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && field.Type == (MsWord.WdFieldType)(int)fieldType)
            {
                count++;
            }
        }

        return count;
    }

    /// <inheritdoc/>
    public List<WdFieldType> GetAllFieldTypes()
    {
        var types = new List<WdFieldType>();

        if (_fields == null)
            return types;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && !types.Contains((WdFieldType)(int)field.Type))
            {
                types.Add((WdFieldType)(int)field.Type);
            }
        }

        return types;
    }


    /// <inheritdoc/>
    public IWordField Add(IWordRange range, WdFieldType type, string text = "", bool preserveFormatting = true)
    {
        if (_fields == null || range == null)
            return null;

        try
        {
            var wordRange = (range as WordRange)?._range;
            if (wordRange != null)
            {
                var field = _fields.Add(wordRange, (MsWord.WdFieldType)(int)type, text, preserveFormatting ? 1 : 0);
                return field != null ? new WordField(field) : null;
            }
            return null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加域类型 {type}。", ex);
        }
    }

    /// <inheritdoc/>
    public List<IWordField> GetDateFields()
    {
        var dateFields = new List<IWordField>();
        var dateTypes = new[]
        {
            MsWord.WdFieldType.wdFieldDate,
            MsWord.WdFieldType.wdFieldCreateDate,
            MsWord.WdFieldType.wdFieldPrintDate,
            MsWord.WdFieldType.wdFieldSaveDate
        };

        if (_fields == null)
            return dateFields;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && dateTypes.Contains(field.Type))
            {
                dateFields.Add(new WordField(field));
            }
        }
        return dateFields;
    }

    /// <inheritdoc/>
    public List<IWordField> GetPageFields()
    {
        var pageFields = new List<IWordField>();
        var pageTypes = new[]
        {
            MsWord.WdFieldType.wdFieldPage,
            MsWord.WdFieldType.wdFieldNumPages,
            MsWord.WdFieldType.wdFieldSectionPages
        };

        if (_fields == null)
            return pageFields;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && pageTypes.Contains(field.Type))
            {
                pageFields.Add(new WordField(field));
            }
        }

        return pageFields;
    }

    /// <inheritdoc/>
    public List<IWordField> GetTOCFields()
    {
        var tocFields = new List<IWordField>();
        var tocTypes = new[]
        {
            MsWord.WdFieldType.wdFieldTOC,
            MsWord.WdFieldType.wdFieldTOA,
            MsWord.WdFieldType.wdFieldIndex,
            MsWord.WdFieldType.wdFieldBibliography
        };

        if (_fields == null)
            return tocFields;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && tocTypes.Contains(field.Type))
            {
                tocFields.Add(new WordField(field));
            }
        }

        return tocFields;
    }

    /// <inheritdoc/>
    public List<IWordField> GetLinkedFields()
    {
        var linkedFields = new List<IWordField>();

        if (_fields == null)
            return linkedFields;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null && field.LinkFormat != null)
            {
                linkedFields.Add(new WordField(field));
            }
        }

        return linkedFields;
    }

    /// <inheritdoc/>
    public void Refresh()
    {
        Update();
    }


    /// <inheritdoc/>
    public List<IWordField> GetFieldsInRange(int startIndex, int endIndex)
    {
        var fields = new List<IWordField>();

        if (_fields == null || startIndex < 1 || endIndex > Count || startIndex > endIndex)
            return fields;

        for (int i = startIndex; i <= endIndex; i++)
        {
            var field = this[i];
            if (field != null)
            {
                fields.Add(field);
            }
        }
        return fields;
    }

    /// <inheritdoc/>
    public int CleanupInvalidFields()
    {
        if (_fields == null)
            return 0;

        int cleanedCount = 0;
        // 从后往前检查和清理
        for (int i = Count; i >= 1; i--)
        {
            var field = _fields[i];
            if (field != null)
            {
                // 检查域是否有效
                if (string.IsNullOrEmpty(field.Code?.Text) ||
                    field.Type == MsWord.WdFieldType.wdFieldEmpty)
                {
                    field.Delete();
                    cleanedCount++;
                }
            }
        }

        return cleanedCount;
    }

    #endregion

    #region IEnumerable<IWordField> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordField> GetEnumerator()
    {
        if (_fields == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var field = _fields[i];
            if (field != null)
                yield return new WordField(field);
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
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fields != null)
        {
            Marshal.ReleaseComObject(_fields);
            _fields = null;
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