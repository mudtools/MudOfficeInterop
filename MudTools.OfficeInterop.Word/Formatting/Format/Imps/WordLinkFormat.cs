//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Core.LinkFormat 的实现类。
/// </summary>
internal class WordLinkFormat : IWordLinkFormat
{
    private MsWord.LinkFormat _linkFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="linkFormat">原始 COM LinkFormat 对象。</param>
    internal WordLinkFormat(MsWord.LinkFormat linkFormat)
    {
        _linkFormat = linkFormat ?? throw new ArgumentNullException(nameof(linkFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _linkFormat != null ? new WordApplication(_linkFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _linkFormat?.Parent;

    /// <inheritdoc/>
    public string SourceFullName
    {
        get => _linkFormat?.SourceFullName ?? string.Empty;
        set
        {
            if (_linkFormat != null)
                _linkFormat.SourceFullName = value;
        }
    }

    /// <inheritdoc/>
    public string SourceName => _linkFormat?.SourceName ?? string.Empty;

    /// <inheritdoc/>
    public string SourcePath => _linkFormat?.SourcePath ?? string.Empty;

    /// <inheritdoc/>
    public bool? AutoUpdate
    {
        get => _linkFormat?.AutoUpdate;
        set
        {
            if (_linkFormat != null)
                _linkFormat.AutoUpdate = value == true;
        }
    }

    /// <inheritdoc/>
    public string ParentType => _linkFormat?.Parent?.GetType().Name ?? string.Empty;

    /// <inheritdoc/>
    public string ParentName
    {
        get
        {
            try
            {
                var parent = _linkFormat?.Parent;
                if (parent != null)
                {
                    var nameProperty = parent.GetType().GetProperty("Name");
                    return nameProperty?.GetValue(parent)?.ToString() ?? string.Empty;
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    /// <inheritdoc/>
    public WdLinkType Type => _linkFormat?.Type != null ? (WdLinkType)(int)_linkFormat?.Type : WdLinkType.wdLinkTypeText;


    /// <inheritdoc/>
    public bool IsEmbedded => _linkFormat?.Parent is MsWord.OLEFormat oleFormat && oleFormat != null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public bool Update()
    {
        if (_linkFormat == null)
            return false;

        try
        {
            _linkFormat.Update();
            return true;
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public bool BreakLink()
    {
        if (_linkFormat == null)
            return false;

        try
        {
            _linkFormat.BreakLink();
            return true;
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public bool Relink(string newSourceFullName)
    {
        if (_linkFormat == null || string.IsNullOrWhiteSpace(newSourceFullName))
            return false;

        try
        {
            _linkFormat.SourceFullName = newSourceFullName;
            return true;
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public bool ValidateLink()
    {
        if (_linkFormat == null)
            return false;

        try
        {
            // 通过尝试访问源文件信息来验证链接
            var sourceName = _linkFormat.SourceName;
            var sourcePath = _linkFormat.SourcePath;
            return !string.IsNullOrEmpty(sourceName);
        }
        catch (COMException)
        {
            return false;
        }
        catch
        {
            return false;
        }
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

        if (disposing && _linkFormat != null)
        {
            Marshal.ReleaseComObject(_linkFormat);
            _linkFormat = null;
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