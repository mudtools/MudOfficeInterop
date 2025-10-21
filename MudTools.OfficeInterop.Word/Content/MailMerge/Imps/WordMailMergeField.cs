//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordMailMergeField : IWordMailMergeField
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordMailMergeField));
    private MsWord.MailMergeField? _mailMergeField;
    private bool _disposedValue;

    internal WordMailMergeField(MsWord.MailMergeField mailMergeField)
    {
        _mailMergeField = mailMergeField ?? throw new ArgumentNullException(nameof(mailMergeField));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _mailMergeField != null ? new WordApplication(_mailMergeField.Application) : null;

    public object? Parent => _mailMergeField?.Parent;

    public IWordRange? Code => _mailMergeField?.Code != null ? new WordRange(_mailMergeField.Code) : null;

    public IWordMailMergeField? Next => _mailMergeField?.Next != null ? new WordMailMergeField(_mailMergeField.Next) : null;

    public IWordMailMergeField? Previous => _mailMergeField?.Previous != null ? new WordMailMergeField(_mailMergeField.Previous) : null;

    public WdFieldType Type => _mailMergeField?.Type.EnumConvert(WdFieldType.wdFieldEmpty) ?? WdFieldType.wdFieldEmpty;

    public bool locked
    {
        get => _mailMergeField?.Locked ?? false;
        set
        {
            try
            {
                if (_mailMergeField != null)
                    _mailMergeField.Locked = value;
            }
            catch (Exception ex)
            {
                log.Error($"设置邮件合并域的锁定状态失败。", ex);
                throw new InvalidOperationException("设置邮件合并域的锁定状态失败。", ex);
            }
        }
    }
    #endregion

    #region 方法实现

    public void Select()
    {
        try
        {
            _mailMergeField?.Select();
        }
        catch (Exception ex)
        {
            log.Error("选择邮件合并域失败。", ex);
            throw new InvalidOperationException("选择邮件合并域失败。", ex);
        }
    }

    public void Copy()
    {
        try
        {
            _mailMergeField?.Copy();
        }
        catch (Exception ex)
        {
            log.Error("复制邮件合并域失败。", ex);
            throw new InvalidOperationException("复制邮件合并域失败。", ex);
        }
    }

    public void Cut()
    {
        try
        {
            _mailMergeField?.Cut();
        }
        catch (Exception ex)
        {
            log.Error("剪切邮件合并域失败。", ex);
            throw new InvalidOperationException("剪切邮件合并域失败。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _mailMergeField?.Delete();
        }
        catch (Exception ex)
        {
            log.Error("删除邮件合并域失败。", ex);
            throw new InvalidOperationException("删除邮件合并域失败。", ex);
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _mailMergeField != null)
        {
            Marshal.ReleaseComObject(_mailMergeField);
            _mailMergeField = null;
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