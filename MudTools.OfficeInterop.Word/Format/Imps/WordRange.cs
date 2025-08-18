//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 范围实现类
/// </summary>
internal class WordRange : IWordRange
{
    private readonly MsWord.Range _range;
    private bool _disposedValue;

    public string Text
    {
        get => _range.Text;
        set => _range.Text = value;
    }

    public int Start
    {
        get => _range.Start;
        set => _range.SetRange(value, _range.End);
    }

    public int End
    {
        get => _range.End;
        set => _range.SetRange(_range.Start, value);
    }

    public int Length => _range.End - _range.Start;

    public object Parent => _range.Parent;

    internal WordRange(MsWord.Range range)
    {
        _range = range ?? throw new ArgumentNullException(nameof(range));
        _disposedValue = false;
    }

    public void Copy()
    {
        try
        {
            _range.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy range.", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _range.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete range.", ex);
        }
    }

    public void Select()
    {
        try
        {
            _range.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select range.", ex);
        }
    }

    public void SetRange(int start, int end)
    {
        try
        {
            _range.SetRange(start, end);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set range.", ex);
        }
    }

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
