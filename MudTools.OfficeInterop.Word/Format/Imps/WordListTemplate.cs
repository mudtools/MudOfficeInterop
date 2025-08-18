//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档列表模板实现类
/// </summary>
internal class WordListTemplate : IWordListTemplate
{
    private readonly MsWord.ListTemplate _listTemplate;
    private bool _disposedValue;
    private IWordListLevels _listLevels;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _listTemplate.Parent;

    /// <summary>
    /// 获取列表级别
    /// </summary>
    public IWordListLevels ListLevels
    {
        get
        {
            if (_listLevels == null)
            {
                _listLevels = new WordListLevels(_listTemplate.ListLevels);
            }
            return _listLevels;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="listTemplate">COM ListTemplate 对象</param>
    internal WordListTemplate(MsWord.ListTemplate listTemplate)
    {
        _listTemplate = listTemplate ?? throw new ArgumentNullException(nameof(listTemplate));
        _disposedValue = false;
    }

    /// <summary>
    /// 删除列表模板
    /// </summary>
    public void Delete()
    {
        try
        {
            // ListTemplate 通常不需要显式删除
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete list template.", ex);
        }
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _listLevels?.Dispose();
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
