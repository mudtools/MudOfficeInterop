//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word 文档列表模板集合实现类
/// </summary>
internal class WordListTemplates : IWordListTemplates
{
    private readonly MsWord.ListTemplates _listTemplates;
    private readonly IWordDocument _document;
    private bool _disposedValue;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _listTemplates != null ? new WordApplication(_listTemplates.Application) : null;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _listTemplates?.Parent;

    /// <summary>
    /// 获取列表模板数量
    /// </summary>
    public int Count => _listTemplates.Count;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="listTemplates">COM ListTemplates 对象</param>
    /// <param name="document">关联的文档对象</param>
    internal WordListTemplates(MsWord.ListTemplates listTemplates, IWordDocument document)
    {
        _listTemplates = listTemplates ?? throw new ArgumentNullException(nameof(listTemplates));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取列表模板
    /// </summary>
    /// <param name="index">列表模板索引</param>
    /// <returns>列表模板对象</returns>
    public IWordListTemplate Item(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var listTemplate = _listTemplates[index];
            return new WordListTemplate(listTemplate);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get list template at index {index}.", ex);
        }
    }

    /// <summary>
    /// 添加列表模板
    /// </summary>
    /// <param name="outlineNumbered">是否大纲编号</param>
    /// <returns>列表模板对象</returns>
    public IWordListTemplate Add(bool outlineNumbered = false)
    {
        try
        {
            var listTemplate = _listTemplates.Add(outlineNumbered);
            return new WordListTemplate(listTemplate);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add list template.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>列表模板枚举器</returns>
    public IEnumerator<IWordListTemplate> GetEnumerator()
    {
        try
        {
            var listTemplates = new List<IWordListTemplate>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    listTemplates.Add(Item(i));
                }
                catch
                {
                    // 忽略获取失败的列表模板
                    continue;
                }
            }
            return listTemplates.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate list templates.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
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