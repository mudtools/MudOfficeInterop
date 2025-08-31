//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档范围集合实现类
/// </summary>
internal class WordStoryRanges : IWordStoryRanges
{
    private readonly MsWord.StoryRanges _storyRanges;
    private bool _disposedValue;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _storyRanges != null ? new WordApplication(_storyRanges.Application) : null;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _storyRanges?.Parent;

    /// <summary>
    /// 获取范围数量
    /// </summary>
    public int Count => _storyRanges.Count;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="storyRanges">COM StoryRanges 对象</param>
    internal WordStoryRanges(MsWord.StoryRanges storyRanges)
    {
        _storyRanges = storyRanges ?? throw new ArgumentNullException(nameof(storyRanges));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取范围
    /// </summary>
    /// <param name="index">范围索引</param>
    /// <returns>范围对象</returns>
    public IWordRange Item(WdStoryType index)
    {
        try
        {
            var range = _storyRanges[(MsWord.WdStoryType)index];
            return new WordRange(range);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get story range at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据类型获取范围
    /// </summary>
    /// <param name="index">范围类型</param>
    /// <returns>范围对象</returns>
    public IWordRange this[WdStoryType index] => Item(index);

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>范围枚举器</returns>
    public IEnumerator<IWordRange> GetEnumerator()
    {
        try
        {
            var ranges = new List<IWordRange>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    ranges.Add(Item((WdStoryType)i));
                }
                catch
                {
                    // 忽略获取失败的范围
                    continue;
                }
            }
            return ranges.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate story ranges.", ex);
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

