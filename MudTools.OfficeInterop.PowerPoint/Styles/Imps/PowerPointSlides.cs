//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 幻灯片集合实现类（基本功能）
/// </summary>
internal class PowerPointSlides : IPowerPointSlides
{
    private readonly MsPowerPoint.Slides _slides;
    private bool _disposedValue;

    /// <summary>
    /// 获取幻灯片数量
    /// </summary>
    public int Count => _slides?.Count ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _slides?.Parent;

    /// <summary>
    /// 根据索引获取幻灯片（从1开始）
    /// </summary>
    public IPowerPointSlide this[int index] => ByIndex(index);

    /// <summary>
    /// 根据名称获取幻灯片
    /// </summary>
    public IPowerPointSlide this[string name] => ByName(name);

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="slides">COM Slides 对象</param>
    internal PowerPointSlides(MsPowerPoint.Slides slides)
    {
        _slides = slides ?? throw new ArgumentNullException(nameof(slides));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取幻灯片（从1开始）
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    /// <returns>幻灯片对象</returns>
    public IPowerPointSlide ByIndex(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var slide = _slides[index];
            return slide != null ? new PowerPointSlide(slide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get slide at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据名称获取幻灯片
    /// </summary>
    /// <param name="name">幻灯片名称</param>
    /// <returns>幻灯片对象</returns>
    public IPowerPointSlide ByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Slide name cannot be null or empty.", nameof(name));

        try
        {
            var slide = _slides[name];
            return slide != null ? new PowerPointSlide(slide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get slide with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 添加新幻灯片
    /// </summary>
    /// <param name="layout">幻灯片布局</param>
    /// <param name="position">插入位置（-1表示末尾）</param>
    /// <returns>新添加的幻灯片</returns>
    public IPowerPointSlide Add(PpSlideLayout layout = PpSlideLayout.ppLayoutText, int position = -1)
    {
        try
        {
            MsPowerPoint.Slide slide;
            if (position > 0 && position <= Count + 1)
            {
                slide = _slides?.Add(position, (MsPowerPoint.PpSlideLayout)layout);
            }
            else
            {
                slide = _slides?.Add(Count + 1, (MsPowerPoint.PpSlideLayout)layout);
            }
            return slide != null ? new PowerPointSlide(slide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add slide.", ex);
        }
    }

    /// <summary>
    /// 从文件插入幻灯片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="position">插入位置</param>
    /// <param name="slideRange">幻灯片范围</param>
    /// <returns>插入的幻灯片</returns>
    //public IPowerPointSlide InsertFromFile(string fileName, int position, int slideRange = -1)
    //{
    //    if (string.IsNullOrEmpty(fileName))
    //        throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

    //    if (!System.IO.File.Exists(fileName))
    //        throw new System.IO.FileNotFoundException("Slide file not found.", fileName);

    //    if (position < 1 || position > Count + 1)
    //        throw new ArgumentOutOfRangeException(nameof(position), $"Position must be between 1 and {Count + 1}.");

    //    try
    //    {
    //        var slide = _slides?.InsertFromFile(fileName, position, 1, slideRange > 0 ? slideRange : 1);
    //        return slide != null ? new PowerPointSlide(slide) : null;
    //    }
    //    catch (Exception ex)
    //    {
    //        throw new InvalidOperationException($"Failed to insert slide from file '{fileName}'.", ex);
    //    }
    //}

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    public void Delete(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            _slides[index]?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete slide at index {index}.", ex);
        }
    }

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="name">幻灯片名称</param>
    public void Delete(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Slide name cannot be null or empty.", nameof(name));

        try
        {
            _slides[name]?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete slide with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 移动幻灯片
    /// </summary>
    /// <param name="fromIndex">源索引</param>
    /// <param name="toIndex">目标索引</param>
    public void Move(int fromIndex, int toIndex)
    {
        if (fromIndex < 1 || fromIndex > Count)
            throw new ArgumentOutOfRangeException(nameof(fromIndex), $"From index must be between 1 and {Count}.");
        if (toIndex < 1 || toIndex > Count)
            throw new ArgumentOutOfRangeException(nameof(toIndex), $"To index must be between 1 and {Count}.");

        try
        {
            _slides[fromIndex]?.MoveTo(toIndex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to move slide from index {fromIndex} to {toIndex}.", ex);
        }
    }

    /// <summary>
    /// 复制幻灯片
    /// </summary>
    /// <param name="sourceIndex">源索引</param>
    /// <param name="targetIndex">目标索引</param>
    /// <returns>复制的幻灯片</returns>
    public IPowerPointSlide Copy(int sourceIndex, int targetIndex = -1)
    {
        if (sourceIndex < 1 || sourceIndex > Count)
            throw new ArgumentOutOfRangeException(nameof(sourceIndex), $"Source index must be between 1 and {Count}.");

        try
        {
            var sourceSlide = _slides[sourceIndex];
            if (sourceSlide == null)
                return null;

            MsPowerPoint.Slide copiedSlide;
            if (targetIndex > 0 && targetIndex <= Count + 1)
            {
                copiedSlide = (MsPowerPoint.Slide)sourceSlide.Duplicate();
                copiedSlide.MoveTo(targetIndex);
            }
            else
            {
                copiedSlide = (MsPowerPoint.Slide)sourceSlide.Duplicate();
            }
            return copiedSlide != null ? new PowerPointSlide(copiedSlide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to copy slide from index {sourceIndex}.", ex);
        }
    }

    /// <summary>
    /// 获取所有幻灯片
    /// </summary>
    /// <returns>幻灯片列表</returns>
    public IEnumerable<IPowerPointSlide> GetAll()
    {
        var slides = new List<IPowerPointSlide>();
        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    slides.Add(ByIndex(i));
                }
                catch
                {
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all slides.", ex);
        }
        return slides;
    }

    /// <summary>
    /// 根据幻灯片编号查找幻灯片
    /// </summary>
    /// <param name="slideNumber">幻灯片编号</param>
    /// <returns>幻灯片对象</returns>
    public IPowerPointSlide FindByNumber(int slideNumber)
    {
        if (slideNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(slideNumber), "Slide number must be greater than 0.");

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    var slide = ByIndex(i);
                    if (slide?.SlideNumber == slideNumber)
                    {
                        return slide;
                    }
                }
                catch
                {
                    continue;
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to find slide by number {slideNumber}.", ex);
        }
    }

    /// <summary>
    /// 获取第一张幻灯片
    /// </summary>
    /// <returns>第一张幻灯片对象</returns>
    public IPowerPointSlide First()
    {
        try
        {
            return Count > 0 ? ByIndex(1) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get first slide.", ex);
        }
    }

    /// <summary>
    /// 获取最后一张幻灯片
    /// </summary>
    /// <returns>最后一张幻灯片对象</returns>
    public IPowerPointSlide Last()
    {
        try
        {
            return Count > 0 ? ByIndex(Count) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get last slide.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>幻灯片枚举器</returns>
    public IEnumerator<IPowerPointSlide> GetEnumerator()
    {
        return GetAll().GetEnumerator();
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
