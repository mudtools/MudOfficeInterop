//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;

/// <summary>
/// PowerPoint 幻灯片实现类
/// </summary>
internal class PowerPointSlide : IPowerPointSlide
{
    internal readonly MsPowerPoint.Slide _slide;
    private bool _disposedValue;
    private IPowerPointShapes? _shapes;
    private IPowerPointHeadersFooters? _headersFooters;
    private IPowerPointShapeRange? _background;
    private IPowerPointSlideShowTransition? _slideShowTransition;
    private IPowerPointTimeLine? _timeLine;
    private IPowerPointTags? _tags;

    /// <summary>
    /// 获取幻灯片名称
    /// </summary>
    public string Name
    {
        get => _slide.Name;
        set => _slide.Name = value;
    }

    /// <summary>
    /// 获取幻灯片索引
    /// </summary>
    public int Index => _slide.SlideIndex;

    /// <summary>
    /// 获取幻灯片布局
    /// </summary>
    public PpSlideLayout Layout
    {
        get => (PpSlideLayout)_slide.Layout;
        set => _slide.Layout = (MsPowerPoint.PpSlideLayout)value;
    }

    /// <summary>
    /// 获取幻灯片标题
    /// </summary>
    public string Title
    {
        get
        {
            try
            {
                var titleShape = _slide.Shapes.Title;
                return titleShape?.TextFrame?.TextRange?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _slide.Parent;


    /// <summary>
    /// 获取幻灯片的形状集合
    /// </summary>
    public IPowerPointShapes? Shapes
    {
        get
        {
            _shapes ??= new PowerPointShapes(_slide.Shapes);
            return _shapes;
        }
    }

    /// <summary>
    /// 获取幻灯片的页眉页脚
    /// </summary>
    public IPowerPointHeadersFooters? HeadersFooters
    {
        get
        {
            _headersFooters ??= new PowerPointHeadersFooters(_slide.HeadersFooters);
            return _headersFooters;
        }
    }

    /// <summary>
    /// 获取幻灯片的背景
    /// </summary>
    public IPowerPointShapeRange? Background
    {
        get
        {
            _background ??= new PowerPointShapeRange(_slide.Background);
            return _background;
        }
    }

    /// <summary>
    /// 获取幻灯片的母版
    /// </summary>
    public IPowerPointMaster? Master
    {
        get
        {
            try
            {
                var master = _slide.Master;
                return master != null ? new PowerPointMaster(master) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 获取幻灯片的幻灯片显示
    /// </summary>
    public IPowerPointSlideShowTransition? SlideShowTransition
    {
        get
        {
            _slideShowTransition ??= new PowerPointSlideShowTransition(_slide.SlideShowTransition);
            return _slideShowTransition;
        }
    }

    /// <summary>
    /// 获取幻灯片的动画设置
    /// </summary>
    public IPowerPointTimeLine? TimeLine
    {
        get
        {
            _timeLine ??= new PowerPointTimeLine(_slide.TimeLine);
            return _timeLine;
        }
    }

    /// <summary>
    /// 获取幻灯片的超链接集合
    /// </summary>
    public IEnumerable<IPowerPointHyperlink> Hyperlinks
    {
        get
        {
            var hyperlinks = new List<IPowerPointHyperlink>();
            try
            {
                foreach (MsPowerPoint.Hyperlink hyperlink in _slide.Hyperlinks)
                {
                    hyperlinks.Add(new PowerPointHyperlink(hyperlink));
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to enumerate hyperlinks.", ex);
            }
            return hyperlinks;
        }
    }

    /// <summary>
    /// 获取幻灯片的标签集合
    /// </summary>
    public IPowerPointTags Tags
    {
        get
        {
            _tags ??= new PowerPointTags(_slide.Tags);
            return _tags;
        }
    }

    /// <summary>
    /// 获取幻灯片的自定义布局
    /// </summary>
    public IPowerPointCustomLayout CustomLayout
    {
        get
        {
            try
            {
                var customLayout = _slide.CustomLayout;
                return customLayout != null ? new PowerPointCustomLayout(customLayout) : null;
            }
            catch
            {
                return null;
            }
        }
        set
        {
            // 设置自定义布局需要更复杂的实现
            throw new NotImplementedException("Setting custom layout is not implemented.");
        }
    }


    /// <summary>
    /// 获取幻灯片的幻灯片ID
    /// </summary>
    public int SlideID => _slide.SlideID;


    public int SlideNumber => _slide.SlideNumber;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="slide">COM Slide 对象</param>
    internal PowerPointSlide(MsPowerPoint.Slide slide)
    {
        _slide = slide ?? throw new ArgumentNullException(nameof(slide));
        _disposedValue = false;
    }

    /// <summary>
    /// 激活幻灯片
    /// </summary>
    public void Select()
    {
        try
        {
            _slide.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to select slide.", ex);
        }
    }

    /// <summary>
    /// 复制幻灯片
    /// </summary>
    public void Copy()
    {
        try
        {
            _slide.Copy();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy slide.", ex);
        }
    }

    /// <summary>
    /// 剪切幻灯片
    /// </summary>
    public void Cut()
    {
        try
        {
            _slide.Cut();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to cut slide.", ex);
        }
    }

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    public void Delete()
    {
        try
        {
            _slide.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete slide.", ex);
        }
    }

    /// <summary>
    /// 移动幻灯片到指定位置
    /// </summary>
    /// <param name="toPosition">目标位置</param>
    public void MoveTo(int toPosition)
    {
        if (toPosition < 1)
            throw new ArgumentOutOfRangeException(nameof(toPosition), "Position must be greater than 0.");

        try
        {
            _slide.MoveTo(toPosition);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to move slide to position {toPosition}.", ex);
        }
    }

    /// <summary>
    /// 应用设计模板
    /// </summary>
    /// <param name="designName">设计模板名称</param>
    public void ApplyDesign(string designName)
    {
        if (string.IsNullOrEmpty(designName))
            throw new ArgumentException("Design name cannot be null or empty.", nameof(designName));

        try
        {
            // 需要根据具体实现来查找和应用设计模板
            throw new NotImplementedException($"Applying design '{designName}' is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to apply design '{designName}'.", ex);
        }
    }

    /// <summary>
    /// 应用主题
    /// </summary>
    /// <param name="themeName">主题名称</param>
    public void ApplyTheme(string themeName)
    {
        if (string.IsNullOrEmpty(themeName))
            throw new ArgumentException("Theme name cannot be null or empty.", nameof(themeName));

        try
        {
            _slide.ApplyTheme(themeName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to apply theme '{themeName}'.", ex);
        }
    }

    /// <summary>
    /// 导出幻灯片为图片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="filterName">图片格式</param>
    /// <param name="scaleWidth">宽度缩放</param>
    /// <param name="scaleHeight">高度缩放</param>
    public void Export(string fileName, string filterName = "PNG", int scaleWidth = 0, int scaleHeight = 0)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            _slide.Export(fileName, filterName, scaleWidth, scaleHeight);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export slide to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 获取幻灯片的缩略图
    /// </summary>
    /// <returns>缩略图数据</returns>
    public byte[] GetThumbnail()
    {
        try
        {
            // 这需要更复杂的实现来生成缩略图
            throw new NotImplementedException("Getting thumbnail is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get thumbnail.", ex);
        }
    }


    /// <summary>
    /// 获取幻灯片的所有文本内容
    /// </summary>
    /// <returns>文本内容列表</returns>
    public IEnumerable<string> GetAllText()
    {
        var texts = new List<string>();
        try
        {
            foreach (MsPowerPoint.Shape shape in _slide.Shapes)
            {
                try
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        var text = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            texts.Add(text);
                        }
                    }
                }
                catch
                {
                    // 忽略获取单个形状文本失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all text from slide.", ex);
        }
        return texts;
    }

    /// <summary>
    /// 获取指定占位符
    /// </summary>
    /// <param name="placeholderIndex">占位符索引</param>
    /// <returns>形状对象</returns>
    public IPowerPointShape GetPlaceholder(int placeholderIndex)
    {
        try
        {
            var placeholder = _slide.Shapes.Placeholders[placeholderIndex];
            return placeholder != null ? new PowerPointShape(placeholder) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get placeholder at index {placeholderIndex}.", ex);
        }
    }

    /// <summary>
    /// 获取所有占位符
    /// </summary>
    /// <returns>形状对象列表</returns>
    public IEnumerable<IPowerPointShape> GetPlaceholders()
    {
        var placeholders = new List<IPowerPointShape>();
        try
        {
            var placeholdersCollection = _slide.Shapes.Placeholders;
            for (int i = 1; i <= placeholdersCollection.Count; i++)
            {
                try
                {
                    var placeholder = placeholdersCollection[i];
                    placeholders.Add(new PowerPointShape(placeholder));
                }
                catch
                {
                    // 忽略获取单个占位符失败的情况
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get placeholders.", ex);
        }
        return placeholders;
    }

    /// <summary>
    /// 刷新幻灯片显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            // PowerPoint 中没有直接的刷新方法，这里模拟刷新
            _slide.Select();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh slide.", ex);
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
            _shapes?.Dispose();
            _headersFooters?.Dispose();
            _background?.Dispose();
            _slideShowTransition?.Dispose();
            _timeLine?.Dispose();
            _tags?.Dispose();
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
