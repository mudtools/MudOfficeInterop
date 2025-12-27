//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 演示文稿实现类
/// </summary>
internal class PowerPointPresentation : IPowerPointPresentation
{
    internal MsPowerPoint.Presentation _presentation;
    private bool _disposedValue;
    private IPowerPointSlides _slides;

    /// <summary>
    /// 获取或设置演示文稿名称
    /// </summary>
    public string Name
    {
        get => _presentation?.Name ?? string.Empty;
    }

    /// <summary>
    /// 获取演示文稿完整路径
    /// </summary>
    public string FullName => _presentation?.FullName ?? string.Empty;

    /// <summary>
    /// 获取演示文稿路径
    /// </summary>
    public string Path => _presentation?.Path ?? string.Empty;

    /// <summary>
    /// 获取幻灯片数量
    /// </summary>
    public int SlideCount => _presentation?.Slides?.Count ?? 0;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _presentation?.Parent;

    /// <summary>
    /// 获取幻灯片集合
    /// </summary>
    public IPowerPointSlides Slides
    {
        get
        {
            if (_slides == null && _presentation?.Slides != null)
            {
                _slides = new PowerPointSlides(_presentation.Slides);
            }
            return _slides;
        }
    }



    /// <summary>
    /// 获取或设置演示文稿是否已修改
    /// </summary>
    public bool Saved
    {
        get => _presentation?.Saved == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_presentation != null)
                _presentation.Saved = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置演示文稿是否只读
    /// </summary>
    public bool ReadOnly => _presentation?.ReadOnly == MsCore.MsoTriState.msoTrue;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="presentation">COM Presentation 对象</param>
    internal PowerPointPresentation(MsPowerPoint.Presentation presentation)
    {
        _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
        _disposedValue = false;
    }

    /// <summary>
    /// 保存演示文稿
    /// </summary>
    /// <param name="fileName">文件名（可选）</param>
    /// <param name="fileFormat">文件格式（可选）</param>
    public void Save(string fileName = null, PpSaveAsFileType fileFormat = PpSaveAsFileType.ppSaveAsDefault)
    {
        try
        {
            if (string.IsNullOrEmpty(fileName))
            {
                _presentation?.Save();
            }
            else
            {
                _presentation?.SaveAs(fileName, (MsPowerPoint.PpSaveAsFileType)fileFormat, MsCore.MsoTriState.msoTrue);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save presentation{(string.IsNullOrEmpty(fileName) ? "" : $" as '{fileName}'")}.", ex);
        }
    }

    /// <summary>
    /// 另存为演示文稿
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="fileFormat">文件格式</param>
    /// <param name="embedTrueTypeFonts">是否嵌入TrueType字体</param>
    public void SaveAs(string fileName, PpSaveAsFileType fileFormat = PpSaveAsFileType.ppSaveAsDefault, bool embedTrueTypeFonts = false)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            var embedFonts = embedTrueTypeFonts ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            _presentation?.SaveAs(fileName, (MsPowerPoint.PpSaveAsFileType)fileFormat, embedFonts);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save presentation as '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 关闭演示文稿
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    public void Close(bool saveChanges = true)
    {
        try
        {
            var saveOption = saveChanges ? MsPowerPoint.PpSaveAsFileType.ppSaveAsDefault : MsPowerPoint.PpSaveAsFileType.ppSaveAsDefault;
            _presentation?.Close();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to close presentation.", ex);
        }
    }

    /// <summary>
    /// 导出演示文稿
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="exportFormat">导出格式</param>
    /// <param name="scaleWidth">缩放宽度</param>
    /// <param name="scaleHeight">缩放高度</param>
    public void Export(string fileName, string exportFormat = "PNG", int scaleWidth = 0, int scaleHeight = 0)
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            _presentation?.Export(fileName, exportFormat, scaleWidth, scaleHeight);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export presentation to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 保护演示文稿
    /// </summary>
    /// <param name="password">密码</param>
    /// <param name="writePassword">写入密码</param>
    /// <param name="readOnlyRecommended">是否推荐只读打开</param>
    public void Protect(string password, string writePassword = null, bool readOnlyRecommended = false)
    {
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty.", nameof(password));

        try
        {
            var readOnly = readOnlyRecommended ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
            var writePwd = string.IsNullOrEmpty(writePassword) ? null : (object)writePassword;
            _presentation.Password = password;
            if (!string.IsNullOrEmpty(writePassword))
            {
                _presentation.WritePassword = writePassword;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to protect presentation.", ex);
        }
    }

    /// <summary>
    /// 取消保护演示文稿
    /// </summary>
    /// <param name="password">密码</param>
    public void Unprotect(string password)
    {
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password cannot be null or empty.", nameof(password));

        try
        {
            _presentation.Password = string.Empty;
            _presentation.WritePassword = string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to unprotect presentation.", ex);
        }
    }

    /// <summary>
    /// 添加幻灯片到演示文稿
    /// </summary>
    /// <param name="layout">幻灯片布局</param>
    /// <param name="position">插入位置</param>
    /// <returns>新添加的幻灯片</returns>
    public IPowerPointSlide AddSlide(PpSlideLayout layout = PpSlideLayout.ppLayoutText, int position = -1)
    {
        try
        {
            MsPowerPoint.Slide slide;
            if (position > 0 && position <= (_presentation?.Slides?.Count ?? 0) + 1)
            {
                slide = _presentation?.Slides?.Add(position, (MsPowerPoint.PpSlideLayout)layout);
            }
            else
            {
                slide = _presentation?.Slides?.Add((_presentation?.Slides?.Count ?? 0) + 1, (MsPowerPoint.PpSlideLayout)layout);
            }
            return slide != null ? new PowerPointSlide(slide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to add slide to presentation.", ex);
        }
    }

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    public void RemoveSlide(int index)
    {
        if (index < 1 || index > (_presentation?.Slides?.Count ?? 0))
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {_presentation?.Slides?.Count ?? 0}.");

        try
        {
            _presentation?.Slides[index]?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to remove slide at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据索引获取幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    /// <returns>幻灯片对象</returns>
    public IPowerPointSlide GetSlide(int index)
    {
        if (index < 1 || index > (_presentation?.Slides?.Count ?? 0))
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {_presentation?.Slides?.Count ?? 0}.");

        try
        {
            var slide = _presentation?.Slides[index];
            return slide != null ? new PowerPointSlide(slide) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get slide at index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取所有幻灯片
    /// </summary>
    /// <returns>幻灯片列表</returns>
    public IEnumerable<IPowerPointSlide> GetAllSlides()
    {
        var slides = new List<IPowerPointSlide>();
        try
        {
            if (_presentation?.Slides != null)
            {
                for (int i = 1; i <= _presentation.Slides.Count; i++)
                {
                    try
                    {
                        var slide = _presentation.Slides[i];
                        slides.Add(new PowerPointSlide(slide));
                    }
                    catch
                    {
                        continue;
                    }
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
    /// 获取演示文稿信息
    /// </summary>
    /// <returns>演示文稿信息字符串</returns>
    public string GetPresentationInfo()
    {
        try
        {
            return $"Presentation: {Name}, Slides: {SlideCount}, Path: {Path}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get presentation info.", ex);
        }
    }

    /// <summary>
    /// 替换演示文稿中的文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    public int ReplaceText(string findText, string replaceText)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        var totalCount = 0;
        try
        {
            if (_presentation?.Slides != null)
            {
                foreach (MsPowerPoint.Slide slide in _presentation.Slides)
                {
                    try
                    {
                        var slideCount = ReplaceTextInSlide(slide, findText, replaceText ?? string.Empty);
                        totalCount += slideCount;
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in presentation.", ex);
        }
        return totalCount;
    }

    /// <summary>
    /// 在单个幻灯片中替换文本
    /// </summary>
    /// <param name="slide">幻灯片</param>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    private int ReplaceTextInSlide(MsPowerPoint.Slide slide, string findText, string replaceText)
    {
        var count = 0;
        try
        {
            if (slide?.Shapes != null)
            {
                foreach (MsPowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.HasTextFrame == MsCore.MsoTriState.msoTrue &&
                            shape.TextFrame?.TextRange != null)
                        {
                            var replaceCount = shape.TextFrame.TextRange.Replace(findText, replaceText);
                            count += replaceCount.Count;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }
        catch
        {
            // 忽略单个幻灯片替换失败
        }
        return count;
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
            _slides?.Dispose();
            _presentation = null;
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