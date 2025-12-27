//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 母版实现类
/// </summary>
internal class PowerPointMaster : IPowerPointMaster
{
    private readonly MsPowerPoint.Master _master;
    private bool _disposedValue;
    private IPowerPointShapes _shapes;
    private IPowerPointHeadersFooters _headersFooters;
    private IPowerPointBackground _background;

    /// <summary>
    /// 获取母版名称
    /// </summary>
    public string Name
    {
        get => _master?.Name ?? string.Empty;
        set
        {
            if (_master != null)
                _master.Name = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取形状集合
    /// </summary>
    public IPowerPointShapes Shapes
    {
        get
        {
            if (_shapes == null && _master?.Shapes != null)
            {
                _shapes = new PowerPointShapes(_master.Shapes);
            }
            return _shapes;
        }
    }

    /// <summary>
    /// 获取页眉页脚
    /// </summary>
    public IPowerPointHeadersFooters HeadersFooters
    {
        get
        {
            if (_headersFooters == null && _master?.HeadersFooters != null)
            {
                _headersFooters = new PowerPointHeadersFooters(_master.HeadersFooters);
            }
            return _headersFooters;
        }
    }

    /// <summary>
    /// 获取背景
    /// </summary>
    public IPowerPointBackground Background
    {
        get
        {
            if (_background == null && _master != null)
            {
                // 获取背景形状范围
                var backgroundRange = _master.Shapes.Range(new object[] { 1 }); // 假设背景是第一个形状
                _background = new PowerPointBackground(backgroundRange);
            }
            return _background;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _master?.Parent;

    /// <summary>
    /// 获取自定义布局集合
    /// </summary>
    public IEnumerable<IPowerPointCustomLayout> CustomLayouts
    {
        get
        {
            var layouts = new List<IPowerPointCustomLayout>();
            try
            {
                if (_master?.CustomLayouts != null)
                {
                    foreach (MsPowerPoint.CustomLayout layout in _master.CustomLayouts)
                    {
                        layouts.Add(new PowerPointCustomLayout(layout));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to enumerate custom layouts.", ex);
            }
            return layouts;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="master">COM Master 对象</param>
    internal PowerPointMaster(MsPowerPoint.Master master)
    {
        _master = master;
        _disposedValue = false;
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
            // 设计模板应用需要通过演示文稿对象实现
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
            _master?.ApplyTheme(themeName);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to apply theme '{themeName}'.", ex);
        }
    }

    /// <summary>
    /// 复制母版
    /// </summary>
    /// <returns>复制的母版</returns>
    public IPowerPointMaster Duplicate()
    {
        try
        {
            // PowerPoint 不直接支持复制母版
            throw new NotImplementedException("Duplicating master is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate master.", ex);
        }
    }

    /// <summary>
    /// 删除母版
    /// </summary>
    public void Delete()
    {
        try
        {
            _master?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete master.", ex);
        }
    }

    /// <summary>
    /// 重置母版
    /// </summary>
    public void Reset()
    {
        try
        {
            // 母版重置需要通过演示文稿对象实现
            throw new NotImplementedException("Resetting master is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset master.", ex);
        }
    }

    /// <summary>
    /// 刷新母版显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            _master?.Shapes?.SelectAll();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh master.", ex);
        }
    }

    /// <summary>
    /// 获取母版的缩略图
    /// </summary>
    /// <returns>缩略图数据</returns>
    public byte[] GetThumbnail()
    {
        try
        {
            // 缩略图获取需要更复杂的实现
            throw new NotImplementedException("Getting thumbnail is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get thumbnail.", ex);
        }
    }

    /// <summary>
    /// 导出母版为图片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="filterName">图片格式</param>
    public void Export(string fileName, string filterName = "PNG")
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            // 母版导出需要通过幻灯片对象实现
            throw new NotImplementedException("Exporting master is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export master to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 获取母版的所有文本内容
    /// </summary>
    /// <returns>文本内容列表</returns>
    public IEnumerable<string> GetAllText()
    {
        var texts = new List<string>();
        try
        {
            if (_master?.Shapes != null)
            {
                foreach (MsPowerPoint.Shape shape in _master.Shapes)
                {
                    try
                    {
                        if (shape.HasTextFrame == MsCore.MsoTriState.msoTrue && shape.TextFrame?.TextRange != null)
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
                        continue;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get all text from master.", ex);
        }
        return texts;
    }

    /// <summary>
    /// 替换母版中的文本
    /// </summary>
    /// <param name="findText">查找文本</param>
    /// <param name="replaceText">替换文本</param>
    /// <returns>替换次数</returns>
    public int ReplaceText(string findText, string replaceText)
    {
        if (string.IsNullOrEmpty(findText))
            return 0;

        var count = 0;
        try
        {
            if (_master?.Shapes != null)
            {
                foreach (MsPowerPoint.Shape shape in _master.Shapes)
                {
                    try
                    {
                        if (shape.HasTextFrame == MsCore.MsoTriState.msoTrue && shape.TextFrame?.TextRange != null)
                        {
                            var replaceCount = shape.TextFrame.TextRange.Replace(findText, replaceText ?? string.Empty);
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
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to replace text in master.", ex);
        }
        return count;
    }

    /// <summary>
    /// 获取母版信息
    /// </summary>
    /// <returns>母版信息字符串</returns>
    public string GetMasterInfo()
    {
        try
        {
            return $"Master - Name: {Name}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get master info.", ex);
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
