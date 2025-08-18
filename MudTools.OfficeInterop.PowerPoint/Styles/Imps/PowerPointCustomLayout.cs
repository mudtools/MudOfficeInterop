//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 自定义布局实现类
/// </summary>
internal class PowerPointCustomLayout : IPowerPointCustomLayout
{
    private readonly MsPowerPoint.CustomLayout _customLayout;
    private bool _disposedValue;
    private IPowerPointShapes _shapes;
    private IPowerPointHeadersFooters _headersFooters;
    private IPowerPointBackground _background;

    /// <summary>
    /// 获取或设置布局名称
    /// </summary>
    public string Name
    {
        get => _customLayout?.Name ?? string.Empty;
        set
        {
            if (_customLayout != null)
                _customLayout.Name = value ?? string.Empty;
        }
    }

    /// <summary>
    /// 获取形状集合
    /// </summary>
    public IPowerPointShapes Shapes
    {
        get
        {
            if (_shapes == null && _customLayout?.Shapes != null)
            {
                _shapes = new PowerPointShapes(_customLayout.Shapes);
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
            if (_headersFooters == null && _customLayout?.HeadersFooters != null)
            {
                _headersFooters = new PowerPointHeadersFooters(_customLayout.HeadersFooters);
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
            if (_background == null && _customLayout != null)
            {
                // 获取背景形状范围
                var backgroundRange = _customLayout.Shapes.Range(new object[] { 1 }); // 假设背景是第一个形状
                _background = new PowerPointBackground(backgroundRange);
            }
            return _background;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _customLayout?.Parent;

    /// <summary>
    /// 获取或设置是否包含主题
    /// </summary>
    public bool FollowMasterBackground
    {
        get => _customLayout?.FollowMasterBackground == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_customLayout != null)
                _customLayout.FollowMasterBackground = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }



    /// <summary>
    /// 获取布局索引
    /// </summary>
    public int Index => _customLayout?.Index ?? 0;


    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="customLayout">COM CustomLayout 对象</param>
    internal PowerPointCustomLayout(MsPowerPoint.CustomLayout customLayout)
    {
        _customLayout = customLayout;
        _disposedValue = false;
    }

    /// <summary>
    /// 应用到幻灯片
    /// </summary>
    /// <param name="slide">目标幻灯片</param>
    public void ApplyTo(IPowerPointSlide slide)
    {
        if (slide == null)
            throw new ArgumentNullException(nameof(slide));

        try
        {
            // 这需要将 IPowerPointSlide 转换为 COM Slide 对象
            throw new NotImplementedException("Applying custom layout is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to apply custom layout.", ex);
        }
    }

    /// <summary>
    /// 复制布局
    /// </summary>
    /// <returns>复制的布局</returns>
    public IPowerPointCustomLayout Duplicate()
    {
        try
        {
            var duplicatedLayout = _customLayout?.Duplicate();
            return duplicatedLayout != null ? new PowerPointCustomLayout(duplicatedLayout) : null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to duplicate custom layout.", ex);
        }
    }

    /// <summary>
    /// 删除布局
    /// </summary>
    public void Delete()
    {
        try
        {
            _customLayout?.Delete();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to delete custom layout.", ex);
        }
    }

    /// <summary>
    /// 重置布局
    /// </summary>
    public void Reset()
    {
        try
        {
            // 布局重置需要通过母版实现
            throw new NotImplementedException("Resetting custom layout is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to reset custom layout.", ex);
        }
    }

    /// <summary>
    /// 刷新布局显示
    /// </summary>
    public void Refresh()
    {
        try
        {
            _customLayout?.Shapes?.SelectAll();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to refresh custom layout.", ex);
        }
    }

    /// <summary>
    /// 获取布局的缩略图
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
    /// 导出布局为图片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="filterName">图片格式</param>
    public void Export(string fileName, string filterName = "PNG")
    {
        if (string.IsNullOrEmpty(fileName))
            throw new ArgumentException("File name cannot be null or empty.", nameof(fileName));

        try
        {
            // 布局导出需要通过幻灯片对象实现
            throw new NotImplementedException("Exporting custom layout is not implemented.");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to export custom layout to '{fileName}'.", ex);
        }
    }

    /// <summary>
    /// 获取布局信息
    /// </summary>
    /// <returns>布局信息字符串</returns>
    public string GetLayoutInfo()
    {
        try
        {
            return $"CustomLayout - Name: {Name}, Index: {Index}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get layout info.", ex);
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
