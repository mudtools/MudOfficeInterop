//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.PowerPoint;

namespace MudTools.OfficeInterop.PowerPoint.Imps;
/// <summary>
/// PowerPoint 页眉页脚实现类
/// </summary>
internal class PowerPointHeadersFooters : IPowerPointHeadersFooters
{
    private readonly MsPowerPoint.HeadersFooters _headersFooters;
    private bool _disposedValue;
    private IPowerPointHeaderFooter _slideNumber;
    private IPowerPointHeaderFooter _dateAndTime;
    private IPowerPointHeaderFooter _footer;
    private IPowerPointHeaderFooter _header;

    /// <summary>
    /// 获取幻灯片编号
    /// </summary>
    public IPowerPointHeaderFooter SlideNumber
    {
        get
        {
            if (_slideNumber == null && _headersFooters?.SlideNumber != null)
            {
                _slideNumber = new PowerPointHeaderFooter(_headersFooters.SlideNumber);
            }
            return _slideNumber;
        }
    }

    /// <summary>
    /// 获取日期和时间
    /// </summary>
    public IPowerPointHeaderFooter DateAndTime
    {
        get
        {
            if (_dateAndTime == null && _headersFooters?.DateAndTime != null)
            {
                _dateAndTime = new PowerPointHeaderFooter(_headersFooters.DateAndTime);
            }
            return _dateAndTime;
        }
    }

    /// <summary>
    /// 获取页脚
    /// </summary>
    public IPowerPointHeaderFooter Footer
    {
        get
        {
            if (_footer == null && _headersFooters?.Footer != null)
            {
                _footer = new PowerPointHeaderFooter(_headersFooters.Footer);
            }
            return _footer;
        }
    }

    /// <summary>
    /// 获取页眉
    /// </summary>
    public IPowerPointHeaderFooter Header
    {
        get
        {
            if (_header == null && _headersFooters?.Header != null)
            {
                _header = new PowerPointHeaderFooter(_headersFooters.Header);
            }
            return _header;
        }
    }

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object? Parent => _headersFooters?.Parent;

    /// <summary>
    /// 获取或设置是否显示幻灯片编号
    /// </summary>
    public bool SlideNumberVisible
    {
        get => _headersFooters?.SlideNumber?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters?.SlideNumber != null)
                _headersFooters.SlideNumber.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否显示日期和时间
    /// </summary>
    public bool DateAndTimeVisible
    {
        get => _headersFooters?.DateAndTime?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters?.DateAndTime != null)
                _headersFooters.DateAndTime.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否显示页脚
    /// </summary>
    public bool FooterVisible
    {
        get => _headersFooters?.Footer?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters?.Footer != null)
                _headersFooters.Footer.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否显示页眉
    /// </summary>
    public bool HeaderVisible
    {
        get => _headersFooters?.Header?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters?.Header != null)
                _headersFooters.Header.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置日期和时间格式
    /// </summary>
    public int DateTimeFormat
    {
        get => (int)_headersFooters?.DateAndTime?.Format;
        set
        {
            if (_headersFooters?.DateAndTime != null)
                _headersFooters.DateAndTime.Format = (PpDateTimeFormat)value;
        }
    }

    /// <summary>
    /// 获取或设置是否使用预设格式
    /// </summary>
    public bool UseDateTimeFormat
    {
        get => _headersFooters?.DateAndTime?.UseFormat == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters?.DateAndTime != null)
                _headersFooters.DateAndTime.UseFormat = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 获取或设置是否显示背景图形
    /// </summary>
    public bool BackgroundVisible
    {
        get => _headersFooters?.DisplayOnTitleSlide == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_headersFooters != null)
                _headersFooters.DisplayOnTitleSlide = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="headersFooters">COM HeadersFooters 对象</param>
    internal PowerPointHeadersFooters(MsPowerPoint.HeadersFooters headersFooters)
    {
        _headersFooters = headersFooters;
        _disposedValue = false;
    }


    /// <summary>
    /// 设置日期和时间文本
    /// </summary>
    /// <param name="dateTimeText">日期时间文本</param>
    public void SetDateTimeText(string dateTimeText)
    {
        try
        {
            if (_headersFooters?.DateAndTime != null)
            {
                _headersFooters.DateAndTime.Text = dateTimeText ?? string.Empty;
                _headersFooters.DateAndTime.UseFormat = MsCore.MsoTriState.msoFalse;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set date and time text.", ex);
        }
    }

    /// <summary>
    /// 设置页脚文本
    /// </summary>
    /// <param name="footerText">页脚文本</param>
    public void SetFooterText(string footerText)
    {
        try
        {
            if (_headersFooters?.Footer != null)
            {
                _headersFooters.Footer.Text = footerText ?? string.Empty;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set footer text.", ex);
        }
    }

    /// <summary>
    /// 设置页眉文本
    /// </summary>
    /// <param name="headerText">页眉文本</param>
    public void SetHeaderText(string headerText)
    {
        try
        {
            if (_headersFooters?.Header != null)
            {
                _headersFooters.Header.Text = headerText ?? string.Empty;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set header text.", ex);
        }
    }

    /// <summary>
    /// 设置幻灯片编号格式
    /// </summary>
    /// <param name="format">编号格式</param>
    public void SetSlideNumberFormat(int format)
    {
        try
        {
            if (_headersFooters?.SlideNumber != null)
            {
                // 幻灯片编号格式通常由幻灯片母版控制
                throw new NotImplementedException("Setting slide number format is not implemented.");
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to set slide number format.", ex);
        }
    }

    /// <summary>
    /// 获取页眉页脚信息
    /// </summary>
    /// <returns>页眉页脚信息字符串</returns>
    public string GetHeadersFootersInfo()
    {
        try
        {
            return $"HeadersFooters -  SlideNumber: {SlideNumberVisible}, DateTime: {DateAndTimeVisible}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to get headers and footers info.", ex);
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
            _slideNumber?.Dispose();
            _dateAndTime?.Dispose();
            _footer?.Dispose();
            _header?.Dispose();
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
