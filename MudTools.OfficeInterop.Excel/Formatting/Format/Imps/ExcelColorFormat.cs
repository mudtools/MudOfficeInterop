//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
internal class ExcelColorFormat : IExcelColorFormat
{
    internal MsExcel.ColorFormat? _colorFormat;
    private bool _disposedValue;

    public object Parent => _colorFormat.Parent;

    public MsoColorType? Type
    {
        get => _colorFormat?.Type.EnumConvert(MsoColorType.msoColorTypeMixed);
    }

    public MsoThemeColorIndex? ObjectThemeColor
    {
        get => _colorFormat?.ObjectThemeColor.EnumConvert(MsoThemeColorIndex.msoThemeColorMixed);
        set
        {
            if (_colorFormat != null)
                _colorFormat.ObjectThemeColor = value.EnumConvert(MsCore.MsoThemeColorIndex.msoThemeColorMixed);
        }
    }

    public int? RGB
    {
        get => _colorFormat?.RGB;
        set
        {
            if (_colorFormat != null && value != null)
                _colorFormat.RGB = value.Value;
        }
    }

    public float? Brightness
    {
        get => _colorFormat?.Brightness;
        set
        {
            if (_colorFormat != null && value != null)
                _colorFormat.Brightness = value.Value;
        }
    }

    public float? TintAndShade
    {
        get => _colorFormat?.TintAndShade;
        set
        {
            if (_colorFormat != null && value != null)
                _colorFormat.TintAndShade = value.Value;
        }
    }


    public IExcelApplication Application => new ExcelApplication(_colorFormat.Application as MsExcel.Application);


    internal ExcelColorFormat(MsExcel.ColorFormat colorFormat)
    {
        _colorFormat = colorFormat ?? throw new ArgumentNullException(nameof(colorFormat));
        _disposedValue = false;
    }

    public string ToHexString()
    {
        try
        {
            var rgb = _colorFormat.RGB;
            return $"#{rgb:X6}";
        }
        catch (COMException)
        {
            return "#000000";
        }
    }

    public void GetHSL(out double hue, out double saturation, out double lightness)
    {
        try
        {
            var rgb = _colorFormat.RGB;
            RGBToHSL(rgb, out hue, out saturation, out lightness);
        }
        catch (COMException)
        {
            hue = 0;
            saturation = 0;
            lightness = 0;
        }
    }

    public void SetHSL(double hue, double saturation, double lightness)
    {
        try
        {
            var rgb = HSLToRGB(hue, saturation, lightness);
            _colorFormat.RGB = rgb;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法设置HSL颜色值。", ex);
        }
    }

    public string GetColorName()
    {
        try
        {
            var rgb = _colorFormat.RGB;
            return GetStandardColorName(rgb);
        }
        catch (COMException)
        {
            return "Unknown";
        }
    }

    public int GetContrastColor()
    {
        try
        {
            var rgb = _colorFormat.RGB;
            return GetContrastingColor(rgb);
        }
        catch (COMException)
        {
            return 0xFFFFFF; // 默认白色
        }
    }

    public int BlendColors(int color1, int color2, double ratio)
    {
        if (ratio < 0 || ratio > 1)
            throw new ArgumentOutOfRangeException(nameof(ratio));

        try
        {
            var r1 = (color1 >> 16) & 0xFF;
            var g1 = (color1 >> 8) & 0xFF;
            var b1 = color1 & 0xFF;

            var r2 = (color2 >> 16) & 0xFF;
            var g2 = (color2 >> 8) & 0xFF;
            var b2 = color2 & 0xFF;

            var r = (int)(r1 * (1 - ratio) + r2 * ratio);
            var g = (int)(g1 * (1 - ratio) + g2 * ratio);
            var b = (int)(b1 * (1 - ratio) + b2 * ratio);

            return (r << 16) | (g << 8) | b;
        }
        catch (COMException)
        {
            return 0x000000;
        }
    }

    public double GetLuminance()
    {
        try
        {
            var rgb = _colorFormat.RGB;
            var r = (rgb >> 16) & 0xFF;
            var g = (rgb >> 8) & 0xFF;
            var b = rgb & 0xFF;

            // 使用相对亮度公式
            return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
        }
        catch (COMException)
        {
            return 0;
        }
    }

    public double GetSaturation()
    {
        try
        {
            var rgb = _colorFormat.RGB;
            var r = (rgb >> 16) & 0xFF;
            var g = (rgb >> 8) & 0xFF;
            var b = rgb & 0xFF;

            var max = Math.Max(Math.Max(r, g), b);
            var min = Math.Min(Math.Min(r, g), b);

            if (max == min)
                return 0;

            var lightness = (max + min) / 2.0;
            var delta = max - min;

            return lightness > 128 ? delta / (510 - max - min) : delta / (max + min);
        }
        catch (COMException)
        {
            return 0;
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _colorFormat != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_colorFormat) > 0) { }
            }
            catch { }
            _colorFormat = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    private void RGBToHSL(int rgb, out double hue, out double saturation, out double lightness)
    {
        var r = ((rgb >> 16) & 0xFF) / 255.0;
        var g = ((rgb >> 8) & 0xFF) / 255.0;
        var b = (rgb & 0xFF) / 255.0;

        var max = Math.Max(Math.Max(r, g), b);
        var min = Math.Min(Math.Min(r, g), b);
        var delta = max - min;

        lightness = (max + min) / 2;

        if (delta == 0)
        {
            hue = 0;
            saturation = 0;
        }
        else
        {
            saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);

            if (max == r)
                hue = (g - b) / delta + (g < b ? 6 : 0);
            else if (max == g)
                hue = (b - r) / delta + 2;
            else
                hue = (r - g) / delta + 4;

            hue /= 6;
        }
    }

    private int HSLToRGB(double hue, double saturation, double lightness)
    {
        if (saturation == 0)
        {
            var gray = (int)(lightness * 255);
            return (gray << 16) | (gray << 8) | gray;
        }

        var q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
        var p = 2 * lightness - q;

        var r = (int)(HueToRGB(p, q, hue + 1.0 / 3) * 255);
        var g = (int)(HueToRGB(p, q, hue) * 255);
        var b = (int)(HueToRGB(p, q, hue - 1.0 / 3) * 255);

        return (r << 16) | (g << 8) | b;
    }

    private double HueToRGB(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    private string GetStandardColorName(int rgb)
    {
        // 简化的标准颜色名称映射
        switch (rgb)
        {
            case 0xFF0000: return "Red";
            case 0x00FF00: return "Green";
            case 0x0000FF: return "Blue";
            case 0xFFFF00: return "Yellow";
            case 0xFF00FF: return "Magenta";
            case 0x00FFFF: return "Cyan";
            case 0x000000: return "Black";
            case 0xFFFFFF: return "White";
            default: return $"RGB({rgb:X6})";
        }
    }

    private int GetContrastingColor(int rgb)
    {
        var r = (rgb >> 16) & 0xFF;
        var g = (rgb >> 8) & 0xFF;
        var b = rgb & 0xFF;

        // 计算亮度
        var luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0;

        // 如果亮度大于0.5，返回黑色；否则返回白色
        return luminance > 0.5 ? 0x000000 : 0xFFFFFF;
    }
}