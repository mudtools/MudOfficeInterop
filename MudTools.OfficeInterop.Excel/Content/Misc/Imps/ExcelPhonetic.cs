//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Excel.Phonetic 的封装实现类
/// </summary>
internal class ExcelPhonetic : IExcelPhonetic
{
    #region 属性封装

    /// <summary>
    /// 获取或设置注音符号的文本内容
    /// </summary>
    public string Text
    {
        get => _phonetic.Text?.ToString();
        set => _phonetic.Text = value;
    }

    /// <summary>
    /// 获取或设置注音符号的可见性
    /// </summary>
    public bool Visible
    {
        get => Convert.ToBoolean(_phonetic.Visible);
        set => _phonetic.Visible = value;
    }

    public IExcelFont? Font => _phonetic != null ? new ExcelFont(_phonetic.Font) : null;

    /// <summary>
    /// 获取或设置注音符号的字体名称
    /// </summary>
    public string FontName
    {
        get => _phonetic.Font.Name?.ToString();
        set => _phonetic.Font.Name = value;
    }

    /// <summary>
    /// 获取或设置注音符号的字体大小
    /// </summary>
    public double FontSize
    {
        get => Convert.ToDouble(_phonetic.Font.Size);
        set => _phonetic.Font.Size = value;
    }

    /// <summary>
    /// 获取或设置注音符号是否粗体
    /// </summary>
    public bool FontBold
    {
        get => Convert.ToBoolean(_phonetic.Font.Bold);
        set => _phonetic.Font.Bold = value;
    }

    /// <summary>
    /// 获取或设置注音符号是否斜体
    /// </summary>
    public bool FontItalic
    {
        get => Convert.ToBoolean(_phonetic.Font.Italic);
        set => _phonetic.Font.Italic = value;
    }

    /// <summary>
    /// 获取或设置注音符号的颜色（RGB值）
    /// </summary>
    public int FontColor
    {
        get => Convert.ToInt32(_phonetic.Font.Color);
        set => _phonetic.Font.Color = value;
    }

    /// <summary>
    /// 获取或设置注音符号的对齐方式
    /// </summary>
    public int Alignment
    {
        get => _phonetic.Alignment;
        set => _phonetic.Alignment = value;
    }

    #endregion

    #region 构造函数与私有字段

    private MsExcel.Phonetic _phonetic;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelPhonetic 实例
    /// </summary>
    /// <param name="phonetic">原始 COM Phonetic 对象</param>
    internal ExcelPhonetic(MsExcel.Phonetic phonetic)
    {
        _phonetic = phonetic ?? throw new ArgumentNullException(nameof(phonetic));
        _disposedValue = false;
    }

    #endregion

    #region IDisposable 模式实现

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否显式调用 Dispose()</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _phonetic != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_phonetic) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _phonetic = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 显式释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}