//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// 对 Microsoft.Office.Interop.Excel.Phonetics 的封装实现类
/// </summary>
internal class ExcelPhonetics : IExcelPhonetics
{
    #region 属性封装

    /// <summary>
    /// 获取注音符号的数量
    /// </summary>
    public int Count => Convert.ToInt32(_phonetics.Count);

    /// <summary>
    /// 获取指定索引的注音符号对象
    /// </summary>
    /// <param name="index">注音符号索引（从1开始）</param>
    /// <returns>注音符号对象</returns>
    public IExcelPhonetic? this[int index]
    {
        get
        {
            try
            {
                if (_phonetics == null)
                    throw new ObjectDisposedException(nameof(ExcelPhonetics));
                if (index < 1 || index > Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), "索引超出范围");
                }
                if (_phonetics[index] is not MsExcel.Phonetic phonetic)
                    return null;

                return new ExcelPhonetic(phonetic);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法获取指定索引的注音符号对象", ex);
            }
        }

    }


    public IExcelFont? Font => _phonetics != null ? new ExcelFont(_phonetics.Font) : null;


    /// <summary>
    /// 获取或设置注音符号的字体名称
    /// </summary>
    public string FontName
    {
        get => _phonetics.Font.Name?.ToString();
        set => _phonetics.Font.Name = value;
    }

    /// <summary>
    /// 获取或设置注音符号的字体大小
    /// </summary>
    public double FontSize
    {
        get => Convert.ToDouble(_phonetics.Font.Size);
        set => _phonetics.Font.Size = value;
    }

    /// <summary>
    /// 获取或设置注音符号是否粗体
    /// </summary>
    public bool FontBold
    {
        get => Convert.ToBoolean(_phonetics.Font.Bold);
        set => _phonetics.Font.Bold = value;
    }

    /// <summary>
    /// 获取或设置注音符号是否斜体
    /// </summary>
    public bool FontItalic
    {
        get => Convert.ToBoolean(_phonetics.Font.Italic);
        set => _phonetics.Font.Italic = value;
    }

    /// <summary>
    /// 获取或设置注音符号的颜色（RGB值）
    /// </summary>
    public int FontColor
    {
        get => Convert.ToInt32(_phonetics.Font.Color);
        set => _phonetics.Font.Color = value;
    }

    /// <summary>
    /// 获取或设置注音符号的可见性
    /// </summary>
    public bool Visible
    {
        get => Convert.ToBoolean(_phonetics.Visible);
        set => _phonetics.Visible = value;
    }

    /// <summary>
    /// 获取或设置注音符号的对齐方式
    /// </summary>
    public int Alignment
    {
        get => Convert.ToInt32(_phonetics.Alignment);
        set => _phonetics.Alignment = value;
    }

    /// <summary>
    /// 获取或设置注音符号的字符类型
    /// </summary>
    public int CharacterType
    {
        get => Convert.ToInt32(_phonetics.CharacterType);
        set => _phonetics.CharacterType = value;
    }

    #endregion

    #region 构造函数与私有字段

    private MsExcel.Phonetics _phonetics;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelPhonetics 实例
    /// </summary>
    /// <param name="phonetics">原始 COM Phonetics 对象</param>
    internal ExcelPhonetics(MsExcel.Phonetics phonetics)
    {
        _phonetics = phonetics ?? throw new ArgumentNullException(nameof(phonetics));
        _disposedValue = false;
    }

    #endregion

    #region 公共方法

    /// <summary>
    /// 向集合中添加新的注音符号
    /// </summary>
    /// <param name="start">开始位置</param>
    /// <param name="length">长度</param>
    /// <param name="text">注音文本</param>
    /// <returns>新创建的注音符号对象</returns>
    public void Add(int start, int length, string text)
    {
        try
        {
            _phonetics.Add(start, length, text);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加注音符号", ex);
        }
    }

    /// <summary>
    /// 删除所有注音符号
    /// </summary>
    public void Delete()
    {
        try
        {
            _phonetics.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除注音符号", ex);
        }
    }

    /// <summary>
    /// 获取所有注音符号的枚举器
    /// </summary>
    public IEnumerable<IExcelPhonetic> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
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

        if (disposing && _phonetics != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_phonetics) > 0) { }
            }
            catch
            {
                // 忽略释放 COM 对象时的异常
            }
            _phonetics = null;
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