//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel ErrorCheckingOptions 对象的二次封装实现类
/// </summary>
internal class ExcelErrorCheckingOptions : IExcelErrorCheckingOptions
{
    /// <summary>
    /// 底层的 COM ErrorCheckingOptions 对象
    /// </summary>
    private MsExcel.ErrorCheckingOptions _errorCheckingOptions;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 初始化 ExcelErrorCheckingOptions 实例
    /// </summary>
    /// <param name="errorCheckingOptions">底层的 COM ErrorCheckingOptions 对象</param>
    internal ExcelErrorCheckingOptions(MsExcel.ErrorCheckingOptions errorCheckingOptions)
    {
        _errorCheckingOptions = errorCheckingOptions;
        _disposedValue = false;
    }

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                if (_errorCheckingOptions != null)
                    Marshal.ReleaseComObject(_errorCheckingOptions);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _errorCheckingOptions = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    /// <summary>
    /// 获取或设置是否检查背景错误
    /// </summary>
    public bool BackgroundChecking
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.BackgroundChecking;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.BackgroundChecking = value;
        }
    }

    /// <summary>
    /// 获取或设置是否检查空单元格引用
    /// </summary>
    public bool EmptyCellReferences
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.EmptyCellReferences;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.EmptyCellReferences = value;
        }
    }

    /// <summary>
    /// 获取或设置是否检查数字存储为文本
    /// </summary>
    public bool NumberAsText
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.NumberAsText;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.NumberAsText = value;
        }
    }

    /// <summary>
    /// 获取或设置是否检查不一致的计算列公式
    /// </summary>
    public bool InconsistentFormula
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.InconsistentFormula;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.InconsistentFormula = value;
        }
    }


    /// <summary>
    /// 获取或设置是否检查文本日期
    /// </summary>
    public bool TextDate
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.TextDate;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.TextDate = value;
        }
    }


    /// <summary>
    /// 获取或设置是否检查锁定单元格
    /// </summary>
    public bool UnlockedFormulaCells
    {
        get => _errorCheckingOptions != null && _errorCheckingOptions.UnlockedFormulaCells;
        set
        {
            if (_errorCheckingOptions != null)
                _errorCheckingOptions.UnlockedFormulaCells = value;
        }
    }



    /// <summary>
    /// 重置所有错误检查选项为默认值
    /// </summary>
    public void Reset()
    {
        if (_errorCheckingOptions == null) return;

        try
        {
            _errorCheckingOptions.BackgroundChecking = true;
            _errorCheckingOptions.EmptyCellReferences = true;
            _errorCheckingOptions.NumberAsText = true;
            _errorCheckingOptions.InconsistentFormula = true;
            _errorCheckingOptions.TextDate = true;
            _errorCheckingOptions.UnlockedFormulaCells = true;
        }
        catch
        {
            // 忽略重置过程中的异常
        }
    }
}