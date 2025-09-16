//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ErrorCheckingOptions 对象的二次封装接口
/// </summary>
public interface IExcelErrorCheckingOptions : IDisposable
{
    /// <summary>
    /// 获取或设置是否检查背景错误
    /// </summary>
    bool BackgroundChecking { get; set; }

    /// <summary>
    /// 获取或设置是否检查空单元格引用
    /// </summary>
    bool EmptyCellReferences { get; set; }

    /// <summary>
    /// 获取或设置是否检查数字存储为文本
    /// </summary>
    bool NumberAsText { get; set; }

    /// <summary>
    /// 获取或设置是否检查不一致的计算列公式
    /// </summary>
    bool InconsistentFormula { get; set; }

    /// <summary>
    /// 获取或设置是否检查文本日期
    /// </summary>
    bool TextDate { get; set; }

    /// <summary>
    /// 获取或设置是否检查锁定单元格
    /// </summary>
    bool UnlockedFormulaCells { get; set; }

    /// <summary>
    /// 重置所有错误检查选项为默认值
    /// </summary>
    void Reset();
}