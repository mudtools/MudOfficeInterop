//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的数学矩阵对象，提供对矩阵行、列和单元格的访问以及矩阵格式设置功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathMat : IDisposable
{
    /// <summary>
    /// 获取与此矩阵关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此矩阵的父级对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取矩阵的行集合，允许访问和操作矩阵中的每一行
    /// </summary>
    IWordOMathMatRows? Rows { get; }

    /// <summary>
    /// 获取矩阵的列集合，允许访问和操作矩阵中的每一列
    /// </summary>
    IWordOMathMatCols? Cols { get; }

    /// <summary>
    /// 获取矩阵的单元格对象，用于访问矩阵中的单个单元格
    /// </summary>
    [IgnoreGenerator]
    IWordOMath? Cell(int row, int col);

    /// <summary>
    /// 获取或设置矩阵的垂直对齐方式
    /// </summary>
    WdOMathVertAlignType Align { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示占位符是否隐藏
    /// </summary>
    bool PlcHoldHidden { get; set; }

    /// <summary>
    /// 获取或设置矩阵行间距的规则
    /// </summary>
    WdOMathSpacingRule RowSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置矩阵行之间的间距值
    /// </summary>
    int RowSpacing { get; set; }

    /// <summary>
    /// 获取或设置矩阵列之间的间距值
    /// </summary>
    int ColSpacing { get; set; }

    /// <summary>
    /// 获取或设置矩阵列间隔的规则
    /// </summary>
    WdOMathSpacingRule ColGapRule { get; set; }

    /// <summary>
    /// 获取或设置矩阵列之间的间隔值
    /// </summary>
    int ColGap { get; set; }
}