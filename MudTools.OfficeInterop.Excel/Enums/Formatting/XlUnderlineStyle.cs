//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 下划线样式枚举
/// 用于指定文本下划线的样式类型
/// </summary>
public enum XlUnderlineStyle
{
    /// <summary>
    /// 双线下划线
    /// 使用两条细线作为下划线
    /// </summary>
    xlUnderlineStyleDouble = -4119,
    
    /// <summary>
    /// 会计用双线下划线
    /// 会计专用的双线下划线样式，与普通双线下划线略有不同
    /// </summary>
    xlUnderlineStyleDoubleAccounting = 5,
    
    /// <summary>
    /// 无下划线
    /// 不应用下划线
    /// </summary>
    xlUnderlineStyleNone = -4142,
    
    /// <summary>
    /// 单线下划线
    /// 使用一条细线作为下划线
    /// </summary>
    xlUnderlineStyleSingle = 2,
    
    /// <summary>
    /// 会计用单线下划线
    /// 会计专用的单线下划线样式
    /// </summary>
    xlUnderlineStyleSingleAccounting = 4
}
