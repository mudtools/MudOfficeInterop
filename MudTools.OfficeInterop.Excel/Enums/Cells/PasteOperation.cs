//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 粘贴操作枚举
/// 用于指定粘贴时要应用的数学运算操作
/// </summary>
public enum PasteOperation
{
    /// <summary>
    /// 无操作
    /// 粘贴时不应用任何数学运算
    /// </summary>
    None = -4142,
    
    /// <summary>
    /// 加法操作
    /// 粘贴时将源数据与目标数据相加
    /// </summary>
    Add = 2,
    
    /// <summary>
    /// 除法操作
    /// 粘贴时将目标数据除以源数据
    /// </summary>
    Divide = 5,
    
    /// <summary>
    /// 乘法操作
    /// 粘贴时将源数据与目标数据相乘
    /// </summary>
    Multiply = 4,
    
    /// <summary>
    /// 减法操作
    /// 粘贴时将目标数据减去源数据
    /// </summary>
    Subtract = 3
}