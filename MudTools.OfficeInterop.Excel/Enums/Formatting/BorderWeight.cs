//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 边框粗细枚举
/// 用于指定单元格边框线条的粗细程度
/// </summary>
public enum BorderWeight
{
    /// <summary>
    /// 发丝线
    /// 最细的边框线条
    /// </summary>
    Hairline = 1,
    
    /// <summary>
    /// 细线
    /// 较细的边框线条
    /// </summary>
    Thin = 2,
    
    /// <summary>
    /// 中等粗细
    /// 中等粗细的边框线条
    /// </summary>
    Medium = 3,
    
    /// <summary>
    /// 粗线
    /// 较粗的边框线条
    /// </summary>
    Thick = 4
}