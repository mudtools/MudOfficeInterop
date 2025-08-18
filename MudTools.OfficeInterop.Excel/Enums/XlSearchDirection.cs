//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 搜索方向枚举
/// 用于指定在单元格区域中进行搜索或查找时的方向
/// </summary>
public enum XlSearchDirection
{
    /// <summary>
    /// 下一个
    /// 按照从上到下、从左到右的顺序搜索下一个匹配项
    /// </summary>
    xlNext = 1,
    
    /// <summary>
    /// 上一个
    /// 按照从下到上、从右到左的顺序搜索上一个匹配项
    /// </summary>
    xlPrevious
}
