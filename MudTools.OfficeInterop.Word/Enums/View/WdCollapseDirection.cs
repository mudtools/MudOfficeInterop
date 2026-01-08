//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定在折叠范围或选择时应将插入点移动到范围的开头还是结尾。
/// 这个枚举用于Microsoft Word对象模型中范围的折叠操作。
/// </summary>
public enum WdCollapseDirection
{
    /// <summary>
    /// 将范围或选择折叠到其开始位置。这会将活动端点移动到静态端点的位置，
    /// 从而创建一个插入点位于范围开始位置的范围。
    /// </summary>
    wdCollapseStart = 1,
    /// <summary>
    /// 将范围或选择折叠到其结束位置。这会将活动端点移动到静态端点的位置，
    /// 从而创建一个插入点位于范围结束位置的范围。
    /// </summary>
    wdCollapseEnd = 0
}