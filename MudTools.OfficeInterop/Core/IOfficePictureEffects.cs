//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureEffects 的接口，用于操作图片效果集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficePictureEffects : IOfficeObject<IOfficePictureEffects, MsCore.PictureEffects>, IEnumerable<IOfficePictureEffect?>, IDisposable
{
    /// <summary>
    /// 获取图片效果集合中的效果数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取图片效果（从1开始）。
    /// </summary>
    IOfficePictureEffect? this[int index] { get; }

    /// <summary>
    /// 根据效果类型获取图片效果。
    /// </summary>
    IOfficePictureEffect? this[[ConvertInt] MsoPictureEffectType effectType] { get; }

    /// <summary>
    /// 添加新的图片效果。
    /// </summary>
    /// <param name="effectType">效果类型。</param>
    /// <param name="position">效果位置（可选）。</param>
    /// <returns>新添加的图片效果。</returns>
    IOfficePictureEffect? Insert(MsoPictureEffectType effectType, int position = -1);

    /// <summary>
    /// 删除指定索引的图片效果。
    /// </summary>
    /// <param name="index">效果索引。</param>
    void Delete(int index = -1);
}