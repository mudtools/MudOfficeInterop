namespace MudTools.OfficeInterop;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.PictureEffects 的接口，用于操作图片效果集合。
/// </summary>
public interface IOfficePictureEffects : IEnumerable<IOfficePictureEffect>, IDisposable
{
    /// <summary>
    /// 获取图片效果集合中的效果数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取图片效果（从1开始）。
    /// </summary>
    IOfficePictureEffect this[int index] { get; }

    /// <summary>
    /// 根据效果类型获取图片效果。
    /// </summary>
    IOfficePictureEffect this[MsoPictureEffectType effectType] { get; }

    /// <summary>
    /// 添加新的图片效果。
    /// </summary>
    /// <param name="effectType">效果类型。</param>
    /// <param name="position">效果位置（可选）。</param>
    /// <returns>新添加的图片效果。</returns>
    IOfficePictureEffect Insert(MsoPictureEffectType effectType, int position = -1);

    /// <summary>
    /// 删除指定索引的图片效果。
    /// </summary>
    /// <param name="index">效果索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定类型的所有图片效果。
    /// </summary>
    /// <param name="effectType">效果类型。</param>
    void DeleteByType(MsoPictureEffectType effectType);

    /// <summary>
    /// 清除所有图片效果。
    /// </summary>
    void Clear();

    /// <summary>
    /// 检查是否包含指定类型的效果。
    /// </summary>
    /// <param name="effectType">效果类型。</param>
    /// <returns>是否包含。</returns>
    bool Contains(MsoPictureEffectType effectType);

    /// <summary>
    /// 获取指定类型的效果数量。
    /// </summary>
    /// <param name="effectType">效果类型。</param>
    /// <returns>效果数量。</returns>
    int GetCountByType(MsoPictureEffectType effectType);

    /// <summary>
    /// 获取所有效果类型的列表。
    /// </summary>
    /// <returns>效果类型列表。</returns>
    List<MsoPictureEffectType> GetAllEffectTypes();

    /// <summary>
    /// 将效果移动到指定位置。
    /// </summary>
    /// <param name="effect">要移动的效果。</param>
    /// <param name="newPosition">新位置。</param>
    void MoveEffect(IOfficePictureEffect effect, int newPosition);

    /// <summary>
    /// 重置所有图片效果为默认设置。
    /// </summary>
    void Reset();
}