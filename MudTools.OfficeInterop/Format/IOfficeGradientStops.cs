//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office应用中渐变停止点的集合接口
/// 渐变停止点定义了渐变颜色过渡的关键点，包括颜色、位置和透明度等属性
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
public interface IOfficeGradientStops : IEnumerable<IOfficeGradientStop?>, IDisposable
{
    /// <summary>
    /// 获取集合中渐变停止点的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定位置的渐变停止点
    /// </summary>
    /// <param name="index">要获取的渐变停止点的从零开始的索引</param>
    /// <returns>位于指定索引处的渐变停止点；如果索引无效，则返回null</returns>
    IOfficeGradientStop? this[int index] { get; }

    /// <summary>
    /// 删除指定索引位置的渐变停止点
    /// </summary>
    /// <param name="index">要删除的渐变停止点的索引；如果为-1，则删除最后一个渐变停止点</param>
    void Delete(int index = -1);

    /// <summary>
    /// 在集合中插入一个新的渐变停止点
    /// </summary>
    /// <param name="RGB">渐变停止点的颜色值，以RGB格式表示（0-16777215）</param>
    /// <param name="position">渐变停止点的位置，以百分比表示（0.0-1.0）</param>
    /// <param name="transparency">渐变停止点的透明度，以百分比表示（0.0-1.0，默认为0）</param>
    /// <param name="index">要插入的位置索引；如果为-1，则在末尾插入</param>
    void Insert(int RGB, float position, float transparency = 0f, int index = -1);

    /// <summary>
    /// 在集合中插入一个新的渐变停止点（包含亮度参数）
    /// </summary>
    /// <param name="RGB">渐变停止点的颜色值，以RGB格式表示（0-16777215）</param>
    /// <param name="position">渐变停止点的位置，以百分比表示（0.0-1.0）</param>
    /// <param name="transparency">渐变停止点的透明度，以百分比表示（0.0-1.0，默认为0）</param>
    /// <param name="index">要插入的位置索引；如果为-1，则在末尾插入</param>
    /// <param name="brightness">渐变停止点的亮度，以百分比表示（-1.0-1.0，默认为0）</param>
    void Insert2(int RGB, float position, float transparency = 0f, int index = -1, float brightness = 0f);
}