//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 幻灯片集合接口
/// </summary>
public interface IPowerPointSlides : IDisposable, IEnumerable<IPowerPointSlide>
{
    /// <summary>
    /// 获取幻灯片数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 根据索引获取幻灯片（从1开始）
    /// </summary>
    IPowerPointSlide this[int index] { get; }

    /// <summary>
    /// 根据名称获取幻灯片
    /// </summary>
    IPowerPointSlide this[string name] { get; }

    /// <summary>
    /// 根据索引获取幻灯片（从1开始）
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    /// <returns>幻灯片对象</returns>
    IPowerPointSlide ByIndex(int index);

    /// <summary>
    /// 根据名称获取幻灯片
    /// </summary>
    /// <param name="name">幻灯片名称</param>
    /// <returns>幻灯片对象</returns>
    IPowerPointSlide ByName(string name);

    /// <summary>
    /// 添加新幻灯片
    /// </summary>
    /// <param name="layout">幻灯片布局</param>
    /// <param name="position">插入位置（-1表示末尾）</param>
    /// <returns>新添加的幻灯片</returns>
    IPowerPointSlide Add(PpSlideLayout layout = PpSlideLayout.ppLayoutText, int position = -1);

    /// <summary>
    /// 从文件插入幻灯片
    /// </summary>
    /// <param name="fileName">文件路径</param>
    /// <param name="position">插入位置</param>
    /// <param name="slideRange">幻灯片范围</param>
    /// <returns>插入的幻灯片</returns>
    //IPowerPointSlide InsertFromFile(string fileName, int position, int slideRange = -1);

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="index">幻灯片索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除幻灯片
    /// </summary>
    /// <param name="name">幻灯片名称</param>
    void Delete(string name);

    /// <summary>
    /// 移动幻灯片
    /// </summary>
    /// <param name="fromIndex">源索引</param>
    /// <param name="toIndex">目标索引</param>
    void Move(int fromIndex, int toIndex);

    /// <summary>
    /// 复制幻灯片
    /// </summary>
    /// <param name="sourceIndex">源索引</param>
    /// <param name="targetIndex">目标索引</param>
    /// <returns>复制的幻灯片</returns>
    IPowerPointSlide Copy(int sourceIndex, int targetIndex = -1);

    /// <summary>
    /// 获取所有幻灯片
    /// </summary>
    /// <returns>幻灯片列表</returns>
    IEnumerable<IPowerPointSlide> GetAll();

    /// <summary>
    /// 根据幻灯片编号查找幻灯片
    /// </summary>
    /// <param name="slideNumber">幻灯片编号</param>
    /// <returns>幻灯片对象</returns>
    IPowerPointSlide FindByNumber(int slideNumber);

    /// <summary>
    /// 获取第一张幻灯片
    /// </summary>
    /// <returns>第一张幻灯片对象</returns>
    IPowerPointSlide First();

    /// <summary>
    /// 获取最后一张幻灯片
    /// </summary>
    /// <returns>最后一张幻灯片对象</returns>
    IPowerPointSlide Last();
}