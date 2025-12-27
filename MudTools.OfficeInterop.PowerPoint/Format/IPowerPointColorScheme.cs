//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 颜色方案接口
/// </summary>
public interface IPowerPointColorScheme : IDisposable
{
    /// <summary>
    /// 获取颜色数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置指定索引的颜色
    /// </summary>
    /// <param name="schemeColorIndex">颜色方案索引</param>
    /// <returns>颜色值</returns>
    int this[int schemeColorIndex] { get; set; }

    /// <summary>
    /// 获取指定索引的颜色
    /// </summary>
    /// <param name="index">颜色索引</param>
    /// <returns>颜色值</returns>
    int Colors(int index);

    /// <summary>
    /// 获取指定名称的颜色
    /// </summary>
    /// <param name="name">颜色名称</param>
    /// <returns>颜色值</returns>
    int Colors(string name);

    /// <summary>
    /// 应用颜色方案
    /// </summary>
    /// <param name="schemeIndex">方案索引</param>
    void Apply(int schemeIndex);

    /// <summary>
    /// 保存颜色方案
    /// </summary>
    /// <param name="fileName">文件路径</param>
    void Save(string fileName);

    /// <summary>
    /// 加载颜色方案
    /// </summary>
    /// <param name="fileName">文件路径</param>
    void Load(string fileName);

    /// <summary>
    /// 重置颜色方案
    /// </summary>
    void Reset();

    /// <summary>
    /// 设置颜色值
    /// </summary>
    /// <param name="index">颜色索引</param>
    /// <param name="color">颜色值</param>
    void SetColor(int index, int color);

    /// <summary>
    /// 应用到所有幻灯片
    /// </summary>
    void ApplyToAll();

    /// <summary>
    /// 获取颜色方案信息
    /// </summary>
    /// <returns>颜色方案信息字符串</returns>
    string GetColorSchemeInfo();
}