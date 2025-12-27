//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// PowerPoint 柔化边缘格式接口
/// </summary>
public interface IPowerPointSoftEdgeFormat : IDisposable
{
    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置半径
    /// </summary>
    float Radius { get; set; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 设置柔化边缘
    /// </summary>
    /// <param name="radius">半径</param>
    void SetSoftEdge(float radius);

    /// <summary>
    /// 重置柔化边缘格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 应用柔化边缘格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    void ApplyTo(IPowerPointShape shape);

    /// <summary>
    /// 设置柔化边缘参数
    /// </summary>
    /// <param name="radius">半径</param>
    void SetParameters(float radius = -1);

    /// <summary>
    /// 获取柔化边缘信息
    /// </summary>
    /// <returns>柔化边缘信息字符串</returns>
    string GetSoftEdgeInfo();
}
