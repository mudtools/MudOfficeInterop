//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 线条格式接口
/// </summary>
public interface IPowerPointLineFormat : IDisposable
{
    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置线条样式
    /// </summary>
    int Style { get; set; }

    /// <summary>
    /// 获取或设置线条粗细
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置前景色
    /// </summary>
    int ForeColor { get; set; }

    /// <summary>
    /// 获取或设置可见性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置虚线样式
    /// </summary>
    int DashStyle { get; set; }

    /// <summary>
    /// 获取或设置起始箭头样式
    /// </summary>
    int BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置结束箭头样式
    /// </summary>
    int EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置起始箭头宽度
    /// </summary>
    int BeginArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置结束箭头宽度
    /// </summary>
    int EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置起始箭头长度
    /// </summary>
    int BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置结束箭头长度
    /// </summary>
    int EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取线条类型
    /// </summary>
    int Type { get; }


    /// <summary>
    /// 设置箭头样式
    /// </summary>
    /// <param name="beginStyle">起始箭头样式</param>
    /// <param name="endStyle">结束箭头样式</param>
    /// <param name="beginWidth">起始箭头宽度</param>
    /// <param name="endWidth">结束箭头宽度</param>
    /// <param name="beginLength">起始箭头长度</param>
    /// <param name="endLength">结束箭头长度</param>
    void SetArrowheads(int beginStyle = 0, int endStyle = 0, int beginWidth = 1, int endWidth = 1, int beginLength = 1, int endLength = 1);


    /// <summary>
    /// 重置线条格式
    /// </summary>
    void Reset();

    /// <summary>
    /// 复制线条格式
    /// </summary>
    /// <returns>复制的线条格式对象</returns>
    IPowerPointLineFormat Duplicate();

    /// <summary>
    /// 应用线条格式到指定形状
    /// </summary>
    /// <param name="shape">目标形状</param>
    void ApplyTo(IPowerPointShape shape);

    /// <summary>
    /// 设置线条粗细
    /// </summary>
    /// <param name="weight">线条粗细</param>
    void SetWeight(float weight);

    /// <summary>
    /// 设置线条颜色
    /// </summary>
    /// <param name="color">线条颜色</param>
    void SetColor(int color);

    /// <summary>
    /// 设置线条样式
    /// </summary>
    /// <param name="style">线条样式</param>
    void SetStyle(int style);

    /// <summary>
    /// 获取线条信息
    /// </summary>
    /// <returns>线条信息字符串</returns>
    string GetLineInfo();
}