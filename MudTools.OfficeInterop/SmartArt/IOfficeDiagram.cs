//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中的图表对象接口，用于管理和操作 SmartArt 图形
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore", ComClassName = "IMsoDiagram")]
public interface IOfficeDiagram : IDisposable
{
    /// <summary>
    /// 获取图表中的节点集合
    /// </summary>
    IOfficeDiagramNodes? Nodes { get; }

    /// <summary>
    /// 获取图表的类型
    /// </summary>
    MsoDiagramType Type { get; }

    /// <summary>
    /// 获取或设置是否自动布局图表节点
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoLayout { get; set; }

    /// <summary>
    /// 获取或设置是否反向显示图表布局
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Reverse { get; set; }

    /// <summary>
    /// 获取或设置是否自动格式化图表
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoFormat { get; set; }

    /// <summary>
    /// 将图表转换为指定类型
    /// </summary>
    /// <param name="Type">目标图表类型</param>
    void Convert(MsoDiagramType Type);

    /// <summary>
    /// 调整图表文本以适应节点大小
    /// </summary>
    void FitText();
}