//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// &lt;summary&gt;
/// 表示渐变停止点的接口，用于定义渐变中的颜色、位置和透明度属性
/// &lt;/summary&gt;
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeGradientStop : IDisposable
{
    /// &lt;summary&gt;
    /// 获取或设置渐变停止点的颜色
    /// &lt;/summary&gt;
    IOfficeColorFormat Color { get; }

    /// &lt;summary&gt;
    /// 获取或设置渐变停止点的位置，通常以0.0到1.0之间的浮点数表示
    /// &lt;/summary&gt;
    float Position { get; set; }
    
    /// &lt;summary&gt;
    /// 获取或设置渐变停止点的透明度，通常以0.0到1.0之间的浮点数表示，其中0.0为完全不透明，1.0为完全透明
    /// &lt;/summary&gt;
    float Transparency { get; set; }
}