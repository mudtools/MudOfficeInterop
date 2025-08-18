//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 定义Excel字体样式的基本接口
/// 该接口提供了操作Excel单元格字体的各种属性和方法
/// </summary>
public interface IExcelFont : IDisposable
{
    /// <summary>
    /// 获取或设置字体名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置字体大小
    /// </summary>
    double Size { get; set; }

    /// <summary>
    /// 获取或设置是否粗体
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置是否斜体
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置是否删除线
    /// </summary>
    bool Strikethrough { get; set; }

    /// <summary>
    /// 获取或设置字体样式
    /// </summary>
    object FontStyle { get; set; }

    /// <summary>
    /// 获取或设置字体颜色索引
    /// </summary>
    int ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置字体颜色（RGB值）
    /// </summary>
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置下划线样式
    /// </summary>
    XlUnderlineStyle Underline { get; set; }

    /// <summary>
    /// 获取或设置是否为上标
    /// </summary>
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置是否为下标
    /// </summary>
    bool Subscript { get; set; }
}
