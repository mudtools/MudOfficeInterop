//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// SmartTagRecognizer 接口及实现类
/// </summary>
public interface IWordSmartTagRecognizer : IDisposable
{
    /// <summary>
    /// 获取代表 Microsoft Word 应用程序的 Application 对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取代表指定对象的父对象的对象 [[10]]。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取标识指定智能标记识别程序的标签 [[4]]。
    /// </summary>
    string Caption { get; }

    /// <summary>
    /// 获取智能标记识别程序的完整名称，包括路径 [[9]]。
    /// </summary>
    string FullName { get; }

    string ProgID { get; }

    /// <summary>
    /// 获取或设置一个 Boolean 类型的值，该值确定是否启用指定的智能标记识别程序 [[7]]。
    /// </summary>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取一个 32 位整数，该整数指示创建对象的应用程序 [[13]]。
    /// </summary>
    int Creator { get; }
}
