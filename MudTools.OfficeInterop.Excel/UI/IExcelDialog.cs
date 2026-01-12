//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示内置的 Microsoft Excel 对话框。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDialog : IOfficeObject<IExcelDialog, MsExcel.Dialog>, IDisposable
{
    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 显示内置对话框并等待用户输入数据。
    /// </summary>
    /// <param name="arg1">命令的初始参数（可选）。</param>
    /// <param name="arg2">命令的初始参数（可选）。</param>
    /// <param name="arg3">命令的初始参数（可选）。</param>
    /// <param name="arg4">命令的初始参数（可选）。</param>
    /// <param name="arg5">命令的初始参数（可选）。</param>
    /// <param name="arg6">命令的初始参数（可选）。</param>
    /// <param name="arg7">命令的初始参数（可选）。</param>
    /// <param name="arg8">命令的初始参数（可选）。</param>
    /// <param name="arg9">命令的初始参数（可选）。</param>
    /// <param name="arg10">命令的初始参数（可选）。</param>
    /// <param name="arg11">命令的初始参数（可选）。</param>
    /// <param name="arg12">命令的初始参数（可选）。</param>
    /// <param name="arg13">命令的初始参数（可选）。</param>
    /// <param name="arg14">命令的初始参数（可选）。</param>
    /// <param name="arg15">命令的初始参数（可选）。</param>
    /// <param name="arg16">命令的初始参数（可选）。</param>
    /// <param name="arg17">命令的初始参数（可选）。</param>
    /// <param name="arg18">命令的初始参数（可选）。</param>
    /// <param name="arg19">命令的初始参数（可选）。</param>
    /// <param name="arg20">命令的初始参数（可选）。</param>
    /// <param name="arg21">命令的初始参数（可选）。</param>
    /// <param name="arg22">命令的初始参数（可选）。</param>
    /// <param name="arg23">命令的初始参数（可选）。</param>
    /// <param name="arg24">命令的初始参数（可选）。</param>
    /// <param name="arg25">命令的初始参数（可选）。</param>
    /// <param name="arg26">命令的初始参数（可选）。</param>
    /// <param name="arg27">命令的初始参数（可选）。</param>
    /// <param name="arg28">命令的初始参数（可选）。</param>
    /// <param name="arg29">命令的初始参数（可选）。</param>
    /// <param name="arg30">命令的初始参数（可选）。</param>
    /// <returns>如果用户单击“确定”按钮，则为 true；如果用户单击“取消”按钮，则为 false。</returns>
    bool? Show(object? arg1 = null, object? arg2 = null, object? arg3 = null, object? arg4 = null,
            object? arg5 = null, object? arg6 = null, object? arg7 = null, object? arg8 = null,
            object? arg9 = null, object? arg10 = null, object? arg11 = null, object? arg12 = null,
            object? arg13 = null, object? arg14 = null, object? arg15 = null, object? arg16 = null,
            object? arg17 = null, object? arg18 = null, object? arg19 = null, object? arg20 = null,
            object? arg21 = null, object? arg22 = null, object? arg23 = null, object? arg24 = null,
            object? arg25 = null, object? arg26 = null, object? arg27 = null, object? arg28 = null,
            object? arg29 = null, object? arg30 = null);
}