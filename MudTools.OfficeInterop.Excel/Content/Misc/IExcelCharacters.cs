//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel中的字符对象接口，提供对单元格中文本的字符级操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelCharacters : IOfficeObject<IExcelCharacters, MsExcel.Characters>, IDisposable
{
    /// <summary>
    /// 获取当前COM对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取当前COM对象的Application对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取字符数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置文本内容
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置对象的标题
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置对象的拼音文本
    /// </summary>
    string PhoneticCharacters { get; set; }

    /// <summary>
    /// 获取字符的字体属性
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 删除字符
    /// </summary>
    void Delete();

    /// <summary>
    /// 插入文本到指定位置
    /// </summary>
    /// <param name="text">要插入的文本</param>
    /// <returns>插入后的字符对象</returns>
    [ValueConvert]
    IExcelCharacters? Insert(string text);
}
