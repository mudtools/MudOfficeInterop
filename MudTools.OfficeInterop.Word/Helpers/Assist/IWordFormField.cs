//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中的一个表单域（FormField）的封装接口。
/// </summary>
public interface IWordFormField : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置表单域的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置表单域的结果文本。
    /// </summary>
    string Result { get; set; }

    /// <summary>
    /// 获取表单域的类型。
    /// </summary>
    WdFieldType Type { get; }

    /// <summary>
    /// 获取表单域的文本输入属性（仅适用于文本输入类型的表单域）。
    /// </summary>
    IWordTextInput? TextInput { get; }

    /// <summary>
    /// 获取表单域的复选框属性（仅适用于复选框类型的表单域）。
    /// </summary>
    IWordCheckBox? CheckBox { get; }

    /// <summary>
    /// 获取表单域的下拉列表属性（仅适用于下拉列表类型的表单域）。
    /// </summary>
    IWordDropDown? DropDown { get; }

    /// <summary>
    /// 获取文档中的下一个表单域。
    /// </summary>
    IWordFormField? Next { get; }

    /// <summary>
    /// 获取文档中的上一个表单域。
    /// </summary>
    IWordFormField? Previous { get; }

    /// <summary>
    /// 获取与表单域关联的范围对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取或设置复选框表单域的状态（仅适用于 CheckBox 类型）。
    /// </summary>
    bool CheckBox_Checked { get; set; }

    /// <summary>
    /// 获取或设置文本表单域的默认值（仅适用于 TextInput 类型）。
    /// </summary>
    string TextInput_Default { get; set; }

    /// <summary>
    /// 获取或设置下拉表单域的默认项索引（仅适用于 DropDown 类型）。
    /// </summary>
    int DropDown_Default { get; set; }

    /// <summary>
    /// 获取或设置下拉表单域的选项列表（仅适用于 DropDown 类型）。
    /// </summary>
    List<string> DropDown_ListEntries { get; }

    /// <summary>
    /// 将表单域复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将表单域剪切到剪贴板。
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除此表单域。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选中该表单域。
    /// </summary>
    void Select();
}