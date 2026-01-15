//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表中的方案。方案是一组已命名并保存的输入值（称为可变单元格）。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelScenario : IOfficeObject<IExcelScenario, MsExcel.Scenario>, IDisposable
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
    /// 更改方案，使其具有新的可变单元格集和（可选）方案值。
    /// </summary>
    /// <param name="changingCells">必需。Range 对象，指定方案的新可变单元格集。可变单元格必须与方案位于同一工作表上。</param>
    /// <param name="values">可选项。数组，包含可变单元格的新方案值。如果省略此参数，则假定方案值为可变单元格中的当前值。</param>
    void ChangeScenario(IExcelRange? changingCells, object? values = null);

    /// <summary>
    /// 获取表示方案的可变单元格的 Range 对象。
    /// </summary>
    IExcelRange? ChangingCells { get; }

    /// <summary>
    /// 获取或设置与方案关联的注释。注释文本不能超过 255 个字符。
    /// </summary>
    string Comment { get; set; }

    /// <summary>
    /// 删除该对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取或设置一个值，该值指示方案是否隐藏。默认值为 False。
    /// </summary>
    bool Hidden { get; set; }

    /// <summary>
    /// 获取对象在相似对象集合中的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示对象是否已锁定。如果工作表受保护，则锁定的对象无法修改。
    /// </summary>
    bool Locked { get; set; }

    /// <summary>
    /// 获取或设置对象的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 通过在工作表上插入方案的值来显示方案。受影响的单元格是方案的可变单元格。
    /// </summary>
    void Show();

    /// <summary>
    /// 获取包含方案可变单元格的当前值的数组。
    /// </summary>
    object Values { get; }
}