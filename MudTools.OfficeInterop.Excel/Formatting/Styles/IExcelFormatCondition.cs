//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel FormatCondition 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.FormatCondition (及 ColorScale, DataBar, IconSetCondition) 的安全访问和操作
/// </summary>
public interface IExcelFormatCondition : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取条件格式规则的父对象 (通常是 FormatConditions 集合)
    /// 对应 FormatCondition.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取条件格式规则所在的Application对象
    /// 对应 FormatCondition.Application 属性
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取条件格式规则的类型
    /// 对应 FormatCondition.Type 属性
    /// </summary>
    int Type { get; } // 使用 int 代表 XlFormatConditionType

    /// <summary>
    /// 获取或设置比较操作符 (对于 xlCellValue 类型)
    /// 对应 FormatCondition.Operator 属性
    /// </summary>
    int Operator { get; set; } // 使用 int 代表 XlFormatConditionOperator

    /// <summary>
    /// 获取或设置公式1
    /// 对应 FormatCondition.Formula1 属性
    /// </summary>
    string Formula1 { get; set; }

    /// <summary>
    /// 获取或设置公式2
    /// 对应 FormatCondition.Formula2 属性
    /// </summary>
    string Formula2 { get; set; }

    /// <summary>
    /// 获取或设置文本 (对于 xlTextString 类型)
    /// 对应 FormatCondition.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置文本比较器 (对于 xlTextString 类型)
    /// 对应 FormatCondition.TextOperator 属性
    /// </summary>
    int TextOperator { get; set; } // 使用 int 代表 XlContainsOperator
    #endregion

    #region 格式设置 (特定于 FormatCondition)
    /// <summary>
    /// 获取条件格式规则的字体对象
    /// 对应 FormatCondition.Font 属性
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取条件格式规则的背景对象
    /// 对应 FormatCondition.Interior 属性
    /// </summary>
    IExcelInterior Interior { get; }

    /// <summary>
    /// 获取条件格式规则的边框对象
    /// 对应 FormatCondition.Borders 属性
    /// </summary>
    IExcelBorders Borders { get; }

    /// <summary>
    /// 获取或设置条件格式规则的编号格式
    /// 对应 FormatCondition.NumberFormat 属性
    /// </summary>
    object NumberFormat { get; set; } // Can be string or other types
    #endregion

    #region 高级属性 (特定类型)

    /// <summary>
    /// 获取或设置颜色刻度对象 (如果此条件是 ColorScale)
    /// </summary>
    IExcelColorScaleCriteria ColorScaleCriteria { get; }

    /// <summary>
    /// 获取或设置数据条对象 (如果此条件是 Databar)
    /// </summary>
    IExcelDataBar DataBar { get; }

    /// <summary>
    /// 获取或设置图标集对象 
    /// </summary>
    IExcelIconSet IconSet { get; }

    /// <summary>
    /// 获取或设置是否显示图标集的过滤器
    /// </summary>
    bool ShowIconOnly { get; set; }

    /// <summary>
    /// 获取或设置是否反向图标集
    /// </summary>
    bool ReverseOrder { get; set; }

    /// <summary>
    /// 获取或设置是否允许 PT 时间 (对于某些高级条件)
    /// </summary>
    bool PTCondition { get; }
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除此条件格式规则
    /// 对应 FormatCondition.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 修改此条件格式规则 (适用于 xlCellValue, xlExpression)
    /// 对应 FormatCondition.Modify 方法
    /// </summary>
    /// <param name="type">条件类型</param>
    /// <param name="operator">比较操作符</param>
    /// <param name="formula1">公式1</param>
    /// <param name="formula2">公式2</param>
    void Modify(int type, int @operator, string formula1, string formula2);

    /// <summary>
    /// 修改此条件格式规则为表达式类型
    /// </summary>
    /// <param name="formula">条件公式</param>
    void ModifyExpression(string formula);
    #endregion

}
