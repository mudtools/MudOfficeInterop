//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel FormatConditions 集合对象的二次封装实现类
/// 实现 IExcelFormatConditions 接口
/// </summary>
internal partial class ExcelFormatConditions : IExcelFormatConditions
{

    #region 创建和添加
    public IExcelFormatCondition? Add(
        XlFormatConditionType type,
        XlFormatConditionOperator? @operator,
        object? formula1 = null,
        object? formula2 = null,
        object? @string = null,
        object? textOperator = null,
        object? dateOperator = null,
        object? scopeType = null)
    {
        if (_formatConditions == null)
            return null;

        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            (MsExcel.XlFormatConditionType)type,
            GetObject(formula1), GetObject(formula2),
            @string ?? Type.Missing,
            textOperator ?? Type.Missing,
            dateOperator ?? Type.Missing,
            scopeType ?? Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }

    private object? GetObject(object? obj)
    {
        if (obj == null)
            return Type.Missing;
        if (obj is IExcelRange range)
            return range.Address;
        return obj;
    }

    public IExcelFormatCondition? AddExpression(string formula)
    {
        if (_formatConditions == null)
            return null;
        MsExcel.FormatCondition newCondition = (MsExcel.FormatCondition)_formatConditions.Add(
            MsExcel.XlFormatConditionType.xlExpression,
            Type.Missing,
            formula,
            Type.Missing
        );
        return new ExcelFormatCondition(newCondition);
    }
    #endregion

}
