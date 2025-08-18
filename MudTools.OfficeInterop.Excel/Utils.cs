//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Excel.Imps;
using log4net;

namespace MudTools.OfficeInterop.Excel;

internal class Utils
{
    /// <summary>
    /// 用于记录此类型运行时日志的 logger 实例。
    /// </summary>
    private static readonly ILog log = LogManager.GetLogger(typeof(Utils));

    /// <summary>
    /// 详细识别Selection对象类型
    /// </summary>
    /// <param name="selection">Selection对象</param>
    /// <returns>类型信息</returns>
    public static object? CreateSelectionType(object selection)
    {
        try
        {
            return selection switch
            {
                MsExcel.Range range => new ExcelRange(range),
                MsExcel.ChartObject chartObject => new ExcelChartObject(chartObject),
                MsExcel.Shape shape => new ExcelShape(shape),
                MsExcel.Chart chart => new ExcelChart(chart),
                MsExcel.PivotTable pivotTable => new ExcelPivotTable(pivotTable),
                MsExcel.ListObject listObject => new ExcelListObject(listObject),
                MsExcel.Comment comment => new ExcelComment(comment),
                MsExcel.DrawingObjects drawingObjs => new ExcelDrawingObjects(drawingObjs),
                _ => selection,// 如果以上类型都不匹配，直接返回原始COM对象。
            };
        }
        catch (COMException cx)
        {
            log.Error("识别Selection对象类型失败：" + cx.Message, cx);
            return null;
        }
        catch (Exception ex)
        {
            log.Error("识别Selection对象类型失败：" + ex.Message, ex);
            return null;
        }
    }
    public static IExcelControl? CreateControl(object? comObj, XlFormControl xlFormControl)
    {
        if (comObj == null)
            return null;

        try
        {
            switch (xlFormControl)
            {
                case XlFormControl.xlButtonControl:
                    break;
                case XlFormControl.xlCheckBox:
                    if (comObj is MsExcel.CheckBox checkbox)
                    {
                        var t = new ExcelCheckBox(checkbox);
                        return t;
                    }
                    break;
                case XlFormControl.xlDropDown:
                    if (comObj is MsExcel.DropDown dropDwon)
                    {
                        var t = new ExcelDropDown(dropDwon);
                        return t;
                    }
                    break;
                case XlFormControl.xlEditBox:
                    if (comObj is MsExcel.EditBox editBox)
                    {
                        var t = new ExcelEditBox(editBox);
                        return t;
                    }
                    break;
                case XlFormControl.xlGroupBox:
                    break;
                case XlFormControl.xlLabel:
                    break;
                case XlFormControl.xlListBox:
                    if (comObj is MsExcel.ListBox listBox)
                    {
                        var t = new ExcelListBox(listBox);
                        return t;
                    }
                    break;
                case XlFormControl.xlOptionButton:
                    break;
                case XlFormControl.xlScrollBar:
                    break;
                case XlFormControl.xlSpinner:
                    break;
                default:
                    break;
            }
            return null;
        }
        catch (COMException cx)
        {
            log.Error("创建Excel控件对象类型失败：" + cx.Message, cx);
            return null;
        }
        catch (Exception ex)
        {
            log.Error("创建Excel控件对象类型失败：" + ex.Message, ex);
            return null;
        }

    }
}
