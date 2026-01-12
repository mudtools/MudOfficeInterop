//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件中。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定Excel中单元格区域的自动格式化选项
/// </summary>
public enum XlRangeAutoFormat
{
    /// <summary>
    /// 经典格式1
    /// </summary>
    xlRangeAutoFormatClassic1 = 1,

    /// <summary>
    /// 经典格式2
    /// </summary>
    xlRangeAutoFormatClassic2 = 2,

    /// <summary>
    /// 经典格式3
    /// </summary>
    xlRangeAutoFormatClassic3 = 3,

    /// <summary>
    /// 会计格式1
    /// </summary>
    xlRangeAutoFormatAccounting1 = 4,

    /// <summary>
    /// 会计格式2
    /// </summary>
    xlRangeAutoFormatAccounting2 = 5,

    /// <summary>
    /// 会计格式3
    /// </summary>
    xlRangeAutoFormatAccounting3 = 6,

    /// <summary>
    /// 颜色格式1
    /// </summary>
    xlRangeAutoFormatColor1 = 7,

    /// <summary>
    /// 颜色格式2
    /// </summary>
    xlRangeAutoFormatColor2 = 8,

    /// <summary>
    /// 颜色格式3
    /// </summary>
    xlRangeAutoFormatColor3 = 9,

    /// <summary>
    /// 列表格式1
    /// </summary>
    xlRangeAutoFormatList1 = 10,

    /// <summary>
    /// 列表格式2
    /// </summary>
    xlRangeAutoFormatList2 = 11,

    /// <summary>
    /// 列表格式3
    /// </summary>
    xlRangeAutoFormatList3 = 12,

    /// <summary>
    /// 3D效果格式1
    /// </summary>
    xlRangeAutoFormat3DEffects1 = 13,

    /// <summary>
    /// 3D效果格式2
    /// </summary>
    xlRangeAutoFormat3DEffects2 = 14,

    /// <summary>
    /// 本地化格式1
    /// </summary>
    xlRangeAutoFormatLocalFormat1 = 15,

    /// <summary>
    /// 本地化格式2
    /// </summary>
    xlRangeAutoFormatLocalFormat2 = 16,

    /// <summary>
    /// 会计格式4
    /// </summary>
    xlRangeAutoFormatAccounting4 = 17,

    /// <summary>
    /// 本地化格式3
    /// </summary>
    xlRangeAutoFormatLocalFormat3 = 19,

    /// <summary>
    /// 本地化格式4
    /// </summary>
    xlRangeAutoFormatLocalFormat4 = 20,

    /// <summary>
    /// 报表格式1
    /// </summary>
    xlRangeAutoFormatReport1 = 21,

    /// <summary>
    /// 报表格式2
    /// </summary>
    xlRangeAutoFormatReport2 = 22,

    /// <summary>
    /// 报表格式3
    /// </summary>
    xlRangeAutoFormatReport3 = 23,

    /// <summary>
    /// 报表格式4
    /// </summary>
    xlRangeAutoFormatReport4 = 24,

    /// <summary>
    /// 报表格式5
    /// </summary>
    xlRangeAutoFormatReport5 = 25,

    /// <summary>
    /// 报表格式6
    /// </summary>
    xlRangeAutoFormatReport6 = 26,

    /// <summary>
    /// 报表格式7
    /// </summary>
    xlRangeAutoFormatReport7 = 27,

    /// <summary>
    /// 报表格式8
    /// </summary>
    xlRangeAutoFormatReport8 = 28,

    /// <summary>
    /// 报表格式9
    /// </summary>
    xlRangeAutoFormatReport9 = 29,

    /// <summary>
    /// 报表格式10
    /// </summary>
    xlRangeAutoFormatReport10 = 30,

    /// <summary>
    /// 经典数据透视表格式
    /// </summary>
    xlRangeAutoFormatClassicPivotTable = 31,

    /// <summary>
    /// 表格格式1
    /// </summary>
    xlRangeAutoFormatTable1 = 32,

    /// <summary>
    /// 表格格式2
    /// </summary>
    xlRangeAutoFormatTable2 = 33,

    /// <summary>
    /// 表格格式3
    /// </summary>
    xlRangeAutoFormatTable3 = 34,

    /// <summary>
    /// 表格格式4
    /// </summary>
    xlRangeAutoFormatTable4 = 35,

    /// <summary>
    /// 表格格式5
    /// </summary>
    xlRangeAutoFormatTable5 = 36,

    /// <summary>
    /// 表格格式6
    /// </summary>
    xlRangeAutoFormatTable6 = 37,

    /// <summary>
    /// 表格格式7
    /// </summary>
    xlRangeAutoFormatTable7 = 38,

    /// <summary>
    /// 表格格式8
    /// </summary>
    xlRangeAutoFormatTable8 = 39,

    /// <summary>
    /// 表格格式9
    /// </summary>
    xlRangeAutoFormatTable9 = 40,

    /// <summary>
    /// 表格格式10
    /// </summary>
    xlRangeAutoFormatTable10 = 41,

    /// <summary>
    /// 数据透视表无格式
    /// </summary>
    xlRangeAutoFormatPTNone = 42,

    /// <summary>
    /// 无格式
    /// </summary>
    xlRangeAutoFormatNone = -4142,

    /// <summary>
    /// 简单格式
    /// </summary>
    xlRangeAutoFormatSimple = -4154
}