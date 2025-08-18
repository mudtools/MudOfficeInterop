//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel Range 地址接口
/// </summary>
public interface IExcelRangeAddress : IDisposable, IEquatable<IExcelRangeAddress>
{
    // ========== 基本属性 ==========

    /// <summary>起始行号 (1-based)</summary>
    int StartRow { get; }

    /// <summary>起始列号 (1-based)</summary>
    int StartColumn { get; }

    /// <summary>结束行号 (1-based)</summary>
    int EndRow { get; }

    /// <summary>结束列号 (1-based)</summary>
    int EndColumn { get; }

    /// <summary>工作表名称</summary>
    string SheetName { get; }

    /// <summary>工作簿名称</summary>
    string WorkbookName { get; }

    /// <summary>是否为绝对引用 ($A$1 格式)</summary>
    bool IsAbsoluteReference { get; }

    /// <summary>是否为单个单元格</summary>
    bool IsSingleCell { get; }

    /// <summary>单元格行数</summary>
    int RowCount { get; }

    /// <summary>单元格列数</summary>
    int ColumnCount { get; }

    // ========== 地址获取方法 ==========

    /// <summary>获取A1样式的地址</summary>
    string GetAddressA1(
        bool absoluteRow = true,
        bool absoluteCol = true,
        bool includeSheet = true,
        bool includeWorkbook = false);

    /// <summary>获取R1C1样式的地址</summary>
    string GetAddressR1C1(
        bool absoluteRow = true,
        bool absoluteCol = true,
        bool includeSheet = true,
        bool includeWorkbook = false);

    /// <summary>获取本地化地址（使用Excel的本地设置）</summary>
    string GetLocalAddress();

    /// <summary>获取外部引用地址（包含工作簿和工作表）</summary>
    string GetExternalAddress();

    /// <summary>获取起始单元格地址 (A1)</summary>
    string GetStartCellAddress();

    /// <summary>获取结束单元格地址 (A1)</summary>
    string GetEndCellAddress();

    // ========== 地址操作方法 ==========

    /// <summary>检查地址是否包含指定单元格</summary>
    bool ContainsCell(int row, int column);

    /// <summary>检查地址是否包含指定单元格 (A1格式)</summary>
    bool ContainsCell(string address);

    /// <summary>获取偏移后的新地址</summary>
    IExcelRangeAddress Offset(int rowOffset, int colOffset);

    /// <summary>获取调整大小后的新地址</summary>
    IExcelRangeAddress Resize(int rows, int columns);

    /// <summary>转换为标准字符串表示形式</summary>
    string ToString();
}