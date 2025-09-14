//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel
{
    /// <summary>
    /// Excel工作表集合接口 (适用于 Sheets 对象)
    /// </summary>
    public interface IExcelSheets : IExcelCommonSheets
    {
        #region 操作方法
        /// <summary>
        /// 将此 Sheets 集合中的所有工作表复制到指定位置。
        /// 这是对 Microsoft.Office.Interop.Excel.Sheets.Copy 方法的封装。
        /// </summary>
        /// <param name="beforeSheet">
        /// 指定应在哪个工作表之前放置复制的工作表。
        /// 如果为 null，则不指定此参数。
        /// </param>
        /// <param name="afterSheet">
        /// 指定应在哪个工作表之后放置复制的工作表。
        /// 如果为 null，则不指定此参数。
        /// </param>
        /// <exception cref="System.InvalidOperationException">
        /// 如果内部的 Sheets 对象为 null。
        /// </exception>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// 如果与 Excel 的交互失败（例如，参数无效，工作表被保护），可能会抛出 COM 异常。
        /// </exception>
        /// <remarks>
        /// 如果 beforeSheet 和 afterSheet 都为 null，则 Excel 通常会创建一个新工作簿来容纳复制的工作表。
        /// 如果同时指定了 beforeSheet 和 afterSheet，行为可能不确定（通常 After 会被忽略）。
        /// </remarks>
        void CopyTo(IExcelCommonSheet? beforeSheet = null, IExcelCommonSheet? afterSheet = null);
        /// <summary>
        /// 将此 Sheets 集合中的所有工作表移动到指定位置。
        /// 这是对 Microsoft.Office.Interop.Excel.Sheets.Move 方法的封装。
        /// </summary>
        /// <param name="beforeSheet">
        /// 指定应在哪个工作表之前放置移动的工作表。
        /// 如果为 null，则不指定此参数。
        /// </param>
        /// <param name="afterSheet">
        /// 指定应在哪个工作表之后放置移动的工作表。
        /// 如果为 null，则不指定此参数。
        /// </param>
        /// <exception cref="System.InvalidOperationException">
        /// 如果内部的 Sheets 对象为 null。
        /// </exception>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// 如果与 Excel 的交互失败（例如，参数无效，工作表被保护），可能会抛出 COM 异常。
        /// </exception>
        /// <remarks>
        /// 如果 beforeSheet 和 afterSheet 都为 null，行为可能不确定（可能移动到新工作簿或失败）。
        /// 如果同时指定了 beforeSheet 和 afterSheet，行为可能不确定（通常 After 会被忽略）。
        /// </remarks>
        void MoveTo(IExcelCommonSheet? beforeSheet = null, IExcelCommonSheet? afterSheet = null);

        /// <summary>
        /// 将指定区域的内容和格式填充到此 Sheets 集合中所有工作表的对应区域。
        /// 这是对 Microsoft.Office.Interop.Excel.Sheets.FillAcrossSheets 方法的封装。
        /// </summary>
        /// <param name="sourceRange">
        /// 代表要填充的源区域的 ExcelRange 对象。
        /// </param>
        /// <param name="fillType">
        /// 指定要填充的内容类型（全部、仅内容、仅格式）。
        /// </param>
        /// <exception cref="System.ArgumentNullException">
        /// 如果 sourceRange 为 null。
        /// </exception>
        /// <exception cref="System.InvalidOperationException">
        /// 如果内部的 Sheets 对象为 null。
        /// </exception>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// 如果与 Excel 的交互失败（例如，源区域无效，工作表被保护），可能会抛出 COM 异常。
        /// </exception>
        void FillAcrossSheets(IExcelRange sourceRange, XlFillWith? fillType = XlFillWith.xlFillWithAll);

        /// <summary>
        /// 删除所有工作表 (注意：Excel通常不允许删除所有工作表)
        /// </summary>
        void Clear();
        #endregion

        #region 导出和导入
        /// <summary>
        /// 导出所有工作表到单独的文件
        /// </summary>
        /// <param name="folderPath">导出文件夹路径</param>
        /// <param name="fileFormat">文件格式 (例如 "xlsx", "xls")</param>
        /// <param name="prefix">文件名前缀</param>
        /// <returns>成功导出的工作表数量</returns>
        int ExportToFolder(string folderPath, string fileFormat = "xlsx", string prefix = "sheet_");

        #endregion

        #region 高级功能
        /// <summary>
        /// 隐藏所有工作表
        /// </summary>
        void HideAll();

        /// <summary>
        /// 显示所有工作表
        /// </summary>
        void ShowAll();
        #endregion
    }
}
