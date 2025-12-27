//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示Excel中的XML映射接口，提供对XML数据映射的操作功能
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelXmlMap : IOfficeObject<IExcelXmlMap>, IDisposable
{
    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取一个值，该值指示此映射的数据是否可以导出到XML文件
    /// </summary>
    bool IsExportable { get; }

    /// <summary>
    /// 获取或设置XML映射的名称
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示在导入或导出过程中是否显示验证错误
    /// </summary>
    bool ShowImportExportValidationErrors { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否保存数据源定义
    /// </summary>
    bool SaveDataSourceDefinition { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否根据导入的数据调整列宽
    /// </summary>
    bool AdjustColumnWidth { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否保留现有的列筛选器
    /// </summary>
    bool PreserveColumnFilter { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否保留现有数字格式
    /// </summary>
    bool PreserveNumberFormatting { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示在导入数据时是否追加到现有数据
    /// </summary>
    bool AppendOnImport { get; set; }

    /// <summary>
    /// 获取根元素的名称
    /// </summary>
    string RootElementName { get; }

    /// <summary>
    /// 获取与该映射关联的XML数据绑定对象
    /// </summary>
    IExcelXmlDataBinding DataBinding { get; }

    /// <summary>
    /// 获取根元素的命名空间
    /// </summary>
    IExcelXmlNamespace RootElementNamespace { get; }

    /// <summary>
    /// 获取与该映射关联的XML架构集合
    /// </summary>
    IExcelXmlSchemas Schemas { get; }

    /// <summary>
    /// 获取与该映射关联的工作簿连接对象
    /// </summary>
    IExcelWorkbookConnection WorkbookConnection { get; }

    /// <summary>
    /// 删除当前XML映射
    /// </summary>
    void Delete();

    /// <summary>
    /// 从指定URL导入XML数据
    /// </summary>
    /// <param name="url">要从中导入数据的XML文件的URL</param>
    /// <param name="overwrite">指定是否覆盖现有数据，默认为null</param>
    /// <returns>XML导入操作的结果</returns>
    XlXmlImportResult Import(string url, bool? overwrite = null);

    /// <summary>
    /// 从XML字符串导入数据
    /// </summary>
    /// <param name="xmlData">包含要导入的XML数据的字符串</param>
    /// <param name="overwrite">指定是否覆盖现有数据，默认为null</param>
    /// <returns>XML导入操作的结果</returns>
    XlXmlImportResult ImportXml(string xmlData, bool? overwrite = null);

    /// <summary>
    /// 将数据从此映射导出到指定的URL
    /// </summary>
    /// <param name="url">导出XML文件的目标URL</param>
    /// <param name="overwrite">指定是否覆盖现有文件的对象</param>
    /// <returns>XML导出操作的结果</returns>
    XlXmlExportResult Export(string url, object overwrite);

    /// <summary>
    /// 将数据从此映射导出为XML字符串
    /// </summary>
    /// <param name="data">包含导出的XML数据的字符串</param>
    /// <returns>XML导出操作的结果</returns>
    XlXmlExportResult ExportXml(out string data);

}