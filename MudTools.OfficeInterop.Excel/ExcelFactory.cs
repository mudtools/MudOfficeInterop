//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Excel;
using MudTools.OfficeInterop.Excel.Imps;

namespace MudTools.OfficeInterop;

/// <summary>
/// Excel应用程序的静态工厂类，提供创建和操作Excel工作簿的便捷方法
/// </summary>
/// <remarks>
/// 这个工厂类封装了Excel应用程序的创建过程，提供了三种主要的工作簿创建方式：
/// 1. 创建空白工作簿
/// 2. 基于模板创建工作簿
/// 3. 打开现有工作簿文件
/// 所有方法都返回统一的IExcelApplication接口，便于统一管理和操作
/// </remarks>
public static class ExcelFactory
{
    /// <summary>
    /// 创建Office COM对象的包装器实例
    /// 此方法通过反射查找与接口T对应的实现类，并将COM对象包装为强类型的接口实例
    /// </summary>
    /// <typeparam name="T">Office对象接口类型，必须实现IOfficeObject&lt;T&gt;接口</typeparam>
    /// <param name="comObj">原始的COM对象，将被包装为接口T的实例</param>
    /// <returns>接口T的实现实例，如果无法创建则返回默认值(null或类型的默认值)</returns>
    public static T? Create<T>(object comObj) where T : IOfficeObject<T>
    {
        return ObjectExtensions.Create<T>(comObj);
    }

    /// <summary>
    /// 通过COM对象连接到现有的Excel应用程序实例
    /// </summary>
    /// <param name="comObj">COM对象，应为Excel应用程序实例</param>
    /// <returns>如果comObj是有效的Excel应用程序实例，则返回封装的IExcelApplication对象；否则返回null</returns>
    public static IExcelApplication? Connection(object comObj)
    {
        MsExcel.Application? excelCom = comObj as MsExcel.Application;
        if (excelCom == null)
            return null;
        return new ExcelApplication(excelCom);
    }

    /// <summary>
    /// 根据ProgID创建Excel应用程序的新实例
    /// </summary>
    /// <param name="progId">Excel应用程序的ProgID，如果为null则可能引发异常</param>
    /// <returns>返回新创建的Excel应用程序实例</returns>
    /// <exception cref="InvalidOperationException">当无法从指定的ProgID获取类型时抛出</exception>
    public static IExcelApplication CreateInstance(string? progId)
    {
        Type type = Type.GetTypeFromProgID(progId);
        if (type == null)
        {
            throw new InvalidOperationException($"无法从 ProgID '{progId}' 获取类型。");
        }

        MsExcel.Application instance = (MsExcel.Application)Activator.CreateInstance(type);
        instance.UserControl = true; // 允许用户控制该实例
        ExcelApplication excel = new(instance);
        return excel;
    }

    /// <summary>
    /// 创建一个新的空白Excel工作簿
    /// </summary>
    /// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
    /// <example>
    /// <code>
    /// // 创建新的空白工作簿
    /// var excelApp = ExcelFactory.BlankWorkbook();
    /// // 现在可以对工作簿进行操作
    /// excelApp.GetActiveSheet().Cells[1, 1].Value = "Hello World";
    /// </code>
    /// </example>
    /// <remarks>
    /// 此方法会启动Excel应用程序并创建一个包含一个工作表的空白工作簿
    /// 使用下划线(_)忽略BlankWorkbook方法的返回值，因为我们只关心创建工作簿的副作用
    /// </remarks>
    public static IExcelApplication BlankWorkbook()
    {
        // 创建ExcelApplication实例，这会启动Excel应用程序进程
        ExcelApplication excel = new ExcelApplication();

        // 调用BlankWorkbook方法创建空白工作簿，使用下划线忽略返回值
        // 下划线表示我们不关心这个方法的返回值，只执行创建工作簿的操作
        _ = excel.BlankWorkbook();

        // 返回已配置好空白工作簿的Excel应用程序实例
        return excel;
    }

    /// <summary>
    /// 基于指定模板创建新的Excel工作簿
    /// </summary>
    /// <param name="templatePath">Excel模板文件的完整路径（.xltx, .xltm等格式）</param>
    /// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
    /// <exception cref="ArgumentNullException">当templatePath参数为null时抛出</exception>
    /// <exception cref="FileNotFoundException">当指定的模板文件不存在时抛出</exception>
    /// <exception cref="InvalidOperationException">当模板文件格式无效时抛出</exception>
    /// <example>
    /// <code>
    /// // 基于模板创建工作簿
    /// var excelApp = ExcelFactory.CreateFrom(@"C:\Templates\ReportTemplate.xltx");
    /// // 新工作簿将继承模板的格式、样式、公式等
    /// </code>
    /// </example>
    /// <remarks>
    /// 此方法会启动Excel应用程序并基于指定模板创建新工作簿
    /// 新工作簿会继承模板的格式设置、样式、命名区域、宏等特性
    /// 模板文件不会被修改，而是作为基础创建新的工作簿
    /// </remarks>
    public static IExcelApplication CreateFrom(string templatePath)
    {
        // 创建ExcelApplication实例，初始化Excel应用程序环境
        var excel = new ExcelApplication();

        // 调用CreateFrom方法基于模板创建工作簿，使用下划线忽略返回值
        // 该方法会加载指定路径的模板文件并创建基于该模板的新工作簿
        _ = excel.CreateFrom(templatePath);

        // 返回已加载模板内容的Excel应用程序实例
        return excel;
    }

    /// <summary>
    /// 打开现有的Excel工作簿文件
    /// </summary>
    /// <param name="filePath">要打开的Excel文件的完整路径（.xlsx, .xls等格式）</param>
    /// <returns>返回实现了IExcelApplication接口的Excel应用程序实例</returns>
    /// <exception cref="ArgumentNullException">当filePath参数为null时抛出</exception>
    /// <exception cref="FileNotFoundException">当指定的Excel文件不存在时抛出</exception>
    /// <exception cref="InvalidOperationException">当文件损坏或格式不支持时抛出</exception>
    /// <example>
    /// <code>
    /// // 打开现有工作簿
    /// var excelApp = ExcelFactory.Open(@"C:\Data\SalesReport.xlsx");
    /// // 现在可以读取和修改现有数据
    /// var value = excelApp.GetActiveSheet().Cells[1, 1].Value;
    /// </code>
    /// </example>
    /// <remarks>
    /// 此方法会启动Excel应用程序并打开指定的现有工作簿文件
    /// 工作簿将以可编辑模式打开，可以进行读写操作
    /// 如果文件被其他程序占用，可能会抛出相应的异常
    /// </remarks>
    public static IExcelApplication Open(string filePath)
    {
        // 创建ExcelApplication实例，准备Excel应用程序运行环境
        var excel = new ExcelApplication();

        // 调用Open方法打开指定路径的Excel文件，使用下划线忽略返回值
        // 该方法会加载并显示指定的Excel工作簿文件
        _ = excel.Open(filePath);

        // 返回已加载指定文件的Excel应用程序实例
        return excel;
    }
}