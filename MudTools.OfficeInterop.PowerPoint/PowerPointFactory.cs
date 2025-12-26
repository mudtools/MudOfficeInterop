//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.PowerPoint;
using MudTools.OfficeInterop.PowerPoint.Imps;

namespace MudTools.OfficeInterop;

/// <summary>
/// PowerPoint应用程序的静态工厂类，提供创建和操作PowerPoint文档的便捷方法
/// </summary>
public static class PowerPointFactory
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
    /// 通过COM对象创建PowerPoint应用程序实例
    /// </summary>
    /// <param name="comObj">COM对象，应为PowerPoint.Application类型的实例</param>
    /// <returns>返回实现了IPowerPointApplication接口的PowerPoint应用程序实例，如果comObj不是有效的PowerPoint应用程序COM对象则返回null</returns>
    public static IPowerPointApplication? Connection(object comObj)
    {
        MsPowerPoint.Application? powerPointCom = comObj as MsPowerPoint.Application;
        if (powerPointCom == null)
            return null;
        return new PowerPointApplication(powerPointCom);
    }

    /// <summary>
    /// 创建一个新的空白PowerPoint文档
    /// </summary>
    /// <returns>返回实现了IPowerPointApplication接口的PowerPoint应用程序实例</returns>
    /// <remarks>
    /// 此方法会启动PowerPoint应用程序并创建一个空白文档
    /// </remarks>
    public static IPowerPointApplication BlankDocument()
    {
        // 创建PowerPointApplication实例，这会启动PowerPoint应用程序
        var ppt = new PowerPointApplication();

        // 调用BlankDocument方法创建空白文档，使用下划线忽略返回值
        // 下划线表示我们不关心这个方法的返回值，只执行创建文档的操作
        _ = ppt.BlankDocument();

        // 返回配置好的PowerPoint应用程序实例
        return ppt;
    }


    /// <summary>
    /// 打开现有的PowerPoint文档文件
    /// </summary>
    /// <param name="filePath">要打开的PowerPoint文档文件的完整路径</param>
    /// <returns>返回实现了IPowerPointApplication接口的PowerPoint应用程序实例</returns>
    /// <exception cref="ArgumentNullException">当filePath为null时抛出</exception>
    /// <exception cref="FileNotFoundException">当指定的文件不存在时抛出</exception>
    /// <remarks>
    /// 此方法会启动PowerPoint应用程序并打开指定的现有文档
    /// 文档将以可编辑模式打开
    /// </remarks>
    public static IPowerPointApplication Open(string filePath)
    {
        // 创建PowerPointApplication实例，这会启动PowerPoint应用程序
        var ppt = new PowerPointApplication();

        // 调用Open方法打开指定路径的文档，使用下划线忽略返回值
        // 该方法会加载并显示指定的PowerPoint文档
        _ = ppt.OpenPresentation(filePath);

        // 返回配置好的PowerPoint应用程序实例
        return ppt;
    }
}