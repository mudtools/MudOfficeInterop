//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;
using MudTools.OfficeInterop.Word.Imps;

namespace MudTools.OfficeInterop;

/// <summary>
/// Word应用程序的静态工厂类，提供创建和操作Word文档的便捷方法
/// </summary>
public static class WordFactory
{
    /// <summary>
    /// 通过COM对象创建Word应用程序实例
    /// </summary>
    /// <param name="comObj">COM对象，应为Microsoft Word Application对象</param>
    /// <returns>如果comObj是有效的Word Application对象，则返回封装后的IWordApplication实例；否则返回null</returns>
    public static IWordApplication? Connection(object comObj)
    {
        MsWord.Application? wordCom = comObj as MsWord.Application;
        if (wordCom == null)
            return null;
        return new WordApplication(wordCom);
    }

    /// <summary>
    /// 创建一个新的空白Word文档
    /// </summary>
    /// <returns>返回实现了IWordApplication接口的Word应用程序实例</returns>
    /// <remarks>
    /// 此方法会启动Word应用程序并创建一个空白文档
    /// </remarks>
    public static IWordApplication BlankDocument()
    {
        // 创建WordApplication实例，这会启动Word应用程序
        var word = new WordApplication();

        // 调用BlankDocument方法创建空白文档，使用下划线忽略返回值
        // 下划线表示我们不关心这个方法的返回值，只执行创建文档的操作
        _ = word.BlankDocument();

        // 返回配置好的Word应用程序实例
        return word;
    }

    /// <summary>
    /// 基于指定模板创建新的Word文档
    /// </summary>
    /// <param name="templatePath">模板文件的完整路径</param>
    /// <returns>返回实现了IWordApplication接口的Word应用程序实例</returns>
    /// <exception cref="ArgumentNullException">当templatePath为null时抛出</exception>
    /// <exception cref="FileNotFoundException">当指定的模板文件不存在时抛出</exception>
    /// <remarks>
    /// 此方法会启动Word应用程序并基于模板创建新文档
    /// 新文档会继承模板的格式、样式和内容
    /// </remarks>
    public static IWordApplication CreateFrom(string templatePath)
    {
        // 创建WordApplication实例，这会启动Word应用程序
        var word = new WordApplication();

        // 调用CreateFrom方法基于模板创建文档，使用下划线忽略返回值
        // 该方法会加载指定路径的模板文件并创建基于该模板的新文档
        _ = word.CreateFrom(templatePath);

        // 返回配置好的Word应用程序实例
        return word;
    }

    /// <summary>
    /// 打开现有的Word文档文件
    /// </summary>
    /// <param name="filePath">要打开的Word文档文件的完整路径</param>
    /// <returns>返回实现了IWordApplication接口的Word应用程序实例</returns>
    /// <exception cref="ArgumentNullException">当filePath为null时抛出</exception>
    /// <exception cref="FileNotFoundException">当指定的文件不存在时抛出</exception>
    /// <remarks>
    /// 此方法会启动Word应用程序并打开指定的现有文档
    /// 文档将以可编辑模式打开
    /// </remarks>
    public static IWordApplication Open(string filePath)
    {
        // 创建WordApplication实例，这会启动Word应用程序
        var word = new WordApplication();

        // 调用Open方法打开指定路径的文档，使用下划线忽略返回值
        // 该方法会加载并显示指定的Word文档
        _ = word.Open(filePath);

        // 返回配置好的Word应用程序实例
        return word;
    }
}