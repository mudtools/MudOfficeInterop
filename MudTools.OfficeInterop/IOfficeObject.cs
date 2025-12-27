//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 定义Office对象的通用接口，用于将COM对象加载为托管对象
/// </summary>
/// <typeparam name="T">实现此接口的类型，遵循自我引用的泛型模式</typeparam>
public interface IOfficeObject<T> where T : IOfficeObject<T>
{
    /// <summary>
    /// 从指定的COM对象加载并创建相应的Office对象实例
    /// </summary>
    /// <param name="comObject">原始的COM对象，通常是从Office应用程序获取的底层对象</param>
    /// <returns>转换后的Office对象实例，如果转换失败则返回null</returns>
    T? LoadFromObject(object comObject);
}