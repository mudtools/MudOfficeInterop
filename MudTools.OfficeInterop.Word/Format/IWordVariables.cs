//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// Word 文档变量集合接口
/// </summary>
public interface IWordVariables : IDisposable, IEnumerable<IWordVariable>
{
    /// <summary>
    /// 获取变量数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取变量
    /// </summary>
    IWordVariable Item(int index);

    /// <summary>
    /// 根据名称获取变量
    /// </summary>
    IWordVariable Item(string name);

    /// <summary>
    /// 添加变量
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <param name="value">变量值</param>
    /// <returns>变量对象</returns>
    IWordVariable Add(string name, string value);

    /// <summary>
    /// 删除变量
    /// </summary>
    /// <param name="name">变量名称</param>
    void Delete(string name);

    /// <summary>
    /// 检查变量是否存在
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <returns>是否存在</returns>
    bool Exists(string name);
}