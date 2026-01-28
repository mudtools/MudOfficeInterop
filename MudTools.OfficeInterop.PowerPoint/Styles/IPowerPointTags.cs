//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 标签集合接口
/// </summary>
public interface IPowerPointTags : IDisposable
{
    /// <summary>
    /// 获取标签数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 根据索引获取标签值
    /// </summary>
    string this[int index] { get; }

    /// <summary>
    /// 根据名称获取标签值
    /// </summary>
    string this[string name] { get; }


    /// <summary>
    /// 获取标签名称
    /// </summary>
    /// <param name="index">标签索引</param>
    /// <returns>标签名称</returns>
    string Name(int index);

    /// <summary>
    /// 获取标签值
    /// </summary>
    /// <param name="index">标签索引</param>
    /// <returns>标签值</returns>
    string Value(int index);

    /// <summary>
    /// 添加标签
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <param name="value">标签值</param>
    void Add(string name, string value);

    /// <summary>
    /// 删除标签
    /// </summary>
    /// <param name="name">标签名称</param>
    void Delete(string name);

    /// <summary>
    /// 清除所有标签
    /// </summary>
    void Clear();

    /// <summary>
    /// 检查标签是否存在
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <returns>是否存在</returns>
    bool Contains(string name);

    /// <summary>
    /// 更新标签值
    /// </summary>
    /// <param name="name">标签名称</param>
    /// <param name="value">新标签值</param>
    void Update(string name, string value);

    /// <summary>
    /// 获取所有标签名称
    /// </summary>
    /// <returns>标签名称列表</returns>
    IEnumerable<string> GetAllNames();

    /// <summary>
    /// 获取所有标签键值对
    /// </summary>
    /// <returns>标签键值对字典</returns>
    IDictionary<string, string> GetAllTags();

    /// <summary>
    /// 导出标签到文件
    /// </summary>
    /// <param name="fileName">文件路径</param>
    void ExportToFile(string fileName);

    /// <summary>
    /// 从文件导入标签
    /// </summary>
    /// <param name="fileName">文件路径</param>
    void ImportFromFile(string fileName);

    /// <summary>
    /// 查找符合名称条件的标签
    /// </summary>
    /// <param name="namePredicate">名称条件</param>
    /// <returns>符合条件的标签列表</returns>
    IEnumerable<KeyValuePair<string, string>> FindByName(Func<string, bool> namePredicate);

    /// <summary>
    /// 查找符合值条件的标签
    /// </summary>
    /// <param name="valuePredicate">值条件</param>
    /// <returns>符合条件的标签列表</returns>
    IEnumerable<KeyValuePair<string, string>> FindByValue(Func<string, bool> valuePredicate);

    /// <summary>
    /// 获取标签集合信息
    /// </summary>
    /// <returns>标签集合信息字符串</returns>
    string GetTagsInfo();
}
