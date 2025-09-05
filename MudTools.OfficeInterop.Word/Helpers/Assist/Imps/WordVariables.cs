//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档变量集合实现类
/// </summary>
internal class WordVariables : IWordVariables
{
    private readonly MsWord.Variables _variables;
    private readonly IWordDocument _document;
    private bool _disposedValue;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _variables != null ? new WordApplication(_variables.Application) : null;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _variables?.Parent;

    /// <summary>
    /// 获取变量数量
    /// </summary>
    public int Count => _variables.Count;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="variables">COM Variables 对象</param>
    /// <param name="document">关联的文档对象</param>
    internal WordVariables(MsWord.Variables variables, IWordDocument document)
    {
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取变量
    /// </summary>
    /// <param name="index">变量索引</param>
    /// <returns>变量对象</returns>
    public IWordVariable Item(int index)
    {
        if (index < 1 || index > Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

        try
        {
            var variable = _variables[index];
            return new WordVariable(variable);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get variable at index {index}.", ex);
        }
    }

    /// <summary>
    /// 根据名称获取变量
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <returns>变量对象</returns>
    public IWordVariable Item(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            var variable = _variables[name];
            return new WordVariable(variable);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to get variable with name '{name}'.", ex);
        }
    }

    /// <summary>
    /// 添加变量
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <param name="value">变量值</param>
    /// <returns>变量对象</returns>
    public IWordVariable Add(string name, string value)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            var variable = _variables.Add(name, value ?? string.Empty);
            return new WordVariable(variable);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add variable '{name}'.", ex);
        }
    }

    /// <summary>
    /// 删除变量
    /// </summary>
    /// <param name="name">变量名称</param>
    public void Delete(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Variable name cannot be null or empty.", nameof(name));

        try
        {
            if (Exists(name))
            {
                _variables[name].Delete();
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete variable '{name}'.", ex);
        }
    }

    /// <summary>
    /// 检查变量是否存在
    /// </summary>
    /// <param name="name">变量名称</param>
    /// <returns>是否存在</returns>
    public bool Exists(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        try
        {

            return _variables[name] != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>变量枚举器</returns>
    public IEnumerator<IWordVariable> GetEnumerator()
    {
        try
        {
            var variables = new List<IWordVariable>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    variables.Add(Item(i));
                }
                catch
                {
                    // 忽略获取失败的变量
                    continue;
                }
            }
            return variables.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate variables.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
