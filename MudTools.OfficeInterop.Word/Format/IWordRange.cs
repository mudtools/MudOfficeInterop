//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 范围接口
/// </summary>
public interface IWordRange : IDisposable
{
    /// <summary>
    /// 获取或设置范围文本
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置起始位置
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取或设置结束位置
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取范围长度
    /// </summary>
    int Length { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 复制范围内容
    /// </summary>
    void Copy();

    /// <summary>
    /// 删除范围内容
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择范围
    /// </summary>
    void Select();

    /// <summary>
    /// 设置范围
    /// </summary>
    /// <param name="start">起始位置</param>
    /// <param name="end">结束位置</param>
    void SetRange(int start, int end);
}