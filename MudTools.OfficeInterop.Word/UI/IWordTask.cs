//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示当前在系统上运行的应用程序任务。
/// <para>注：Task 对象是 Tasks 集合的成员。</para>
/// <para>注：使用 Tasks(index) 可返回单个 Task 对象，其中 index 是应用程序名称或索引号。</para>
/// </summary>
public interface IWordTask : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取或设置任务窗口的水平位置（以磅为单位）。
    /// </summary>
    int Left { get; set; }

    /// <summary>
    /// 获取或设置任务窗口的垂直位置（以磅为单位）。
    /// </summary>
    int Top { get; set; }

    /// <summary>
    /// 获取或设置任务窗口的宽度（以磅为单位）。
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 获取或设置任务窗口的高度（以磅为单位）。
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 获取或设置任务窗口的状态。
    /// </summary>
    WdWindowState WindowState { get; set; }

    /// <summary>
    /// 获取任务的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 激活指定的任务。 对应于“窗口”菜单底部的命令（在任务列表中）。
    /// </summary>
    void Activate();

    void Move(int Left, int Top);

    void Resize(int Width, int Height);


    void SendWindowMessage(int Message, int wParam, int lParam);

    /// <summary>
    /// 关闭指定的任务。
    /// </summary>
    void Close();

}