//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示活动电子邮件。
/// <para>注：使用 Application.MailMessage 属性可返回 MailMessage 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordMailMessage : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    #endregion


    #region 邮件消息方法 (Mail Message Methods)

    /// <summary>
    /// 检查邮件中的收件人姓名是否有效
    /// </summary>
    void CheckName();


    /// <summary>
    /// 删除当前邮件
    /// </summary>
    void Delete();


    /// <summary>
    /// 显示移动邮件对话框
    /// </summary>
    void DisplayMoveDialog();


    /// <summary>
    /// 显示邮件属性对话框
    /// </summary>
    void DisplayProperties();


    /// <summary>
    /// 显示选择收件人名称对话框
    /// </summary>
    void DisplaySelectNamesDialog();


    /// <summary>
    /// 转发邮件
    /// </summary>
    void Forward();


    /// <summary>
    /// 导航到下一个邮件项目
    /// </summary>
    void GoToNext();


    /// <summary>
    /// 导航到上一个邮件项目
    /// </summary>
    void GoToPrevious();


    /// <summary>
    /// 回复邮件
    /// </summary>
    void Reply();


    /// <summary>
    /// 回复所有收件人
    /// </summary>
    void ReplyAll();


    /// <summary>
    /// 切换邮件头的显示状态
    /// </summary>
    void ToggleHeader();

    #endregion
}