//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的协同认证锁定对象，用于管理文档中特定范围的锁定状态
/// 协同认证锁定用于在多用户协作编辑文档时控制对特定部分的编辑权限
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCoAuthLock : IOfficeObject<IWordCoAuthLock>, IDisposable
{
    /// <summary>
    /// 获取与当前协同认证锁定关联的Word应用程序实例
    /// </summary>
    /// <value>返回IWordApplication接口的实例，如果未关联则返回null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前协同认证锁定的父对象
    /// </summary>
    /// <value>返回父对象的引用</value>
    object? Parent { get; }

    /// <summary>
    /// 获取协同认证锁定的类型
    /// </summary>
    /// <value>返回WdLockType枚举值，表示锁定的类型</value>
    WdLockType Type { get; }

    /// <summary>
    /// 获取拥有此锁定的协同认证作者
    /// </summary>
    /// <value>返回IWordCoAuthor接口的实例，表示锁定的拥有者</value>
    IWordCoAuthor? Owner { get; }

    /// <summary>
    /// 获取被此协同认证锁定保护的文档范围
    /// </summary>
    /// <value>返回IWordRange接口的实例，表示被锁定的文档范围</value>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取一个值，指示此协同认证锁定是否应用于页眉或页脚
    /// </summary>
    /// <value>如果锁定应用于页眉或页脚则返回true，否则返回false</value>
    bool HeaderFooter { get; }

    /// <summary>
    /// 解除当前协同认证锁定，允许其他用户编辑被锁定的文档范围
    /// </summary>
    void Unlock();
}