//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的共同作者锁定集合，用于管理文档中各种范围的锁定状态
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordCoAuthLocks : IEnumerable<IWordCoAuthLock?>, IOfficeObject<IWordCoAuthLocks>, IDisposable
{
    /// <summary>
    /// 获取与当前锁定集合关联的Word应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取当前锁定集合的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中锁定的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 根据索引获取指定的共同作者锁定对象
    /// </summary>
    /// <param name="index">锁定对象的索引，从0开始</param>
    /// <returns>指定位置的共同作者锁定对象，如果不存在则返回null</returns>
    IWordCoAuthLock? this[int index] { get; }

    /// <summary>
    /// 在指定的文档范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的文档范围</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordRange range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 在指定的段落范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的段落</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordParagraph range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 在指定的列范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的列</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordColumn range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 在指定的单元格范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的单元格</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordCell range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 在指定的行范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的行</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordRow range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 在指定的表格范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的表格</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordTable range, WdLockType type = WdLockType.wdLockReservation);


    /// <summary>
    /// 在指定的选择范围内添加一个共同作者锁定
    /// </summary>
    /// <param name="range">要锁定的选择范围</param>
    /// <param name="type">锁定类型，默认为保留锁定</param>
    /// <returns>创建的锁定对象，如果创建失败则返回null</returns>
    IWordCoAuthLock? Add(IWordSelection range, WdLockType type = WdLockType.wdLockReservation);

    /// <summary>
    /// 移除临时锁定，这些锁定通常是临时性的，不会持久保存
    /// </summary>
    void RemoveEphemeralLocks();
}