//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

using System;
using System.Runtime.InteropServices;

internal class OfficeCommandBarButton : OfficeCommandBarControl, IOfficeCommandBarButton
{
    private MsCore.CommandBarButton _button;

    public int Priority
    {
        get => _button.Priority;
        set => _button.Priority = value;
    }

    public MsoCommandBarButtonHyperlinkType HyperlinkType
    {
        get => (MsoCommandBarButtonHyperlinkType)_button.HyperlinkType;
        set => _button.HyperlinkType = (MsCore.MsoCommandBarButtonHyperlinkType)value;
    }

    public int FaceId
    {
        get => _button.FaceId;
        set => _button.FaceId = value;
    }

    public string ShortcutText
    {
        get => _button.ShortcutText;
        set => _button.ShortcutText = value;
    }

    public MsoButtonState State
    {
        get => (MsoButtonState)(int)_button.State;
        set => _button.State = (MsCore.MsoButtonState)value;
    }

    public MsoButtonStyle Style
    {
        get => (MsoButtonStyle)_button.Style;
        set => _button.Style = (MsCore.MsoButtonStyle)value;
    }

    public string DescriptionText
    {
        get => _button.DescriptionText;
        set => _button.DescriptionText = value;
    }

    public bool BuiltInFace => _button.BuiltInFace;

    public bool IsPriorityDropped
    {
        get => _button.IsPriorityDropped;
    }

    internal OfficeCommandBarButton(MsCore.CommandBarButton button) : base(button)
    {
        _button = button ?? throw new ArgumentNullException(nameof(button));
    }

    public void SetIcon(string imageFile)
    {
        try
        {
            // 这里需要根据实际需求实现图标设置逻辑
            // 可能需要加载图片并转换为IPictureDisp对象
            // 伪代码实现
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法设置按钮图标。", ex);
        }
    }

    public void ResetIcon()
    {
        try
        {
            _button.Reset();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法重置按钮图标。", ex);
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && _button != null)
        {
            try
            {
                while (Marshal.ReleaseComObject(_button) > 0) { }
            }
            catch { }
            _button = null;
        }
        base.Dispose(disposing);
    }
}