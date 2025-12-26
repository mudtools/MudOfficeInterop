//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

partial class WordFont
{
    public bool Bold
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Bold == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Bold = value ? 1 : 0;
        }
    }

    public bool Italic
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Italic == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Italic = value ? 1 : 0;
        }
    }
    public bool Superscript
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Superscript == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Superscript = value ? 1 : 0;
        }
    }

    public bool Subscript
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Subscript == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Subscript = value ? 1 : 0;
        }
    }

    public bool Outline
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Outline == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Outline = value ? 1 : 0;
        }
    }

    ///  <inheritdoc/>
    public bool Shadow
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Shadow == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Shadow = value ? 1 : 0;
        }
    }

    public bool Hidden
    {
        get
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            return _font.Hidden == 1;
        }
        set
        {
            if (_font == null)
                throw new ObjectDisposedException(nameof(_font));
            _font.Hidden = value ? 1 : 0;
        }
    }
}