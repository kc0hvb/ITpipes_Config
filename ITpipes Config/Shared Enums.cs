using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ITpipes_Config.SharedEnums {

    public enum videoCaptureFormats {
        WMV7,
        WMV9,
        MPEG1,
        MPEG2,
        MPEG4
    }

    public enum videoCaptureResolutions {
        CD,
        DVD
    }

    public enum ValidVideoRenderers {
        Default,
        GDI,
        VMR7,
        VMR9,
        OverlaySurface
    }

    public enum TtsSpeeds {
        Normal,
        Fast,
        Slow
    }

    public enum TextCases {
        lower,
        UPPER,
        Proper,
        No_Restriction
    }

    public enum TextDisplays {
        No_Restriction,
        Code,
        Description,
        Code_And_Description,
        Alias
    }

    public enum VideoCapturePin {
        Composite,
        Svideo
    }

    #region Contact Book

    public enum ContactType {
        Contractor,
        Client,
        None
    }

    public enum ContactSourceType {
        AddressBook,
        TemplateFile
    }


    #endregion Contact Book

}
