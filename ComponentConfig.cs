using System;
using System.Runtime.InteropServices;
using GME.Util;
using GME.MGA;

namespace GME.CSharp
{

    abstract class ComponentConfig
    {
        // Set paradigm name. Provide * if you want to register it for all paradigms.
        public const string paradigmName = "*";

        // Set the human readable name of the addon. You can use white space characters.
        public const string componentName = "Visualizer Integration";

        // Select the object events you want the addon to listen to.
        public const int eventMask = (int)(objectevent_enum.OBJEVENT_OPENMODEL | objectevent_enum.OBJEVENT_CLOSEMODEL);

        // Uncomment the flag if your component is paradigm independent.
        public static componenttype_enum componentType = componenttype_enum.COMPONENTTYPE_ADDON;

        public const regaccessmode_enum registrationMode = regaccessmode_enum.REGACCESS_SYSTEM;
        public const string progID = "MGA.Addon.VisualizerIntegration";
        public const string guid = "39922EC7-3BDB-442D-AC21-D91C5F2893B1";
    }
}
