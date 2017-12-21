//---------------------------------------------------------------------------
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
//---------------------------------------------------------------------------

using System;

#if WINDOWS_BASE
namespace MS.Internal.WindowsBase.Interop
#elif PRESENTATION_CORE
namespace System.Windows.Interop
#elif PRESENTATIONFRAMEWORK
namespace MS.Internal.PresentationFramework.Interop
#elif REACHFRAMEWORK
namespace MS.Internal.ReachFramework.Interop
#else
namespace Microsoft.Internal.Interop
#endif
{
    // This is not a complete list of operating system versions and service packs.
    // These are the interesting versions where features or behaviors were introduced 
    // or changed, and code needs to detect those points and do different things.
    //
    // This list has been expanded in order to support our new OSVersionHelper lib.
    //
    // If you need to add an OS, do the following steps:
    //     Create the appropriate native function in OSVersionHelper (see shared/OSVersionHelper/inc/OSVersionHelper.cpp)
    //     Create the appropriate export in PresentationNative (see native/dll/dllentry.cpp)
    //     Add it to the OperatingSystemVersion enumeration (Congrats, you're here already!)
    //     Create the appropriate managed wrappers (see OSVersionHelper.cs)
    //     Detect your freshly minted OS!
    internal enum OperatingSystemVersion
    {
        /// <summary>
        /// WPF minimum requirement.
        /// </summary>
        WindowsXPSP2,

        /// <summary>
        /// DevDiv:1158540
        /// Added for completeness
        /// </summary>
        WindowsXPSP3,

        /// <summary>
        /// Introduced Aero glass, DWM, new common dialogs.
        /// </summary>
        WindowsVista,

        /// <summary>
        /// DevDiv:1158540
        /// Added for completeness
        /// </summary>
        WindowsVistaSP1,

        /// <summary>
        /// DevDiv:1158540
        /// Added for completeness
        /// </summary>
        WindowsVistaSP2,

        /// <summary>
        /// Introduced multi-touch.
        /// </summary>
        Windows7,

        /// <summary>
        /// DevDiv:1158540
        /// Added for completeness
        /// </summary>
        Windows7SP1,

        /// <summary>
        /// Introduced feature on demand
        /// </summary>
        Windows8,

        /// <summary>
        /// Downlevel font fallback differentiation
        /// </summary>
        Windows8Point1,

        /// <summary>
        /// New font fallbacks
        /// </summary>
        Windows10,
    }
}
