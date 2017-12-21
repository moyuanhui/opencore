//------------------------------------------------------------------------------
// <copyright file="LocalAppContextSwitches.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>                                                                
//------------------------------------------------------------------------------

namespace System.Windows.Forms {
    using System.Runtime.CompilerServices;

    internal static class LocalAppContextSwitches {
        internal const string DontSupportReentrantFilterMessageSwitchName = @"Switch.System.Windows.Forms.DontSupportReentrantFilterMessage";
        internal const string DoNotSupportSelectAllShortcutInMultilineTextBoxSwitchName = @"Switch.System.Windows.Forms.DoNotSupportSelectAllShortcutInMultilineTextBox";
        private static int _dontSupportReentrantFilterMessage;
        private static int _doNotSupportSelectAllShortcutInMultilineTextBox;

        public static bool DontSupportReentrantFilterMessage {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get {
                return LocalAppContext.GetCachedSwitchValue(LocalAppContextSwitches.DontSupportReentrantFilterMessageSwitchName, ref _dontSupportReentrantFilterMessage);
            }
        }

        public static bool DoNotSupportSelectAllShortcutInMultilineTextBox {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get {
                return LocalAppContext.GetCachedSwitchValue(LocalAppContextSwitches.DoNotSupportSelectAllShortcutInMultilineTextBoxSwitchName, ref _doNotSupportSelectAllShortcutInMultilineTextBox);
            }
        }

    }
}
