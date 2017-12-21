//------------------------------------------------------------------------------
//  Microsoft Avalon
//  Copyright (c) Microsoft Corporation, All Rights Reserved
//
//  File: BmpBitmapEncoder.cs
//
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Security;
using System.Security.Permissions;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Reflection;
using MS.Internal;
using MS.Win32.PresentationCore;
using System.Diagnostics;
using System.Windows.Media;
using System.Globalization;
using System.Windows.Media.Imaging;

namespace System.Windows.Media.Imaging
{
    #region BmpBitmapEncoder

    /// <summary>
    /// Built-in Encoder for Bmp files.
    /// </summary>
    public sealed class BmpBitmapEncoder : BitmapEncoder
    {
        #region Constructors

        /// <summary>
        /// Constructor for BmpBitmapEncoder
        /// </summary>
        /// <SecurityNote>
        /// Critical - will eventually create unmanaged resources
        /// PublicOK - all inputs are verified
        /// </SecurityNote>
        [SecurityCritical ]
        public BmpBitmapEncoder() :
            base(true)
        {
            _supportsPreview = false;
            _supportsGlobalThumbnail = false;
            _supportsGlobalMetadata = false;
            _supportsFrameThumbnails = false;
            _supportsMultipleFrames = false;
            _supportsFrameMetadata = false;
        }

        #endregion

        #region Internal Properties / Methods

        /// <summary>
        /// Returns the container format for this encoder
        /// </summary>
        /// <SecurityNote>
        /// Critical - uses guid to create unmanaged resources
        /// </SecurityNote>
        internal override Guid ContainerFormat
        {
            [SecurityCritical]
            get
            {
                return _containerFormat;
            }
        }

        /// <summary>
        /// Setups the encoder and other properties before encoding each frame
        /// </summary>
        /// <SecurityNote>
        /// Critical - calls Critical Initialize()
        /// </SecurityNote>
       [SecurityCritical]
        internal override void SetupFrame(SafeMILHandle frameEncodeHandle, SafeMILHandle encoderOptions)
        {
            HRESULT.Check(UnsafeNativeMethods.WICBitmapFrameEncode.Initialize(
                frameEncodeHandle,
                encoderOptions
                ));
        }

        #endregion

        #region Internal Abstract

        /// Need to implement this to derive from the "sealed" object
        internal override void SealObject()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Data Members

        /// <SecurityNote>
        /// Critical - CLSID used for creation of critical resources
        /// </SecurityNote>
        [SecurityCritical]
        private Guid _containerFormat = MILGuidData.GUID_ContainerFormatBmp;

        #endregion
    }

    #endregion // BmpBitmapEncoder
}


