//---------------------------------------------------------------------------
//
// <copyright file="ReaderWriterLockWrapper.cs" company="Microsoft">
//    Copyright (C) Microsoft Corporation.  All rights reserved.
// </copyright>
//
//
// Description:
// Wrapper that allows a ReaderWriterLock to work with C#'s using() clause
//
// History:
//  07/23/2003 : BrendanM Ported to WCP
//
//---------------------------------------------------------------------------


using System;
using System.Threading;
using System.Windows.Threading;
using MS.Internal.WindowsBase;

namespace MS.Internal
{
    // Wrapper that allows a ReaderWriterLock to work with C#'s using() clause
    [FriendAccessAllowed] // Built into Base, used by Core and Framework.
    internal class ReaderWriterLockWrapper
    {
        //------------------------------------------------------
        //
        //  Constructors
        //
        //------------------------------------------------------

        #region Constructors

        internal ReaderWriterLockWrapper()
        {
            _rwLock = new ReaderWriterLock();
            _awr = new AutoWriterRelease(_rwLock);
            _arr = new AutoReaderRelease(_rwLock);
        }

        #endregion Constructors

        //------------------------------------------------------
        //
        //  Internal Properties
        //
        //------------------------------------------------------

        #region Internal Properties

        internal IDisposable WriteLock
        {
            get
            {
                // if there's a dispatcher on this thread, disable its
                // processing.  This avoids unwanted re-entrancy while waiting
                // for the lock (DevDiv2 1177236).
                Dispatcher dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
                if (dispatcher != null)
                {
                    dispatcher._disableProcessingCount++;
                }

                _rwLock.AcquireWriterLock(Timeout.Infinite);
                return _awr;
            }
        }

        internal IDisposable ReadLock
        {
            get
            {
                // if there's a dispatcher on this thread, disable its
                // processing.  This avoids unwanted re-entrancy while waiting
                // for the lock (DevDiv2 1177236).
                Dispatcher dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
                if (dispatcher != null)
                {
                    dispatcher._disableProcessingCount++;
                }

                _rwLock.AcquireReaderLock(Timeout.Infinite);
                return _arr;
            }
        }

        #endregion Internal Properties

        //------------------------------------------------------
        //
        //  Private Fields
        //
        //------------------------------------------------------

        #region Private Fields

        private ReaderWriterLock _rwLock;
        private AutoReaderRelease _arr;
        private AutoWriterRelease _awr;

        #endregion Private Fields

        //------------------------------------------------------
        //
        //  Private Classes & Structs
        //
        //------------------------------------------------------

        #region Private Classes & Structs

        private struct AutoWriterRelease : IDisposable
        {
            public AutoWriterRelease(ReaderWriterLock rwLock)
            {
                _lock = rwLock;
            }

            public void Dispose()
            {
                // if there's a dispatcher on this thread, re-enable its
                // processing.  (DevDiv2 1177236).
                Dispatcher dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
                if (dispatcher != null)
                {
                    dispatcher._disableProcessingCount--;
                }

                _lock.ReleaseWriterLock();
            }

            private ReaderWriterLock _lock;
        }

        private struct AutoReaderRelease : IDisposable
        {
            public AutoReaderRelease(ReaderWriterLock rwLock)
            {
                _lock = rwLock;
            }

            public void Dispose()
            {
                // if there's a dispatcher on this thread, re-enable its
                // processing.  (DevDiv2 1177236).
                Dispatcher dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
                if (dispatcher != null)
                {
                    dispatcher._disableProcessingCount--;
                }

                _lock.ReleaseReaderLock();
            }

            private ReaderWriterLock _lock;
        }
        #endregion Private Classes
    }
}




