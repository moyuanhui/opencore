//---------------------------------------------------------------------------
//
// File: WinRTSpellerInterop.cs
//
// Description: Custom COM marshalling code and interfaces for interaction
//                  with the WinRT wordbreaker API and ISpellChecker 
//                  spell-checker API
//
//---------------------------------------------------------------------------

namespace System.Windows.Documents
{

    using MS.Internal;
    using MS.Internal.WindowsRuntime.Windows.Data.Text;

    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.CompilerServices;
    using System.Security;
    using System.Security.Permissions;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;

    internal class WinRTSpellerInterop: SpellerInteropBase
    {
        #region Constructors

        /// <exception cref="PlatformNotSupportedException">
        /// The OS platform is not supported
        /// </exception>
        /// <exception cref="NotSupportedException">
        /// The OS platform is supportable, but spellchecking services are currently unavailable
        /// </exception>
        /// <SecurityNote>
        /// Critical:
        ///     Asserts permissions, instantiates COM objects
        /// Safe:
        ///     Takes no input, does not give the caller access to any 
        ///     critical (COM) resources directly.
        /// </SecurityNote>
        [SecuritySafeCritical]
        internal WinRTSpellerInterop()
        {
            // When the CLR consumes an unmanaged COM object, it invokes 
            // System.ComponentModel.LicenseManager.LicenseInteropHelper.GetCurrentContextInfo
            // which in turn calls Assembly.GetName. Assembly.GetName requires FileIOPermission for
            // access to the path of the assembly. 
            FileIOPermission fiop = new FileIOPermission(PermissionState.None);
            fiop.AllLocalFiles = FileIOPermissionAccess.PathDiscovery;
            fiop.Assert();

            try
            {
                _spellCheckerFactory = new SpellCheckerFactory();
            }
            catch (Exception ex)
                // Sometimes, InvalidCastException is thrown when SpellCheckerFactory fails to instantiate correctly
                when (ex is InvalidCastException || ex is COMException ) 
            {
                Dispose();
                throw new PlatformNotSupportedException(string.Empty, ex);
            }
            finally 
            {
                CodeAccessPermission.RevertAssert();
            }

            _spellCheckers = new Dictionary<CultureInfo, Tuple<WordsSegmenter, ISpellChecker>>();
            _customDictionaryFiles = new Dictionary<CultureInfo, List<string>>();

            _defaultCulture = InputLanguageManager.Current?.CurrentInputLanguage ?? Thread.CurrentThread.CurrentCulture;
            _culture = null;

            _customDictionaryFilesLock = new Semaphore(1, 1);
            _spellCheckerFactoryLock = new ReaderWriterLockSlimWrapper(LockRecursionPolicy.NoRecursion);

            try
            {
                EnsureWordBreakerAndSpellCheckerForCulture(_defaultCulture, throwOnError: true);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PlatformNotSupportedException)
            {
                _spellCheckers = null;
                Dispose();

                if ((ex is PlatformNotSupportedException) || (ex is NotSupportedException))
                {
                    throw;
                }
                else
                {
                    throw new NotSupportedException(string.Empty, ex);
                }
            }

            WeakEventManager<AppDomain, UnhandledExceptionEventArgs>
                .AddHandler(AppDomain.CurrentDomain, "UnhandledException", ProcessUnhandledException);
        }

        ~WinRTSpellerInterop()
        {
            Dispose(false);
        }

        #endregion Constructors

        #region IDispose

        public override void  Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Internal interop resource cleanup
        /// </summary>
        /// <SecurityNote>
        /// Critical:
        ///     Calls into Marshal.ReleaseComObject
        /// Safe:
        ///     Called by transparent callers, and does not expose
        ///     critical resources (COM objects) to the callers.
        /// </SecurityNote>
        /// <param name="disposing"></param>
        [SecuritySafeCritical]
        protected override void Dispose(bool disposing)
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(SR.Get(SRID.TextEditorSpellerInteropHasBeenDisposed));
            }


            if (_spellCheckers != null)
            {
                foreach(Tuple<WordsSegmenter, ISpellChecker> item in _spellCheckers.Values)
                {
                    ISpellChecker spellChecker = item?.Item2;
                    if (spellChecker != null)
                    {
                        Marshal.ReleaseComObject(spellChecker);
                    }
                }

                _spellCheckers = null; 
            }

            ClearDictionaries(isDisposeOrFinalize:true);

            if (_spellCheckerFactory != null)
            {
                // After this point, _spellCheckerFactory cannot be used elsewhere.
                _spellCheckerFactoryLock.WithWriteLock(ReleaseSpellCheckerFactory);
            }

            // Locks may not be initialized if Dispose() is called from the 
            // constructor. Only call _spellCheckerFactoryLock.Dispose() 
            // if the lock has been initialized.
            _spellCheckerFactoryLock?.Dispose();

            _isDisposed = true;
        }

        #endregion 

        #region Internal Methods

        internal override void SetLocale(CultureInfo culture)
        {
            Culture = culture;
        }

        /// <summary>
        /// Sets the mode in which the spell-checker operates
        /// We care about 3 different modes here: 
        /// 
        /// 1. Shallow spellchecking - i.e., wordbreaking +      spellchecking + NOT (suggestions)
        /// 2. Deep spellchecking    - i.e., wordbreaking +      spellchecking +      suggestions
        /// 3. Wordbreaking only     - i.e., wordbreaking + NOT (spellchcking) + NOT (suggestions)
        /// </summary>
        internal override SpellerMode Mode
        {
            set
            {
                _mode = value;
            }
        }

        /// <summary>
        /// If true, multi-word spelling errors would be detected
        /// This flag is ignored by WinRTSpellerInterop
        /// </summary>
        internal override bool MultiWordMode
        {
            set
            {
                // do nothing - multi-word mode specification is not supported
                // _multiWordMode = value;
            }
        }

        /// <summary>
        /// Sets spelling reform mode
        /// WinRTSpellerInterop doesn't support spelling reform
        /// </summary>
        /// <param name="culture"></param>
        /// <param name="spellingReform"></param>
        internal override void SetReformMode(CultureInfo culture, SpellingReform spellingReform)
        {
            // Do nothing - spelling reform is not supported
            // _spellingReformInfos[culture] =  spellingReform;
        }

        /// <summary>
        /// Returns true if we have an engine capable of proofing the specified language.
        /// </summary>
        /// <param name="culture"></param>
        /// <returns></returns>
        internal override bool CanSpellCheck(CultureInfo culture)
        {
            return EnsureWordBreakerAndSpellCheckerForCulture(culture);
        }



        #region Dictionary Methods

        /// <summary>
        /// Unloads a given custom dictionary
        /// </summary>
        /// <param name="token"></param>
        /// <SecurityNote>
        /// Critical - 
        ///     Demands FileIOPermission
        ///     Calls into COM API
        /// </SecurityNote>
        [SecurityCritical]
        internal override void UnloadDictionary(object token)
        {
            var data = (Tuple<CultureInfo, String>)token;
            CultureInfo culture = data.Item1;
            string filePath = data.Item2;

            new FileIOPermission(FileIOPermissionAccess.AllAccess, filePath).Demand();

            _customDictionaryFilesLock.WaitOne();
            try
            {
                _customDictionaryFiles[culture].RemoveAll((str) => str == filePath);
            }
            finally
            {
                _customDictionaryFilesLock.Release();
            }

            _spellCheckerFactoryLock.WithReadLock(()=> 
            {
                IUserDictionariesRegistrar registrar = (IUserDictionariesRegistrar)_spellCheckerFactory;
                registrar.UnregisterUserDictionary(filePath, culture.IetfLanguageTag);
            });

            File.Delete(filePath);
        }

        /// <summary>
        /// Loads a custom dictionary
        /// </summary>
        /// <param name="lexiconFilePath"></param>
        /// <returns></returns>
        /// <SecurityNote>
        /// Critical
        ///     Calls into LoadDictionaryImpl which is Critical
        /// </SecurityNote>
        [SecurityCritical]
        internal override object LoadDictionary(string lexiconFilePath)
        {
            return LoadDictionaryImpl(lexiconFilePath);
        }

        /// <summary>
        /// Loads a custom dictionary
        /// </summary>
        /// <param name="item"></param>
        /// <param name="trustedFolder"></param>
        /// <param name="dictionaryLoadedCallback"></param>
        /// <returns></returns>
        /// <SecurityNote>
        /// Critical - 
        ///     Calls into LoadDictionaryImp which is Critical
        ///     Asserts FileIOPermission
        /// </SecurityNote>
        [SecurityCritical]
        internal override object LoadDictionary(Uri item, string trustedFolder)
        {
            // Assert neccessary security to load trusted files.
            new FileIOPermission(FileIOPermissionAccess.Read, trustedFolder).Assert();
            try
            {
                return LoadDictionaryImpl(item.LocalPath);
            }
            finally
            {
                FileIOPermission.RevertAssert();
            }
        }

        /// <summary>
        /// Releases all currently loaded custom dictionaries
        /// </summary>
        /// <SecurityNote>
        /// Critical -
        ///     Calls ClearDictionaries which is Critical
        /// </SecurityNote>
        [SecurityCritical]
        internal override void ReleaseAllLexicons()
        {
            ClearDictionaries();
        }

        #endregion

        #endregion Internal Methods


        #region Private Methods

        /// <summary>
        /// <SecurityNote>
        /// Critical - Calls WordsSegmenter.Create which is Critical
        /// </SecurityNote>
        /// </summary>
        /// <param name="culture"></param>
        /// <param name="throwOnError"></param>
        /// <returns></returns>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private bool EnsureWordBreakerAndSpellCheckerForCulture(CultureInfo culture, bool throwOnError = false)
        {
            if (culture == null)
            {
                return false; 
            }

            if(!_spellCheckers.ContainsKey(culture))
            {
                WordsSegmenter wordBreaker = null; 
                
                try
                {
                    wordBreaker = WordsSegmenter.Create(culture.Name);
                }
                catch when (!throwOnError)
                {
                    // ArgumentException: culture name is malformed - unlikely given we use culture.Name
                    // PlatformNotSupportedException: OS is not supported
                    // NotSupportedException: culture name is likely well-formed, but not available currently for wordbreaking
                    wordBreaker = null;
                }

                // Even if wordBreaker.ResolvedLanguage == WordsSegmenter.Undetermined, we will use it 
                // as an appropriate fallback wordbreaker as long as a corresponding ISpellChecker is found. 
                if (wordBreaker == null)
                {
                    _spellCheckers[culture] = null;
                    return false; 
                }

                ISpellChecker spellChecker = null;

                _spellCheckerFactoryLock.WithReadLock(() =>
                {
                    try
                    {
                        spellChecker = _spellCheckerFactory.CreateSpellChecker(culture.Name);
                    }
                    catch (Exception ex)
                    {
                        spellChecker = null;

                        // ArgumentException: 
                        // Either the language name is malformed (unlikely given we use culture.Name)
                        //   or this language is not supported. It might be supported if the appropriate 
                        //   input language is added by the user, but it is not available at this time. 

                        if (throwOnError && ex is ArgumentException)
                        {
                            throw new NotSupportedException(string.Empty, ex);
                        }
                    }
                });

                if (spellChecker == null)
                {
                    _spellCheckers[culture] = null;
                }
                else
                {
                    _spellCheckers[culture] = new Tuple<WordsSegmenter, ISpellChecker>(wordBreaker, spellChecker);
                }
            }

            return (_spellCheckers[culture] == null ? false : true);
        }

        /// <summary>
        /// foreach(sentence in text.sentences)
        ///      foreach(segment in sentence)
        ///          continueIteration = segmentCallback(segment, data)
        ///      endfor
        ///
        ///      if (sentenceCallback != null) 
        ///          continueIteration = sentenceCallback(sentence, data)
        ///      endif
        ///
        ///      if (!continueIteration) 
        ///          break
        ///      endif
        ///  endfor 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="count"></param>
        /// <param name="sentenceCallback"></param>
        /// <param name="segmentCallback"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        internal override int EnumTextSegments(char[] text, int count, 
            EnumSentencesCallback sentenceCallback, EnumTextSegmentsCallback segmentCallback, object data)
        {
            var wordBreaker = CurrentWordBreaker ?? DefaultCultureWordBreaker;
            var spellChecker = CurrentSpellChecker;

            bool spellCheckerNeeded = _mode.HasFlag(SpellerMode.SpellingErrors) || _mode.HasFlag(SpellerMode.Suggestions);
            if ((wordBreaker == null) || (spellCheckerNeeded && spellChecker == null)) return 0;

            int segmentCount = 0;
            bool continueIteration = true;

            // WinRT WordsSegmenter doesn't have the ability to break down text into segments (sentences). 
            // Treat the whole text as a single segment for now. 
            foreach(string strSentence in new string[]{string.Join(string.Empty, text)})
            {
                SpellerSentence sentence = new SpellerSentence(strSentence, wordBreaker, CurrentSpellChecker);
                segmentCount += sentence.Segments.Count;

                if (segmentCallback != null)
                {
                    for (int i = 0; continueIteration && (i < sentence.Segments.Count); i++)
                    {
                        continueIteration = segmentCallback(sentence.Segments[i], data);
                    }
                }

                if (sentenceCallback != null)
                {
                    continueIteration = sentenceCallback(sentence, data);
                }
                
                if (!continueIteration) break;
            }

            return segmentCount;
        }

        /// <summary>
        ///     Actual implementation of loading a dictionary
        /// </summary>
        /// <param name="lexiconFilePath"></param>
        /// <param name="dictionaryLoadedCallback"></param>
        /// <param name="callbackParam"></param>
        /// <returns>
        ///     A tuple of cultureinfo detected from <paramref name="lexiconFilePath"/> and 
        ///     a temp file path which holds a copy of <paramref name="lexiconFilePath"/>
        /// 
        ///     If no culture is specified in the first line of <paramref name="lexiconFilePath"/>
        ///     in the format #LID nnnn (where nnnn = decimal LCID of the culture), then invariant 
        ///     culture is returned. 
        /// </returns>
        /// <remarks>
        ///     At the end of this method, we guarantee that <paramref name="lexiconFilePath"/> 
        ///     can be reclaimed (i.e., potentially deleted) by the caller. 
        /// </remarks>
        /// <SecurityNote>
        /// Critical - 
        ///     Demands and Asserts permissions
        ///     Calls into WinRTSpellerInterop.GetTempFileName which is Critical
        ///     Calls into IUserDictionariesRegistrar COM API's
        /// </SecurityNote>
        [SecurityCritical]
        private Tuple<CultureInfo, String> LoadDictionaryImpl(string lexiconFilePath)
        {
            try
            {
                new FileIOPermission(FileIOPermissionAccess.Read, lexiconFilePath).Demand();
            }
            catch (SecurityException se)
            {
                throw new ArgumentException(SR.Get(SRID.CustomDictionaryFailedToLoadDictionaryUri, lexiconFilePath), se);
            }

            if (!File.Exists(lexiconFilePath))
            {
                throw new ArgumentException(SR.Get(SRID.CustomDictionaryFailedToLoadDictionaryUri, lexiconFilePath));
            }

            bool fileCopied = false;
            string lexiconPrivateCopyPath = null; 

            try
            {
                CultureInfo culture = null;

                // Read the first line of the file and detect culture, if specified
                using (FileStream stream = new FileStream(lexiconFilePath, FileMode.Open, FileAccess.Read))
                {
                    string line = null;
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        line = reader.ReadLine();
                        culture = WinRTSpellerInterop.TryParseLexiconCulture(line);
                    }
                }

                // Make a temp file and copy the original file over. 
                // Ensure that the copy has Unicode (UTF16-LE) encoding
                lexiconPrivateCopyPath = WinRTSpellerInterop.GetTempFileName(extension: "dic");

                new FileIOPermission(FileIOPermissionAccess.Read | FileIOPermissionAccess.Write, lexiconPrivateCopyPath).Assert();
                try
                {
                    WinRTSpellerInterop.CopyToUnicodeFile(lexiconFilePath, lexiconPrivateCopyPath);
                    fileCopied = true;
                }
                finally
                {
                    CodeAccessPermission.RevertAssert();
                }

                // Add the temp file (with .dic extension) just created to a cache, 
                // then pass it along to IUserDictionariesRegistrar

                _customDictionaryFilesLock.WaitOne();
                try
                {
                    if (!_customDictionaryFiles.ContainsKey(culture))
                    {
                        _customDictionaryFiles[culture] = new List<string>();
                    }

                    _customDictionaryFiles[culture].Add(lexiconPrivateCopyPath);
                }
                finally
                {
                    _customDictionaryFilesLock.Release();
                }

                _spellCheckerFactoryLock.WithReadLock(() =>
                {
                    IUserDictionariesRegistrar registrar = (IUserDictionariesRegistrar)_spellCheckerFactory;
                    registrar.RegisterUserDictionary(lexiconPrivateCopyPath, culture.IetfLanguageTag);
                });

                return new Tuple<CultureInfo, string>(culture, lexiconPrivateCopyPath);
            }
            catch (Exception e) when ((e is SecurityException) || (e is ArgumentException) || !fileCopied)
            {
                // IUserDictionariesRegistrar.RegisterUserDictionary can 
                // throw ArgumentException on failure. Cleanup the temp file if 
                // we successfully created one. 
                if (lexiconPrivateCopyPath != null)
                {
                    File.Delete(lexiconPrivateCopyPath);
                }

                throw new ArgumentException(SR.Get(SRID.CustomDictionaryFailedToLoadDictionaryUri, lexiconFilePath), e);
            }
        }

        /// <summary>
        ///     Actual implementation of clearing all dictionaries
        /// </summary>
        /// <remarks>
        ///     ClearDictionaries() can be called from the following methods/threads
        ///         Dispose(bool):              UI thread or the finalizer thread
        ///         ReleaseAllLexicons:         UI thread
        ///         ProcessUnhandledException:  Any thread
        /// 
        ///     In order to avoid contentions between potentially reentrant threads trying to 
        ///     call into ClearDictionaries, we use a semaphore (_customDictionaryFilesLock) to 
        ///     control all write accesses to _customDictionaryFiles cache.
        /// </remarks>
        /// <SecurityNote>
        /// Critical -
        ///     Calls into IUserDictionariesRegistrar COM API's
        ///     Demands FileIOPermission
        /// </SecurityNote>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        [SecurityCritical]
        private void ClearDictionaries(bool isDisposeOrFinalize = false)
        {
            if ((_customDictionaryFilesLock == null) || (_spellCheckerFactoryLock == null))
            {
                // Locks are not initialized => Dispose called from within the constructor. 
                // Likely this platform is not supported - do not process further. 
                return;
            }

            _customDictionaryFilesLock.WaitOne();
            try
            {
                _spellCheckerFactoryLock.WithReadLock(() =>
                {

                    if ((_spellCheckerFactory != null) && (_customDictionaryFiles != null))
                    {
                        IUserDictionariesRegistrar registrar = (IUserDictionariesRegistrar)_spellCheckerFactory;

                        foreach (KeyValuePair<CultureInfo, List<string>> items in _customDictionaryFiles)
                        {
                            CultureInfo culture = items.Key;
                            foreach (string filePath in items.Value)
                            {
                                try
                                {
                                    new FileIOPermission(FileIOPermissionAccess.AllAccess, filePath).Demand();

                                    registrar.UnregisterUserDictionary(filePath, culture.IetfLanguageTag);
                                    File.Delete(filePath);
                                }
                                catch
                                {
                                    // Do nothing - Continue to make a best effort 
                                    // attempt at unregistering custom dictionaries
                                }
                            }
                        }

                        _customDictionaryFiles.Clear();
                    }
                });
            }
            finally
            {
                if(isDisposeOrFinalize)
                {
                    _customDictionaryFiles = null;
                }

                _customDictionaryFilesLock.Release();
            }
        }

        /// <summary>
        ///     Detect whether the <paramref name="line"/> is of the form #LID nnnn, 
        ///     and if it is, try to instantiate a CultureInfo object with LCID nnnn. 
        /// </summary>
        /// <param name="line"></param>
        /// <returns>
        ///     The CultureInfo object corresponding to the LCID specified in the <paramref name="line"/>
        /// </returns>
        private static CultureInfo TryParseLexiconCulture(string line)
        {
            const string regexPattern = @"\s*\#LID\s+(\d+)\s*";
            RegexOptions regexOptions = RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled;

            CultureInfo result = CultureInfo.InvariantCulture;

            if (line == null)
            {
                return result; 
            }

            string[] matches = Regex.Split(line.Trim(), regexPattern, regexOptions);

            // We expect 1 exact match, which implies matches.Length == 3 (before, match, after)
            if (matches.Length != 3)
            {
                return result;
            }

            string before = matches[0];
            string match  = matches[1];
            string after  = matches[2];

            // We expect 1 exact match, which implies the following:
            //      before == after == string.Emtpy
            //      match is parsable into an integer
            int lcid;
            if ((before != string.Empty) || (after != string.Empty) || (!Int32.TryParse(match, out lcid)))
            {
                return result;
            }

            try
            {
                result = new CultureInfo(lcid);
            }
            catch (CultureNotFoundException)
            {
                result = CultureInfo.InvariantCulture;
            }

            return result;
        }

        /// <summary>
        ///     Creates a temp file with extension <paramref name="extension"/>
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        /// <remarks>
        ///     We try to create a temp file under %temp% by calling Path.GetRandomFileName(), 
        ///     changing its extension to <paramref name="extension"/>, and attempt to create a 0 byte file 
        ///     with this full path. This has the potential for collisions, so we retry this 10 times, 
        ///     after which we fail.
        /// </remarks>
        [SecurityCritical]
        private static string GetTempFileName(string extension)
        {
            const int maxTries = 10; 

            string tempFolderPath = Path.GetTempPath();
            new FileIOPermission(FileIOPermissionAccess.Read | FileIOPermissionAccess.Write, tempFolderPath).Demand();

            int attempts = 0;

            while (true)
            {
                ++attempts;
                string filename = Path.Combine(tempFolderPath, Path.ChangeExtension(Path.GetRandomFileName(), extension));
                try
                {
                    using (new FileStream(filename, FileMode.CreateNew)) { }
                    return filename;
                }
                catch (IOException) when (attempts <= maxTries)
                {
                    // do nothing
                }
            }
        }

        /// <summary>
        ///     Copies <paramref name="sourcePath"/> to <paramref name="targetPath"/>. During the copy, it transcodes 
        ///     <paramref name="sourcePath"/> to Unicode (UTL16-LE) if necessary and ensures that <paramref name="targetPath"/>
        ///     has the right BOM (Byte Order Mark) for UTF16-LE (FF FE) 
        /// </summary>
        /// <see cref = "// See http://www.unicode.org/faq/utf_bom.html" />
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <SecurityNote>
        /// Critical - 
        ///     Demands FileIOPermission permissions
        /// </SecurityNote>
        [SecurityCritical]
        private static void CopyToUnicodeFile(string sourcePath, string targetPath)
        {
            new FileIOPermission(FileIOPermissionAccess.Read, sourcePath).Demand();
            new FileIOPermission(FileIOPermissionAccess.Write, targetPath).Demand();

            bool utf16LEEncoding = false;
            using (FileStream sourceStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            {
                // Check that the first two bytes indicate the BOM for UTF16-LE
                // If found, we can directly copy the file over without additional transcoding.
                utf16LEEncoding = ((sourceStream.ReadByte() == 0xFF) && (sourceStream.ReadByte() == 0xFE));

                if (!utf16LEEncoding)
                {
                    sourceStream.Seek(0, SeekOrigin.Begin);
                    using (StreamReader reader = new StreamReader(sourceStream))
                    {
                        using (FileStream targetStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
                        {
                            // Create the StreamWriter with encoding = Unicode to ensure that the new file 
                            // contains the BOM for UTF16-LE, and also ensures that the file contents are 
                            // encoded correctly
                            using (StreamWriter writer = new StreamWriter(targetStream, Text.Encoding.Unicode))
                            {
                                string line = null;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    writer.WriteLine(line);
                                }
                            }
                        }
                    }
                }
            }

            if (utf16LEEncoding)
            {
                File.Copy(sourcePath, targetPath, true);
            }
        }

        /// <summary>
        /// Attempts to unregister all custom dictionaries if an unhandled exception is raised
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <SecurityNote>
        /// Critical:
        ///     Calls ClearDictionaries which is Critical
        /// Safe:
        ///     Called by transparent methods, and does not expose any 
        ///     critical resources (COM objects) to callers.
        /// </SecurityNote>
        [SecuritySafeCritical]
        private void ProcessUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ClearDictionaries();   
        }

        /// <summary>
        /// Releases the ISpellCheckerFactory object and sets the reference to null
        /// </summary>
        /// <SecurityNote>
        /// Critical - 
        ///     Calls into Marshal.ReleaseComObject
        /// </SecurityNote>
        /// <remarks>
        /// This is only called from within Dispose(bool) and passed as a 
        /// parameter to ReaderWriterLockSlimWrapper.WithWriteLock(Action action). 
        /// We cannot simply pass an anonymous delegate to WithWriteLock because 
        /// anon. delegates cannot be [SecurityCritical] 
        /// </remarks>
        [SecurityCritical]
        private void ReleaseSpellCheckerFactory()
        {
            Marshal.ReleaseComObject(_spellCheckerFactory);
            _spellCheckerFactory = null;
        }

        #endregion 

        #region Private Properties

        private CultureInfo Culture
        {
            get
            {
                return _culture;
            }

            set
            {
                _culture = value;
                EnsureWordBreakerAndSpellCheckerForCulture(_culture);
            }
        }

        private WordsSegmenter CurrentWordBreaker
        {
            get
            {
                if (Culture == null)
                {
                    return null;
                }
                else
                {
                    EnsureWordBreakerAndSpellCheckerForCulture(Culture);
                    return _spellCheckers[Culture]?.Item1;
                }
            }
        }

        private WordsSegmenter DefaultCultureWordBreaker
        {
            get
            {
                if (_defaultCulture == null)
                {
                    return null;
                }
                else
                {
                    return _spellCheckers[_defaultCulture]?.Item1;
                }
            }
        }

        private ISpellChecker CurrentSpellChecker
        {
            get
            {
                if (Culture == null)
                {
                    return null;
                }
                else 
                {
                    EnsureWordBreakerAndSpellCheckerForCulture(Culture);
                    return _spellCheckers[Culture]?.Item2;
                }
            }
        }

        #endregion

        #region Private Fields

        private bool _isDisposed = false;
        private SpellerMode _mode = SpellerMode.None;

        private SpellCheckerFactory _spellCheckerFactory;

        // Cache of word-breakers and spellcheckers
        private Dictionary<CultureInfo, Tuple<WordsSegmenter, ISpellChecker>> _spellCheckers;

        private CultureInfo _defaultCulture;
        private CultureInfo _culture;

        // Cache of private dictionaries
        private Dictionary<CultureInfo, List<string>> _customDictionaryFiles;

        /// <remarks>
        ///     See remarks in ClearDictionaries method
        /// </remarks>
        private Semaphore _customDictionaryFilesLock;

        /// <summary>
        /// Ensures that we do not call Marshal.ReleaseComObject(_spellCheckerFactory)
        /// while it is still in use. This will avoid the potential activation of
        /// ----OnRCWCleanup MDA. 
        /// 
        /// ----OnRCWCleanup MDA can be activated if, for e.g., an unhandled exception 
        /// sets up a ---- between a normal operation in the UI thread (for e.g., LoadDictionary), 
        /// and the finalizer being called from another thread. 
        /// </summary>
        private ReaderWriterLockSlimWrapper _spellCheckerFactoryLock;
        
        #endregion Private Fields

        #region Private Types

        private struct TextRange: SpellerInteropBase.ITextRange
        {
            public TextRange(MS.Internal.WindowsRuntime.Windows.Data.Text.TextSegment textSegment)
            {
                _length = (int)textSegment.Length;
                _start = (int)textSegment.StartPosition;
            }

            public static explicit operator TextRange(MS.Internal.WindowsRuntime.Windows.Data.Text.TextSegment textSegment)
            {
                return new TextRange(textSegment);
            }

            #region SpellerInteropBase.ITextRange

            public int Start
            {
                get { return _start;  }
            }

            public int Length
            {
                get { return _length; }
            }

            #endregion 

            private readonly int _start;
            private readonly int _length;
        }

        [DebuggerDisplay("SubSegments.Count = {SubSegments.Count} TextRange = {TextRange.Start},{TextRange.Length}")]
        private class SpellerSegment: ISpellerSegment
        {
            #region Constructor

            public SpellerSegment(WordSegment segment, ISpellChecker spellChecker)
            {
                _segment = segment;
                _spellChecker = spellChecker;
                _suggestions = null;
            }

            static SpellerSegment()
            {
                _empty = new List<ISpellerSegment>().AsReadOnly();
            }

            #endregion 

            #region Private Methods

            private void EnumerateSuggestions()
            {
                List<string> result = new List<string>();
                _isClean = true;

                if (_spellChecker == null)
                {
                    _suggestions = result.AsReadOnly(); 
                    return;
                }

                var spellingErrors = _spellChecker.ComprehensiveCheck(_segment.Text);
                if (spellingErrors == null)
                {
                    _suggestions = result.AsReadOnly();
                    return;
                }

                try
                {
                    ISpellingError error;
                    while(true)
                    {
                        error = spellingErrors.Next();
                        if (error == null) break;

                        // Once _clean has been set to false, it will not be updated again. 
                        if (_isClean.Value)
                        {
                            _isClean = (error.CorrectiveAction == CORRECTIVE_ACTION.CORRECTIVE_ACTION_NONE);
                        }

                        try
                        {
                            if ((error.CorrectiveAction == CORRECTIVE_ACTION.CORRECTIVE_ACTION_GET_SUGGESTIONS) || 
                                (error.CorrectiveAction == CORRECTIVE_ACTION.CORRECTIVE_ACTION_REPLACE))
                            {
                                var suggestions = _spellChecker.Suggest(_segment.Text);
                                if (suggestions == null) break;

                                try
                                {
                                    uint fetched = 0;
                                    string suggestion = string.Empty;

                                    do
                                    {
                                        suggestions.RemoteNext(1, out suggestion, out fetched);
                                        if (fetched > 0) result.Add(suggestion);
                                    }
                                    while (fetched > 0);
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(suggestions);
                                }
                            } 
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(error);
                        }
                    } 
                }
                finally
                {
                    Marshal.ReleaseComObject(spellingErrors);
                }

                _suggestions = result.AsReadOnly();
            }

            #endregion 

            #region SpellerInteropBase.ISpellerSegment

            /// <summary>
            /// Returns a read-only list of sub-segments of this segment
            /// WinRT word-segmenter doesn't really support sub-segments,
            ///   so we always return an empty list
            /// </summary>
            public IReadOnlyList<ISpellerSegment> SubSegments
            {
                get
                {
                    return SpellerSegment._empty;
                }
            }

            public ITextRange TextRange
            {
                get
                {
                    return new TextRange(_segment.SourceTextSegment);
                }
            }

            public IReadOnlyList<string> Suggestions
            {
                get
                {
                    if (_suggestions == null)
                    {
                        EnumerateSuggestions();
                    }

                    return _suggestions;
                }
            }

            public bool IsClean
            {
                get
                {
                    if (_isClean == null)
                    {
                        EnumerateSuggestions();
                    }

                    return _isClean.Value;
                }
            }

            public void EnumSubSegments(EnumTextSegmentsCallback segmentCallback, object data)
            {
                bool result = true;

                for (int i = 0; result && (i < SubSegments.Count); i++)
                {
                    result = segmentCallback(SubSegments[i], data);
                }
            }

            #endregion SpellerInteropBase.ISpellerSegment

            #region Private Fields

            private WordSegment _segment;

            ISpellChecker _spellChecker;
            private IReadOnlyList<string> _suggestions;
            private bool? _isClean = null; 

            private static readonly IReadOnlyList<ISpellerSegment> _empty;

            #endregion Private Fields
        }

        [DebuggerDisplay("Sentence = {_sentence}")]
        private class SpellerSentence: ISpellerSentence
        {
            public SpellerSentence(string sentence, WordsSegmenter wordBreaker, ISpellChecker spellChecker)
            {
                _sentence = sentence;
                _wordBreaker = wordBreaker;
                _spellChecker = spellChecker;
                _segments = null;
            }

            #region SpellerInteropBase.ISpellerSentence

            public IReadOnlyList<ISpellerSegment> Segments
            {
                get
                {
                    if (_segments == null)
                    {
                        List<SpellerSegment> segments = new List<SpellerSegment>();

                        foreach (var wordSegment in _wordBreaker.GetTokens(_sentence))
                        {
                            segments.Add(new SpellerSegment(wordSegment, _spellChecker));
                        }

                        _segments = segments.AsReadOnly();
                    }

                    return _segments;
                }
            }

            public int EndOffset
            {
                get
                {
                    int endOffset = -1;

                    if (Segments.Count > 0)
                    {
                        ITextRange textRange = Segments[Segments.Count - 1].TextRange;
                        endOffset = textRange.Start + textRange.Length;
                    }

                    return endOffset;
                }
            }

            #endregion 

            private string _sentence;
            private WordsSegmenter _wordBreaker;
            private ISpellChecker _spellChecker;
            private IReadOnlyList<SpellerSegment> _segments;

        }

        #endregion Private Types

        #region Private Interfaces

        /// <summary>
        /// RCW for spellcheck.idl found in Windows SDK
        /// This is generated code with minor manual edits. 
        /// i.  Generate TLB
        ///      MIDL /TLB MsSpellCheckLib.tlb SpellCheck.IDL //SpellCheck.IDL found in Windows SDK
        /// ii. Generate RCW in a DLL
        ///      TLBIMP MsSpellCheckLib.tlb // Generates MsSpellCheckLib.dll
        /// iii.Decompile the DLL and copy out the RCW by hand.
        ///      ILDASM MsSpellCheckLib.dll
        /// </summary>
        #region MsSpellCheckLib RCW

        #region WORDLIST_TYPE

        // Types of user custom wordlists
        // Custom wordlists are language-specific
        private enum WORDLIST_TYPE : int
        {
            WORDLIST_TYPE_IGNORE = 0, // Ignore wordlist - words that should be considered correctly spelled in a single spell checking session
            WORDLIST_TYPE_ADD = 1, // Added words wordlist - words that should be considered correctly spelled - permanent and applies to all clients
            WORDLIST_TYPE_EXCLUDE = 2, // Excluded words wordlist - words that should be considered misspelled - permanent and applies to all clients
            WORDLIST_TYPE_AUTOCORRECT = 3, // Autocorrect wordlit - pairs of words with a word that should be automatically substituted by the other word in the pair - permanent and applies to all clients
        }

        #endregion // WORDLIST_TYPE

        #region CORRECTIVE_ACTION

        // Action that a client should take on a specific spelling error(obtained from ISpellingError::get_CorrectiveAction)
        private enum CORRECTIVE_ACTION : int
        {
            CORRECTIVE_ACTION_NONE = 0, // None - there's no error
            CORRECTIVE_ACTION_GET_SUGGESTIONS = 1, // GetSuggestions - the client should show a list of suggestions (obtained through ISpellChecker::Suggest) to the user
            CORRECTIVE_ACTION_REPLACE = 2, // Replace - the client should autocorrect the word to the word obtained from ISpellingError::get_Replacement
            CORRECTIVE_ACTION_DELETE = 3, // Delete - the client should delete this word
        }

        #endregion // CORRECTIVE_ACTION

        #region ISpellingError

        // This interface represents a spelling error - you can get information like the range that comprises the error, or the suggestions for that misspelled word
        // Should be implemented by any spell check provider (someone who provides a spell checking engine), and used by clients of spell checking
        // It is obtained through IEnumSpellingError::Next
        [ComImport]
        [Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ISpellingError
        {
            uint StartIndex
            {
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            uint Length
            {
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            CORRECTIVE_ACTION CorrectiveAction
            {
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            string Replacement
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }
        }

        #endregion //ISpellingError

        #region IEnumSpellingError

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("803E3BD4-2828-4410-8290-418D1D73C762")]
        private interface IEnumSpellingError
        {
            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            ISpellingError Next();
        }

        #endregion // IEnumSpellingError

        #region IEnumString

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00000101-0000-0000-C000-000000000046")]
        private interface IEnumString
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void RemoteNext([In] uint celt, [MarshalAs(UnmanagedType.LPWStr)] out string rgelt, out uint pceltFetched);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Skip([In] uint celt);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Reset();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumString ppenum);
        }

        #endregion // IEnumString

        #region IOptionDescription

        [ComImport]
        [Guid("432E5F85-35CF-4606-A801-6F70277E1D7A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOptionDescription
        {
            string Id
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            string Heading
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            string Description
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            IEnumString Labels
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }
        }

        #endregion // IOptionDescription

        #region ISpellCheckerChangedEventHandler

        [ComImport]
        [Guid("0B83A5B0-792F-4EAB-9799-ACF52C5ED08A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ISpellCheckerChangedEventHandler
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Invoke([In, MarshalAs(UnmanagedType.Interface)] ISpellChecker sender);
        }

        #endregion // #region ISpellCheckerChangedEventHandler

        #region ISpellChecker

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B")]
        private interface ISpellChecker
        {
            string languageTag
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IEnumSpellingError Check([In, MarshalAs(UnmanagedType.LPWStr)] string text);

            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IEnumString Suggest([In, MarshalAs(UnmanagedType.LPWStr)] string word);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Add([In, MarshalAs(UnmanagedType.LPWStr)] string word);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Ignore([In, MarshalAs(UnmanagedType.LPWStr)] string word);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void AutoCorrect([In, MarshalAs(UnmanagedType.LPWStr)] string from, [In, MarshalAs(UnmanagedType.LPWStr)] string to);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            byte GetOptionValue([In, MarshalAs(UnmanagedType.LPWStr)] string optionId);

            IEnumString OptionIds
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            string Id
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            string LocalizedName
            {
                [return: MarshalAs(UnmanagedType.LPWStr)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            uint add_SpellCheckerChanged([In, MarshalAs(UnmanagedType.Interface)] ISpellCheckerChangedEventHandler handler);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void remove_SpellCheckerChanged([In] uint eventCookie);

            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IOptionDescription GetOptionDescription([In, MarshalAs(UnmanagedType.LPWStr)] string optionId);

            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IEnumSpellingError ComprehensiveCheck([In, MarshalAs(UnmanagedType.LPWStr)] string text);
        }

        #endregion // ISpellChecker

        #region ISpellCheckerFactory

        [ComImport]
        [Guid("8E018A9D-2415-4677-BF08-794EA61F94BB")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ISpellCheckerFactory
        {
            IEnumString SupportedLanguages
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int IsSupported([In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            ISpellChecker CreateSpellChecker([In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);
        }

        #endregion // ISpellCheckerFactory

        #region IUserDictionariesRegistrar

        [ComImport]
        [Guid("AA176B85-0E12-4844-8E1A-EEF1DA77F586")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IUserDictionariesRegistrar
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void RegisterUserDictionary([In, MarshalAs(UnmanagedType.LPWStr)] string dictionaryPath, [In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void UnregisterUserDictionary([In, MarshalAs(UnmanagedType.LPWStr)] string dictionaryPath, [In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);
        }

        #endregion // IUserDictionariesRegistrar

        #region SpellCheckerFactoryCoClass

        [ComImport]
        [Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
        [TypeLibType(TypeLibTypeFlags.FCanCreate)]
        [ClassInterface(ClassInterfaceType.None)]
        private class SpellCheckerFactoryCoClass : ISpellCheckerFactory, SpellCheckerFactory, IUserDictionariesRegistrar
        {
            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            public virtual extern ISpellChecker CreateSpellChecker([In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            public virtual extern int IsSupported([In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            public virtual extern void RegisterUserDictionary([In, MarshalAs(UnmanagedType.LPWStr)] string dictionaryPath, [In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            public virtual extern void UnregisterUserDictionary([In, MarshalAs(UnmanagedType.LPWStr)] string dictionaryPath, [In, MarshalAs(UnmanagedType.LPWStr)] string languageTag);

            public virtual extern IEnumString SupportedLanguages
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
                get;
            }
        }

        #endregion // SpellCheckerFactoryCoClass

        #region SpellCheckerFactory

        [ComImport]
        [CoClass(typeof(SpellCheckerFactoryCoClass))]
        [Guid("8E018A9D-2415-4677-BF08-794EA61F94BB")]
        private interface SpellCheckerFactory : ISpellCheckerFactory
        {
        }

        #endregion // SpellCheckerFactory
        
        #endregion MsSpellCheckLib RCW

        #endregion Private Interfaces
    }

}
