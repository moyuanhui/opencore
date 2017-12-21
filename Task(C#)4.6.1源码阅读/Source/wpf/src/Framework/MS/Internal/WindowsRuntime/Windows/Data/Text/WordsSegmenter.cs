//-----------------------------------------------------------------------------------------------------
//
// File: WordsSegmenter.cs
//
// Description: Provides a class describing the WinRT type 
//                  Windows.Data.Text.WordsSegmenter. In addition to this, it also 
//                  defines a related delegate type used to support methods in this class.
//
//              WordsSegmenter is a class able to segment provided text into words.  
//              
//              The language supplied when these objects are constructed is matched against the 
//                  languages with word breakers on the system, and the best word segmentation rules 
//                  available are used. The language need not be one of the appliation's supported 
//                  languages. If there are no supported language rules available specifically for 
//                  that language, the language-neutral rules are used (an implementation of Unicode 
//                  Standard Annex #29 Unicode Text Segmentation), and the ResolvedLanguage property is 
//                  set to "und" (undetermined language).
//
//-----------------------------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Security.Permissions;


namespace MS.Internal.WindowsRuntime
{
    namespace Windows.Data.Text
    {
        internal class WordsSegmenter
        {
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
            static WordsSegmenter()
            {
                ConstructorInfo constructor = null;

                // We don't want to throw here - so wrap in try..catch
                try
                {
                    s_WinRTType = Type.GetType(s_TypeName);
                    if (s_WinRTType != null)
                    {
                        constructor = s_WinRTType.GetConstructor(new Type[] { typeof(string) });
                    }
                }
                catch
                {
                    s_WinRTType = null;
                }

                s_PlatformSupported = (s_WinRTType != null) && (constructor != null);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="language"></param>
            /// <exception cref="System.ArgumentException"></exception>
            /// <exception cref="System.NotSupportedException"></exception>
            /// <exception cref="PlatformNotSupportedException"></exception>
            private WordsSegmenter(string language)
            {
                if (!s_PlatformSupported)
                {
                    throw new PlatformNotSupportedException();
                }

                try
                {
                    _wordsSegmenter = s_WinRTType.ReflectionNew<string>(language);
                }
                catch (TargetInvocationException tiex) when (tiex.InnerException is ArgumentException)
                {
                    throw new ArgumentException(string.Empty, "language", tiex);
                }

                if ((_wordsSegmenter == null) || (ResolvedLanguage == Undetermined))
                {
                    throw new NotSupportedException();
                }
            }

            /// <summary>
            /// </summary>
            /// <param name="language"></param>
            /// <returns></returns>
            /// <exception cref="System.ArgumentException"><paramref name="language"/> is not a well-formed language identifier</exception>
            /// <exception cref="System.NotSupportedException"><paramref name="language"/> is not supported</exception>
            /// <exception cref="PlatformNotSupportedException">The OS platform is not supported</exception>
            public static WordsSegmenter Create(string language)
            {
                return new WordsSegmenter(language);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="text"></param>
            /// <param name="startIndex"></param>
            /// <returns></returns>
            /// <exception cref="System.IndexOutOfRangeException"></exception>
            /// <exception cref="ArgumentNullException">
            /// <i>text</i> is null; or internal reference to WinRT WordsSegmenter object is null
            /// </exception>
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
            public WordSegment GetTokenAt(string text, uint startIndex)
            {
                const int E_BOUNDS = -2147483637; //0x8000000B;

                if (text == null)
                {
                    throw new ArgumentNullException("text");
                }

                try
                {
                    object token = _wordsSegmenter.ReflectionCall<string, uint>("GetTokenAt", text, startIndex);

                    if (token != null)
                    {
                        return new WordSegment(token);
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (TargetInvocationException tex) when (tex.InnerException?.HResult == E_BOUNDS)
                {
                    // E_BOUNDS is not an HRESULT with an exception mapping defined
                    // at https://github.com/dotnet/coreclr/blob/master/src/vm/rexcep.h
                    // This results in a System.Exception being thrown with HRESULT == E_BOUNDS
                    
                    throw new IndexOutOfRangeException(string.Empty, tex);
                }
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="text"></param>
            /// <returns></returns>
            /// <exception cref="ArgumentNullException"><i>text</i> is null</exception>
            /// <exception cref="NullReferenceException">Internal reference to WinRT WordsSegmenter object is null</exception>
            public IReadOnlyList<WordSegment> GetTokens(string text)
            {
                object tokens = _wordsSegmenter.ReflectionCall<string>("GetTokens", text);

                List<WordSegment> result = new List<WordSegment>();
                foreach (var item in (IEnumerable)tokens)
                {
                    if (item != null)
                    {
                        result.Add(new WordSegment(item));
                    }
                }

                return result.AsReadOnly();
            }

            /// <summary>
            /// Not implemented. 
            /// 
            /// Including a stub definition here for completeness because this interface
            /// exists in Winodws.Data.Text.WordsSegmenter. 
            /// 
            /// Implementing this using reflection based calls is non-trivial due to the 
            /// need to pass the WinRT object a delegate. Given this is not used in
            /// spell-checking, we leave this unimplemented for now. 
            /// </summary>
            /// <param name="text"></param>
            /// <param name="startIndex"></param>
            /// <param name="handler"></param>
            /// <exception cref="NotImplementedException">Always throws</exception>
            public void Tokenize(string text, uint startIndex, WordSegmentsTokenizingHandler handler)
            {
                throw new NotImplementedException();
            }

            public string ResolvedLanguage
            {
                get
                {
                    if (_resolvedLanguage == null)
                    {
                        _resolvedLanguage = _wordsSegmenter.ReflectionGetProperty<string>("ResolvedLanguage");
                    }

                    return _resolvedLanguage;
                }
            }

            public static Type WinRTType
            {
                get
                {
                    return s_WinRTType;
                }
            }


            public static readonly string Undetermined = "und";

            private static string s_TypeName = "Windows.Data.Text.WordsSegmenter, Windows, ContentType=WindowsRuntime";

            private static Type s_WinRTType;
            private static bool s_PlatformSupported;

            private object _wordsSegmenter;
            private string _resolvedLanguage = null;
        }

        #region Internal Delegates

        // unused
        internal delegate void WordSegmentsTokenizingHandler(
            IEnumerable<WordSegment> precedingWords,
            IEnumerable<WordSegment> words);

        #endregion 
    }
}
