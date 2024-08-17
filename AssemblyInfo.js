@cc_on
@set @MAJOR = 0
@set @MINOR = 0
@set @BUILD = 1
@set @REVISION = 0

import System;
import System.Runtime.InteropServices;
import System.Reflection;
@if (@StdRegProvWim || @Win32ProcessWim)
import System.Diagnostics;
import WbemScripting;
@else
import Shell32;
import IWshRuntimeLibrary;
import Scriptlet;
import ROOT.CIMV2;
import ROOT.CIMV2.WIN32;
@end

[assembly: AssemblyProduct('MarkdownToHtml Shortcut')]
[assembly: AssemblyInformationalVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
[assembly: AssemblyCopyright('\u00A9 2024 sangafabrice')]
[assembly: AssemblyCompany('sangafabrice')]
[assembly: AssemblyVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
@if (@StdRegProvWim)
[assembly: AssemblyTitle('StdRegProv WIM')]
@elif (@Win32ProcessWim)
[assembly: AssemblyTitle('Win32_Process WIM')]
@else
[assembly: AssemblyTitle('CvMd2Html Launcher')]
@end

@if (@StdRegProvWim)
package ROOT.CIMV2 {
@elif (@Win32ProcessWim)
package ROOT.CIMV2.WIN32 {
@end
@if (@StdRegProvWim || @Win32ProcessWim)
  internal abstract class Util {

    /**
     * Get the name of the method calling this method.
     * NOTE: the method should initialize the stackTrace variable in its scope
     * before calling GetMethodName. So avoid GetMethodName(new stackTrace()).
     * @param stackTrace is the stack trace from the calling method.
     */
    public static function GetMethodName(stackTrace: StackTrace): String {
      return stackTrace.GetFrame(0).GetMethod().Name;
    }
  }
}
@end