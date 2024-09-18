@cc_on
import System;
import System.Runtime.InteropServices;
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