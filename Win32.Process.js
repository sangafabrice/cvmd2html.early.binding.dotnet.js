/**
 * @file WMI classes as inspired by mgmclassgen.exe.
 * @version 0.0.1
 */
import System;
import System.Diagnostics;
import System.Runtime.InteropServices;
import System.Reflection;
import WbemScripting;

[assembly: AssemblyTitle('Win32_Process WIM')]

package ROOT.CIMV2.WIN32 {

  abstract class Process {

    /**
     * Initializing the Process.CurrentDirectory causes an issue.
     * That is why I removed it from the method signature.
     */
    public static function Create(CommandLine: String, ProcessStartupInformation: SWbemObject): uint {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get('Win32_Process');
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('CommandLine').Value = CommandLine;
      inParams.Properties_.Item('ProcessStartupInformation').Value = ProcessStartupInformation;
      try {
        return Convert.ToUInt32(classObj.ExecMethod_(methodName, inParams).Properties_.Item('ProcessId').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    /**
     * Use-case specific method. ProcessStartupInformation is made option.
     */
    public static function Create(CommandLine: String): uint {
      return Create(CommandLine, ProcessStartup.CreateInstance());
    }
  }

  abstract class ProcessStartup {

    /**
     * We change the default value of the ProcessStartup instance to HIDDEN.
     */
    public static function CreateInstance(): SWbemObject {
      var WINDOW_STYLE_HIDDEN = 0xC;
      var startInfo = null;
      startInfo = (new SWbemLocatorClass()).ConnectServer().Get('Win32_ProcessStartup').SpawnInstance_();
      startInfo.Properties_.Item('ShowWindow').Value = WINDOW_STYLE_HIDDEN;
      return startInfo;
    }
  }

  internal abstract class Util {

    public static function GetMethodName(stackTrace: StackTrace): String {
      return stackTrace.GetFrame(0).GetMethod().Name;
    }
  }
}