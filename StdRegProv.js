/**
 * @file WMI classes as inspired by mgmclassgen.exe.
 * @version 0.0.1
 */

package ROOT.CIMV2 {

  abstract class StdRegProv {

    private static var CreatedClassName: String = 'StdRegProv';

    public static function EnumKey(hDefKey: uint, sSubKeyName: String): String[] {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      try {
        var sNames = classObj.ExecMethod_(methodName, inParams).Properties_.Item('sNames').Value;
        if (sNames == null) {
          return null;
        }
        var sNameStr: String[] = new String[sNames.length];
        for (var index = 0; index < sNames.length; index++) {
          sNameStr[index] = sNames[index];
        }
        return sNameStr;
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    /**
     * Convert the UInt32 key hive value to negative integer.
     */
    static function GetKeyHiveValue(hDefKey: uint): int {
      // Append 33 1's to the 32-bit representation of hDefKey
      // to get the value read by the parameter hDefKey.
      return Convert.ToInt32(Convert.ToInt64(hDefKey) | 0x1FFFFFFFF00000000);
    }
  }

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