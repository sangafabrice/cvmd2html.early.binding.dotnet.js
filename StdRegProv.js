/**
 * @file WMI classes as inspired by mgmclassgen.exe.
 * @version 0.0.1
 */

package ROOT.CIMV2 {

  abstract class StdRegProv {

    private static var CreatedClassName: String = 'StdRegProv';

    /**
     * uRequired is not needed.
     * That is why I removed it from the method signature.
     */
    public static function CheckAccess(hDefKey: uint, sSubKeyName: String): Boolean {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      try {
        return Convert.ToBoolean(classObj.ExecMethod_(methodName, inParams).Properties_.Item('bGranted').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    public static function CreateKey(hDefKey: uint, sSubKeyName: String): uint {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      try {
        return Convert.ToUInt32(classObj.ExecMethod_(methodName, inParams).Properties_.Item('ReturnValue').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    public static function DeleteKey(hDefKey: uint, sSubKeyName: String): uint {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      try {
        return Convert.ToUInt32(classObj.ExecMethod_(methodName, inParams).Properties_.Item('ReturnValue').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    public static function DeleteValue(hDefKey: uint, sSubKeyName: String, sValueName: String): uint {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      inParams.Properties_.Item('sValueName').Value = sValueName;
      try {
        return Convert.ToUInt32(classObj.ExecMethod_(methodName, inParams).Properties_.Item('ReturnValue').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

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

    public static function GetStringValue(hDefKey: uint, sSubKeyName: String, sValueName: String): String {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      if (hDefKey) {
        inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      }
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      inParams.Properties_.Item('sValueName').Value = sValueName;
      try {
        return Convert.ToString(classObj.ExecMethod_(methodName, inParams).Properties_.Item('sValue').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    public static function SetStringValue(hDefKey: uint, sSubKeyName: String, sValueName: String, sValue: String): uint {
      var stackTrace: StackTrace = new StackTrace();
      var methodName: String = Util.GetMethodName(stackTrace);
      var classObj: SWbemObject = (new SWbemLocatorClass()).ConnectServer().Get(CreatedClassName);
      var inParams = null;
      inParams = classObj.Methods_.Item(methodName).InParameters.SpawnInstance_();
      inParams.Properties_.Item('hDefKey').Value = GetKeyHiveValue(hDefKey);
      inParams.Properties_.Item('sSubKeyName').Value = sSubKeyName;
      inParams.Properties_.Item('sValueName').Value = sValueName;
      inParams.Properties_.Item('sValue').Value = sValue;
      try {
        return Convert.ToUInt32(classObj.ExecMethod_(methodName, inParams).Properties_.Item('ReturnValue').Value);
      } finally {
        Marshal.FinalReleaseComObject(classObj);
        Marshal.FinalReleaseComObject(inParams);
        classObj = null;
        inParams = null;
      }
    }

    /**
     * Remove the key and all descendant subkeys.
     * @borrows DeleteKey as DeleteKeyTree
     */
    public static function DeleteKeyTree(hDefKey: uint, sSubKeyName: String): uint {
      var returnValue: uint = 0;
      var sNames = EnumKey(hDefKey, sSubKeyName);
      if (sNames != null) {
        for (var index = 0; index < sNames.length; index++) {
          returnValue += DeleteKeyTree(hDefKey, sSubKeyName + '\\' + sNames[index]);
        }
      }
      return (returnValue += DeleteKey(hDefKey, sSubKeyName));
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
}