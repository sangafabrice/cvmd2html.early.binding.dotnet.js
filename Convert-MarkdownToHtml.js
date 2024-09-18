/**
 * @file Launches the shortcut target PowerShell
 * script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window
 * when the user clicks on the shortcut menu.
 * @version 0.0.1
 */

RequestAdminPrivileges(Environment.GetCommandLineArgs())

/**
 * The parameters and arguments.
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @type {ParamHash} */
var param = GetParameters(Environment.GetCommandLineArgs());

if (param.Markdown) {
  StartWith(param.Markdown);
  Quit(0);
}

if (param.Set || param.Unset) {
  var HKCU = 0x80000001;
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  if (param.Set) {
    // Configure the shortcut menu in the registry.
    var COMMAND_KEY = VERB_KEY + '\\command';
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    StdRegProv.CreateKey(HKCU, COMMAND_KEY);
    StdRegProv.SetStringValue(HKCU, COMMAND_KEY, null, command);
    StdRegProv.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
    var iconValueName = 'Icon';
    if (param.NoIcon) {
      StdRegProv.DeleteValue(HKCU, VERB_KEY, iconValueName);
    } else {
      StdRegProv.SetStringValue(HKCU, VERB_KEY, iconValueName, ChangeScriptExtension('.ico'));
    }
  } else if (param.Unset) {
    // Remove the shortcut menu.
    // Remove the verb key and subkeys.
    StdRegProv.DeleteKeyTree(HKCU, VERB_KEY);
  }
  Quit(0);
}

Quit(1);

/**
 * Start the shortcut target PowerShell script with
 * the path of the selected markdown file as an argument.
 * @param {string} markdown is the input markdown path argument.
 */
function StartWith(markdown) {
  var linkPath = GetDynamicLinkPathWith(markdown);
  if (Process.WaitForExit(Process.Create(Format('C:\\Windows\\System32\\cmd.exe /d /c "{0}"', linkPath)))) {
    var MESSAGEBOX_TITLE: Object = 'Convert to HTML';
    var OKONLY_BUTTON: Object = 0;
    var ERROR_ICON: Object = 16;
    var NO_TIMEOUT: Object = 0;
    (new WshShellClass()).Popup('An unhandled exception occured.', NO_TIMEOUT, MESSAGEBOX_TITLE, OKONLY_BUTTON + ERROR_ICON)
  }
  (new FileSystemObjectClass()).DeleteFile(linkPath, true);
}

/**
 * Save and get the dynamic link.
 * @param {string} markdown is the input markdown path argument.
 * @returns {string} the link path.
 */
function GetDynamicLinkPathWith(markdown) {
  var tempDir = (new WshShellClass()).ExpandEnvironmentStrings('%TEMP%');
  var tempLinkName = IGenScriptletTLib.GUID + '.tmp.lnk';
  var linkPath = Format('{0}\\{1}', [tempDir, tempLinkName]);
  (new FileSystemObjectClass()).CreateTextFile(linkPath).Close();
  var link: ShellLinkObject = (new ShellClass()).NameSpace(tempDir).ParseName(tempLinkName).GetLink;
  link.Path = GetPwshPath();
  link.Arguments = Format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown "{1}"', [ChangeScriptExtension('.ps1'), markdown]);
  link.SetIconLocation(ChangeScriptExtension('.ico'), 0);
  link.Save();
  Marshal.FinalReleaseComObject(link);
  link = null;
  return linkPath;
}

/**
 * Change the launcher script path extension.
 * This change implies that the launcher script and the resulting
 * path file reside in the same directory and have the same name.
 * @param {string} extension is the new extension.
 * @returns {string} a file path with the new extension.
 */
function ChangeScriptExtension(extension) {
  return param.ApplicationPath.replace(/\.exe$/i, extension);
}

/**
 * Get the PowerShell Core application path from the registry.
 * @returns {string} the pwsh.exe full path or an empty string.
 */
function GetPwshPath() {
  // The HKLM registry subkey stores the PowerShell Core application path.
  return StdRegProv.GetStringValue(null, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe', null);
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} format the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function Format(format, args) {
  if (args.constructor !== Array) {
    return format.replace(/\{0\}/g, args);
  }
  while (args.length > 0) {
    format = format.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return format;
}

/**
 * Get the input arguments and parameters.
 * @param {string[]} args is the list of command line arguments including the command path.
 * @returns {ParamHash|undefined} a value-name pair of arguments.
 */
function GetParameters(args) {
  var applicationPath = args[0];
  if (args.length == 2) {
    var arg = args[1];
    var param = { ApplicationPath: applicationPath };
    var paramName = arg.split(':', 1)[0].toLowerCase();
    if (paramName == '/markdown') {
      param.Markdown = arg.replace(new RegExp('^' + paramName + ':?', 'i'), '')
      if (param.Markdown.length > 0) {
        return param;
      }
    }
    switch (arg.toLowerCase()) {
      case '/set':
        param.Set = true;
        param.NoIcon = false;
        return param;
      case '/set:noicon':
        param.Set = true;
        param.NoIcon = true;
        return param;
      case '/unset':
        param.Unset = true;
        return param;
      case '/help':
        param.Help = true;
        return param;
    }
  } else if (args.length == 1) {
    return {
      Set: true,
      NoIcon: false,
      ApplicationPath: applicationPath
    }
  }
  ShowHelp();
}

/**
 * Show help and quit.
 */
function ShowHelp() {
  var helpText = '';
  helpText += 'The MarkdownToHtml shortcut launcher.\n';
  helpText += 'It starts the shortcut menu target script in a hidden window.\n\n';
  helpText += 'Syntax:\n';
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml.js [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml.js /Unset\n';
  helpText += '  Convert-MarkdownToHtml.js /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  (new WshShellClass()).Popup(helpText, 0);
  Quit(1);
}

/**
 * Clean up and quit.
 * @param {number} exitCode .
 */
function Quit(exitCode) {
  GC.Collect();
  Environment.Exit(exitCode);
}

/**
 * Request administrator privileges is standard user.
 * @param {string[]} args is the list of command line arguments including the command path.
 */
function RequestAdminPrivileges(args) {
  var HKU: uint = 0x80000003;
  if (StdRegProv.CheckAccess(HKU, 'S-1-5-19\\Environment')) {
    return;
  }
  var inputCommand = '';
  for (var index = 1; index < args.length; index++) {
    inputCommand += Format(' "{0}"', args[index]);
  }
  var WINDOW_STYLE_HIDDEN = 0;
  (new ShellClass()).ShellExecute(args[0], inputCommand, null, "runas", WINDOW_STYLE_HIDDEN)
  Quit(0);
}