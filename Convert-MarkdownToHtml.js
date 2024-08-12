/**
 * @file Launches the shortcut target PowerShell
 * script with the selected markdown as an argument.
 * It aims to eliminate the flashing console window
 * when the user clicks on the shortcut menu.
 * @version 0.0.1
 */
import System;
import System.Runtime.InteropServices;
import System.Reflection;
import Shell32;
import IWshRuntimeLibrary;
import ROOT.CIMV2;
import ROOT.CIMV2.WIN32;

[assembly: AssemblyTitle('CvMd2Html Launcher')]

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
    var shortcutIconPath = ChangeScriptExtension('.ico');
    // Create the link with the partial "arguments" string.
    var linkPath = ChangeScriptExtension('.lnk');
    (new FileSystemObjectClass()).CreateTextFile(linkPath).Close();
    var link: ShellLinkObject = GetCustomIconLink(linkPath);
    link.Path = GetPwshPath();
    link.Arguments = GetCustomIconLinkArguments();
    link.SetIconLocation(shortcutIconPath, 0);
    link.Save();
    Marshal.FinalReleaseComObject(link);
    link = null;
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
      StdRegProv.SetStringValue(HKCU, VERB_KEY, iconValueName, shortcutIconPath);
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
  var linkPath = ChangeScriptExtension('.lnk');
  if (!IsLinkReady(GetCustomIconLink(linkPath))) {
    return;
  }
  Process.Create(Format('C:\\Windows\\System32\\cmd.exe /d /c ""{0}" "{1}""', [linkPath, markdown]));
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
 * Check the link target command.
 * @param {object} link is the shortcut link.
 * @param {string} link.Path is the path to the runner.
 * @param {string} link.Arguments is the target command line list of arguments.
 * @returns {boolean} True if the target command is as expected, false otherwise.
 */
function IsLinkReady(link: ShellLinkObject) {
  var format = '{0} {1}';
  return Format(format, [link.Path, link.Arguments]).toLowerCase() == Format(format, [GetPwshPath(), GetCustomIconLinkArguments()]).toLowerCase();
}

/**
 * Get the custom icon link object from its path.
 * @param {string} linkPath is the custom icon link path.
 * @returns {object} the specified link file object.
 */
function GetCustomIconLink(linkPath): ShellLinkObject {
  var fs: FileSystemObject = new FileSystemObjectClass();
  try {
    return (new ShellClass()).NameSpace(fs.GetParentFolderName(linkPath)).ParseName(fs.GetFileName(linkPath)).GetLink;
  } finally {
    Marshal.FinalReleaseComObject(fs);
    fs = null;
  }
}

/**
 * Get the partial "arguments" property string of the custom icon link.
 * The command is partial because it does not include the markdown file path string.
 * The markdown file path string will be input when calling the shortcut link.
 * @returns {string} the "arguments" property of the custom icon link.
 */
function GetCustomIconLinkArguments() {
  return Format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown', ChangeScriptExtension('.ps1'));
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