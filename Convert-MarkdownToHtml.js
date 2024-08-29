/**
 * @file Converts markdown file content to html document.
 * @version 1.0.0
 */

/**
 * The parameters and arguments.
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 * @property {boolean} RunLink restarts the runtime with the custom icon.
 */

/** @type {ParamHash} */
var param = GetParameters(Environment.GetCommandLineArgs());

if (param.Markdown) {
(function() {
/* The app module */

// The file system object must be initialized first.
var fs: FileSystemObject = new FileSystemObjectClass();
// Append 1 to the left of the binary representation of Exception HResult to force
// 64-bit computation. The value of the HResults here should be read as a 64-bit 
// negative integer instead of 32-bit unsigned in JScript. This is for clarity.
/** @constant {int64} */
var PERMISSION_DENIED = 0x800A0046 | 0x100000000;
/** @constant {int64} */
var FILE_NOT_FOUND = 0x800A0035 | 0x100000000;
/** @class */
var MessageBox = GetMessageBoxType();
/** @constant {regexp} */
var MARKDOWN_REGEX = /\.md$/i;
CheckMarkdown();
ConvertTo(GetHtmlPath());
ReleaseFileSystem();
Quit(0);

/**
 * Validate the input markdown path string.
 */
function CheckMarkdown() {
  if (!MARKDOWN_REGEX.test(param.Markdown)) {
    MessageBox.Show(Format('"{0}" is not a markdown (.md) file.', param.Markdown));
  }
}

/**
 * Convert the content of the markdown file to html.
 * @param {string} htmlPath is the output html path.
 */
function ConvertTo(htmlPath) {
  SetHtmlContent(htmlPath, ConvertFrom(GetContent(param.Markdown)));
}

/**
 * Write the html text to the output HTML file.
 * It notifies the user when the operation did not complete with success.
 * @param {string} htmlPath is the output html path.
 * @param {string} htmlContent is the content of the html file.
 */
function SetHtmlContent(htmlPath, htmlContent) {
  var FOR_WRITING = 2;
  try {
    var txtStream: TextStream = fs.OpenTextFile(htmlPath, FOR_WRITING, true);
    txtStream.Write(htmlContent);
  } catch (error: Exception) {
    if (error.HResult == PERMISSION_DENIED) {
      MessageBox.Show(Format('Access to the path "{0}" is denied.', htmlPath));
    } else {
      MessageBox.Show(Format('Unspecified error trying to write to "{0}".', htmlPath));
    }
  } finally {
    if (txtStream != undefined) {
      txtStream.Close();
      Marshal.FinalReleaseComObject(txtStream);
      txtStream = null;
    }
  }
}

/**
 * This function returns the output path when it is unique
 * without prompts or when the user accepts to overwrite an
 * existing HTML file. Otherwise, it exits the script.
 * @returns {string} the output html path.
 */
function GetHtmlPath() {
  var htmlPath = param.Markdown.replace(MARKDOWN_REGEX, '.html');
  if (fs.FileExists(htmlPath)) {
    MessageBox.Show(Format('The file "{0}" already exists.\n\nDo you want to overwrite it?', htmlPath), MessageBox.WARNING);
  } else if (fs.FolderExists(htmlPath)) {
    MessageBox.Show(Format('"{0}" cannot be overwritten because it is a directory.', htmlPath));
  }
  return htmlPath;
}

/**
 * Get the content of a file.
 * @param {string} filePath is path that is read.
 * @returns {string} the content of the file.
 */
function GetContent(filePath) {
  var FOR_READING = 1;
  try {
    with (fs.OpenTextFile(filePath, FOR_READING)) {
      var content = ReadAll();
      Close();
    }
    return content;
  } catch (error: Exception) {
    switch (error.HResult) {
      case PERMISSION_DENIED:
        if (!fs.FolderExists(filePath)) {
          MessageBox.Show(Format('Access to the path "{0}" is denied.', filePath));
        }
      case FILE_NOT_FOUND:
        MessageBox.Show(Format('File "{0}" is not found.', filePath));
      default:
        MessageBox.Show(Format('Unspecified error trying to read from "{0}".', filePath));
    }
  } 
}

/**
 * Convert a markdown content to an html document.
 * @param {string} mardownContent is the content to convert.
 * @returns {string} the output html document content. 
 */
function ConvertFrom(markdownContent) {
  // Build the HTML document that will load the showdown.js library.
  var document: HTMLDocumentClass = new HTMLDocumentClass();
  document.open();
  document.IHTMLDocument2_write(GetContent(ChangeScriptExtension('.html')));
  document.close();
  document.body.innerText = markdownContent;
  document.parentWindow.execScript('convertMarkdown()', 'javascript');
  try {
    return document.body.innerHTML;
  } finally {
    if (document != undefined) {
      Marshal.FinalReleaseComObject(document);
      document = null;
    }
  }
}

/**
 * @returns the MessageBox type.
 */
function GetMessageBoxType() {
  /** @private @constant {number} */
  var MESSAGE_BOX_TITLE: Object = 'Convert to HTML';
  /** @private @constant {number} */
  var NO_MESSAGE_TIMEOUT: Object = 0;
  /** @private @constant {number} */
  var YESNO_BUTTON: Object = 4;
  /** @private @constant {number} */
  var OK_BUTTON: Object = 0;
  /** @private @constant {number} */
  var OK_POPUPRESULT = 1;
  /** @private @constant {number} */
  var NO_POPUPRESULT = 7;
  /** @private @constant {number} */
  var ERROR_MESSAGE: Object = 16;
  /** @private @constant {number} */
  var WARNING_MESSAGE: Object = 48;

  /**
   * Represents the markdown conversion message box.
   * @typedef {object} MessageBox
   * @property {number} WARNING specifies that the dialog shows a warning message.
   */
  var MessageBox = { };

  /** @public @static @readonly @property {number} */
  MessageBox.WARNING = WARNING_MESSAGE;
  // Object.defineProperty() method does not work in WSH.
  // It is not possible in this implementation to make the
  // property non-writable.

  /**
   * Show a warning message or an error message box.
   * @public @static @method Show @memberof MessageBox
   * @param {string} message is the message text.
   * @param {number} [messageType = ERROR_MESSAGE] message box type (Warning/Error).
   */
  MessageBox.Show = function(message, messageType) {
    if (messageType != ERROR_MESSAGE && messageType != WARNING_MESSAGE) {
      messageType = ERROR_MESSAGE;
    }
    // The error message box shows the OK button alone.
    // The warning message box shows the alternative Yes or No buttons.
    messageType += messageType == ERROR_MESSAGE ? OK_BUTTON:YESNO_BUTTON;
    switch ((new WshShellClass()).Popup(message, NO_MESSAGE_TIMEOUT, MESSAGE_BOX_TITLE, messageType)) {
      case OK_POPUPRESULT:
      case NO_POPUPRESULT:
        ReleaseFileSystem();
        Quit(1);
    }
  }

  return MessageBox;
}

/**
 * Release the file system com object.
 */
function ReleaseFileSystem() {
  Marshal.FinalReleaseComObject(fs);
  fs = null;
}
})();
}

if (param.Set || param.Unset) {
  var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
  if (param.Set) {
    var VERB_KEY = Format('HKCU\\{0}\\', VERB_KEY);
    var COMMAND_KEY = VERB_KEY + 'command\\';
    var VERBICON_VALUENAME = VERB_KEY + 'Icon';
    var registry: WshShell = new WshShellClass();
    // Configure the shortcut menu in the registry.
    var command = Format('"{0}" /Markdown:"%1"', param.ApplicationPath);
    registry.RegWrite(COMMAND_KEY, command);
    registry.RegWrite(VERB_KEY, 'Convert to &HTML');
    if (param.NoIcon) {
      try {
        registry.RegDelete(VERBICON_VALUENAME);
      } catch (error) { }
    } else {
      registry.RegWrite(VERBICON_VALUENAME, param.ApplicationPath);
    }
    Marshal.FinalReleaseComObject(registry);
    registry = null;
  } else if (param.Unset) {
    // Remove the shortcut menu.
    // Remove the verb key and subkeys.
    var HKCU = 0x80000001;
    StdRegProv.DeleteKeyTree(HKCU, VERB_KEY);
  }
  Quit(0);
}

Quit(1);

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
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path> [/RunLink]\n';
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