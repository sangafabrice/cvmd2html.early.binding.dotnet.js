@cc_on
@set @MAJOR = 0
@set @MINOR = 0
@set @BUILD = 1
@set @REVISION = 0

import System;
import System.Runtime.InteropServices;
import System.Reflection;
@if (@StdRegProvWim)
import System.Diagnostics;
import WbemScripting;
@else
import IWshRuntimeLibrary;
import Scriptlet;
import ROOT.CIMV2;
@end

[assembly: AssemblyProduct('MarkdownToHtml Shortcut')]
[assembly: AssemblyInformationalVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
[assembly: AssemblyCopyright('\u00A9 2024 sangafabrice')]
[assembly: AssemblyCompany('sangafabrice')]
[assembly: AssemblyVersion(@MAJOR + '.' + @MINOR + '.' + @BUILD + '.' + @REVISION)]
@if (@StdRegProvWim)
[assembly: AssemblyTitle('StdRegProv WIM')]
@else
[assembly: AssemblyTitle('CvMd2Html Launcher')]
@end