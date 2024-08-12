#Requires -Version 6.1
using namespace System.IO
using namespace System.Management.Automation
[CmdletBinding()]
Param ()

Function Build-MarkdownToHtmlShortcut {
  <#
  .SYNOPSIS
  Build the project around one or more executables.
  .DESCRIPTION
  This function creates the binary and resource files that are ready for distribution.
  .OUTPUTS
  The executable path string.
  #>
  [CmdletBinding()]
  Param ()

  $HostColorArgs = @{
    ForegroundColor = 'Black'
    BackgroundColor = 'Green'
    NoNewline = $True
  }
  # Set the extension of the file with base name Convert-MarkdownToHtml and return full path.
  Function Private:Set-ConvertMd2HtmlExtension([string] $Extension) {
    Return "$PSScriptRoot\Convert-MarkdownToHtml$Extension"
  }
  Try {
    Remove-Item ($ConvertExe = Set-ConvertMd2HtmlExtension '.exe') -ErrorAction Stop
  } Catch [ItemNotFoundException] {
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
  } Catch {
    $HostColorArgs.BackgroundColor = 'Red'
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
    Return
  }
  Remove-Item ($StdRegProvDll = "$PSScriptRoot\StdRegProv.dll") -ErrorAction SilentlyContinue
  Remove-Item ($SWbemDllPath = "$PSScriptRoot\Interop.WbemScripting.dll") -ErrorAction SilentlyContinue
  Remove-Item ($WshDllPath = "$PSScriptRoot\Interop.IWshRuntimeLibrary.dll") -ErrorAction SilentlyContinue
  # Import the dependency libraries.
  & "$PSScriptRoot\TlbImp.exe" /nologo /silent 'C:\Windows\System32\wshom.ocx'  /out:$WshDllPath /namespace:IWshRuntimeLibrary
  & "$PSScriptRoot\TlbImp.exe" /nologo /silent 'C:\Windows\System32\wbem\wbemdisp.tlb' /out:$SWbemDllPath /namespace:WbemScripting
  # Compile the launcher script into an .exe file of the same base name.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:windir\Microsoft.NET\Framework$(If ([Environment]::Is64BitOperatingSystem) { '64' })\v4.0.30319\;$Env:Path"
  jsc.exe /nologo /target:library /reference:$SWbemDllPath /out:$StdRegProvDll "$PSScriptRoot\AssemblyInfo.js" "$PSScriptRoot\StdRegProv.js"
  jsc.exe /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /reference:$SWbemDllPath /reference:$WshDllPath /reference:$StdRegProvDll /out:$(($ConvertExe = Set-ConvertMd2HtmlExtension '.exe')) "$PSScriptRoot\AssemblyInfo.js" $(Set-ConvertMd2HtmlExtension '.js')
  $Env:Path = $EnvPath
  If ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs
    (Get-Item $ConvertExe).VersionInfo | Format-List * -Force
  }
}

Build-MarkdownToHtmlShortcut -Debug:$($DebugPreference -eq 'Continue')