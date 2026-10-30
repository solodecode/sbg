# Balanced Items mod installer for Super Battle Golf (BepInEx 5, Steam app 4069520).
#
# Usage (friends):   irm https://<domain>/sbg/install.ps1 | iex
# With parameters:   & ([scriptblock]::Create((irm https://<domain>/sbg/install.ps1))) -Uninstall
#
# Source-file notes (for developers):
#  * Windows PowerShell 5.1 compatible (no ternary, no &&, no PS7-only features), no admin rights needed.
#  * Everything except the MessagesJson block is pure ASCII. This file is saved as UTF-8 WITH BOM so that
#    "powershell -File" reads the Russian text correctly. pack.ps1 turns every non-ASCII character of the
#    block into a \uXXXX escape, so the published dist\install.ps1 is pure ASCII and survives any web-server
#    charset / BOM / line-ending handling by "irm | iex".
#  * No top-level variables or functions: the whole script runs inside "& { }" so that "iex" does not leak
#    anything into the caller's session. "exit" is only used when the script runs from a file.
#  * Exit codes: 0 ok, 1 unexpected error, 2 game not found, 3 game is running, 4 network / download /
#    checksum / archive error, 5 file-system error (all changes rolled back), 6 bad parameters.
#  * The package layout accepted by this installer is mirrored in pack.ps1 (Test-AllowedPackagePath).

<#
.SYNOPSIS
    Installs, updates, disables or removes the "Balanced Items" mod (BepInEx 5) for Super Battle Golf.
.PARAMETER Uninstall
    Removes the mod plugin folder and the install receipt (BepInEx itself stays).
.PARAMETER Full
    Together with -Uninstall: also removes BepInEx (core, winhttp.dll, doorstop files) if this installer created it.
.PARAMETER Disable
    Renames winhttp.dll to winhttp.dll.off: the game starts without any mods (vanilla).
.PARAMETER Enable
    Renames winhttp.dll.off back to winhttp.dll.
.PARAMETER DisableConflicts
    Renames DLLs of conflicting plugins (IssaMod, SBGItemFramework, ...) to *.disabled.
.PARAMETER DryRun
    Prints what would be done, changes nothing.
.PARAMETER Quiet
    Prints only warnings, errors and the final result line.
.PARAMETER GamePath
    Game folder (the one containing "Super Battle Golf.exe"). Skips the Steam lookup.
.PARAMETER BaseUrl
    Server folder with latest.json (baked in by pack.ps1).
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$Full,
    [switch]$Disable,
    [switch]$Enable,
    [switch]$DisableConflicts,
    [switch]$DryRun,
    [switch]$Quiet,
    [string]$GamePath,
    [string]$BaseUrl = 'https://solodecode.github.io/sbg',
    [Parameter(DontShow = $true)][string]$SteamPathOverride,
    [Parameter(DontShow = $true)][string]$GameProcessName = 'Super Battle Golf',
    [Parameter(DontShow = $true)][switch]$AllowInsecureHttp
)

& {
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'

    # ======================================================================================== constants
    $AppId             = '4069520'
    $GameExe           = 'Super Battle Golf.exe'
    $DefaultFolderName = 'Super Battle Golf'
    $PluginName        = 'SbgBalancedItems'
    $PluginRel         = 'BepInEx/plugins/SbgBalancedItems'
    $ReceiptRel        = 'BepInEx/sbgbi-install.json'
    $SupportedProtocol = 1
    $UserAgent         = 'SbgBalancedItems-Installer/1 (Windows PowerShell)'
    $RootFiles         = @('winhttp.dll', 'doorstop_config.ini', '.doorstop_version')
    # Plugins that add their own items / lobby version checks and may clash with this mod (warning only).
    $ConflictPattern   = '(?i)issa|itemframework|sbgitem|customitem|ballswap|moreitems|extraitems|newitems'
    $Exit = @{ Ok = 0; Error = 1; GameNotFound = 2; GameRunning = 3; Network = 4; FileSystem = 5; Usage = 6 }
    $Utf8Bom   = New-Object System.Text.UTF8Encoding($true)
    $Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $State = @{
        ExitCode = 0
        LogPath  = (Join-Path ([IO.Path]::GetTempPath()) 'sbgbi-install.log')
    }

    # ======================================================================================== messages
    # All user-facing text (Russian). Placeholders {0}, {1}, ... are filled with -f. Keep this block free of
    # backslashes and braces other than placeholders. pack.ps1 escapes non-ASCII characters here to \uXXXX.
    $MessagesJson = @'
{
  "title": "Super Battle Golf: \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a \u043c\u043e\u0434\u0430 \u00abBalanced Items\u00bb",
  "dryrun_banner": "\u041f\u0420\u041e\u0411\u041d\u042b\u0419 \u0417\u0410\u041f\u0423\u0421\u041a: \u043d\u0438\u0447\u0435\u0433\u043e \u043d\u0435 \u0431\u0443\u0434\u0435\u0442 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u043e, \u043d\u0438\u0436\u0435 \u0442\u043e\u043b\u044c\u043a\u043e \u043f\u043b\u0430\u043d \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u0439.",
  "step_latest": "[1/5] \u041f\u0440\u043e\u0432\u0435\u0440\u044f\u044e \u043f\u043e\u0441\u043b\u0435\u0434\u043d\u044e\u044e \u0432\u0435\u0440\u0441\u0438\u044e \u043c\u043e\u0434\u0430 \u043d\u0430 \u0441\u0435\u0440\u0432\u0435\u0440\u0435...",
  "latest_version": "      \u0414\u043e\u0441\u0442\u0443\u043f\u043d\u0430 \u0432\u0435\u0440\u0441\u0438\u044f: {0}",
  "latest_tested": "      \u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e \u043d\u0430 Super Battle Golf {0} (\u0441\u0431\u043e\u0440\u043a\u0430 Steam {1}).",
  "latest_notes": "      \u0427\u0442\u043e \u043d\u043e\u0432\u043e\u0433\u043e:",
  "step_game": "[2/5] \u0418\u0449\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u0443\u044e \u0438\u0433\u0440\u0443...",
  "game_found": "      \u0418\u0433\u0440\u0430 \u043d\u0430\u0439\u0434\u0435\u043d\u0430: {0}",
  "game_from_arg": "      \u0418\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u044e \u043f\u0430\u043f\u043a\u0443 \u0438\u0433\u0440\u044b \u0438\u0437 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u0430 -GamePath: {0}",
  "game_not_found": "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043d\u0430\u0439\u0442\u0438 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043d\u0443\u044e \u0438\u0433\u0440\u0443 Super Battle Golf.\n\u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u044b: \u0440\u0435\u0435\u0441\u0442\u0440 Steam \u0438 \u0432\u0441\u0435 \u043f\u0430\u043f\u043a\u0438 \u0431\u0438\u0431\u043b\u0438\u043e\u0442\u0435\u043a Steam.\n\u0423\u043a\u0430\u0436\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443 \u0441 \u0438\u0433\u0440\u043e\u0439 \u0432\u0440\u0443\u0447\u043d\u0443\u044e (\u0432 Steam: \u043f\u0440\u0430\u0432\u044b\u0439 \u043a\u043b\u0438\u043a \u043f\u043e \u0438\u0433\u0440\u0435, \u00ab\u0423\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u0438\u0435\u00bb, \u00ab\u041f\u0440\u043e\u0441\u043c\u043e\u0442\u0440\u0435\u0442\u044c \u043b\u043e\u043a\u0430\u043b\u044c\u043d\u044b\u0435 \u0444\u0430\u0439\u043b\u044b\u00bb), \u043d\u0430\u043f\u0440\u0438\u043c\u0435\u0440:\n  {0}",
  "game_path_bad": "\u0412 \u043f\u0430\u043f\u043a\u0435 \u00ab{0}\u00bb \u043d\u0435\u0442 \u0444\u0430\u0439\u043b\u0430 Super Battle Golf.exe. \u0423\u043a\u0430\u0436\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443, \u0432 \u043a\u043e\u0442\u043e\u0440\u043e\u0439 \u043b\u0435\u0436\u0438\u0442 \u0438\u0433\u0440\u0430.",
  "game_build_differs": "      \u0412\u043d\u0438\u043c\u0430\u043d\u0438\u0435: \u0443 \u0432\u0430\u0441 \u0438\u0433\u0440\u0430 \u0441\u0431\u043e\u0440\u043a\u0438 Steam {0}, \u0430 \u043c\u043e\u0434 \u043f\u0440\u043e\u0432\u0435\u0440\u044f\u043b\u0441\u044f \u043d\u0430 \u0441\u0431\u043e\u0440\u043a\u0435 {1}. \u041e\u0431\u044b\u0447\u043d\u043e \u044d\u0442\u043e \u043d\u0435 \u0441\u0442\u0440\u0430\u0448\u043d\u043e, \u043d\u043e \u0435\u0441\u043b\u0438 \u043f\u043e\u0441\u043b\u0435 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f \u0438\u0433\u0440\u044b \u0447\u0442\u043e-\u0442\u043e \u0441\u043b\u043e\u043c\u0430\u043b\u043e\u0441\u044c, \u0432\u044b\u043f\u043e\u043b\u043d\u0438\u0442\u0435 {2} \u0438 \u043d\u0430\u043f\u0438\u0448\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443 \u043c\u043e\u0434\u0430.",
  "game_running": "\u0418\u0433\u0440\u0430 Super Battle Golf \u0441\u0435\u0439\u0447\u0430\u0441 \u0437\u0430\u043f\u0443\u0449\u0435\u043d\u0430. \u0417\u0430\u043a\u0440\u043e\u0439\u0442\u0435 \u0435\u0451 \u043f\u043e\u043b\u043d\u043e\u0441\u0442\u044c\u044e \u0438 \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0435\u0449\u0451 \u0440\u0430\u0437.",
  "game_running_dry": "      \u0412\u043d\u0438\u043c\u0430\u043d\u0438\u0435: \u0438\u0433\u0440\u0430 \u0441\u0435\u0439\u0447\u0430\u0441 \u0437\u0430\u043f\u0443\u0449\u0435\u043d\u0430. \u0414\u043b\u044f \u043d\u0430\u0441\u0442\u043e\u044f\u0449\u0435\u0439 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438 \u0435\u0451 \u043d\u0443\u0436\u043d\u043e \u0431\u0443\u0434\u0435\u0442 \u0437\u0430\u043a\u0440\u044b\u0442\u044c.",
  "installed_now": "      \u0421\u0435\u0439\u0447\u0430\u0441 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u0430 \u0432\u0435\u0440\u0441\u0438\u044f: {0}",
  "installed_none": "      \u041c\u043e\u0434 \u0443 \u0432\u0430\u0441 \u0435\u0449\u0451 \u043d\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d.",
  "vanilla_on": "      \u041c\u043e\u0434 \u0441\u0435\u0439\u0447\u0430\u0441 \u043e\u0442\u043a\u043b\u044e\u0447\u0451\u043d (\u0440\u0435\u0436\u0438\u043c \u00ab\u0432\u0430\u043d\u0438\u043b\u044c\u00bb). \u041f\u043e\u0441\u043b\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438 \u043e\u043d \u0441\u043d\u043e\u0432\u0430 \u0432\u043a\u043b\u044e\u0447\u0438\u0442\u0441\u044f.",
  "bepinex_foreign": "      BepInEx \u0443\u0436\u0435 \u0435\u0441\u0442\u044c \u0432 \u043f\u0430\u043f\u043a\u0435 \u0438\u0433\u0440\u044b (\u043f\u043e\u0441\u0442\u0430\u0432\u043b\u0435\u043d \u043d\u0435 \u044d\u0442\u0438\u043c \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a\u043e\u043c): \u0435\u0433\u043e \u0444\u0430\u0439\u043b\u044b \u0431\u0443\u0434\u0443\u0442 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u044b \u0434\u043e \u0432\u0435\u0440\u0441\u0438\u0438 \u0438\u0437 \u043f\u0430\u043a\u0435\u0442\u0430, \u0430 \u043f\u0440\u0438 \u0443\u0434\u0430\u043b\u0435\u043d\u0438\u0438 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u044b.",
  "bepinex_too_old": "\u0412 \u043f\u0430\u043f\u043a\u0435 \u0438\u0433\u0440\u044b \u043b\u0435\u0436\u0438\u0442 BepInEx \u0434\u0440\u0443\u0433\u043e\u0439 \u043c\u0430\u0436\u043e\u0440\u043d\u043e\u0439 \u0432\u0435\u0440\u0441\u0438\u0438 ({0}). \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a \u043d\u0435 \u0431\u0443\u0434\u0435\u0442 \u0435\u0433\u043e \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u044b\u0432\u0430\u0442\u044c, \u0447\u0442\u043e\u0431\u044b \u043d\u0438\u0447\u0435\u0433\u043e \u043d\u0435 \u0441\u043b\u043e\u043c\u0430\u0442\u044c. \u0423\u0434\u0430\u043b\u0438\u0442\u0435 \u0441\u0442\u0430\u0440\u044b\u0439 BepInEx \u0432\u0440\u0443\u0447\u043d\u0443\u044e \u0438\u043b\u0438 \u043d\u0430\u043f\u0438\u0448\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443.",
  "conflicts_found": "\u041d\u0430\u0439\u0434\u0435\u043d\u044b \u043c\u043e\u0434\u044b, \u043a\u043e\u0442\u043e\u0440\u044b\u0435 \u043c\u043e\u0433\u0443\u0442 \u043a\u043e\u043d\u0444\u043b\u0438\u043a\u0442\u043e\u0432\u0430\u0442\u044c \u0441 Balanced Items (\u0443 \u043d\u0438\u0445 \u0441\u0432\u043e\u0438 \u043f\u0440\u0435\u0434\u043c\u0435\u0442\u044b \u0438 \u0441\u0432\u043e\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0432\u0435\u0440\u0441\u0438\u0438 \u043b\u043e\u0431\u0431\u0438):",
  "conflict_item": "      - {0}",
  "conflicts_hint": "\u0420\u0435\u043a\u043e\u043c\u0435\u043d\u0434\u0443\u0435\u043c \u0438\u0445 \u043e\u0442\u043a\u043b\u044e\u0447\u0438\u0442\u044c: \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0435\u0449\u0451 \u0440\u0430\u0437 \u0441 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u043c -DisableConflicts. \u0424\u0430\u0439\u043b\u044b \u043f\u043e\u043b\u0443\u0447\u0430\u0442 \u0440\u0430\u0441\u0448\u0438\u0440\u0435\u043d\u0438\u0435 .disabled, \u0430 -Uninstall \u0432\u0435\u0440\u043d\u0451\u0442 \u0438\u0445 \u043e\u0431\u0440\u0430\u0442\u043d\u043e.",
  "conflict_disabled": "      \u041e\u0442\u043a\u043b\u044e\u0447\u0451\u043d: {0}",
  "step_download": "[3/5] \u0421\u043a\u0430\u0447\u0438\u0432\u0430\u044e \u043c\u043e\u0434 ({0} \u041a\u0411)...",
  "verifying": "      \u041f\u0440\u043e\u0432\u0435\u0440\u044f\u044e \u043a\u043e\u043d\u0442\u0440\u043e\u043b\u044c\u043d\u0443\u044e \u0441\u0443\u043c\u043c\u0443 SHA-256...",
  "hash_ok": "      \u041a\u043e\u043d\u0442\u0440\u043e\u043b\u044c\u043d\u0430\u044f \u0441\u0443\u043c\u043c\u0430 \u0441\u043e\u0432\u043f\u0430\u043b\u0430.",
  "hash_bad": "\u041a\u043e\u043d\u0442\u0440\u043e\u043b\u044c\u043d\u0430\u044f \u0441\u0443\u043c\u043c\u0430 \u0441\u043a\u0430\u0447\u0430\u043d\u043d\u043e\u0433\u043e \u0444\u0430\u0439\u043b\u0430 \u043d\u0435 \u0441\u043e\u0432\u043f\u0430\u043b\u0430 \u0441 \u043e\u0436\u0438\u0434\u0430\u0435\u043c\u043e\u0439. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430.\n\u041e\u0436\u0438\u0434\u0430\u043b\u0430\u0441\u044c: {0}\n\u041f\u043e\u043b\u0443\u0447\u0435\u043d\u0430:   {1}\n\u0427\u0430\u0449\u0435 \u0432\u0441\u0435\u0433\u043e \u044d\u0442\u043e \u043e\u0431\u0440\u044b\u0432 \u0437\u0430\u0433\u0440\u0443\u0437\u043a\u0438 \u0438\u043b\u0438 \u0443\u0441\u0442\u0430\u0440\u0435\u0432\u0448\u0438\u0439 \u043a\u0435\u0448 \u043d\u0430 \u0441\u0435\u0440\u0432\u0435\u0440\u0435: \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0435\u0449\u0451 \u0440\u0430\u0437. \u0415\u0441\u043b\u0438 \u043f\u043e\u0432\u0442\u043e\u0440\u044f\u0435\u0442\u0441\u044f, \u0441\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443.",
  "size_bad": "\u0420\u0430\u0437\u043c\u0435\u0440 \u0441\u043a\u0430\u0447\u0430\u043d\u043d\u043e\u0433\u043e \u0444\u0430\u0439\u043b\u0430 ({0} \u0431\u0430\u0439\u0442) \u043d\u0435 \u0441\u043e\u0432\u043f\u0430\u043b \u0441 \u043e\u0436\u0438\u0434\u0430\u0435\u043c\u044b\u043c ({1} \u0431\u0430\u0439\u0442). \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430. \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0435\u0449\u0451 \u0440\u0430\u0437.",
  "step_extract": "[4/5] \u0420\u0430\u0441\u043f\u0430\u043a\u043e\u0432\u044b\u0432\u0430\u044e \u0432\u043e \u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0443\u044e \u043f\u0430\u043f\u043a\u0443 \u0438 \u043f\u0440\u043e\u0432\u0435\u0440\u044f\u044e \u0441\u043e\u0434\u0435\u0440\u0436\u0438\u043c\u043e\u0435...",
  "pkg_bad_entry": "\u0412 \u0430\u0440\u0445\u0438\u0432\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u043d\u0435\u0434\u043e\u043f\u0443\u0441\u0442\u0438\u043c\u044b\u0439 \u043f\u0443\u0442\u044c \u00ab{0}\u00bb. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430. \u0421\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443 \u043c\u043e\u0434\u0430.",
  "pkg_incomplete": "\u0412 \u0430\u0440\u0445\u0438\u0432\u0435 \u043d\u0435 \u0445\u0432\u0430\u0442\u0430\u0435\u0442 \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u043d\u043e\u0433\u043e \u0444\u0430\u0439\u043b\u0430 \u00ab{0}\u00bb. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430. \u0421\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443 \u043c\u043e\u0434\u0430.",
  "pkg_too_big": "\u0410\u0440\u0445\u0438\u0432 \u043f\u043e\u0434\u043e\u0437\u0440\u0438\u0442\u0435\u043b\u044c\u043d\u043e \u0431\u043e\u043b\u044c\u0448\u043e\u0439 \u0438\u043b\u0438 \u0441\u043e\u0434\u0435\u0440\u0436\u0438\u0442 \u0441\u043b\u0438\u0448\u043a\u043e\u043c \u043c\u043d\u043e\u0433\u043e \u0444\u0430\u0439\u043b\u043e\u0432. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430.",
  "pkg_corrupt": "\u0410\u0440\u0445\u0438\u0432 \u043f\u043e\u0432\u0440\u0435\u0436\u0434\u0451\u043d \u0438 \u043d\u0435 \u043e\u0442\u043a\u0440\u044b\u0432\u0430\u0435\u0442\u0441\u044f. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u0430, \u0438\u0433\u0440\u0430 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u0430. \u0421\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443 \u043c\u043e\u0434\u0430.",
  "step_install": "[5/5] \u041a\u043e\u043f\u0438\u0440\u0443\u044e \u0444\u0430\u0439\u043b\u044b \u0432 \u043f\u0430\u043f\u043a\u0443 \u0438\u0433\u0440\u044b...",
  "config_kept": "      \u0412\u0430\u0448\u0438 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043d\u0435 \u0442\u0440\u043e\u043d\u0443\u0442\u044b: {0}",
  "files_written": "      \u0417\u0430\u043f\u0438\u0441\u0430\u043d\u043e \u0444\u0430\u0439\u043b\u043e\u0432: {0}",
  "stale_removed": "      \u0423\u0434\u0430\u043b\u0435\u043d\u043e \u0443\u0441\u0442\u0430\u0440\u0435\u0432\u0448\u0438\u0445 \u0444\u0430\u0439\u043b\u043e\u0432 \u043f\u0440\u043e\u0448\u043b\u043e\u0439 \u0432\u0435\u0440\u0441\u0438\u0438: {0}",
  "vanilla_off": "      \u0420\u0435\u0436\u0438\u043c \u00ab\u0432\u0430\u043d\u0438\u043b\u044c\u00bb \u0432\u044b\u043a\u043b\u044e\u0447\u0435\u043d: \u043c\u043e\u0434 \u0441\u043d\u043e\u0432\u0430 \u0430\u043a\u0442\u0438\u0432\u0435\u043d.",
  "install_failed": "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c \u043c\u043e\u0434: {0}",
  "access_denied": "\u041d\u0435\u0442 \u0434\u043e\u0441\u0442\u0443\u043f\u0430 \u043a \u0444\u0430\u0439\u043b\u0430\u043c \u0438\u0433\u0440\u044b. \u0423\u0431\u0435\u0434\u0438\u0442\u0435\u0441\u044c, \u0447\u0442\u043e \u0438\u0433\u0440\u0430 \u0437\u0430\u043a\u0440\u044b\u0442\u0430, \u0430\u043d\u0442\u0438\u0432\u0438\u0440\u0443\u0441 \u043d\u0435 \u0431\u043b\u043e\u043a\u0438\u0440\u0443\u0435\u0442 \u043f\u0430\u043f\u043a\u0443, \u0430 \u0435\u0441\u043b\u0438 \u0438\u0433\u0440\u0430 \u043b\u0435\u0436\u0438\u0442 \u0432 \u0437\u0430\u0449\u0438\u0449\u0451\u043d\u043d\u043e\u0439 \u043f\u0430\u043f\u043a\u0435 (\u043d\u0430\u043f\u0440\u0438\u043c\u0435\u0440, Program Files), \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 PowerShell \u043e\u0442 \u0438\u043c\u0435\u043d\u0438 \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0430.",
  "rolled_back": "\u0412\u0441\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u044b: \u043f\u0430\u043f\u043a\u0430 \u0438\u0433\u0440\u044b \u043e\u0441\u0442\u0430\u043b\u0430\u0441\u044c \u0442\u0430\u043a\u043e\u0439, \u043a\u0430\u043a\u043e\u0439 \u0431\u044b\u043b\u0430.",
  "rollback_partial": "\u041d\u0435 \u0432\u0441\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u0442\u043c\u0435\u043d\u0438\u0442\u044c (\u043f\u043e\u0434\u0440\u043e\u0431\u043d\u043e\u0441\u0442\u0438 \u0432 \u0436\u0443\u0440\u043d\u0430\u043b\u0435). \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a \u0435\u0449\u0451 \u0440\u0430\u0437, \u043e\u043d \u0432\u043e\u0441\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442 \u043d\u0435\u0434\u043e\u0441\u0442\u0430\u044e\u0449\u0435\u0435.",
  "summary_installed": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u00abBalanced Items\u00bb {0} \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d.",
  "summary_updated": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u00abBalanced Items\u00bb \u043e\u0431\u043d\u043e\u0432\u043b\u0451\u043d: \u0441 {0} \u0434\u043e {1}.",
  "summary_reinstalled": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u00abBalanced Items\u00bb {0} \u043f\u0435\u0440\u0435\u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d: \u0444\u0430\u0439\u043b\u044b \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u044b \u0438 \u0432\u043e\u0441\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u044b.",
  "summary_next": "\u0417\u0430\u043f\u0443\u0441\u043a\u0430\u0439\u0442\u0435 \u0438\u0433\u0440\u0443 \u0447\u0435\u0440\u0435\u0437 Steam \u043a\u0430\u043a \u043e\u0431\u044b\u0447\u043d\u043e.",
  "summary_same_version": "\u0412\u0410\u0416\u041d\u041e: \u0443 \u0432\u0441\u0435\u0445 \u0438\u0433\u0440\u043e\u043a\u043e\u0432 \u0432 \u043b\u043e\u0431\u0431\u0438 \u0434\u043e\u043b\u0436\u043d\u0430 \u0431\u044b\u0442\u044c \u043e\u0434\u043d\u0430 \u0438 \u0442\u0430 \u0436\u0435 \u0432\u0435\u0440\u0441\u0438\u044f \u043c\u043e\u0434\u0430 (\u0441\u0435\u0439\u0447\u0430\u0441 {0}). \u0415\u0441\u043b\u0438 \u0438\u0433\u0440\u0430 \u043f\u0438\u0448\u0435\u0442 \u043f\u0440\u043e \u043d\u0435\u0432\u0435\u0440\u043d\u0443\u044e \u0432\u0435\u0440\u0441\u0438\u044e, \u043f\u0443\u0441\u0442\u044c \u044d\u0442\u043e\u0442 \u0438\u0433\u0440\u043e\u043a \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442 \u0442\u0443 \u0436\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438.",
  "summary_hints": "\u0415\u0441\u043b\u0438 \u043d\u0443\u0436\u043d\u043e \u043f\u043e\u0438\u0433\u0440\u0430\u0442\u044c \u0431\u0435\u0437 \u043c\u043e\u0434\u0430 (\u043d\u0430\u043f\u0440\u0438\u043c\u0435\u0440, \u0432 \u0447\u0443\u0436\u043e\u043c \u043b\u043e\u0431\u0431\u0438):\n  {0}\n\u0412\u0435\u0440\u043d\u0443\u0442\u044c \u043c\u043e\u0434:\n  {1}\n\u0423\u0434\u0430\u043b\u0438\u0442\u044c \u043c\u043e\u0434:\n  {2}",
  "log_location": "\u041f\u043e\u0434\u0440\u043e\u0431\u043d\u044b\u0439 \u0436\u0443\u0440\u043d\u0430\u043b: {0}",
  "dry_download": "      \u0421\u043a\u0430\u0447\u0430\u043b \u0431\u044b {0} ({1} \u041a\u0411) \u0438 \u043f\u0440\u043e\u0432\u0435\u0440\u0438\u043b \u0431\u044b SHA-256.",
  "dry_extract": "      \u0420\u0430\u0441\u043f\u0430\u043a\u043e\u0432\u0430\u043b \u0431\u044b \u0432\u043e \u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0443\u044e \u043f\u0430\u043f\u043a\u0443, \u043f\u0440\u043e\u0432\u0435\u0440\u0438\u043b \u0441\u043e\u0434\u0435\u0440\u0436\u0438\u043c\u043e\u0435 \u0438 \u0441\u043a\u043e\u043f\u0438\u0440\u043e\u0432\u0430\u043b \u0431\u044b \u0432 {0} (\u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044e\u0449\u0438\u0435 \u0444\u0430\u0439\u043b\u044b .cfg \u043d\u0435 \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0438\u0441\u044b\u0432\u0430\u044e\u0442\u0441\u044f).",
  "dry_receipt": "      \u0417\u0430\u043f\u0438\u0441\u0430\u043b \u0431\u044b \u0444\u0430\u0439\u043b \u0443\u0447\u0451\u0442\u0430 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438 (BepInEx, sbgbi-install.json).",
  "dry_conflicts": "      \u041e\u0442\u043a\u043b\u044e\u0447\u0438\u043b \u0431\u044b \u043a\u043e\u043d\u0444\u043b\u0438\u043a\u0442\u0443\u044e\u0449\u0438\u0435 \u043c\u043e\u0434\u044b (\u043f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0432 .disabled): {0}",
  "dry_done": "\u041f\u0440\u043e\u0431\u043d\u044b\u0439 \u0437\u0430\u043f\u0443\u0441\u043a \u0437\u0430\u0432\u0435\u0440\u0448\u0451\u043d: \u043d\u0438\u0447\u0435\u0433\u043e \u043d\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u043e.",
  "mode_uninstall": "\u0423\u0434\u0430\u043b\u044f\u044e \u043c\u043e\u0434 \u00abBalanced Items\u00bb...",
  "mode_disable": "\u041e\u0442\u043a\u043b\u044e\u0447\u0430\u044e \u043c\u043e\u0434 (\u0440\u0435\u0436\u0438\u043c \u00ab\u0432\u0430\u043d\u0438\u043b\u044c\u00bb)...",
  "mode_enable": "\u0412\u043a\u043b\u044e\u0447\u0430\u044e \u043c\u043e\u0434...",
  "uninstall_nothing": "\u041c\u043e\u0434 \u00abBalanced Items\u00bb \u0432 \u044d\u0442\u043e\u0439 \u043f\u0430\u043f\u043a\u0435 \u0438\u0433\u0440\u044b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d: \u0443\u0434\u0430\u043b\u044f\u0442\u044c \u043d\u0435\u0447\u0435\u0433\u043e.",
  "uninstall_plan_item": "      \u0423\u0434\u0430\u043b\u0438\u043b \u0431\u044b: {0}",
  "uninstall_plan_more": "      ...\u0438 \u0435\u0449\u0451 \u0444\u0430\u0439\u043b\u043e\u0432: {0}",
  "uninstall_done": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u00abBalanced Items\u00bb \u0443\u0434\u0430\u043b\u0451\u043d (\u0444\u0430\u0439\u043b\u043e\u0432 \u0443\u0434\u0430\u043b\u0435\u043d\u043e: {0}).",
  "uninstall_bepinex_left": "BepInEx \u043e\u0441\u0442\u0430\u0432\u043b\u0435\u043d \u0432 \u043f\u0430\u043f\u043a\u0435 \u0438\u0433\u0440\u044b (\u0438\u043c \u043c\u043e\u0433\u0443\u0442 \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u044c\u0441\u044f \u0434\u0440\u0443\u0433\u0438\u0435 \u043c\u043e\u0434\u044b).\n\u0423\u0431\u0440\u0430\u0442\u044c \u0438 \u0435\u0433\u043e:\n  {0}\n\u041f\u0440\u043e\u0441\u0442\u043e \u0438\u0433\u0440\u0430\u0442\u044c \u0431\u0435\u0437 \u043c\u043e\u0434\u043e\u0432:\n  {1}",
  "uninstall_full_done": "BepInEx \u0442\u043e\u0436\u0435 \u0443\u0434\u0430\u043b\u0451\u043d: \u043f\u0430\u043f\u043a\u0430 \u0438\u0433\u0440\u044b \u043e\u0447\u0438\u0449\u0435\u043d\u0430 \u043e\u0442 \u0432\u0441\u0435\u0433\u043e, \u0447\u0442\u043e \u0441\u0442\u0430\u0432\u0438\u043b \u044d\u0442\u043e\u0442 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a.",
  "uninstall_full_skipped_foreign": "BepInEx \u043d\u0435 \u0443\u0434\u0430\u043b\u0451\u043d: \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d \u043d\u0435 \u044d\u0442\u0438\u043c \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a\u043e\u043c (\u0438\u043b\u0438 \u043d\u0435\u0442 \u0444\u0430\u0439\u043b\u0430 \u0443\u0447\u0451\u0442\u0430 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438).",
  "uninstall_full_skipped_plugins": "BepInEx \u043d\u0435 \u0443\u0434\u0430\u043b\u0451\u043d: \u0432 \u043d\u0451\u043c \u043e\u0441\u0442\u0430\u043b\u0438\u0441\u044c \u0447\u0443\u0436\u0438\u0435 \u043c\u043e\u0434\u044b, \u0438 \u0431\u0435\u0437 \u043d\u0435\u0433\u043e \u043e\u043d\u0438 \u043f\u0435\u0440\u0435\u0441\u0442\u0430\u043d\u0443\u0442 \u0440\u0430\u0431\u043e\u0442\u0430\u0442\u044c.",
  "uninstall_dir_left": "\u0412 \u043f\u0430\u043f\u043a\u0435 BepInEx \u043e\u0441\u0442\u0430\u043b\u0438\u0441\u044c \u0447\u0443\u0436\u0438\u0435 \u0444\u0430\u0439\u043b\u044b, \u043f\u043e\u044d\u0442\u043e\u043c\u0443 \u0441\u0430\u043c\u0430 \u043f\u0430\u043f\u043a\u0430 \u043d\u0435 \u0443\u0434\u0430\u043b\u0435\u043d\u0430.",
  "uninstall_full_needs_uninstall": "\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 -Full \u0440\u0430\u0431\u043e\u0442\u0430\u0435\u0442 \u0442\u043e\u043b\u044c\u043a\u043e \u0432\u043c\u0435\u0441\u0442\u0435 \u0441 -Uninstall.",
  "conflict_restored": "      \u0412\u043a\u043b\u044e\u0447\u0451\u043d \u043e\u0431\u0440\u0430\u0442\u043d\u043e: {0}",
  "toggle_disabled": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u043e\u0442\u043a\u043b\u044e\u0447\u0451\u043d: \u0438\u0433\u0440\u0430 \u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0441\u044f \u0431\u0435\u0437 \u043c\u043e\u0434\u043e\u0432 (\u0440\u0435\u0436\u0438\u043c \u00ab\u0432\u0430\u043d\u0438\u043b\u044c\u00bb). \u0412\u0435\u0440\u043d\u0443\u0442\u044c \u043c\u043e\u0434: {0}",
  "toggle_already_disabled": "\u041c\u043e\u0434 \u0443\u0436\u0435 \u043e\u0442\u043a\u043b\u044e\u0447\u0451\u043d. \u0412\u0435\u0440\u043d\u0443\u0442\u044c \u043c\u043e\u0434: {0}",
  "toggle_enabled": "\u0413\u043e\u0442\u043e\u0432\u043e! \u041c\u043e\u0434 \u0432\u043a\u043b\u044e\u0447\u0451\u043d. \u041d\u0435 \u0437\u0430\u0431\u0443\u0434\u044c\u0442\u0435: \u0443 \u0432\u0441\u0435\u0445 \u0438\u0433\u0440\u043e\u043a\u043e\u0432 \u0432 \u043b\u043e\u0431\u0431\u0438 \u0434\u043e\u043b\u0436\u043d\u0430 \u0431\u044b\u0442\u044c \u043e\u0434\u043d\u0430 \u0438 \u0442\u0430 \u0436\u0435 \u0432\u0435\u0440\u0441\u0438\u044f \u043c\u043e\u0434\u0430.",
  "toggle_already_enabled": "\u041c\u043e\u0434 \u0443\u0436\u0435 \u0432\u043a\u043b\u044e\u0447\u0451\u043d.",
  "toggle_not_installed": "\u041c\u043e\u0434 \u043d\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d: \u0444\u0430\u0439\u043b winhttp.dll \u0432 \u043f\u0430\u043f\u043a\u0435 \u0438\u0433\u0440\u044b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u0435 \u043c\u043e\u0434 \u043a\u043e\u043c\u0430\u043d\u0434\u043e\u0439 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438.",
  "toggle_dry": "      \u041f\u0435\u0440\u0435\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043b \u0431\u044b {0} \u0432 {1}.",
  "usage_exclusive": "\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b -Uninstall, -Disable \u0438 -Enable \u043d\u0435\u043b\u044c\u0437\u044f \u0438\u0441\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u044c \u0432\u043c\u0435\u0441\u0442\u0435. \u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043e\u0434\u0438\u043d.",
  "no_base_url": "\u0412 \u044d\u0442\u043e\u043c \u0441\u043a\u0440\u0438\u043f\u0442\u0435 \u043d\u0435 \u0437\u0430\u0434\u0430\u043d \u0430\u0434\u0440\u0435\u0441 \u0441\u0435\u0440\u0432\u0435\u0440\u0430. \u0418\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u0439\u0442\u0435 \u0444\u0430\u0439\u043b install.ps1, \u0441\u043e\u0431\u0440\u0430\u043d\u043d\u044b\u0439 \u0447\u0435\u0440\u0435\u0437 pack.ps1, \u0438\u043b\u0438 \u0434\u043e\u0431\u0430\u0432\u044c\u0442\u0435 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 -BaseUrl.",
  "url_not_https": "\u0410\u0434\u0440\u0435\u0441 {0} \u043d\u0435 \u043f\u043e\u0434\u0445\u043e\u0434\u0438\u0442: \u0440\u0430\u0437\u0440\u0435\u0448\u0435\u043d\u044b \u0442\u043e\u043b\u044c\u043a\u043e \u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0435 \u0430\u0434\u0440\u0435\u0441\u0430 https.",
  "net_failed": "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043a\u0430\u0447\u0430\u0442\u044c {0}: {1}",
  "net_dns": "\u043d\u0435 \u0443\u0434\u0430\u0451\u0442\u0441\u044f \u043d\u0430\u0439\u0442\u0438 \u0441\u0435\u0440\u0432\u0435\u0440 (\u043f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0438\u043d\u0442\u0435\u0440\u043d\u0435\u0442, VPN \u0438 \u0430\u0434\u0440\u0435\u0441)",
  "net_connect": "\u043d\u0435\u0442 \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f \u0441 \u0441\u0435\u0440\u0432\u0435\u0440\u043e\u043c (\u043f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0438\u043d\u0442\u0435\u0440\u043d\u0435\u0442, VPN, \u0431\u0440\u0430\u043d\u0434\u043c\u0430\u0443\u044d\u0440)",
  "net_timeout": "\u0441\u0435\u0440\u0432\u0435\u0440 \u043d\u0435 \u043e\u0442\u0432\u0435\u0442\u0438\u043b \u0432\u043e\u0432\u0440\u0435\u043c\u044f",
  "net_tls": "\u043e\u0448\u0438\u0431\u043a\u0430 \u0437\u0430\u0449\u0438\u0449\u0451\u043d\u043d\u043e\u0433\u043e \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f (\u043f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0434\u0430\u0442\u0443 \u0438 \u0432\u0440\u0435\u043c\u044f \u043d\u0430 \u043a\u043e\u043c\u043f\u044c\u044e\u0442\u0435\u0440\u0435, \u0430\u043d\u0442\u0438\u0432\u0438\u0440\u0443\u0441 \u0438 VPN)",
  "net_http": "\u0441\u0435\u0440\u0432\u0435\u0440 \u043e\u0442\u0432\u0435\u0442\u0438\u043b \u043e\u0448\u0438\u0431\u043a\u043e\u0439 {0}",
  "net_insecure_redirect": "\u0441\u0435\u0440\u0432\u0435\u0440 \u043f\u0435\u0440\u0435\u043d\u0430\u043f\u0440\u0430\u0432\u0438\u043b \u0437\u0430\u043f\u0440\u043e\u0441 \u043d\u0430 \u043d\u0435\u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u044b\u0439 \u0430\u0434\u0440\u0435\u0441, \u044d\u0442\u043e \u0437\u0430\u043f\u0440\u0435\u0449\u0435\u043d\u043e",
  "net_too_big": "\u0444\u0430\u0439\u043b \u043e\u043a\u0430\u0437\u0430\u043b\u0441\u044f \u0441\u043b\u0438\u0448\u043a\u043e\u043c \u0431\u043e\u043b\u044c\u0448\u0438\u043c",
  "latest_bad": "\u0424\u0430\u0439\u043b latest.json \u043d\u0430 \u0441\u0435\u0440\u0432\u0435\u0440\u0435 \u043f\u043e\u0432\u0440\u0435\u0436\u0434\u0451\u043d \u0438\u043b\u0438 \u0438\u043c\u0435\u0435\u0442 \u043d\u0435\u043e\u0436\u0438\u0434\u0430\u043d\u043d\u044b\u0439 \u0432\u0438\u0434 ({0}). \u0421\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443 \u043c\u043e\u0434\u0430.",
  "installer_outdated": "\u042d\u0442\u043e\u0442 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a \u0441\u043b\u0438\u0448\u043a\u043e\u043c \u0441\u0442\u0430\u0440\u044b\u0439 \u0434\u043b\u044f \u0442\u0435\u043a\u0443\u0449\u0435\u0439 \u0432\u0435\u0440\u0441\u0438\u0438 \u043c\u043e\u0434\u0430 (\u043f\u0440\u043e\u0442\u043e\u043a\u043e\u043b {0}). \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438 \u0437\u0430\u043d\u043e\u0432\u043e; \u0435\u0441\u043b\u0438 \u043d\u0435 \u043f\u043e\u043c\u043e\u0436\u0435\u0442, \u0441\u043e\u043e\u0431\u0449\u0438\u0442\u0435 \u0430\u0432\u0442\u043e\u0440\u0443.",
  "unexpected": "\u041d\u0435\u043f\u0440\u0435\u0434\u0432\u0438\u0434\u0435\u043d\u043d\u0430\u044f \u043e\u0448\u0438\u0431\u043a\u0430: {0}"
}
'@
    $M = $MessagesJson | ConvertFrom-Json

    # ======================================================================================== output
    function Write-Log {
        param([string]$Level, [string]$Text)
        try {
            $line = '{0} [{1}] {2}{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Text, [Environment]::NewLine
            [IO.File]::AppendAllText($State.LogPath, $line, $Utf8Bom)
        } catch { $null = $_ }
    }
    function Write-Line {
        param([string]$Text, [string]$Color, [string]$Level, [bool]$Always)
        Write-Log $Level $Text
        if ($Always -or -not $Quiet) {
            if ($Color) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
        }
    }
    function Write-Info   { param([string]$Text) Write-Line $Text '' 'INFO' $false }
    function Write-Step   { param([string]$Text) Write-Line $Text 'Cyan' 'STEP' $false }
    function Write-Ok     { param([string]$Text) Write-Line $Text 'Green' 'OK' $false }
    function Write-Warn   { param([string]$Text) Write-Line $Text 'Yellow' 'WARN' $true }
    function Write-Note   { param([string]$Text) Write-Line $Text 'Yellow' 'NOTE' $false }
    function Write-Err    { param([string]$Text) Write-Line $Text 'Red' 'ERROR' $true }
    function Write-Result { param([string]$Text) Write-Line $Text 'Green' 'RESULT' $true }

    function Stop-Installer {
        param([int]$Code, [string]$Text)
        $ex = New-Object System.Exception($Text)
        $ex.Data['SbgbiExit'] = $Code
        throw $ex
    }

    function Initialize-Log {
        try {
            $p = $State.LogPath
            if ((Test-Path -LiteralPath $p -PathType Leaf) -and ((Get-Item -LiteralPath $p).Length -gt 1MB)) {
                Move-Item -LiteralPath $p -Destination ($p + '.old') -Force
            }
        } catch { $null = $_ }
        Write-Log 'INFO' ('==== installer start; PowerShell {0}; OS {1}; user {2}; args: Uninstall={3} Full={4} Disable={5} Enable={6} DisableConflicts={7} DryRun={8} Quiet={9}' -f
            $PSVersionTable.PSVersion, [Environment]::OSVersion.VersionString, [Environment]::UserName,
            $Uninstall.IsPresent, $Full.IsPresent, $Disable.IsPresent, $Enable.IsPresent, $DisableConflicts.IsPresent, $DryRun.IsPresent, $Quiet.IsPresent)
    }

    # ======================================================================================== small helpers
    function Get-Prop {
        param($Obj, [string]$Name, $Default)
        $p = $Obj.PSObject.Properties[$Name]
        if ($null -eq $p) { return $Default }
        return $p.Value
    }
    function Format-Kb { param([long]$Bytes) return [long][math]::Ceiling([double]$Bytes / 1024) }
    function ConvertTo-RelPath { param([string]$Full, [string]$Root) return $Full.Substring($Root.TrimEnd('\').Length + 1).Replace('\', '/') }
    function ConvertTo-LocalPath { param([string]$Root, [string]$Rel) return (Join-Path $Root ($Rel.Replace('/', '\'))) }
    function Get-CommandHint {
        param([string]$Switches)
        return ('& ([scriptblock]::Create((irm {0}/install.ps1))) {1}' -f $BaseUrl, $Switches)
    }
    function Clear-ReadOnly {
        param([string]$Path)
        if ([IO.File]::Exists($Path)) { [IO.File]::SetAttributes($Path, [IO.FileAttributes]::Normal) }
    }

    function ConvertTo-JsonString {
        param([string]$Text)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('"')
        foreach ($ch in $Text.ToCharArray()) {
            $c = [int]$ch
            if ($c -eq 34) { [void]$sb.Append('\"') }
            elseif ($c -eq 92) { [void]$sb.Append('\\') }
            elseif ($c -eq 10) { [void]$sb.Append('\n') }
            elseif ($c -eq 13) { [void]$sb.Append('\r') }
            elseif ($c -eq 9) { [void]$sb.Append('\t') }
            elseif ($c -lt 32 -or $c -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $c)) }
            else { [void]$sb.Append($ch) }
        }
        [void]$sb.Append('"')
        return $sb.ToString()
    }
    function ConvertTo-JsonText {
        param($Value, [int]$Level = 0)
        $pad = '  ' * $Level
        $pad1 = '  ' * ($Level + 1)
        if ($null -eq $Value) { return 'null' }
        if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
        if ($Value -is [string]) { return (ConvertTo-JsonString $Value) }
        if ($Value -is [int] -or $Value -is [long]) { return $Value.ToString([Globalization.CultureInfo]::InvariantCulture) }
        if ($Value -is [System.Collections.IDictionary]) {
            if ($Value.Count -eq 0) { return '{}' }
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($k in $Value.Keys) {
                $parts.Add($pad1 + (ConvertTo-JsonString ([string]$k)) + ': ' + (ConvertTo-JsonText $Value[$k] ($Level + 1)))
            }
            return "{`n" + ($parts -join ",`n") + "`n" + $pad + '}'
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($v in $Value) { $parts.Add($pad1 + (ConvertTo-JsonText $v ($Level + 1))) }
            if ($parts.Count -eq 0) { return '[]' }
            return "[`n" + ($parts -join ",`n") + "`n" + $pad + ']'
        }
        return (ConvertTo-JsonString ([string]$Value))
    }

    # ======================================================================================== package rules
    # Mirrors pack.ps1. Returns a normalized forward-slash relative path, or $null when the path is unsafe.
    function Get-SafeRelPath {
        param([string]$Raw)
        if ([string]::IsNullOrEmpty($Raw)) { return $null }
        $Raw = $Raw.Replace('\', '/')
        if ($Raw.StartsWith('/') -or $Raw.Contains(':')) { return $null }
        $invalid = [IO.Path]::GetInvalidFileNameChars()
        $segs = $Raw.Split('/')
        foreach ($s in $segs) {
            if ($s -eq '' -or $s -eq '.' -or $s -eq '..') { return $null }
            if ($s.EndsWith('.') -or $s.EndsWith(' ')) { return $null }
            if ($s.IndexOfAny($invalid) -ge 0) { return $null }
            if ($s -match '(?i)^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$') { return $null }
        }
        return ($segs -join '/')
    }
    function Test-AllowedPackagePath {
        param([string]$Rel)
        if ($RootFiles -contains $Rel) { return $true }
        if ($Rel -match '^BepInEx/core/.+') { return $true }
        if ($Rel -match '^BepInEx/config/[^/]+\.cfg$') { return $true }
        if ($Rel -match '^BepInEx/plugins/SbgBalancedItems/.+') { return $true }
        return $false
    }

    # ======================================================================================== Steam / game
    # Paths coming from the registry, libraryfolders.vdf (which may still list an unplugged drive) or -GamePath are
    # untrusted: Join-Path / Test-Path / Split-Path throw on a missing drive, so only [IO.Path] / [IO.File] /
    # [IO.Directory] (which never throw for "does not exist") are used for them.
    function ConvertFrom-VdfString { param([string]$Text) return [regex]::Replace($Text, '\\(.)', '$1') }
    function Join-Safe {
        param([string]$Base, [string]$Child)
        try { return [IO.Path]::Combine($Base, $Child) } catch { return '' }
    }

    function Get-SteamRoots {
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $roots = New-Object System.Collections.Generic.List[string]
        $candidates = New-Object System.Collections.Generic.List[string]
        if ($SteamPathOverride) {
            $candidates.Add($SteamPathOverride)
        } else {
            $regKeys = @(
                @('HKCU:\Software\Valve\Steam', 'SteamPath'),
                @('HKLM:\SOFTWARE\WOW6432Node\Valve\Steam', 'InstallPath'),
                @('HKLM:\SOFTWARE\Valve\Steam', 'InstallPath')
            )
            foreach ($k in $regKeys) {
                try { $v = Get-ItemPropertyValue -Path $k[0] -Name $k[1] -ErrorAction Stop } catch { continue }
                if ($v) { $candidates.Add([string]$v) }
            }
            foreach ($pf in @(${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
                if ($pf) { $candidates.Add((Join-Safe $pf 'Steam')) }
            }
        }
        foreach ($c in $candidates) {
            if (-not $c) { continue }
            try { $full = [IO.Path]::GetFullPath($c.Trim().Trim('"')) } catch { continue }
            if ($seen.Add($full)) { $roots.Add($full) }
        }
        return $roots.ToArray()
    }

    function Get-SteamLibraries {
        param([string]$SteamRoot)
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $libs = New-Object System.Collections.Generic.List[string]
        [void]$seen.Add($SteamRoot)
        $libs.Add($SteamRoot)
        foreach ($vdf in @((Join-Safe $SteamRoot 'steamapps\libraryfolders.vdf'), (Join-Safe $SteamRoot 'config\libraryfolders.vdf'))) {
            if (-not $vdf -or -not [IO.File]::Exists($vdf)) { continue }
            try { $text = [IO.File]::ReadAllText($vdf, [Text.Encoding]::UTF8) } catch { Write-Log 'WARN' "cannot read ${vdf}: $($_.Exception.Message)"; continue }
            $found = New-Object System.Collections.Generic.List[string]
            # new format:  "path"  "D:\\SteamLibrary"
            foreach ($m in [regex]::Matches($text, '(?im)^\s*"path"\s+"((?:[^"\\]|\\.)*)"')) { $found.Add((ConvertFrom-VdfString $m.Groups[1].Value)) }
            # old format:  "1"  "D:\\SteamLibrary"  (only values that look like absolute paths)
            foreach ($m in [regex]::Matches($text, '(?im)^\s*"\d+"\s+"((?:[^"\\]|\\.)*)"')) {
                $v = ConvertFrom-VdfString $m.Groups[1].Value
                if ($v -match '^[A-Za-z]:[\\/]' -or $v -match '^\\\\') { $found.Add($v) }
            }
            foreach ($p in $found) {
                try { $full = [IO.Path]::GetFullPath($p) } catch { continue }
                if ($seen.Add($full)) { $libs.Add($full) }
            }
        }
        return $libs.ToArray()
    }

    function Read-AcfValue {
        param([string]$Text, [string]$Key)
        $m = [regex]::Match($Text, '(?im)^\s*"' + [regex]::Escape($Key) + '"\s+"((?:[^"\\]|\\.)*)"')
        if ($m.Success) { return (ConvertFrom-VdfString $m.Groups[1].Value) }
        return ''
    }

    function Find-GameInstall {
        foreach ($root in @(Get-SteamRoots)) {
            Write-Log 'INFO' "steam root: $root"
            foreach ($lib in @(Get-SteamLibraries $root)) {
                try {
                    $steamapps = Join-Safe $lib 'steamapps'
                    $acf = Join-Safe $steamapps "appmanifest_$AppId.acf"
                    $installDir = $DefaultFolderName
                    $build = ''
                    if ($acf -and [IO.File]::Exists($acf)) {
                        try {
                            $text = [IO.File]::ReadAllText($acf, [Text.Encoding]::UTF8)
                            $d = Read-AcfValue $text 'installdir'
                            if ($d -and $d -notmatch '[\\/:]' -and $d -ne '..') { $installDir = $d }
                            $build = Read-AcfValue $text 'buildid'
                        } catch { Write-Log 'WARN' "cannot parse ${acf}: $($_.Exception.Message)" }
                    }
                    $dir = Join-Safe (Join-Safe $steamapps 'common') $installDir
                    Write-Log 'INFO' "library: $lib ; candidate: $dir"
                    $exe = Join-Safe $dir $GameExe
                    if ($exe -and [IO.File]::Exists($exe)) {
                        return [pscustomobject]@{ Dir = $dir; BuildId = $build; Library = $lib }
                    }
                } catch { Write-Log 'WARN' "skipping library ${lib}: $($_.Exception.Message)" }
            }
        }
        return $null
    }

    function Get-BuildIdForDir {
        param([string]$Dir)
        try {
            $common = [IO.Path]::GetDirectoryName($Dir)
            if ([IO.Path]::GetFileName($common) -ne 'common') { return '' }
            $acf = Join-Safe ([IO.Path]::GetDirectoryName($common)) "appmanifest_$AppId.acf"
            if (-not $acf -or -not [IO.File]::Exists($acf)) { return '' }
            return (Read-AcfValue ([IO.File]::ReadAllText($acf, [Text.Encoding]::UTF8)) 'buildid')
        } catch { return '' }
    }

    function Resolve-GamePathArg {
        param([string]$Given)
        $p = $Given.Trim().Trim('"').Trim("'").TrimEnd('\', '/')
        $rooted = $false
        try { $rooted = [IO.Path]::IsPathRooted($p) } catch { $rooted = $false }
        if ($p -and -not $rooted) {
            # relative to the PowerShell location (the .NET current directory is not kept in sync with it)
            try { $p = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) } catch { $null = $_ }
        }
        if ($p -like '*.exe' -and [IO.File]::Exists($p)) { $p = [IO.Path]::GetDirectoryName($p) }
        if ($p) {
            $candidates = @($p,
                (Join-Safe $p $DefaultFolderName),
                (Join-Safe $p ('common\' + $DefaultFolderName)),
                (Join-Safe $p ('steamapps\common\' + $DefaultFolderName)))
            foreach ($c in $candidates) {
                if (-not $c) { continue }
                $exe = Join-Safe $c $GameExe
                if ($exe -and [IO.File]::Exists($exe)) {
                    try { return [IO.Path]::GetFullPath($c) } catch { continue }
                }
            }
        }
        Stop-Installer $Exit.GameNotFound ($M.game_path_bad -f $Given)
    }
    function Resolve-Game {
        if ($GamePath) {
            $dir = Resolve-GamePathArg $GamePath
            Write-Info ($M.game_from_arg -f $dir)
            return [pscustomobject]@{ Dir = $dir; BuildId = (Get-BuildIdForDir $dir); Library = '' }
        }
        $found = Find-GameInstall
        if ($null -eq $found) {
            Stop-Installer $Exit.GameNotFound ($M.game_not_found -f ('& ([scriptblock]::Create((irm ' + $BaseUrl + '/install.ps1))) -GamePath "D:\SteamLibrary\steamapps\common\Super Battle Golf"'))
        }
        Write-Info ($M.game_found -f $found.Dir)
        return $found
    }

    function Test-GameRunning {
        $p = @(Get-Process -Name $GameProcessName -ErrorAction SilentlyContinue)
        return ($p.Count -gt 0)
    }
    function Assert-GameClosed {
        if (Test-GameRunning) {
            if ($DryRun) { Write-Warn $M.game_running_dry; return }
            Stop-Installer $Exit.GameRunning $M.game_running
        }
    }

    # ======================================================================================== network
    function Test-UrlAllowed {
        param([string]$Url)
        $u = $null
        if (-not [Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$u)) { return $false }
        if ($u.Scheme -eq 'https') { return $true }
        if ($u.Scheme -eq 'http' -and $AllowInsecureHttp) { return $true }
        return $false
    }
    function Get-WebErrorText {
        param([System.Exception]$Ex)
        $we = $Ex
        while ($null -ne $we -and -not ($we -is [System.Net.WebException])) { $we = $we.InnerException }
        if ($null -ne $we) {
            $status = [string]$we.Status
            if ($status -eq 'NameResolutionFailure') { return $M.net_dns }
            if ($status -eq 'ConnectFailure' -or $status -eq 'ConnectionClosed' -or $status -eq 'ReceiveFailure' -or $status -eq 'SendFailure') { return $M.net_connect }
            if ($status -eq 'Timeout') { return $M.net_timeout }
            if ($status -eq 'TrustFailure' -or $status -eq 'SecureChannelFailure') { return $M.net_tls }
            if ($status -eq 'ProtocolError' -and $null -ne $we.Response) { return ($M.net_http -f [int]$we.Response.StatusCode) }
        }
        return $Ex.Message
    }
    # Downloads $Url into $OutFile (or returns the bytes when no file is given). HTTPS only (http only with the test switch).
    function Get-RemoteData {
        param([string]$Url, [string]$OutFile, [int]$TimeoutSec = 30, [long]$MaxBytes = 209715200)
        if (-not (Test-UrlAllowed $Url)) { Stop-Installer $Exit.Network ($M.url_not_https -f $Url) }
        Write-Log 'INFO' "GET $Url"
        try {
            $req = [System.Net.HttpWebRequest][System.Net.WebRequest]::Create($Url)
            $req.UserAgent = $UserAgent
            $req.Timeout = $TimeoutSec * 1000
            $req.ReadWriteTimeout = $TimeoutSec * 1000
            $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
            $req.CachePolicy = New-Object System.Net.Cache.HttpRequestCachePolicy([System.Net.Cache.HttpRequestCacheLevel]::NoCacheNoStore)
            $req.Headers['Cache-Control'] = 'no-cache'
            $resp = $req.GetResponse()
            try {
                if ($resp.ResponseUri.Scheme -ne 'https' -and -not $AllowInsecureHttp) { throw (New-Object System.InvalidOperationException('insecure-redirect')) }
                $in = $resp.GetResponseStream()
                if ($OutFile) { $out = [IO.File]::Create($OutFile) } else { $out = New-Object System.IO.MemoryStream }
                try {
                    $buf = New-Object byte[] 81920
                    $total = 0L
                    while (($n = $in.Read($buf, 0, $buf.Length)) -gt 0) {
                        $total += $n
                        if ($total -gt $MaxBytes) { throw (New-Object System.InvalidOperationException('too-big')) }
                        $out.Write($buf, 0, $n)
                    }
                    if (-not $OutFile) { return ,$out.ToArray() }
                } finally { $out.Dispose(); $in.Dispose() }
            } finally { $resp.Close() }
        } catch {
            $ex = $_.Exception
            if ($ex.Data.Contains('SbgbiExit')) { throw }
            Write-Log 'ERROR' ($_ | Out-String)
            if ($ex.Message -eq 'insecure-redirect') { $why = $M.net_insecure_redirect }
            elseif ($ex.Message -eq 'too-big') { $why = $M.net_too_big }
            else { $why = Get-WebErrorText $ex }
            Stop-Installer $Exit.Network ($M.net_failed -f $Url, $why)
        }
    }

    function Get-LatestInfo {
        $url = '{0}/latest.json?t={1}' -f $BaseUrl, [DateTime]::UtcNow.Ticks
        $bytes = Get-RemoteData -Url $url -TimeoutSec 30 -MaxBytes 1048576
        $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        try { $o = $text | ConvertFrom-Json } catch { Stop-Installer $Exit.Network ($M.latest_bad -f 'json') }
        if ($null -eq $o) { Stop-Installer $Exit.Network ($M.latest_bad -f 'empty') }
        $name = [string](Get-Prop $o 'name' '')
        $version = [string](Get-Prop $o 'version' '')
        $protocol = 0
        try { $protocol = [int](Get-Prop $o 'protocol' 0) } catch { $protocol = 0 }
        $zipUrl = [string](Get-Prop $o 'zipUrl' '')
        $sha = [string](Get-Prop $o 'sha256' '')
        $size = 0L
        try { $size = [long](Get-Prop $o 'sizeBytes' 0) } catch { $size = 0L }
        if ($name -ne $PluginName) { Stop-Installer $Exit.Network ($M.latest_bad -f 'name') }
        if ($version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+([-+][0-9A-Za-z.+-]{1,30})?$') { Stop-Installer $Exit.Network ($M.latest_bad -f 'version') }
        if ($protocol -lt 1) { Stop-Installer $Exit.Network ($M.latest_bad -f 'protocol') }
        if ($protocol -gt $SupportedProtocol) { Stop-Installer $Exit.Network ($M.installer_outdated -f $protocol) }
        if ($sha -notmatch '^[0-9a-fA-F]{64}$') { Stop-Installer $Exit.Network ($M.latest_bad -f 'sha256') }
        if ([string]::IsNullOrWhiteSpace($zipUrl)) { Stop-Installer $Exit.Network ($M.latest_bad -f 'zipUrl') }
        if (-not [Uri]::IsWellFormedUriString($zipUrl, [UriKind]::Absolute)) { $zipUrl = $BaseUrl + '/' + $zipUrl.TrimStart('/') }
        if (-not (Test-UrlAllowed $zipUrl)) { Stop-Installer $Exit.Network ($M.url_not_https -f $zipUrl) }
        # the archive must come from the same server as latest.json: a tampered latest.json cannot redirect the download elsewhere
        $baseUri = [Uri]$BaseUrl
        $zipUri = [Uri]$zipUrl
        if ($zipUri.Scheme -ne $baseUri.Scheme -or $zipUri.Host -ne $baseUri.Host -or $zipUri.Port -ne $baseUri.Port) { Stop-Installer $Exit.Network ($M.latest_bad -f 'zipUrl host') }
        $notes = New-Object System.Collections.Generic.List[string]
        foreach ($n in @(Get-Prop $o 'notes' @())) {
            if ($null -eq $n) { continue }
            $line = ([string]$n) -replace '[\x00-\x1F\x7F]', ' '     # no terminal control sequences from the server
            if ($line.Trim()) { $notes.Add($line.Trim()) }
        }
        return [pscustomobject]@{
            Version      = $version
            ZipUrl       = $zipUrl
            Sha256       = $sha.ToUpperInvariant()
            SizeBytes    = $size
            GameVersion  = [string](Get-Prop $o 'gameVersionTested' '')
            GameBuild    = [string](Get-Prop $o 'gameBuildTested' '')
            Notes        = $notes.ToArray()
        }
    }

    # ======================================================================================== package
    function Expand-Package {
        param([string]$ZipPath, [string]$StageDir)
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $items = New-Object System.Collections.Generic.List[object]
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $totalBytes = 0L
        try { $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath) } catch {
            Write-Log 'ERROR' ($_ | Out-String)
            Stop-Installer $Exit.Network $M.pkg_corrupt
        }
        try {
            if ($zip.Entries.Count -gt 2000) { Stop-Installer $Exit.Network $M.pkg_too_big }
            foreach ($e in $zip.Entries) {
                $raw = $e.FullName.Replace('\', '/')
                if ($raw.EndsWith('/')) { continue }
                $rel = Get-SafeRelPath $raw
                if ($null -eq $rel -or -not (Test-AllowedPackagePath $rel) -or -not $seen.Add($rel)) { Stop-Installer $Exit.Network ($M.pkg_bad_entry -f $raw) }
                $totalBytes += $e.Length
                if ($totalBytes -gt 209715200) { Stop-Installer $Exit.Network $M.pkg_too_big }
                $dest = ConvertTo-LocalPath $StageDir $rel
                [void][IO.Directory]::CreateDirectory((Split-Path -Parent $dest))
                $in = $e.Open()
                try {
                    $fs = [IO.File]::Create($dest)
                    try { $in.CopyTo($fs) } finally { $fs.Dispose() }
                } finally { $in.Dispose() }
                $items.Add([pscustomobject]@{
                    Rel    = $rel
                    Path   = $dest
                    Size   = (New-Object IO.FileInfo($dest)).Length
                    Sha256 = (Invoke-Retry { (Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash })
                })
            }
        } catch [System.IO.InvalidDataException] {
            Write-Log 'ERROR' ($_ | Out-String)
            Stop-Installer $Exit.Network $M.pkg_corrupt
        } finally { $zip.Dispose() }
        foreach ($need in @('winhttp.dll', 'doorstop_config.ini', 'BepInEx/core/BepInEx.dll', ($PluginRel + '/' + $PluginName + '.dll'))) {
            if (-not $seen.Contains($need)) { Stop-Installer $Exit.Network ($M.pkg_incomplete -f $need) }
        }
        return $items.ToArray()
    }

    # ======================================================================================== receipt
    function Read-Receipt {
        param([string]$Dir)
        $p = ConvertTo-LocalPath $Dir $ReceiptRel
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return $null }
        try {
            $r = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json
            $files = New-Object System.Collections.Generic.List[string]
            foreach ($f in @(Get-Prop $r 'files' @())) {
                $rel = Get-SafeRelPath ([string]$f)
                if ($null -ne $rel -and (Test-AllowedPackagePath $rel)) { $files.Add($rel) } else { Write-Log 'WARN' "ignoring bad receipt path: $f" }
            }
            $dis = New-Object System.Collections.Generic.List[string]
            foreach ($f in @(Get-Prop $r 'disabledConflicts' @())) {
                $rel = Get-SafeRelPath ([string]$f)
                if ($null -ne $rel -and $rel -match '^BepInEx/plugins/.+\.dll\.disabled$') { $dis.Add($rel) } else { Write-Log 'WARN' "ignoring bad receipt conflict path: $f" }
            }
            $created = Get-Prop $r 'createdBepInEx' $null
            return [pscustomobject]@{
                Version           = [string](Get-Prop $r 'version' '')
                InstalledAt       = [string](Get-Prop $r 'installedAt' '')
                GameBuildId       = [string](Get-Prop $r 'gameBuildId' '')
                CreatedBepInEx    = $(if ($created -is [bool]) { [bool]$created } else { $null })
                Files             = $files.ToArray()
                DisabledConflicts = $dis.ToArray()
            }
        } catch {
            Write-Log 'WARN' "cannot parse receipt: $($_.Exception.Message)"
            return $null
        }
    }

    # ======================================================================================== journal (rollback)
    # Every file-system change goes through these helpers; Undo-Journal reverts them in reverse order.
    function Add-Journal {
        param($Journal, [string]$Kind, [string]$Path, [string]$Backup = '', [string]$To = '')
        [void]$Journal.Add([pscustomobject]@{ Kind = $Kind; Path = $Path; Backup = $Backup; To = $To })
    }
    function New-DirectoryJ {
        param($Journal, [string]$Dir)
        $missing = New-Object System.Collections.Generic.List[string]
        $d = $Dir
        while ($d -and -not [IO.Directory]::Exists($d)) { $missing.Add($d); $d = Split-Path -Parent $d }
        for ($i = $missing.Count - 1; $i -ge 0; $i--) {
            Add-Journal $Journal 'mkdir' $missing[$i]
            [void][IO.Directory]::CreateDirectory($missing[$i])
        }
    }
    # Antivirus scanners sometimes hold a freshly written file for a moment: retry short-lived sharing / access errors.
    function Invoke-Retry {
        param([scriptblock]$Action, [int]$Attempts = 3, [int]$DelayMs = 300)
        for ($i = 1; $i -le $Attempts; $i++) {
            try { return (& $Action) } catch {
                $ex = $_.Exception
                $transient = ($ex -is [System.IO.IOException] -or $ex -is [System.UnauthorizedAccessException]) -and
                    -not ($ex -is [System.IO.FileNotFoundException] -or $ex -is [System.IO.DirectoryNotFoundException])
                if (-not $transient -or $i -ge $Attempts) { throw }
                Write-Log 'WARN' ('retry {0}/{1} after: {2}' -f $i, $Attempts, $ex.Message)
                Start-Sleep -Milliseconds ($DelayMs * $i)
            }
        }
    }
    function Copy-FileJ {
        param($Journal, [string]$Src, [string]$Dst, [string]$BackupDir)
        New-DirectoryJ $Journal (Split-Path -Parent $Dst)
        if ([IO.File]::Exists($Dst)) {
            $bak = Join-Path $BackupDir ([guid]::NewGuid().ToString('N'))
            Invoke-Retry { [IO.File]::Copy($Dst, $bak, $true) }
            Add-Journal $Journal 'replaced' $Dst $bak
            Clear-ReadOnly $Dst
            Invoke-Retry { [IO.File]::Copy($Src, $Dst, $true) }
        } else {
            if ([IO.Directory]::Exists($Dst)) { throw "destination is a directory: $Dst" }
            Add-Journal $Journal 'created' $Dst
            Invoke-Retry { [IO.File]::Copy($Src, $Dst, $true) }
        }
    }
    function Remove-FileJ {
        param($Journal, [string]$Path, [string]$BackupDir)
        if (-not [IO.File]::Exists($Path)) { return }
        $bak = Join-Path $BackupDir ([guid]::NewGuid().ToString('N'))
        Invoke-Retry { [IO.File]::Copy($Path, $bak, $true) }
        Add-Journal $Journal 'deleted' $Path $bak
        Clear-ReadOnly $Path
        Invoke-Retry { [IO.File]::Delete($Path) }
    }
    function Move-FileJ {
        param($Journal, [string]$From, [string]$To, [string]$BackupDir)
        if ([IO.File]::Exists($To)) { Remove-FileJ $Journal $To $BackupDir }
        Add-Journal $Journal 'renamed' $From '' $To
        [IO.File]::Move($From, $To)
    }
    function Remove-DirIfEmptyJ {
        param($Journal, [string]$Dir)
        if (-not [IO.Directory]::Exists($Dir)) { return $false }
        if ([IO.Directory]::GetFileSystemEntries($Dir).Length -gt 0) { return $false }
        Add-Journal $Journal 'rmdir' $Dir
        [IO.Directory]::Delete($Dir)
        return $true
    }
    function Undo-Journal {
        param($Journal)
        $failed = 0
        for ($i = $Journal.Count - 1; $i -ge 0; $i--) {
            $e = $Journal[$i]
            try {
                switch ($e.Kind) {
                    'created'  { if ([IO.File]::Exists($e.Path)) { Clear-ReadOnly $e.Path; Invoke-Retry { [IO.File]::Delete($e.Path) } } }
                    'replaced' { Invoke-Retry { [IO.File]::Copy($e.Backup, $e.Path, $true) } }
                    'deleted'  { [void][IO.Directory]::CreateDirectory((Split-Path -Parent $e.Path)); Invoke-Retry { [IO.File]::Copy($e.Backup, $e.Path, $true) } }
                    'mkdir'    { if ([IO.Directory]::Exists($e.Path) -and [IO.Directory]::GetFileSystemEntries($e.Path).Length -eq 0) { [IO.Directory]::Delete($e.Path) } }
                    'rmdir'    { [void][IO.Directory]::CreateDirectory($e.Path) }
                    'renamed'  { if ([IO.File]::Exists($e.To)) { [IO.File]::Move($e.To, $e.Path) } }
                }
            } catch {
                $failed++
                Write-Log 'ERROR' ('rollback step failed: {0} {1}: {2}' -f $e.Kind, $e.Path, $_.Exception.Message)
            }
        }
        return $failed
    }
    function Get-FailureText {
        param($ErrorRecord)
        $ex = $ErrorRecord.Exception
        if ($ex -is [System.UnauthorizedAccessException] -or $ex.InnerException -is [System.UnauthorizedAccessException]) { return $M.access_denied }
        return ($M.install_failed -f $ex.Message)
    }

    # ======================================================================================== conflicts
    function Get-ConflictingPlugins {
        param([string]$Dir)
        $plugins = Join-Path $Dir 'BepInEx\plugins'
        $res = New-Object System.Collections.Generic.List[object]
        if (-not [IO.Directory]::Exists($plugins)) { return $res.ToArray() }
        $ours = (Join-Path $plugins $PluginName) + '\'
        foreach ($f in @(Get-ChildItem -LiteralPath $plugins -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            if ($f.Extension -ne '.dll') { continue }
            if ($f.FullName.StartsWith($ours, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $rel = $f.FullName.Substring($plugins.Length + 1)
            if ($rel -match $ConflictPattern -or $f.Name -like ($PluginName + '*')) {
                $res.Add([pscustomobject]@{ Path = $f.FullName; Rel = ('BepInEx/plugins/' + $rel.Replace('\', '/')); Display = $rel })
            }
        }
        return $res.ToArray()
    }

    # ======================================================================================== install
    function Install-Package {
        param($Game, $Latest, $Prev, $Conflicts, [string]$Work)
        $dir = $Game.Dir
        $zip = Join-Path $Work 'package.zip'
        $stage = Join-Path $Work 'stage'
        $backup = Join-Path $Work 'backup'
        [void][IO.Directory]::CreateDirectory($backup)

        Write-Step ($M.step_download -f (Format-Kb $Latest.SizeBytes))
        Write-Log 'INFO' ('package: {0} expected sha256 {1} size {2}' -f $Latest.ZipUrl, $Latest.Sha256, $Latest.SizeBytes)
        Get-RemoteData -Url $Latest.ZipUrl -OutFile $zip -TimeoutSec 60
        $actualSize = (New-Object IO.FileInfo($zip)).Length
        if ($Latest.SizeBytes -gt 0 -and $actualSize -ne $Latest.SizeBytes) { Stop-Installer $Exit.Network ($M.size_bad -f $actualSize, $Latest.SizeBytes) }
        Write-Info $M.verifying
        $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
        if ($actual -ne $Latest.Sha256) { Stop-Installer $Exit.Network ($M.hash_bad -f $Latest.Sha256, $actual) }
        Write-Info $M.hash_ok

        Write-Step $M.step_extract
        $items = @(Expand-Package -ZipPath $zip -StageDir $stage)
        Write-Log 'INFO' ('package files: {0}' -f $items.Count)

        # existing BepInEx of another major version: do not mix files
        $existingCore = Join-Path $dir 'BepInEx\core\BepInEx.dll'
        if ([IO.File]::Exists($existingCore)) {
            try {
                $fv = [Diagnostics.FileVersionInfo]::GetVersionInfo($existingCore)
                if ($fv.FileMajorPart -gt 0 -and $fv.FileMajorPart -ne 5) { Stop-Installer $Exit.Error ($M.bepinex_too_old -f $fv.FileVersion) }
            } catch {
                if ($_.Exception.Data.Contains('SbgbiExit')) { throw }
            }
        }

        # ---- plan
        $prevFiles = @{}
        $createdBep = $false
        if ($null -ne $Prev) {
            foreach ($f in $Prev.Files) { $prevFiles[$f] = $true }
            if ($null -ne $Prev.CreatedBepInEx) { $createdBep = [bool]$Prev.CreatedBepInEx }
        } else {
            $createdBep = -not ([IO.Directory]::Exists((Join-Path $dir 'BepInEx')) -or [IO.File]::Exists((Join-Path $dir 'winhttp.dll')) -or
                [IO.File]::Exists((Join-Path $dir 'winhttp.dll.off')) -or [IO.File]::Exists((Join-Path $dir 'doorstop_config.ini')))
        }
        $writes = New-Object System.Collections.Generic.List[object]
        $owned = New-Object System.Collections.Generic.List[string]
        $kept = New-Object System.Collections.Generic.List[string]
        $newSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($it in $items) {
            [void]$newSet.Add($it.Rel)
            $dst = ConvertTo-LocalPath $dir $it.Rel
            if ($it.Rel -match '^BepInEx/config/[^/]+\.cfg$' -and [IO.File]::Exists($dst)) {
                $kept.Add($it.Rel)
                if ($prevFiles.ContainsKey($it.Rel)) { $owned.Add($it.Rel) }
                continue
            }
            $writes.Add([pscustomobject]@{ Item = $it; Dst = $dst })
            $owned.Add($it.Rel)
        }
        $ordered = @($writes | Where-Object { $_.Item.Rel -ne 'winhttp.dll' } | Sort-Object { $_.Item.Rel }) + @($writes | Where-Object { $_.Item.Rel -eq 'winhttp.dll' })
        $stale = New-Object System.Collections.Generic.List[string]
        foreach ($rel in $prevFiles.Keys) {
            if ($newSet.Contains($rel) -or $rel -match '\.cfg$') { continue }
            $p = ConvertTo-LocalPath $dir $rel
            if ([IO.File]::Exists($p)) { $stale.Add($p) }
        }
        $disabledNow = New-Object System.Collections.Generic.List[string]
        if ($null -ne $Prev) { foreach ($d in $Prev.DisabledConflicts) { if ([IO.File]::Exists((ConvertTo-LocalPath $dir $d))) { $disabledNow.Add($d) } } }

        # ---- apply (journaled, rolled back on any failure)
        Write-Step $M.step_install
        $journal = New-Object System.Collections.ArrayList
        $finished = $false
        $offRemoved = $false
        try {
            try {
                if ($DisableConflicts) {
                    foreach ($c in $Conflicts) {
                        Move-FileJ $journal $c.Path ($c.Path + '.disabled') $backup
                        $disabledNow.Add($c.Rel + '.disabled')
                        Write-Warn ($M.conflict_disabled -f $c.Display)
                    }
                }
                foreach ($w in $ordered) { Copy-FileJ $journal $w.Item.Path $w.Dst $backup }
                foreach ($s in $stale) { Remove-FileJ $journal $s $backup }
                $off = Join-Path $dir 'winhttp.dll.off'
                if ([IO.File]::Exists($off)) { Remove-FileJ $journal $off $backup; $offRemoved = $true }

                $receipt = [ordered]@{
                    name              = $PluginName
                    receiptVersion    = 1
                    version           = $Latest.Version
                    installedAt       = [DateTimeOffset]::Now.ToString('o')
                    createdBepInEx    = $createdBep
                    gameBuildId       = $Game.BuildId
                    files             = @($owned | Sort-Object -Unique)
                    disabledConflicts = @($disabledNow | Sort-Object -Unique)
                }
                $receiptStage = Join-Path $Work 'receipt.json'
                [IO.File]::WriteAllText($receiptStage, (ConvertTo-JsonText $receipt) + "`n", $Utf8NoBom)
                Copy-FileJ $journal $receiptStage (ConvertTo-LocalPath $dir $ReceiptRel) $backup

                # verify what was written
                foreach ($w in $writes) {
                    $h = Invoke-Retry { (Get-FileHash -LiteralPath $w.Dst -Algorithm SHA256).Hash }
                    if ($h -ne $w.Item.Sha256) { throw ('verification failed for ' + $w.Item.Rel) }
                }
                $finished = $true
            } catch {
                $err = $_
                Write-Log 'ERROR' ($err | Out-String)
                $bad = Undo-Journal $journal
                $finished = $true
                if ($bad -gt 0) { Write-Warn $M.rollback_partial } else { Write-Warn $M.rolled_back }
                if ($err.Exception.Data.Contains('SbgbiExit')) { throw $err }
                Stop-Installer $Exit.FileSystem (Get-FailureText $err)
            }
        } finally {
            if (-not $finished) { [void](Undo-Journal $journal) }
        }

        foreach ($k in $kept) { Write-Info ($M.config_kept -f $k) }
        Write-Info ($M.files_written -f $writes.Count)
        if ($stale.Count -gt 0) { Write-Info ($M.stale_removed -f $stale.Count) }
        if ($offRemoved) { Write-Info $M.vanilla_off }
    }

    function Invoke-InstallFlow {
        if ($BaseUrl -match 'CHANGE[-]ME') { Stop-Installer $Exit.Usage $M.no_base_url }
        if (-not (Test-UrlAllowed $BaseUrl)) { Stop-Installer $Exit.Usage ($M.url_not_https -f $BaseUrl) }

        Write-Step $M.step_latest
        $latest = Get-LatestInfo
        Write-Ok ($M.latest_version -f $latest.Version)
        if ($latest.GameVersion -and $latest.GameBuild) { Write-Info ($M.latest_tested -f $latest.GameVersion, $latest.GameBuild) }
        if ($latest.Notes.Count -gt 0) {
            Write-Info $M.latest_notes
            foreach ($n in ($latest.Notes | Select-Object -First 8)) { Write-Info ('        - ' + $n) }
        }

        Write-Step $M.step_game
        $game = Resolve-Game
        if ($game.BuildId -and $latest.GameBuild -and $game.BuildId -ne $latest.GameBuild) {
            Write-Warn ($M.game_build_differs -f $game.BuildId, $latest.GameBuild, (Get-CommandHint '-Disable'))
        }
        Assert-GameClosed

        $prev = Read-Receipt $game.Dir
        if ($null -ne $prev -and $prev.Version) { Write-Info ($M.installed_now -f $prev.Version) } else { Write-Info $M.installed_none }
        if ([IO.File]::Exists((Join-Path $game.Dir 'winhttp.dll.off'))) { Write-Info $M.vanilla_on }
        if ($null -eq $prev -and [IO.File]::Exists((Join-Path $game.Dir 'BepInEx\core\BepInEx.dll'))) { Write-Info $M.bepinex_foreign }

        $conflicts = @(Get-ConflictingPlugins $game.Dir)
        if ($conflicts.Count -gt 0) {
            Write-Warn $M.conflicts_found
            foreach ($c in $conflicts) { Write-Warn ($M.conflict_item -f $c.Display) }
            if (-not $DisableConflicts) { Write-Warn $M.conflicts_hint }
        }

        if ($DryRun) {
            Write-Info ($M.dry_download -f $latest.ZipUrl, (Format-Kb $latest.SizeBytes))
            Write-Info ($M.dry_extract -f $game.Dir)
            Write-Info $M.dry_receipt
            if ($conflicts.Count -gt 0 -and $DisableConflicts) { Write-Info ($M.dry_conflicts -f (($conflicts | ForEach-Object { $_.Display }) -join ', ')) }
            Write-Result $M.dry_done
            return
        }

        $work = Join-Path ([IO.Path]::GetTempPath()) ('sbgbi-' + [guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($work)
        Write-Log 'INFO' "work dir: $work"
        try {
            Install-Package -Game $game -Latest $latest -Prev $prev -Conflicts $conflicts -Work $work
        } finally {
            try { Invoke-Retry { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction Stop } } catch { Write-Log 'WARN' "cannot remove ${work}: $($_.Exception.Message)" }
        }

        if ($null -eq $prev -or -not $prev.Version) { Write-Result ($M.summary_installed -f $latest.Version) }
        elseif ($prev.Version -eq $latest.Version) { Write-Result ($M.summary_reinstalled -f $latest.Version) }
        else { Write-Result ($M.summary_updated -f $prev.Version, $latest.Version) }
        Write-Info ''
        Write-Info $M.summary_next
        Write-Note ($M.summary_same_version -f $latest.Version)
        Write-Info ($M.summary_hints -f (Get-CommandHint '-Disable'), (Get-CommandHint '-Enable'), (Get-CommandHint '-Uninstall'))
        Write-Info ($M.log_location -f $State.LogPath)
    }

    # ======================================================================================== uninstall
    function Invoke-UninstallFlow {
        Write-Step $M.mode_uninstall
        $game = Resolve-Game
        Assert-GameClosed
        $dir = $game.Dir
        $receipt = Read-Receipt $dir
        $pluginDir = ConvertTo-LocalPath $dir $PluginRel
        $bepDir = Join-Path $dir 'BepInEx'
        $receiptPath = ConvertTo-LocalPath $dir $ReceiptRel
        $backupRoot = Join-Path ([IO.Path]::GetTempPath()) ('sbgbi-' + [guid]::NewGuid().ToString('N'))
        $createdByUs = ($null -ne $receipt -and $receipt.CreatedBepInEx -eq $true)

        # 1. files of the mod itself: everything inside our plugin folder
        $modFiles = New-Object System.Collections.Generic.List[string]
        if ([IO.Directory]::Exists($pluginDir)) {
            foreach ($f in @(Get-ChildItem -LiteralPath $pluginDir -Recurse -File -Force)) { $modFiles.Add($f.FullName) }
        }

        # 2. conflicting plugins disabled earlier by -DisableConflicts
        $restore = New-Object System.Collections.Generic.List[string]
        if ($null -ne $receipt) {
            foreach ($d in $receipt.DisabledConflicts) {
                $disabledPath = ConvertTo-LocalPath $dir $d
                $original = $disabledPath.Substring(0, $disabledPath.Length - '.disabled'.Length)
                if ([IO.File]::Exists($disabledPath) -and -not [IO.File]::Exists($original)) { $restore.Add($d) }
            }
        }

        # 3. BepInEx itself: only with -Full, only when this installer created it, and only when no foreign plugins would be broken
        $frameworkFiles = New-Object System.Collections.Generic.List[string]
        $removeFramework = $false
        $fullNote = ''
        if ($Full) {
            if (-not $createdByUs) {
                $fullNote = $M.uninstall_full_skipped_foreign
            } else {
                $foreign = New-Object System.Collections.Generic.List[string]
                foreach ($sub in @('plugins', 'patchers', 'monomod')) {
                    $sd = Join-Path $bepDir $sub
                    if (-not [IO.Directory]::Exists($sd)) { continue }
                    foreach ($f in @(Get-ChildItem -LiteralPath $sd -Recurse -File -Force)) {
                        if ($f.FullName.StartsWith($pluginDir + '\', [StringComparison]::OrdinalIgnoreCase)) { continue }
                        $foreign.Add($f.FullName)
                    }
                }
                if ($foreign.Count -gt 0) {
                    $fullNote = $M.uninstall_full_skipped_plugins
                } else {
                    $removeFramework = $true
                    foreach ($rel in $receipt.Files) {
                        if ($rel -match '^BepInEx/plugins/') { continue }
                        $frameworkFiles.Add((ConvertTo-LocalPath $dir $rel))
                    }
                    $frameworkFiles.Add((Join-Path $dir 'winhttp.dll.off'))
                    foreach ($g in @('BepInEx\LogOutput.log', 'BepInEx\LogOutput.log.1', 'BepInEx\LogOutput.log.2')) { $frameworkFiles.Add((Join-Path $dir $g)) }
                    $cache = Join-Path $bepDir 'cache'
                    if ([IO.Directory]::Exists($cache)) { foreach ($f in @(Get-ChildItem -LiteralPath $cache -Recurse -File -Force)) { $frameworkFiles.Add($f.FullName) } }
                    $cfgDir = Join-Path $bepDir 'config'
                    if ([IO.Directory]::Exists($cfgDir)) {
                        foreach ($f in @(Get-ChildItem -LiteralPath $cfgDir -File -Force)) { if ($f.Name -match '(?i)sbgbalanceditems|sbgbi') { $frameworkFiles.Add($f.FullName) } }
                    }
                }
            }
        }

        # 4. the receipt: deleted when nothing of ours stays in the game folder; otherwise it is rewritten
        #    so that it only tracks the BepInEx files we created (a later "-Uninstall -Full" still knows them)
        $receiptAction = ''
        if ([IO.File]::Exists($receiptPath)) {
            if ($removeFramework -or -not $createdByUs) { $receiptAction = 'delete' }
            elseif ($null -ne $receipt -and $receipt.Version) { $receiptAction = 'rewrite' }
        }

        $allFiles = New-Object System.Collections.Generic.List[string]
        foreach ($f in $modFiles) { $allFiles.Add($f) }
        foreach ($f in $frameworkFiles) { if ([IO.File]::Exists($f)) { $allFiles.Add($f) } }
        if ($receiptAction -eq 'delete') { $allFiles.Add($receiptPath) }
        if ($allFiles.Count -eq 0 -and $restore.Count -eq 0 -and $receiptAction -ne 'rewrite') {
            Write-Warn $M.uninstall_nothing
            if ($fullNote) { Write-Info $fullNote }
            return
        }

        if ($DryRun) {
            foreach ($f in ($allFiles | Select-Object -First 15)) { Write-Info ($M.uninstall_plan_item -f (ConvertTo-RelPath $f $dir)) }
            if ($allFiles.Count -gt 15) { Write-Info ($M.uninstall_plan_more -f ($allFiles.Count - 15)) }
            foreach ($r in $restore) { Write-Info ($M.conflict_restored -f $r.Substring(0, $r.Length - '.disabled'.Length)) }
            if ($fullNote) { Write-Info $fullNote }
            Write-Result $M.dry_done
            return
        }

        [void][IO.Directory]::CreateDirectory($backupRoot)
        $journal = New-Object System.Collections.ArrayList
        $finished = $false
        $dirLeft = $false
        try {
            try {
                foreach ($f in $allFiles) { Remove-FileJ $journal $f $backupRoot }
                foreach ($r in $restore) {
                    $disabledPath = ConvertTo-LocalPath $dir $r
                    $original = $disabledPath.Substring(0, $disabledPath.Length - '.disabled'.Length)
                    Move-FileJ $journal $disabledPath $original $backupRoot
                    Write-Info ($M.conflict_restored -f $r.Substring(0, $r.Length - '.disabled'.Length))
                }
                if ($receiptAction -eq 'rewrite') {
                    $newReceipt = [ordered]@{
                        name              = $PluginName
                        receiptVersion    = 1
                        version           = ''
                        installedAt       = $receipt.InstalledAt
                        uninstalledAt     = [DateTimeOffset]::Now.ToString('o')
                        createdBepInEx    = $true
                        gameBuildId       = $receipt.GameBuildId
                        files             = @($receipt.Files | Where-Object { $_ -notmatch '^BepInEx/plugins/' } | Sort-Object -Unique)
                        disabledConflicts = @()
                    }
                    $receiptStage = Join-Path $backupRoot 'receipt.json'
                    [IO.File]::WriteAllText($receiptStage, (ConvertTo-JsonText $newReceipt) + "`n", $Utf8NoBom)
                    Copy-FileJ $journal $receiptStage $receiptPath $backupRoot
                }
                # remove empty directories: our plugin folder tree always, the rest of BepInEx only after -Full
                $dirs = New-Object System.Collections.Generic.List[string]
                if ([IO.Directory]::Exists($pluginDir)) {
                    foreach ($d in @(Get-ChildItem -LiteralPath $pluginDir -Recurse -Directory -Force)) { $dirs.Add($d.FullName) }
                    $dirs.Add($pluginDir)
                }
                if ($removeFramework -and [IO.Directory]::Exists($bepDir)) {
                    foreach ($d in @(Get-ChildItem -LiteralPath $bepDir -Recurse -Directory -Force)) { $dirs.Add($d.FullName) }
                    $dirs.Add($bepDir)
                }
                $uniq = @($dirs | Sort-Object -Unique | Sort-Object -Property Length -Descending)
                foreach ($d in $uniq) { [void](Remove-DirIfEmptyJ $journal $d) }
                if ($removeFramework -and [IO.Directory]::Exists($bepDir)) { $dirLeft = $true }
                $finished = $true
            } catch {
                $err = $_
                Write-Log 'ERROR' ($err | Out-String)
                $bad = Undo-Journal $journal
                $finished = $true
                if ($bad -gt 0) { Write-Warn $M.rollback_partial } else { Write-Warn $M.rolled_back }
                if ($err.Exception.Data.Contains('SbgbiExit')) { throw $err }
                Stop-Installer $Exit.FileSystem (Get-FailureText $err)
            }
        } finally {
            if (-not $finished) { [void](Undo-Journal $journal) }
            try { Invoke-Retry { Remove-Item -LiteralPath $backupRoot -Recurse -Force -ErrorAction Stop } } catch { Write-Log 'WARN' "cannot remove ${backupRoot}: $($_.Exception.Message)" }
        }

        Write-Result ($M.uninstall_done -f $allFiles.Count)
        if ($removeFramework) {
            if ($dirLeft) { Write-Info $M.uninstall_dir_left } else { Write-Info $M.uninstall_full_done }
        } elseif ($fullNote) {
            Write-Warn $fullNote
        } elseif ($createdByUs -and [IO.Directory]::Exists($bepDir)) {
            Write-Info ($M.uninstall_bepinex_left -f (Get-CommandHint '-Uninstall -Full'), (Get-CommandHint '-Disable'))
        }
        Write-Info ($M.log_location -f $State.LogPath)
    }
    # ======================================================================================== enable / disable
    function Invoke-ToggleFlow {
        param([bool]$WantEnabled)
        if ($WantEnabled) { Write-Step $M.mode_enable } else { Write-Step $M.mode_disable }
        $game = Resolve-Game
        Assert-GameClosed
        $dll = Join-Path $game.Dir 'winhttp.dll'
        $off = $dll + '.off'
        $hasDll = [IO.File]::Exists($dll)
        $hasOff = [IO.File]::Exists($off)
        try {
            if ($WantEnabled) {
                if ($hasOff -and $hasDll) {
                    if ($DryRun) { Write-Info ($M.toggle_dry -f 'winhttp.dll.off', '(delete stale copy)') }
                    else { Clear-ReadOnly $off; [IO.File]::Delete($off); Write-Result $M.toggle_already_enabled }
                } elseif ($hasOff) {
                    if ($DryRun) { Write-Info ($M.toggle_dry -f 'winhttp.dll.off', 'winhttp.dll') }
                    else { [IO.File]::Move($off, $dll); Write-Result $M.toggle_enabled }
                } elseif ($hasDll) {
                    Write-Result $M.toggle_already_enabled
                } else {
                    Write-Warn $M.toggle_not_installed
                }
            } else {
                if ($hasDll) {
                    if ($DryRun) { Write-Info ($M.toggle_dry -f 'winhttp.dll', 'winhttp.dll.off') }
                    else {
                        if ($hasOff) { Clear-ReadOnly $off; [IO.File]::Delete($off) }
                        [IO.File]::Move($dll, $off)
                        Write-Result ($M.toggle_disabled -f (Get-CommandHint '-Enable'))
                    }
                } elseif ($hasOff) {
                    Write-Result ($M.toggle_already_disabled -f (Get-CommandHint '-Enable'))
                } else {
                    Write-Warn $M.toggle_not_installed
                }
            }
        } catch {
            if ($_.Exception.Data.Contains('SbgbiExit')) { throw }
            Write-Log 'ERROR' ($_ | Out-String)
            Stop-Installer $Exit.FileSystem (Get-FailureText $_)
        }
        if ($DryRun) { Write-Result $M.dry_done }
    }

    # ======================================================================================== main
    $oldProtocol = $null
    try {
        Initialize-Log
        try {
            $oldProtocol = [Net.ServicePointManager]::SecurityProtocol
            [Net.ServicePointManager]::SecurityProtocol = $oldProtocol -bor [Net.SecurityProtocolType]::Tls12
        } catch { Write-Log 'WARN' "cannot set TLS 1.2: $($_.Exception.Message)" }

        $BaseUrl = $BaseUrl.Trim().TrimEnd('/')
        $modes = @($Uninstall.IsPresent, $Disable.IsPresent, $Enable.IsPresent) | Where-Object { $_ }
        if (@($modes).Count -gt 1) { Stop-Installer $Exit.Usage $M.usage_exclusive }
        if ($Full -and -not $Uninstall) { Stop-Installer $Exit.Usage $M.uninstall_full_needs_uninstall }

        Write-Step ('=== ' + $M.title + ' ===')
        if ($DryRun) { Write-Warn $M.dryrun_banner }

        if ($Uninstall) { Invoke-UninstallFlow }
        elseif ($Disable) { Invoke-ToggleFlow $false }
        elseif ($Enable) { Invoke-ToggleFlow $true }
        else { Invoke-InstallFlow }
    } catch {
        $ex = $_.Exception
        if ($ex.Data.Contains('SbgbiExit')) {
            $State.ExitCode = [int]$ex.Data['SbgbiExit']
            Write-Err $ex.Message
        } else {
            $State.ExitCode = $Exit.Error
            Write-Log 'ERROR' ($_ | Out-String)
            Write-Err ($M.unexpected -f $ex.Message)
        }
        Write-Err ($M.log_location -f $State.LogPath)
    } finally {
        if ($null -ne $oldProtocol) { try { [Net.ServicePointManager]::SecurityProtocol = $oldProtocol } catch { $null = $_ } }
    }
    Write-Log 'INFO' ('==== installer end; exit code {0}' -f $State.ExitCode)
    if ($PSCommandPath) { exit $State.ExitCode }
    $global:LASTEXITCODE = $State.ExitCode
}
