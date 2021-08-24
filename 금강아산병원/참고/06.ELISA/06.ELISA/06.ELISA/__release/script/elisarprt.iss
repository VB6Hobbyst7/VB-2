; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyCompany "SD BIOSENSOR"
#define MyAppPublisher "SD BIOSENSOR, Inc."
#define MyAppURL "http://www.sdbiosensor.com/"

#define MyAppName "ELISA REPORT"
#define MyAppExeName "elisarprt.exe"
#define MyAppVersion GetFileVersion("..\files\elisarprt.exe")
#define MyRegKey "SDBIOSENSOR\elisarprt"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{48DD9E4B-3ADF-4E12-96DA-AE8DEB097F6C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={sd}\{#MyCompany}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=yes
AllowNoIcons=yes
OutputDir=..\bin
OutputBaseFilename=Setup.{#MyAppName}(v{#MyAppVersion})
UninstallDisplayIcon={app}\{#MyAppExeName}   
TimeStampsInUTC=True
LanguageDetectionMethod=locale
DisableProgramGroupPage=auto
VersionInfoCompany=SD BIOSENSOR, Inc.
;LicenseFile=..\eula\eula_en.txt
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
SignTool=sign
SignedUninstaller=True
DisableWelcomePage=no

[Registry]
Root: HKCU; Subkey: "{#MyRegKey}"

;[Languages]
;Name: "english"; MessagesFile: "Default.isl"; LicenseFile: "..\eula\eula_en.txt";
;Name: "korean"; MessagesFile: "Korean.isl"; LicenseFile: "..\eula\eula_kr.txt";  

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkablealone checkedonce
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkablealone checkedonce; OnlyBelowVersion: 0,6.1

;[Components]
;Name: "App"; Description: "Vcheck"; Types: full; Flags: fixed

[Files]
; install target files
Source: "..\files\elisarprt.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\files\elisarprt.i18n"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"; 
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
;Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon


[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent; 

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[InstallDelete]
Type: filesandordirs; Name: "{app}\{#MyAppExeName}"

[Code]


{
  // Unzip: see - https://gist.github.com/jakoch/33ac13800c17eddb2dd4
}
{
	Unzip Helper for executing 7zip without blocking the InnoSetup GUI

    ----

	The main procedure is the non-blocking Unzip().
	Your GUI will remain responsive during the unzip operation.

    Written by Rik and Jens A. Koch (@jakoch) on StackOverflow:
    http://stackoverflow.com/questions/32256432/how-to-execute-7zip-without-blocking-the-innosetup-ui

    ----

    Usage:

	1. Include this ISS with

	   // #include "..\some\where\unzip.iss"

	2. Add the unzip tool "7za.exe" to the [Files] section of your installer
	   and copy it to the temp folder during installation.

       // [Files]
       // Source: ..\some\where\7za.exe; DestDir: {tmp}; Flags: dontcopy

	3. Finally, extract your files using Unzip(source, target); in the [Code] section.
}

#IFDEF UNICODE
  #DEFINE AW "W"
#ELSE
  #DEFINE AW "A"
#ENDIF

// --- Start "ShellExecuteEx" Helper

const
  WAIT_TIMEOUT = $00000102;
  SEE_MASK_NOCLOSEPROCESS = $00000040;
  INFINITE = $FFFFFFFF;     { Infinite timeout }

type
  TShellExecuteInfo = record
    cbSize: DWORD;
    fMask: Cardinal;
    Wnd: HWND;
    lpVerb: string;
    lpFile: string;
    lpParameters: string;
    lpDirectory: string;
    nShow: Integer;
    hInstApp: THandle;
    lpIDList: DWORD;
    lpClass: string;
    hkeyClass: THandle;
    dwHotKey: DWORD;
    hMonitor: THandle;
    hProcess: THandle;
  end;

function ShellExecuteEx(var lpExecInfo: TShellExecuteInfo): BOOL;
  external 'ShellExecuteEx{#AW}@shell32.dll stdcall';
function WaitForSingleObject(hHandle: THandle; dwMilliseconds: DWORD): DWORD;
  external 'WaitForSingleObject@kernel32.dll stdcall';
function CloseHandle(hObject: THandle): BOOL; external 'CloseHandle@kernel32.dll stdcall';

// --- End "ShellExecuteEx" Helper

// --- Start "Application.ProcessMessage" Helper
{
   InnoSetup does not provide Application.ProcessMessage().
   This is "generic" code to recreate a "Application.ProcessMessages"-ish procedure,
   using the WinAPI function PeekMessage(), TranslateMessage() and DispatchMessage().
}
type
  TMsg = record
    hwnd: HWND;
    message: UINT;
    wParam: Longint;
    lParam: Longint;
    time: DWORD;
    pt: TPoint;
  end;

const
  PM_REMOVE = 1;

function PeekMessage(var lpMsg: TMsg; hWnd: HWND; wMsgFilterMin, wMsgFilterMax, wRemoveMsg: UINT): BOOL; external 'PeekMessageA@user32.dll stdcall';
function TranslateMessage(const lpMsg: TMsg): BOOL; external 'TranslateMessage@user32.dll stdcall';
function DispatchMessage(const lpMsg: TMsg): Longint; external 'DispatchMessageA@user32.dll stdcall';

procedure AppProcessMessage;
var
  Msg: TMsg;
begin
  while PeekMessage(Msg, WizardForm.Handle, 0, 0, PM_REMOVE) do begin
    TranslateMessage(Msg);
    DispatchMessage(Msg);
  end;
end;

// --- End "Application.ProcessMessage" Helper

procedure Unzip(source: String; targetdir: String);
var
  unzipTool, unzipParams : String; // path and param for the unzip tool
  ExecInfo: TShellExecuteInfo;     // info object for ShellExecuteEx()
begin
    // source and targetdir might contain {tmp} or {app} constant, so expand/resolve it to path names
    source := ExpandConstant(source);
    targetdir := ExpandConstant(targetdir);

    // prepare 7zip execution
    unzipTool := ExpandConstant('{tmp}\7za.exe');
    unzipParams := ' x "' + source + '" -o"' + targetdir + '" -y';

    // prepare information about the application being executed by ShellExecuteEx()
    ExecInfo.cbSize := SizeOf(ExecInfo);
    ExecInfo.fMask := SEE_MASK_NOCLOSEPROCESS;
    ExecInfo.Wnd := 0;
    ExecInfo.lpFile := unzipTool;
    ExecInfo.lpParameters := unzipParams;
    ExecInfo.nShow := SW_HIDE;

    if not FileExists(unzipTool)
    then MsgBox('UnzipTool not found: ' + unzipTool, mbError, MB_OK)
    else if not FileExists(source)
    then MsgBox('File was not found while trying to unzip: ' + source, mbError, MB_OK)
    else begin

          {
             The unzip tool is executed via ShellExecuteEx()
             Then the installer uses a while loop with the condition
             WaitForSingleObject and a very minimal timeout
             to execute AppProcessMessage.

             AppProcessMessage is itself a helper function, because
             Innosetup does not provide Application.ProcessMessages().
             Its job is to be the message pump to the InnoSetup GUI.

             This trick makes the window responsive/dragable again,
             while the extraction is done in the background.
          }

          if ShellExecuteEx(ExecInfo) then
          begin
            while WaitForSingleObject(ExecInfo.hProcess, 100) = WAIT_TIMEOUT
            do begin
                AppProcessMessage;
                WizardForm.Refresh();
            end;
          CloseHandle(ExecInfo.hProcess);
          end;

    end;
end;

// --- End "procedure Unzip(source: String; targetdir: String);" Helper

function IsAppRunning(const FileName : string): Boolean;
var
  FSWbemLocator: Variant;
  FWMIService   : Variant;
  FWbemObjectSet: Variant;
begin
  Result := false;
  FSWbemLocator := CreateOleObject('WBEMScripting.SWBEMLocator');
  FWMIService := FSWbemLocator.ConnectServer('', 'root\CIMV2', '', '');
  FWbemObjectSet := FWMIService.ExecQuery(Format('SELECT Name FROM Win32_Process Where Name="%s"',[FileName]));
  Result := (FWbemObjectSet.Count > 0);
  FWbemObjectSet := Unassigned;
  FWMIService := Unassigned;
  FSWbemLocator := Unassigned;
end;

function ProductTerminate: Integer;
begin    
  ShellExec('open','taskkill.exe','/f /im {#MyAppExeName}', '', SW_HIDE, ewNoWait, Result);
  ShellExec('open','tskill.exe',' {#MyAppName}', '', SW_HIDE, ewNoWait, Result);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    if RegKeyExists(HKEY_CURRENT_USER, '{#MyRegKey}') then
      //if MsgBox('Do you want to delete the overlay filter registry key ?',
      //  mbConfirmation, MB_YESNO) = IDYES
      //then
        RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, '{#MyRegKey}');
  end;
end;

function InitializeSetup: Boolean; 
var
  Cnt: Integer;
  ResultCode: Integer;
begin
  while IsAppRunning('{#MyAppExeName}') do
  begin
    if MsgBox( '{#MyAppName} is running. Click Yes to shut it down and continue installation, or click No to exit.', mbConfirmation, MB_YESNO ) = IDNO then
    begin
      Result := False;
      Exit;
    end;
    ProductTerminate;

    Cnt := 0;
    while IsAppRunning('{#MyAppExeName}') do
    begin
      Sleep(100);
      Cnt := Cnt +1;
      if Cnt > 10 then
        Break;
    end;
  end;
  Result := True
end;

function InitializeUninstall: Boolean;
var
  ResultCode: Integer;
begin
  ProductTerminate;  
  Result := True;
end;