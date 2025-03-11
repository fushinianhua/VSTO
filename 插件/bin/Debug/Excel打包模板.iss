; 基础变量

; 插件名称
#define MyAppName "Excel插件"
; 插件的Dll名称
#define MyAppDllName "插件"
; 插件版本
#define MyAppVersion "1.0.0"
; 发布者名 
#define MyAppPublisher "浮世年华"

; 程序唯一ID,很关键,(若要生成新的 GUID，可在菜单中点击 "tools|Generate GUID"。)
#define MyAppId "{CB7FE072-42EA-4059-B2BC-F73B0210F4B8}"
#define MyAppInclusionId    "B2855094-A608-4BCD-995F-F66A97C90683"
#define MyAppMetadataId     "4F441337-73FF-42B4-ACF2-F0C88A22BCC4"
; 文件夹目录配置
; 安装包输出目录
#define SteupOutputDir "."
 

; 要打包的文件夹
#define SteupPackDir "*"
; 要排除的文件和文件夹
#define ExcludesFile "*.lic,*.pdb,*.xml,*.ohh,\log\*,\Obfuscated\*,*.iss"

[Setup]
; 安装配置
; 管理员身份运行
; PrivilegesRequired=admin
; 程序唯一ID
AppId={{#MyAppId}
; 程序名称
AppName={#MyAppName}
; 程序版本
AppVersion={#MyAppVersion}
; 安装包版本
VersionInfoVersion = {#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
; 发布者名称
AppPublisher={#MyAppPublisher}



; 安装到的目录
DefaultDirName= "d:\Program Files\ETools"
//DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; 安装权限：管理员
PrivilegesRequired=admin
; 安装包输出目录
OutputDir=.
; 安装包输出名称
OutputBaseFilename={#MyAppName}_V{#MyAppVersion}
; 安装包图标

; 协议
LicenseFile= 许可协议.txt
; 自述文件,txt,rtf
; InfoBeforeFile=
Compression=lzma
SolidCompression=yes
; 样式配置
 ;WizardStyle=modern 
 //这个模式，窗口太大，不好看
; 向导页面,如果设置为 yes，安装程序将不显示欢迎 。
DisableWelcomePage=no
;不显示准备安装界面
DisableReadyPage=yes
;InfoBeforeFile=infobefore.rtf
;InfoAfterFile=InfoAfterFile.rtf
WizardImageStretch=yes


; 安装需要输入密码  
; Password={#MyDate}
; Encryption=yes 
; 可以让用户忽略选择语言相关  
ShowLanguageDialog =no 
; 允许安装的系统架构
ArchitecturesAllowed=x86 x64
; 添加控制面板卸载按钮
UninstallDisplayIcon={app}\A.ico
Uninstallable=yes
UninstallDisplayName=卸载{#MyAppName}
[Messages]
;SetupWindowTitle={#MyAppName}安装向导
;ClickNext=为确保本软件安装成功，请先关闭360或者电脑管家、金山毒霸之类，然后再安装本软件。

[Run]
; 根据产品ID,静默卸载之前使用Advanced Installer创建的安装包程序
;Filename: "cmd.exe"; Parameters: "/c MsiExec.exe /x {{2D7ADA3D-B870-4189-ADB8-4DEA2E7C457E}/quiet"; WorkingDir: "{app}"; Flags: runhidden
 
//Filename: "{app}/MyProg.exe"; Description: "{cm:LaunchProgram,我的程序}"; Flags: nowait postinstall skipifsilent
 Filename: "https://www.vbashuo.top/questions"; Description: "查看安装问题指南，必看！"; Flags: postinstall shellexec skipifsilent
 
[Code]
//欢迎页默认选中同意按钮
var
LabelDate:Tlabel;
procedure InitializeWizard();
begin
WizardForm.LicenseAcceptedRadio.Checked := true;
WizardForm.WELCOMELABEL2.hide
LabelDate:=Tlabel.Create(WizardForm);
LabelDate.Caption:='1.插件支持office2007~office365、WPS。'#13''#13'2.安装过程中，若360出现提示，请允许操作。'#13''#13'2.若安装失败，请联系微信：vbashuo3。'#13''#13'' ;
LabelDate.Left:=WizardForm.WELCOMELABEL1.Left;
LabelDate.Top:= WizardForm.WELCOMELABEL1.Top + WizardForm.WELCOMELABEL1.Height;
LabelDate.Width:=WizardForm.WELCOMELABEL1.Width;
LabelDate.AutoSize := False;
LabelDate.Parent := WizardForm.WelcomePage;
LabelDate.Font.Color:=$0000FF;
WizardForm.WELCOMELABEL1.Font.Name := '微软雅黑';
LabelDate.Font.Name := '微软雅黑';
end;

// ========================下面为自定义函数===================
// 检测系统是否已安装大于 .net 4.0
function checkDot4Net():Boolean;
var
  isSucc: Boolean;
  dotNetVersion: DWord;
begin

  isSucc :=false;

  if RegValueExists(HKLM,'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full','Install') then
    begin
      if RegQueryDWordValue(HKLM,'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full','Install',dotNetVersion) then
        begin
          if dotNetVersion = 1 then //判断依据见:https://learn.microsoft.com/zh-cn/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed
            isSucc := true
          else
            isSucc := false;
        end;
    end;

  if RegValueExists(HKLM,'SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full','Install') then
    begin
      if RegQueryDWordValue(HKLM,'SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full','Install',dotNetVersion) then
        begin
          if dotNetVersion >= 1 then //判断依据见:https://learn.microsoft.com/zh-cn/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed
            isSucc := true
          else
            isSucc := false;
        end;
    end;

  result :=isSucc
end;

// 字符串替换
function ReplaceString(const Source, OldValue, NewValue: string): string;
var
  TempStr: string;
  PosStart: Integer;
begin
  Result := Source;
  PosStart := Pos(OldValue, Result);
  while PosStart > 0 do
  begin
    TempStr := Copy(Result, 1, PosStart-1) + NewValue + 
               Copy(Result, PosStart+Length(OldValue), Length(Result)-PosStart-Length(OldValue)+1);
    Result := TempStr;
    PosStart := Pos(OldValue, Result);
  end;
end;

// 检测Vsto Runtime 4.0是否已安装       https://learn.microsoft.com/zh-cn/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=vs-2022
function checkVstoRunTime():Boolean;
var
  isSucc: Boolean;
  vstoVersion: String;
  resultVsreion: Integer;
begin

  isSucc :=false;

  if RegValueExists(HKLM,'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R','Version') then
    begin
     if RegQueryStringValue(HKLM,'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R','Version',vstoVersion) then
       begin
          resultVsreion := StrToInt(ReplaceString(vstoVersion,'.',''));
          if resultVsreion >= 10000000 then
            isSucc :=true
          else
            isSucc :=false;
       end;
    end;

  if RegValueExists(HKLM,'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R','Version') then
    begin
     if RegQueryStringValue(HKLM,'SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R','Version',vstoVersion) then
      begin
        resultVsreion := StrToInt(ReplaceString(vstoVersion,'.',''));
        if resultVsreion >= 10000000 then
            isSucc :=true
        else
            isSucc :=false;
      end;
    end;

  result :=isSucc;
end;
//检测Excel 是否正在运行。
function IsAppRunning(const FileName: string): Boolean;
var
  FWMIService: Variant;
  FSWbemLocator: Variant;
  FWbemObjectSet: Variant;
begin
  Result := false;
  FSWbemLocator := CreateOleObject('WBEMScripting.SWBEMLocator');
  FWMIService := FSWbemLocator.ConnectServer('', 'root\CIMV2', '', '');
  FWbemObjectSet := FWMIService.ExecQuery(Format('SELECT Name FROM Win32_Process Where Name="%s"',[FileName]));
  Result := not VarIsNull(FWbemObjectSet)and (FWbemObjectSet.Count > 0)
  FWbemObjectSet := Unassigned;
  FWMIService := Unassigned;
  FSWbemLocator := Unassigned;
end;
//;  如在运行，安装或卸载时，需要关闭。
const wbemFlagForwardOnly = $00000020;
procedure CloseApp(AppName: String);
var
  WbemLocator : Variant;
  WMIService   : Variant;
  WbemObjectSet: Variant;
  WbemObject   : Variant;
begin;
  WbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  WMIService := WbemLocator.ConnectServer('localhost', 'root\CIMV2');
  WbemObjectSet := WMIService.ExecQuery('SELECT * FROM Win32_Process Where Name="' + AppName + '"');
  if not VarIsNull(WbemObjectSet) and (WbemObjectSet.Count > 0) then
  begin
    WbemObject := WbemObjectSet.ItemIndex(0);
    if not VarIsNull(WbemObject) then
    begin
      WbemObject.Terminate();
      WbemObject := Unassigned;
    end;
  end;
end;

// ========================下面为安装环境检测和安装事件函数===================
// 一,初始安装函数


function InitializeSetup(): boolean;
var
  ResultStr: String;
  ResultCode: Integer;
  ErrorCode: Integer;
 // ========================检测Excel或者WPS是否打开===================
begin//https://stackoverflow.com/questions/5545077/inno-setup-kill-a-running-process
   // 1.安装时关闭Excel,word,wps
  Result := not (IsAppRunning('EXCEL.exe') or IsAppRunning('et.exe') or IsAppRunning('wps.exe') );
  if not Result then
  begin 
      if MsgBox('为保证安装成功，请关闭Excel/WPS再继续。'#13#13'【是】强制关闭，继续安装'#13#13'【否】手动关闭，稍后再安装 ', mbConfirmation, MB_YESNO) = IDYES then
        begin 
        //;关闭excel进程。
        try
             CloseApp('EXCEL.exe') 
             CloseApp('et.exe') 
             CloseApp('wps.exe')
        except
        end           
             //安装程序继续
            Result:=true;

        end else  
        //安装取消。 
         begin
            Result:=false;  
         end;  
  end;
  if  Result then
// 2.安装前先卸载原程序
  // 安装在全部用户时
  //if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppId}_is1', 'UninstallString', ResultStr) then
  // 安装在当前用户时
  if RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppId}_is1', 'UninstallString', ResultStr) then
    begin
      ResultStr := RemoveQuotes(ResultStr);
      Exec(ResultStr, '/silent', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    end;
    //result := true; //继续安装

  // 3.检测.net运行环境
  if not checkDot4Net then
                     begin
                      MsgBox('安装必备运行环境Framework 4.0',mbInformation,MB_OK);
                       ExtractTemporaryFile('dotNetFx40_Full_setup.exe');
                        Exec(ExpandConstant('{tmp}\dotNetFx40_Full_setup.exe'), '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
                      end;

  // 4.检测office runtime 2010 运行时安装情况
  if not checkVstoRunTime then
         begin
                          MsgBox('先安装必备运行环境Visual Studio Tools for Office runtime',mbInformation,MB_OK);
                          if IsWin64 then
                                 begin
                                    ExtractTemporaryFile('vstor40_x64.exe');
                                   //Exec(ExpandConstant('{tmp}\vstor40_x64.exe'), '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
                                    Exec(ExpandConstant('{tmp}\vstor40_x64.exe'), '/VERYSILENT', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
                                 end
                          else
                                 begin
                                    ExtractTemporaryFile('vstor40_x86.exe');
                                    //Exec(ExpandConstant('{tmp}\vstor40_x86.exe'), '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
                                    Exec(ExpandConstant('{tmp}\vstor40_x86.exe'), '/VERYSILENT', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
                                 end;
         end;
  //Result := true; //全部通过后,开始安装插件
end;


// 二,删除旧文件
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usDone then
    begin
    DelTree(ExpandConstant('{app}'), True, True, True);
    end;
end;



function GetPublicKey(Param: String): String;
// 提取.vsto文件中的公钥
var s: AnsiString;
begin
    LoadStringFromFile(Param, s);
    Delete(s, 1, Pos('<RSAKeyValue>', s) - 1);
    Delete(s, Pos('</RSAKeyValue>', s) + 14, 32767);
    Result := s;
end;


[Registry]
; 注册表管理
; VSTO注册信息
; Excel注册表
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\{#MyAppDllName}";ValueType: string; ValueName: "Description"; ValueData:"{#MyAppName}" ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\{#MyAppDllName}";ValueType: string; ValueName: "FriendlyName"; ValueData:"{#MyAppName}" ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\{#MyAppDllName}";ValueType: dword; ValueName: "LoadBehavior"; ValueData:3 ; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\{#MyAppDllName}";ValueType: string; ValueName: "Manifest"; ValueData:"{app}\{#MyAppDllName}.vsto|vstolocal" ; Flags: uninsdeletekey 
; Wps注册表,注意卸载时必须时:uninsdeletevalue标记,否则会误删别的插件
Root: HKCU; Subkey: "Software\Kingsoft\Office\ET\AddinsWL";ValueType: string; ValueName: "{#MyAppDllName}"; ValueData:"" ; Flags: uninsdeletevalue
 ; VSTO安全提示白名单
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VSTO\Security\Inclusion\{#MyAppInclusionId}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VSTO\Security\Inclusion\{#MyAppInclusionId}"; ValueName: "Url"; ValueType: String; ValueData: "file:///{app}\{#MyAppDllName}.vsto"
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VSTO\Security\Inclusion\{#MyAppInclusionId}"; ValueName: "PublicKey"; ValueType: String; ValueData: "{code:GetPublicKey|{app}\{#MyAppDllName}.vsto}"
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VSTO\SolutionMetadata"; ValueName: "{app}\{#MyAppDllName}.vsto"; ValueType: String; ValueData: "{#MyAppMetadataId}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VSTO\SolutionMetadata\{#MyAppMetadataId}"; Flags: uninsdeletekey


[Languages]
; 语言配置
Name: "english"; MessagesFile: "compiler:Default.isl"
;Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"
[Files]
; 安装包的打包文件,忽略掉的文件和文件夹:Excludes
Source: "{#SteupPackDir}" ; Excludes: "{#ExcludesFile}";  DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
