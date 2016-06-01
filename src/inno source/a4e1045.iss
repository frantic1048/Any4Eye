; 脚本用 Inno Setup 脚本向导 生成。
; 查阅文档获取创建 INNO SETUP 脚本文件的详细资料！

#define MyAppName "Any 4 Eye"
#define MyAppVersion "1.1.1"
#define MyAppPublisher "Frantic Black"
#define MyAppURL "http://blog.163.com/frantic_hao/"
#define MyAppExeName "Any 4 Eye 1.1.1.exe"

[ISFormDesigner]
WizardForm=FF0A005457495A415244464F524D003010BC02000054504630F10B5457697A617264466F726D0A57697A617264466F726D0C436C69656E74486569676874034C010B436C69656E74576964746803F101134F6E436F6E73747261696E6564526573697A6507105F4E657874427574746F6E436C69636B0D4578706C69636974576964746803FB010E4578706C69636974486569676874036C010D506978656C73506572496E636802600A54657874486569676874020C00F10654426576656C05426576656C03546F70031F010B4578706C69636974546F70031F010000F10A544E6577427574746F6E0A4E657874427574746F6E074F6E436C69636B07105F4E657874427574746F6E436C69636B0000F10C544E65774E6F7465626F6F6B0D4F757465724E6F7465626F6F6B00F110544E65774E6F7465626F6F6B506167650B57656C636F6D655061676500F10C544269746D6170496D6167651157697A6172644269746D6170496D61676505576964746803AD0006486569676874031E010D4578706C69636974576964746803AD000E4578706C69636974486569676874031E01000000F110544E65774E6F7465626F6F6B5061676509496E6E65725061676500F10C544E65774E6F7465626F6F6B0D496E6E65724E6F7465626F6F6B00F110544E65774E6F7465626F6F6B506167651453656C656374436F6D706F6E656E74735061676500F10C544E6577436F6D626F426F780A5479706573436F6D626F0648656967687402140000000000F110544E65774E6F7465626F6F6B506167650C46696E6973686564506167650B4578706C69636974546F7002F800F10C544269746D6170496D6167651257697A6172644269746D6170496D6167653206486569676874033A010E4578706C69636974486569676874033A010000F10F544E6577526164696F427574746F6E08596573526164696F074F6E436C69636B070D596573526164696F436C69636B0000000000

[Setup]
WizardImageFile=datiaofu副本.bmp
WizardSmallImageFile=E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\disc.bmp
; 注意: AppId 的值是唯一识别这个程序的标志。
; 不要在其他程序中使用相同的 AppId 值。
; (在编译器中点击菜单“工具 -> 产生 GUID”可以产生一个新的 GUID)
AppId={{732C547D-9CC4-4F97-AA40-AFE06B31FDB6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
InfoBeforeFile=E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\1.1.1.txt
OutputDir=E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\包
OutputBaseFilename=Any4Eye1101setup2nd
SetupIconFile=E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\disc.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: default; MessagesFile: compiler:Default.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]


Source: E:\郭珈豪\A4E\A4E source code 1101\A4E source code 1100\Any 4 Eye 1.1.1.exe; DestDir: {app}; Flags: ignoreversion
Source: E:\郭珈豪\A4E\A4E source code 1101\A4E source code 1100\BackGround.exe; DestDir: {app}; Flags: ignoreversion
Source: C:\Windows\SysWOW64\FM20.DLL; DestDir: {sys}; Flags: uninsneveruninstall 32bit regtypelib
Source: E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\lib\FM20.DLL; DestDir: {win}\SysWOW64\; Flags: noregerror uninsneveruninstall regserver 64bit
Source: E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\lib\MSCOMCTL.OCX; DestDir: {sys}; Flags: noregerror uninsneveruninstall regserver
Source: E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\lib\MSMASK32.OCX; DestDir: {sys}; Flags: noregerror uninsneveruninstall regserver
Source: E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\lib\MSMSKCHS.DLL; DestDir: {sys}; Flags: noregerror uninsneveruninstall regserver
Source: E:\郭珈豪\A4E\A4E 1.1.1 inno source\A4E 1.1.1 inno source\lib\VB6CHS.DLL; DestDir: {sys}; Flags: noregerror uninsneveruninstall regserver


; 注意: 不要在任何共享的系统文件使用 "Flags: ignoreversion"


[Icons]
Name: {group}\{#MyAppName}; Filename: {app}\{#MyAppExeName}
Name: {group}\{cm:UninstallProgram,{#MyAppName}}; Filename: {uninstallexe}
Name: {commondesktop}\{#MyAppName}; Filename: {app}\{#MyAppExeName}; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}; Filename: {app}\{#MyAppExeName}; Tasks: quicklaunchicon

[Run]
Filename: {app}\{#MyAppExeName}; Description: {cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}; Flags: nowait postinstall skipifsilent

[Code]
{ RedesignWizardFormBegin } // 不要删除这一行代码。
// 不要修改这一段代码，它是自动生成的。
var
  OldEvent_NextButtonClick: TNotifyEvent;

procedure _NextButtonClick(Sender: TObject); forward;
procedure YesRadioClick(Sender: TObject); forward;

procedure RedesignWizardForm;
begin
  with WizardForm.Bevel do
  begin
    Top := ScaleY(287);
  end;

  with WizardForm.NextButton do
  begin
    OldEvent_NextButtonClick := OnClick;
    OnClick := @_NextButtonClick;
  end;

  with WizardForm.WizardBitmapImage do
  begin
    Width := ScaleX(173);
    Height := ScaleY(286);
  end;

  with WizardForm.WizardBitmapImage2 do
  begin
    Height := ScaleY(314);
  end;

  with WizardForm.YesRadio do
  begin
    OnClick := @YesRadioClick;
  end;

{ ReservationBegin }
  // 这一部分是提供给你的，你可以在这里输入一些补充代码。

{ ReservationEnd }
end;
// 不要修改这一段代码，它是自动生成的。
{ RedesignWizardFormEnd } // 不要删除这一行代码。

procedure YesRadioClick(Sender: TObject);
begin

end;

procedure _NextButtonClick(Sender: TObject);
begin
  OldEvent_NextButtonClick(Sender);
end;

procedure InitializeWizard();
begin
  RedesignWizardForm;
end;


