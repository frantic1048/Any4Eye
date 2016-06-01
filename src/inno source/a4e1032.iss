; 脚本用 Inno Setup 脚本向导 生成。
; 查阅文档获取创建 INNO SETUP 脚本文件的详细资料！

#define MyAppName "Any 4 Eye"
#define MyAppVersion "1.0.44"
#define MyAppPublisher "Frantic Black"
#define MyAppURL "http://blog.163.com/frantic_hao/"
#define MyAppExeName "Any 4 Eye 1.0.44.exe"

[Setup]
WizardImageFile=datiaofu副本.bmp
WizardSmallImageFile=I:\My Documents\VB6\自己的源码\Any4Eye\icon副本.bmp
; 注意: AppId 的值是唯一识别这个程序的标志。
; 不要在其他程序中使用相同的 AppId 值。
; (在编译器中点击菜单“工具 -> 产生 GUID”可以产生一个新的 GUID)
AppId={732C547D-9CC4-4F97-AA40-AFE06B31FDB6}
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
InfoBeforeFile=I:\My Documents\VB6\自己的源码\Any4Eye\EXEs\1.0.44.txt
OutputDir=I:\My Documents\VB6\包
OutputBaseFilename=Any4Eye1032setup
SetupIconFile=I:\My Documents\VB6\自己的源码\Any4Eye\icon.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "default"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
;{ ISFormDesignerFilesBegin } // 不要删除这一行代码。
;// 不要修改这一段代码，它是自动生成的。
DestName: "WizardForm.SelectDirBitmapImage.bmp"; Source: "I:\My Documents\VB6\自己的源码\Any4Eye\folder.bmp"; Flags: dontcopy solidbreak
DestName: "WizardForm.SelectGroupBitmapImage.bmp"; Source: "I:\My Documents\VB6\自己的源码\Any4Eye\folder.bmp"; Flags: dontcopy solidbreak
;// 不要修改这一段代码，它是自动生成的。
;{ ISFormDesignerFilesEnd } // 不要删除这一行代码。

Source: "I:\My Documents\VB6\自己的源码\Any4Eye\EXEs\Any 4 Eye 1.0.44.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "I:\My Documents\VB6\自己的源码\Any4Eye\EXEs\BackGround.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "I:\My Documents\VB6\lib\a4e\comcat.dll"; DestDir: "{sys}"; Flags: ignoreversion noregerror uninsneveruninstall regserver
Source: "I:\My Documents\VB6\lib\a4e\FM20.DLL"; DestDir: "{sys}"; Flags:  noregerror uninsneveruninstall regserver
Source: "I:\My Documents\VB6\lib\a4e\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: noregerror uninsneveruninstall regserver
Source: "I:\My Documents\VB6\lib\a4e\MSMASK32.OCX"; DestDir: "{sys}"; Flags: noregerror uninsneveruninstall regserver
Source: "I:\My Documents\VB6\lib\a4e\MSMSKCHS.DLL"; DestDir: "{sys}"; Flags: noregerror uninsneveruninstall regserver
Source: "I:\My Documents\VB6\lib\a4e\VB6CHS.DLL"; DestDir: "{sys}"; Flags: noregerror uninsneveruninstall regserver


; 注意: 不要在任何共享的系统文件使用 "Flags: ignoreversion"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[ISFormDesigner]
WizardForm=FF0A005457495A415244464F524D003010BB03000054504630F10B5457697A617264466F726D0A57697A617264466F726D0C436C69656E74486569676874034C010B436C69656E74576964746803F1010D4578706C69636974576964746803FB010E4578706C69636974486569676874036A010D506978656C73506572496E636802600A54657874486569676874020C00F10654426576656C05426576656C03546F70031F010B4578706C69636974546F70031F010000F10A544E6577427574746F6E0A4E657874427574746F6E074F6E436C69636B07105F4E657874427574746F6E436C69636B0000F10C544E65774E6F7465626F6F6B0D4F757465724E6F7465626F6F6B00F110544E65774E6F7465626F6F6B506167650B57656C636F6D655061676500F10C544269746D6170496D6167651157697A6172644269746D6170496D61676505576964746803AD0006486569676874031E010D4578706C69636974576964746803AD000E4578706C69636974486569676874031E01000000F110544E65774E6F7465626F6F6B5061676509496E6E65725061676500F10C544E65774E6F7465626F6F6B0D496E6E65724E6F7465626F6F6B00F110544E65774E6F7465626F6F6B506167650D53656C6563744469725061676500F10C544269746D6170496D6167651453656C6563744469724269746D6170496D6167650A4269746D617046696C651436000000493A5C4D7920446F63756D656E74735C5642365CE887AAE5B7B1E79A84E6BA90E7A0815C416E79344579655C666F6C6465722E626D70000000F110544E65774E6F7465626F6F6B506167651453656C656374436F6D706F6E656E74735061676500F10C544E6577436F6D626F426F780A5479706573436F6D626F064865696768740214000000F110544E65774E6F7465626F6F6B506167651653656C65637450726F6772616D47726F75705061676500F10C544269746D6170496D6167651653656C65637447726F75704269746D6170496D6167650A4269746D617046696C651436000000493A5C4D7920446F63756D656E74735C5642365CE887AAE5B7B1E79A84E6BA90E7A0815C416E79344579655C666F6C6465722E626D700000000000F110544E65774E6F7465626F6F6B506167650C46696E6973686564506167650B4578706C69636974546F7002F800F10C544269746D6170496D6167651257697A6172644269746D6170496D6167653206486569676874033A010E4578706C69636974486569676874033A010000F10F544E6577526164696F427574746F6E08596573526164696F074F6E436C69636B070D596573526164696F436C69636B0000000000

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

  with WizardForm.SelectDirBitmapImage do
  begin
    ExtractTemporaryFile('WizardForm.SelectDirBitmapImage.bmp');
    Bitmap.LoadFromFile(ExpandConstant('{tmp}\WizardForm.SelectDirBitmapImage.bmp'));
  end;

  with WizardForm.SelectGroupBitmapImage do
  begin
    ExtractTemporaryFile('WizardForm.SelectGroupBitmapImage.bmp');
    Bitmap.LoadFromFile(ExpandConstant('{tmp}\WizardForm.SelectGroupBitmapImage.bmp'));
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




