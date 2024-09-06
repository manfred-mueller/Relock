; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Relock"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "NASS e.K."
#define MyAppURL "https://www.nass-ek.de"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{E220DEF6-BDC0-4057-8F4E-CA6CDF71E94D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=D:\Dokumente\lizenz-ado.txt
; Uncomment the following line to run in non administrative install mode (install for current user only.)
PrivilegesRequired=lowest
OutputDir=bin\Release
OutputBaseFilename={#MyAppName}_setup-{#MyAppVersion}
SetupIconFile=D:\Bilder\nass-ek.ico
UninstallDisplayIcon={uninstallexe}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SignTool=CertumVS

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "bin\Release\Relock.exe"; DestDir: "{userpf}"; DestName: "Relock.exe"; Flags: confirmoverwrite; MinVersion: 0,6.0sp2

[Run]
Filename: "{userpf}\Relock.exe"; Parameters: "/register"; Flags: postinstall

[UninstallRun]
Filename: "{userpf}\Relock.exe"; Parameters: "/unregister"

[UninstallDelete]
Type: files; Name: "{userpf}\relock.exe"