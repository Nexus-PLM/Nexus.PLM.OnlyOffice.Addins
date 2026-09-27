; Inno Setup script for the Nexus PLM plugin for ONLYOFFICE Desktop Editors.
;
; There is no build step. The plugin is JavaScript with no bundler, so what ships is what is in
; `plugin\` — this script copies that folder and writes one file that cannot be in source control.
; Compile it with:
;
;   "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" installer\Nexus.PLM.OnlyOffice.Addin.iss
;
; PER-USER, NO ELEVATION. ONLYOFFICE reads plugins from two places: one under Program Files, which
; needs admin rights, and one under %LOCALAPPDATA%, which does not. This installs into the second.
; That is not only about elevation: the user folder has its own `v1`, so the `../v1/plugins.js`
; every plugin page loads still resolves. Installing into the Program Files copy would need
; elevation AND would be undone by the next ONLYOFFICE update.
;
; THE FOLDER NAME IS LOAD-BEARING. ONLYOFFICE matches a plugin folder to its config by the guid,
; without the "asc." prefix the config carries. Install it under any other name and the editor
; either ignores it or registers it twice. `tests\installer.test.js` holds the name here against
; the guid in `plugin\config.json`, because nothing about getting this wrong fails loudly.
;
; WHAT THIS DOES NOT INSTALL: the Nexus PLM tray application, which hosts the Addin Service the
; plugin talks to on localhost. It comes from Nexus.PLM.WPF.Addins. This installer deliberately
; does not carry it — two installers writing the same files is how one of them silently wins — and
; the plugin says plainly when the service is not running.

#define AppName "Nexus PLM for ONLYOFFICE"
#define AppPublisher "NexusPLM"

; Keep in step with plugin\config.json. `tests\installer.test.js` fails when they drift.
#define AppVersion "0.2.0"

; The guid from plugin\config.json with its "asc." prefix removed. Doubled leading brace because a
; single "{" opens an Inno constant; "{{" is a literal one.
#define PluginFolder "{{A8F3C512-7D64-4E1B-9C2A-6B0E5D831F47}"

#define PluginsRoot "{localappdata}\ONLYOFFICE\DesktopEditors\data\sdkjs-plugins"
#define PluginTarget PluginsRoot + "\" + PluginFolder

; Where the Addin Service keeps this machine's per-install secret. Written by the tray
; application, read here, and never transmitted anywhere.
#define SecretSource "{localappdata}\NexusPLM\browser-addin-secret.txt"

[Setup]
AppId={{3F7A9D25-4C18-4B6E-A031-8E52C7D4B9A6}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Nothing is installed here — the plugin goes into ONLYOFFICE's own folder — but Inno wants an
; application directory for its uninstall record, so it gets one it will not fill.
DefaultDirName={localappdata}\Programs\Nexus PLM ONLYOFFICE Addin
DefaultGroupName=Nexus PLM
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=NexusPlmOnlyOfficeAddinSetup
OutputDir=Output
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; ONLYOFFICE reads the plugin list once, at startup. Installing underneath a running editor leaves
; the user looking at a window that will never show the tab, and concluding the installer failed.
CloseApplications=yes
RestartApplications=no
UninstallDisplayName={#AppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Everything in plugin\, which is the whole plugin. `excludes` keeps out the one file that must
; never come from source control: secret.js is written below, per machine, from this machine's own
; secret. A secret.js checked in and shipped would hand every seat the same proof of locality.
Source: "..\plugin\*"; DestDir: "{#PluginTarget}"; \
  Excludes: "secret.js"; Flags: recursesubdirs createallsubdirs ignoreversion

[UninstallDelete]
; The plugin folder only. Its parent holds every other ONLYOFFICE plugin, including the ones that
; ship with the editor.
Type: filesandordirs; Name: "{#PluginTarget}"

[Code]

var
  SecretWasWritten: Boolean;

// ONLYOFFICE creates this on first run. Absent, either the editor is not installed or it has
// never been started — and a plugin dropped into a folder nothing reads looks like a failed
// install with no error. Warn, and let the user decide: installing ONLYOFFICE afterwards and
// starting it picks the plugin up, so this is a caution rather than a refusal.
function InitializeSetup(): Boolean;
var
  EditorData: String;
begin
  Result := True;
  EditorData := ExpandConstant('{localappdata}\ONLYOFFICE\DesktopEditors');
  if not DirExists(EditorData) then
    Result := MsgBox(
      'ONLYOFFICE Desktop Editors was not found for this user.' + #13#10#13#10 +
      'Its data folder does not exist yet, which means it is either not installed or has never ' +
      'been started. The plugin can still be installed now and will appear the first time ' +
      'ONLYOFFICE runs.' + #13#10#13#10 +
      'Continue?',
      mbConfirmation, MB_YESNO) = IDYES;
end;

// The plugin's page is loaded from disk, so its origin is opaque, and the Addin Service will not
// trust an opaque origin on its own: any website can obtain one through a sandboxed iframe. The
// secret is the proof that this caller really is on this machine — a page loaded from disk can
// read a file beside itself, a remote site cannot read the user's disk.
//
// Written as a script rather than as data because a page with an opaque origin may not fetch a
// file:// URL, but will load a sibling <script> quite happily.
procedure WriteSecretFile();
var
  SecretPath, Secret: String;
  // LoadStringFromFile and SaveStringToFile deal in AnsiString even in the Unicode compiler.
  // The secret is hex and the file written beside it is ASCII, so nothing is lost crossing over.
  Raw, Contents: AnsiString;
begin
  SecretWasWritten := False;
  SecretPath := ExpandConstant('{#SecretSource}');
  if not FileExists(SecretPath) then
    Exit;
  if not LoadStringFromFile(SecretPath, Raw) then
    Exit;

  Secret := Trim(String(Raw));
  if Secret = '' then
    Exit;

  Contents :=
    '/* Written by the Nexus PLM installer from this machine''s Addin Service secret. Not source,' + #13#10 +
    '   and not secret in any cryptographic sense - it never leaves this machine. It exists so the' + #13#10 +
    '   service can tell a local plugin from a website that forged an opaque origin. */' + #13#10 +
    'window.NexusPlmSecret = "' + Secret + '";' + #13#10;

  SecretWasWritten := SaveStringToFile(
    ExpandConstant('{#PluginTarget}\secret.js'), Contents, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    WriteSecretFile();
    // Said plainly, because without it the plugin installs, draws its tab, and answers 403 to
    // every command — which reads as a broken add-in rather than a missing step.
    if not SecretWasWritten then
      MsgBox(
        'The plugin is installed, but Nexus PLM has not yet created this machine''s secret.' + #13#10#13#10 +
        'Install the Nexus PLM tray application and start it once, then run this installer ' +
        'again. Until then the Nexus PLM tab will appear but every command will be refused.',
        mbInformation, MB_OK);
  end;
end;
