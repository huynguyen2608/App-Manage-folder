; Manage Folder Installer Script

[Setup]
AppName=Manage Folder
AppVersion=1.0
DefaultDirName={pf}\Manage Folder
DefaultGroupName=Manage Folder
; THAY ĐỔI 1: Đặt icon gỡ cài đặt là icon của ứng dụng.
UninstallDisplayIcon={app}\Manage Folder.exe 
OutputDir=installer
OutputBaseFilename=ManageFolder
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; --- ICON CÀI ĐẶT ---
; SỬ DỤNG FILE ICON ĐÃ CHUYỂN ĐỔI
SetupIconFile=folder_icon.ico 

[Files]
; Thêm icon vào thư mục ứng dụng để EXE có thể truy cập nó, 
; mặc dù icon này đã được nhúng trong bước build Python.
; Nếu icon này không được nhúng, đây là nơi bạn cần thêm nó.
Source: "dist\Manage Folder.exe"; DestDir: "{app}"; Flags: ignoreversion
; THÊM: Tùy chọn, nếu bạn muốn đảm bảo tệp icon nằm trong thư mục cài đặt.
; Source: "folder_icon.ico"; DestDir: "{app}"; Flags: hidden

[Icons]
; THAY ĐỔI 2: Sử dụng tham số IconFilename để chỉ định icon tùy chỉnh
; cho các shortcut. Icon này thường sẽ là icon NHÚNG SẴN trong EXE 
; (được đặt trong bước build PyInstaller/py2exe).
Name: "{group}\Manage Folder"; Filename: "{app}\Manage Folder.exe"; IconFilename: "{app}\Manage Folder.exe"
Name: "{commondesktop}\Manage Folder"; Filename: "{app}\Manage Folder.exe"; IconFilename: "{app}\Manage Folder.exe"

[Run]
Filename: "{app}\Manage Folder.exe"; Description: "Launch Manage Folder"; Flags: nowait postinstall skipifsilent