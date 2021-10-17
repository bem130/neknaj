install_folder = "C:\neknaj\janken\"
install_name = "atdjanken1.0"

WScript.Echo("インストールフォルダは "+install_folder+" です")

Set objWShell = CreateObject("Wscript.Shell")
set file_system = createObject("Scripting.FileSystemObject")
set xHttp = createobject("Microsoft.XMLHTTP")
set bStrm = createobject("Adodb.Stream")

folder_array = Split(install_folder,"\")
folder_name = ""
for cnt=0 to UBound(folder_array)-1
    folder_name = folder_name + folder_array(cnt)+ "\"
    if file_system.FolderExists(folder_name) = False Then
        file_system.CreateFolder(folder_name)
    end if
next

xHttp.Open "GET", "https://haruki1234.github.io/neknaj/janken_win_vbs/atd_Jankenv1.0.vbs", False
xHttp.Send
with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile install_folder+"\"+install_name+".vbs", 2
    .close
end with
xHttp.Open "GET", "https://haruki1234.github.io/neknaj/janken_win_vbs/n.bmp", False
xHttp.Send
with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile install_folder+"\n.ico", 2
    .close
end with

shortcut_name = objWShell.SpecialFolders("desktop")+"\"+install_name+".lnk"
set sc = objWShell.CreateShortcut(shortcut_name)
sc.TargetPath = install_folder+"\"+install_name+".vbs"
sc.IconLocation = install_folder+"\n.ico"
sc.WorkingDirectory = "C:\neknaj\"
sc.save