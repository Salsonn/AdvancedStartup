#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
full_command_line := DllCall("GetCommandLine", "str")
if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    try
    {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
            Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
    }
    ExitApp
}
shellMeans := "C:\Users\caleb\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
iconList := IL_Create(1, 1, false)
Gui, Add, ListView, r36 w1000 gStartupItems AltSubmit, Name|Command Line|Miscellaneous Information
Gui, Add, Button, Default w80 gloadItems, Refresh
Gui, Add, Button, Default w80 x+m gbtnDelete, Delete
Gui, Add, Button, Default w80 x+m gbtnOpenFolder Default, Open Folder

iconListDefault := IL_Create(10)                         ;
iconListLarge := IL_Create(10, 10, true)                 ;IconList Creation and Configuration
LV_SetImageList(iconListDefault)                         ;
LV_SetImageList(iconListLarge)                           ;
sfi_size := A_PtrSize + 8 + (A_IsUnicode ? 680 : 340)    ;
VarSetCapacity(sfi, sfi_size)

loadItems:                                               ;Populates the StartupItems ListView
LV_Delete()
root := "HKEY_LOCAL_MACHINE"
path := "Software\Microsoft\Windows\CurrentVersion\Run"
gosub loadRegistry
root := "HKEY_LOCAL_MACHINE"
path := "Software\Microsoft\Windows\CurrentVersion\RunOnce"
gosub loadRegistry
root := "HKEY_LOCAL_MACHINE"
path := "Software\Microsoft\Windows\CurrentVersion\RunServices"
gosub loadRegistry
root := "HKEY_LOCAL_MACHINE"
path := "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
gosub loadRegistry
root := "HKEY_CURRENT_USER"
path := "Software\Microsoft\Windows\CurrentVersion\Run"
gosub loadRegistry
root := "HKEY_CURRENT_USER"
path := "Software\Microsoft\Windows\CurrentVersion\RunOnce"
gosub loadRegistry
root := "HKEY_CURRENT_USER"
path := "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
gosub loadRegistry
path := "C:\Users\caleb\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
gosub loadFolder
path := "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
gosub loadFolder
LV_ModifyCol()
Gui, Show
return

btnDelete:
ControlGet, rowText, List, selected, SysListView321, a
Loop, Parse, rowText, %A_Tab% 
    column%A_Index% := A_LoopField
IfInString, column3, HKEY
{
    RegDelete, %column3%, %column1%
}
else
{
    delPath := StrReplace(column2, "shell:startup", shellMeans)
    FileDelete, %delPath%
}
gosub loadItems
VarSetCapacity(delpath,0)
VarSetCapacity(rowText,0)
VarSetCapacity(column1,0)
VarSetCapacity(column2,0)
VarSetCapacity(column3,0)
return

btnOpenFolder:
ControlGet, rowText, List, selected, SysListView321, a
Loop, Parse, rowText, %A_Tab% 
    column%A_Index% := A_LoopField
SplitPath, column2,, viewPath, viewExt, viewName
StringReplace, viewExt, viewExt,%A_Space%,&,All
StringReplace, viewExt, viewExt,%A_Space%,&,All
viewExt := RegExReplace(viewExt, "&[^&]+$", "")
viewPath = %viewPath%\%viewName%.%viewExt%
viewPath := "explorer /select, " StrReplace(viewPath, "shell:startup", shellMeans)
Transform, viewPath, Deref, %viewPath%
run, %viewPath%
return

StartupItems:
if (A_GuiEvent = "DoubleClick")
{
    gosub btnOpenFolder
}
if (A_GuiEvent = "K")
{
    key := GetKeyName(Format("vk{:x}", A_EventInfo))
    if (key = "NumpadDel" or key = "Delete")
        gosub btnDelete
}
return

loadRegistry:
Loop, %root%, %path%, kv
{
	RegRead, key
	SplitPath, key, fileName, filePath, fileExt, fileNameNoExt
	StringReplace, fileName, fileName,",,All
	StringReplace, fileExt, fileExt,",,All
	StringReplace, filePath, filePath,",,All
	StringReplace, fileName, fileName,%A_Space%,&,All
	StringReplace, fileExt, fileExt,%A_Space%,&,All
	fileExt := RegExReplace(fileExt, "&[^&]+$", "")
	iconPath = %filePath%\%fileNameNoExt%.%fileExt%
	Transform, iconPath, Deref, %iconPath%
	keyPath = %A_LoopRegKey%\%A_LoopRegSubkey%
	GoSub getIcon
    LV_Add("Icon" . IconNumber, A_LoopRegName, key, keyPath)
	sumLoopIndex++
}
return
loadFolder:
Loop, %path%\*.*
{
	SplitPath, A_LoopFilePath, fileName, filePath, fileExt, fileNameNoExt
	iconPath = %filePath%\%fileNameNoExt%.%fileExt%
	StringReplace, annotatedFilePath, iconPath, %shellMeans%, shell:startup, All
	GoSub getIcon
	FormatTime, creation, %A_LoopFileTimeCreated%, 'Created: 'y/MM/dd' @ 'HH:mm:ss
    LV_Add("Icon" . IconNumber, fileName, annotatedFilePath, creation)
}
return

GuiClose:
GuiEscape:
    ExitApp

;$Esc::
;IfWinActive, Advanced Startup.ahk
;	ExitApp

getIcon:
	if FileExt in EXE,ICO,ANI,CUR
    {
        ExtID := FileExt  ; Special ID as a placeholder.
        IconNumber := 0  ; Flag it as not found so that these types can each have a unique icon.
    }
    else  ; Some other extension/file-type, so calculate its unique ID.
    {
        ExtID := 0  ; Initialize to handle extensions that are shorter than others.
        Loop 7     ; Limit the extension to 7 characters so that it fits in a 64-bit value.
        {
            StringMid, ExtChar, FileExt, A_Index, 1
            if not ExtChar  ; No more characters.
                break
            ExtID := ExtID | (Asc(ExtChar) << (8 * (A_Index - 1))) ; Derive a Unique ID by assigning a different bit position to each character
        }
        IconNumber := IconArray%ExtID%  ; Check if this file extension already has an icon in the ImageLists. If it does, several calls can be avoided and loading performance is greatly improved, especially for a folder containing hundreds of files
    }
    if not IconNumber  ; There is not yet any icon for this extension, so load it.
    {
        if not DllCall("Shell32\SHGetFileInfo" . (A_IsUnicode ? "W":"A"), "Str", iconPath ; Get the high-quality small-icon associated with this file extension:
            , "UInt", 0, "Ptr", &sfi, "UInt", sfi_size, "UInt", 0x101)  ; 0x101 is SHGFI_ICON+SHGFI_SMALLICON
            IconNumber := 9999999  ; Set it out of bounds to display a blank icon.
        else ; Icon successfully loaded.
        {
            hIcon := NumGet(sfi, 0) ; Extract the hIcon member from the structure
            IconNumber := DllCall("ImageList_ReplaceIcon", "Ptr", iconListDefault, "Int", -1, "Ptr", hIcon) + 1 ; Add the HICON directly to the small-icon and large-icon lists.
            DllCall("ImageList_ReplaceIcon", "Ptr", iconListLarge5, "Int", -1, "Ptr", hIcon) ; Below uses +1 to convert the returned index from zero-based to one-based
            DllCall("DestroyIcon", "Ptr", hIcon) ; Now that it's been copied into the ImageLists, the original should be destroyed
            ; IconArray%ExtID% := IconNumber ; Cache the icon to save memory and improve loading performance
        }
    }
return
