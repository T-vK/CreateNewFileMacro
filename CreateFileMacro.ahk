Menu, Tray, Icon, Imageres.dll, 3 ; new file icon
TrayTip, CreateFileMacro, Ctrl+Shift+M to create a new file in the current explorer window
#If WinActive("ahk_class CabinetWClass") || WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW")
    ^+m::CurrentWindowNewFile() ;Ctrl+Shift+M Create new, empty, entensionless file in the active explorer window or on the desktop
#If

CurrentWindowNewFile() {
    shellWindows := ComObjCreate("Shell.Application").Windows
    If (WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW")) { 
        VarSetCapacity(_hwnd, 4, 0)
        explorerWin := shellWindows.FindWindowSW(0, "", 8, ComObj(0x4003, &_hwnd), 1)
    } Else {
        hwndActive := HexToDecimal(WinExist())
        Loop % shellWindows.Count+1 { ; +1?
	        explorerWin := shellWindows.Item(A_Index-1)
            If (hwndActive = explorerWin.Hwnd)
                Break
        }
    }
    sfv := explorerWin.Document
    items := sfv.SelectedItems
    Loop % items.Count
        sfv.SelectItem(items.Item(A_Index-1), 0)
    newFileName := "New File"
    If FileExist(sfv.Folder.Self.Path "\" newFileName) {
        Loop {
            newFileName := "New File (" A_Index+1 ")"
            If !FileExist(sfv.Folder.Self.Path "\" newFileName)
                Break
        }
    }
    FileAppend,, % sfv.Folder.Self.Path "\" newFileName
    explorerWin.Refresh()
    sfv.SelectItem(items.Item(newFileName), 1)
    Loop, 5 { ; Loop usually isn't necessary, but adds a bit of reliability.
        Sleep, 50
        Send {F2}
        ; Verify we're now in rename mode.
        ControlGetFocus focus
    } until (focus = "Edit1")
}

HexToDecimal(str){
    static _0:=0,_1:=1,_2:=2,_3:=3,_4:=4,_5:=5,_6:=6,_7:=7,_8:=8,_9:=9,_a:=10,_b:=11,_c:=12,_d:=13,_e:=14,_f:=15
    str:=ltrim(str,"0x `t`n`r"),   len := StrLen(str),  ret:=0
    Loop,Parse,str
      ret += _%A_LoopField%*(16**(len-A_Index))
    return ret
}
