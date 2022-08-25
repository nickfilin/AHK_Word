Explorer_GetSelection(hwnd="") {
        WinGet, process, processName, % "ahk_id" hwnd := hwnd? hwnd:WinExist("A")
        WinGetClass class, ahk_id %hwnd%
        if (process = "explorer.exe")
            if (class ~= "Progman|WorkerW") {
                ControlGet, files, List, Selected Col1, SysListView321, ahk_class %class%
                Loop, Parse, files, `n, `r
                ToReturn .= A_Desktop "\" A_LoopField "`n"
        } else if (class ~= "(Cabinet|Explore)WClass") {
            for window in ComObjCreate("Shell.Application").Windows
                if (window.hwnd==hwnd)
                    sel := window.Document.SelectedItems
            for item in sel
                ToReturn .= item.path "`n"
        }
    return Trim(ToReturn,"`n")
}

