#Requires AutoHotkey v2.0
#SingleInstance Force
^!r::Reload  ; Ctrl+Alt+R

; ! alt
; ^ ctrl
; + shift
; # win


#f::RunActivateOrHide("Quick Search2",'"C:/Users/ryanw/.virtualenvs/quick_search-HvDIYltV/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/quick_search/__main__.py"')
#b::RunActivateOrHide("New_Quick_Search", '"C:/Users/ryanw/.virtualenvs/quick_search-HvDIYltV/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/quick_search/quick_search.py"')
; !t::send "!{Tab}"
#t::send "!{Tab}"

LAlt & j::ShiftAltTab
LAlt & k::AltTab


RunOrHide(prog_name,path)
{
    if (hWnd := WinExist(prog_name)) {
        WinActivate("ahk_id " hWnd)
    } else {
        ; Run '"C:/Users/ryanw/.virtualenvs/quick_search-HvDIYltV/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/quick_search/main_sqlite.py"'
        Run path
        ; Run '"C:/Users/ryanw/.virtualenvs/dabbler-F30Abi41/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/dabbler/dabbler/test_files/df_view_working.py"'
        WinWait(prog_name)
        if (!WinActive(prog_name)) {
            WinActivate(prog_name)
        }
    }
    
}


RunActivateOrHide(prog_name,path)
{
    DetectHiddenWindows true
    if (hWnd := WinExist(prog_name)) {
        ; if hidden
        if (!detm_winIsVisible(hWnd)) {
            WinActivate("ahk_id " hWnd)
            WinShow("ahk_id " hWnd)
        } else {
            if (WinActive(prog_name)) {
                WinHide("ahk_id " hWnd)
                Send "!{Tab}"
            } else {
                WinActivate("ahk_id " hWnd)
            }
        }
        ; WinActivate("ahk_id " hWnd)
        
    } else {
        ; Run '"C:/Users/ryanw/.virtualenvs/quick_search-HvDIYltV/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/quick_search/main_sqlite.py"'
        Run path
        ; Run '"C:/Users/ryanw/.virtualenvs/dabbler-F30Abi41/Scripts/pythonw.exe" "c:/Users/ryanw/python_projects/dabbler/dabbler/test_files/df_view_working.py"'
        WinWait(prog_name)
        if (!WinActive(prog_name)) {
            WinActivate(prog_name)
        }
    }
    
    DetectHiddenWindows false
}

WS_VISIBLE                := 0x10000000

detm_winIsVisible(hwnd) {
    try {
      det_winIsVisible := WinGetStyle("ahk_id " . hwnd) & WS_VISIBLE
      return det_winIsVisible
    } catch TargetError as e {
      return ""
  }
}

