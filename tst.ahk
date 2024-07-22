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

k_wins := ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""] 
k_win_names := ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
k_win_keys := ["y", "u", "i", "o", "p", "h", "j", "k", "l", ";", "n", "m", ",", ".", "/", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x", "c", "v", "b"]
ignore_windows := [
    "Program Manager"
]



listGui := ""

; j & k::send "!{Tab}"
; j::send "{j}"




$y::ActivateOrCapture(1,"{y}")
$u::ActivateOrCapture(2,"{u}")
$i::ActivateOrCapture(3,"{i}")
$o::ActivateOrCapture(4,"{o}")
$p::ActivateOrCapture(5,"{p}")
$h::ActivateOrCapture(6,"{h}")
$j::ActivateOrCapture(7,"{j}")
$k::ActivateOrCapture(8,"{k}")
$l::ActivateOrCapture(9,"{l}")
$;::ActivateOrCapture(10,"{;}")
$n::ActivateOrCapture(11,"{n}")
$m::ActivateOrCapture(12,"{m}")
$,::ActivateOrCapture(13,"{,}")
$.::ActivateOrCapture(14,"{.}")
$/::ActivateOrCapture(15,"{/}")
$q::ActivateOrCapture(16,"{q}")
$w::ActivateOrCapture(17,"{w}")
$e::ActivateOrCapture(18,"{e}")
$r::ActivateOrCapture(19,"{r}")
$t::ActivateOrCapture(20,"{t}")
$a::ActivateOrCapture(21,"{a}")
$s::ActivateOrCapture(22,"{s}")
$d::ActivateOrCapture(23,"{d}")
$f::ActivateOrCapture(24,"{f}")
$g::ActivateOrCapture(25,"{g}")
$z::ActivateOrCapture(26,"{z}")
$x::ActivateOrCapture(27,"{x}")
$c::ActivateOrCapture(28,"{c}")
$v::ActivateOrCapture(29,"{v}")
$b::ActivateOrCapture(30,"{b}")


#HotIf (listGui != "")
Esc::
{
    HideWindowList()
    
}

#HotIf (A_PriorKey = "RShift")
$RShift::
{
    global k_wins
    global k_win_keys
    global k_win_names
    PopulateWindowInfo(&k_wins, &k_win_names, &k_win_keys)
    ShowWindowList()
}


#HotIf (A_PriorKey = "LShift")
$RShift::
{
        global k_wins
        global k_win_keys
        global k_win_names
        PopulateWindowInfo(&k_wins, &k_win_names, &k_win_keys)
}


HasVal(haystack, needle) {
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}

PopulateWindowInfo(&windowIDs, &windowNames, &winKeys, reset := false) {
    ; Get a list of all windows
    windowList := WinGetList()
    
    if (reset) {
        ; Clear all arrays if reset is true
        for i, _ in winKeys {
            windowIDs[i] := ""
            windowNames[i] := ""
        }
    }
    
    ; First, check existing windows and update their titles
    for i, id in windowIDs {
        if (id != "" && WinExist("ahk_id " . id)) {
            windowNames[i] := WinGetTitle("ahk_id " . id)
        } else {
            windowIDs[i] := ""
            windowNames[i] := ""
        }
    }
    
    ; Then, add new windows to empty slots
    for windowID in windowList {
        if (windowID == "") {
            continue
        }
        if (HasVal(windowIDs, windowID) != 0) {
            continue
        }

        title := WinGetTitle("ahk_id " . windowID)
        if (HasVal(ignore_windows, title) != 0) {
            continue
        }


        if (title != "") {
            ; Find an empty slot
            emptyIndex := 0
            for i, id in windowIDs {
                if (id == "") {
                    emptyIndex := i
                    break
                }
            }
            
            ; If an empty slot is found, add the window
            if (emptyIndex > 0) {
                windowIDs[emptyIndex] := windowID
                windowNames[emptyIndex] := title
            }
        }
    }
}



ShowWindowList() {
    global listGui
    if (listGui == "") {
        listGui := Gui("-Caption +ToolWindow +AlwaysOnTop +Resize")
        listGui.BackColor := "FFFFFF"  ; White background
        listGui.SetFont("s14", "Arial")
        
        listBox := listGui.Add("ListView", "w1200 h900 -Hdr -E0x200", ["Key", "Window Name"])
        listBox.Opt("+Grid -Multi")
        
        for index, key in k_win_keys {
            listBox.Add(, key, k_win_names[index])
        }
        
        listBox.ModifyCol(1, 50)
        listBox.ModifyCol(2, 1100)
        
        ; Add GuiResize event to handle window resizing
        ; listGui.OnEvent("Size", (*) => GuiResize(listGui))
        
        listGui.Show("w1200 h900")
    } else {
        listGui.Destroy()
        listGui := ""
    }
}

HideWindowList() {
    global listGui
    if (listGui != "") {
        listGui.Destroy()
        listGui := ""
    }
}



ActivateOrCapture(slot,key)
{
    global k_wins
    global capture_mode
    if (A_PriorKey = "RShift") {
            HideWindowList()
            if (k_wins[slot] != "") {

            if (WinExist("ahk_id " k_wins[slot])) {
                WinActivate("ahk_id " k_wins[slot])
            } else {
                PopulateWindowInfo(&k_wins, &k_win_names, &k_win_keys)
            }
        }
    } else {
        send key
    }
}




; capture_mode := false

; dot := ""

; ShowRedDot() {
;     ; Create a GUI window for the dot if it doesn't exist
;     global dot
;     if (dot == "") {
;         dot := Gui()
;         dot.Opt("+AlwaysOnTop -Caption +ToolWindow")
;         dot.BackColor := "Red"

;         ; Set the size and position of the dot
;         dotSize := 20
;         screenWidth := A_ScreenWidth
;         screenHeight := A_ScreenHeight
;         xPos := 0
;         yPos := 0

;         dot.Show(Format("x{} y{} w{} h{}", xPos, yPos, dotSize, dotSize))
;     } else {
;         ; If the dot already exists, just show it
;         dot.Show("NoActivate")
;     }
; }

; HideRedDot() {
;     global dot
;     if (dot != "") {
;         dot.Destroy()
;         dot := ""
;     }
; }

; Call the function to show the red dot for 2 seconds

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

