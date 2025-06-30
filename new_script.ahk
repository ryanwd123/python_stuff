#Requires AutoHotkey v2.0
#SingleInstance Force
^!r:: {
    MsgBox("Reloaded")
    Reload  
}
; ! alt
; ^ ctrl
; + shift
; # win

LVN_KEYDOWN := -155 
global keypadmode := true

!F10::{
    global keypadmode := !keypadmode
}




^+v::{
    A_Clipboard := A_Clipboard
    Sleep(100)
    send '^v'
}

; !v:: {
;     wins := GetWinInfo()
;     for w in wins {
;         if (w.proc = "code.exe")
;             WinActivate(w.id)
;     }
; }

getActive(){
    aid := WinExist("A")
    t := WinGetTitle(aid)
    p := WinGetProcessName(aid)
    if (t = "") 
        return
    if (process_ignore.Has(p))
        return
    if (title_ignore.Has(t))
        return
    return aid
}


Media_Play_Pause::{
    aid := getActive()
    if (!aid)
        return
    min_max := WinGetMinMax(aid)
    try {
        if (min_max)
            WinRestore(aid)
        WinMove(0,0,1920,1080,aid)
    }
}

getLeftOrRight(x, aid) {
    all_wins := GetWinInfo()
    QuickSortWindows(all_wins, 1, all_wins.Length, (winA, winB) => winA.x - winB.x)
    wins := []
    for w in all_wins {
        if (w.min_max != -1)
            wins.Push(w)
    }

    for w in wins {
        if (w.id = aid) {
            if (wins.Has(A_Index + x)) {
                return wins[A_Index + x].id
            }
        }
    }
    return 0
}
activateLeftorRight(x) {
    aid := "ahk_id " WinExist("A")
    w := getLeftOrRight(x, aid)
    if (w)
        WinActivate(w)

}


swapLeftOrRight(x) {
    aid := "ahk_id " WinExist("A")
    WinGetPos &aX, &aY, &aW, &aH, aid
    tid := getLeftOrRight(x, aid)
    if (tid)
        try {
            WinGetPos &tX, &tY, &tW, &tH, tid
            WinMove(aX,aY,,,tid)
            WinMove(tX,tY,,,aid)
        }
}
!m:: WinMinimize(WinExist("A"))
; !n:: GetWindowSelector()

; Media_Play_Pause::return
; Numpad7::return



#HotIf  keypadmode
Numpad5::return
; Numpad6::return
NumpadAdd::return
Numpad4::return
Numpad8::return
Numpad9::return
NumpadSub::return
; Media_Prev::return
Media_Next::GetWindowSelector()
Numpad7::cascaseWindows()
Numpad1::move_active_win(-A_ScreenWidth*.05,  0)  ;left
Numpad2::move_active_win(  0, A_ScreenHeight*.05)  ;down
Numpad3::move_active_win(  0,-A_ScreenHeight*.05)  ;up
NumpadEnter::move_active_win( A_ScreenWidth*.05,  0)    ;right
!Numpad1::resize_active_win(-A_ScreenWidth*.05,  0)  ;left
!Numpad2::resize_active_win(  0, A_ScreenHeight*.05)   ;down
!Numpad3::resize_active_win(  0,-A_ScreenHeight*.05)   ;up 
!NumpadEnter::resize_active_win( A_ScreenWidth*.05,  0)   ;right
Media_Prev:: activateLeftorRight(-1)
Numpad6:: activateLeftorRight(1)
!Media_Prev:: swapLeftOrRight(-1)
!Numpad6:: swapLeftOrRight(1)
; Numpad1::move_active_win(-A_ScreenWidth*.05,  0)  ;left
; Numpad2::move_active_win(  0, A_ScreenHeight*.05)  ;down
; Numpad3::move_active_win(  0,-A_ScreenHeight*.05)  ;up
; NumpadEnter::move_active_win( A_ScreenWidth*.05,  0)    ;right
; !Numpad1::resize_active_win(-A_ScreenWidth*.05,  0)  ;left
; !Numpad2::resize_active_win(  0, A_ScreenHeight*.05)   ;down
; !Numpad3::resize_active_win(  0,-A_ScreenHeight*.05)   ;up 
; !NumpadEnter::resize_active_win( A_ScreenWidth*.05,  0)   ;right
; Numpad4:: activateLeftorRight(-1)
; NumpadAdd:: activateLeftorRight(1)
; !Numpad4:: swapLeftOrRight(-1)
; !NumpadAdd:: swapLeftOrRight(1)
#HotIf 


move_active_win(x,y){
    aid := getActive()
    if (!aid)
        return
    try {
        WinGetPos &aX, &aY, &aW, &aH, aid
        new_y := Min(Max(aY + y,0), A_ScreenHeight-aH-30)
        new_x := Min(Ax +x, A_ScreenWidth - aW)
        WinMove new_x, new_y, , ,aid
    }
}

resize_active_win(x,y){
    aid := getActive()
    if (!aid)
        return
    try {
        WinGetPos &aX, &aY, &aW, &aH, aid
        WinMove , ,aW + x,ah +y ,aid
    }
}




GetWindowSelector() {
    if (WinExist('window_selector')) {
        WinActivate('window_selector')
    } else {
        MakeWindowSelector()
    }
}

MakeWindowSelector() {

    wins := GetWinInfo()
    QuickSortWindows(wins, 1, wins.Length, (winA, winB) => winA.rid - winB.rid)
    QuickSortWindows(wins, 1, wins.Length, CompareByProcess)
    win_map := Map()
    for w in wins {
        win_map[w.id] := w
    }

    g := Gui("-Caption +ToolWindow")
    g.Title := 'window_selector'
    il := IL_Create(wins.Length)
    lv := g.AddListView("r30 w500 -Hdr", ["window", "state", "x",  "id", "process"])
    lv.SetImageList(il)
    lv.SetFont("s12", "Arial")

    for w in wins {
        icon := GetIconIndex(w.exePath, il)
        lv.Add(
            "Icon" . icon, 
            w.title,
            w.min_max = -1? "min":"",
            "",
            w.id,
            w.proc
        )

    }

    lv.ModifyCol(1, "390")
    lv.ModifyCol(2, " 50")
    lv.ModifyCol(3, "50")
    lv.ModifyCol(4, "0")
    lv.ModifyCol(5, "0")

    if (lv.GetCount() > 0) {
        lv.Modify(1, "Select Focus")
    }

    lv.Focus()
    g.OnEvent("Escape", g.Destroy)  ; Close GUI with Escape key
    btn := g.AddButton("Default w80", "select")
    btn.OnEvent("Click", SelectWindow)  ; Call MyBtn_Click when clicked.
    Send "{RShift}"
    g.Show()

    lv.OnNotify(LVN_KEYDOWN, LV_KeyHandler)

    LV_KeyHandler(ctrl, lParam) {
        offsetVK  := A_PtrSize*2 + 4            ; hwndFrom + idFrom + code
        offsetFlg := offsetVK + 4               ; next DWORD
        ; flags := NumGet(lParam, offsetVK , "UShort")
        vkCode  := NumGet(lParam, offsetFlg, "UInt")
        
        keyName := GetKeyName("VK" Format("{:X}", vkCode))
        ; MsgBox "k: " keyName   "  vkCode:" vkCode  ;  "  f:" flags  
        i := lv.GetNext()
        id := lv.GetText(i,4)
        state := lv.GetText(i,2)
        proc := lv.GetText(i,5)


        if (keyName = "left" and state != "min"){
            if (i > 0) {
                WinMinimize(id)
                lv.Modify(i, , , "min")
                WinActivate(g.Hwnd)
                return
            }
        }
        if (keyName = "Delete" and proc != "EXCEL.EXE"){
            if (i > 0 and id != "") {
                WinClose(id)
                lv.Modify(i, "", "", "", "", "", "")
                return
            }
        }
        if (keyName = "right"){
            if (i > 0) {
                WinRestore(id)
                lv.Modify(i, , , "")
                WinActivate(g.Hwnd)
                return
            }
        }


        return 0
    }

    SelectWindow(*) { 
        i := lv.GetNext()
        if (i > 0) {
            id := lv.GetText(i, 4)
            if WinExist(id)
                WinActivate(id)
        }
        g.Destroy()
    }
}

GetIconIndex(path, il) {
    for type in [1, 3] {
        try {
            idx := IL_Add(il, path, type)
            if (idx)
                return idx
        } catch {
            ; ignore and try next
        }
    }
    return IL_Add(il, "shell32.dll", 3)
}




cascaseWindows() {
    wins := GetWinInfo()
    QuickSortWindows(wins, 1, wins.Length, (winA, winB) => winA.rid - winB.rid)
    QuickSortWindows(wins, 1, wins.Length, CompareByProcess)
    monitorCount := MonitorGetCount()
    monitorPrimary := MonitorGetPrimary()
    MonitorGetWorkArea monitorPrimary, &pL, &pT, &pR, &pB
    win_stacks := Map()
    win_stacks[1] := []
    win_stacks[2] := []
    win_stacks[3] := []
    x_y := Map()
    x_y[1] := [10,10, A_ScreenWidth * 0.65, (pB - pT)* 0.9, pb]
    x_y[2] := [A_ScreenWidth * .6 ,10, A_ScreenWidth * 0.35, (pB - pT)* 0.8, pb]
    x_y[3] := [10,10, A_ScreenWidth * 0.65, (pB - pT)* 0.9, pb]
    if (monitorCount > 1) {
        if (monitorPrimary = 1) {
            s2ndMon := 2 
        } else {
            s2ndMon := 1
        }
        MonitorGetWorkArea s2ndMon, &sL, &sT, &sR, &sB
        x_y[3] := [sL + 10,sT + 10, (sR-sL)*0.8, (sB-sT)*0.8, sB]
    }
    stacks := [3,2,1]

    for w in wins {
        if (w.proc = "Everything.exe")
            continue
        stack := StackMap.Get(w.proc,1)
        if (monitorCount = 1 and stack > 2)
            stack := 1
        win_stacks[stack].Push(w)
    }

    for s in stacks {
        stack_wins := win_stacks[s]
        for w in stack_wins {
            new_x := x_y[s][1]
            new_y := x_y[s][2]
            new_w := x_y[s][3]
            target_h := x_y[s][4]
            area_bottom := x_y[s][5]
            max_h := area_bottom - new_y
            new_h := Min(max_h, target_h)

            try {
                WinMove(new_x, new_y, new_w, new_h, w.id)
                ; wina(w.id)
                x_y[s][1] += 20
                x_y[s][2] += 25
                ; x_y[s][3]
                ; x_y[s][4]

            }
        }
    }
}

monitorInfo() {
    monitorCount := MonitorGetCount()
    monitorPrimary := MonitorGetPrimary()
    MonitorGetWorkArea 1, &pL, &pT, &pR, &pB
    MonitorGetWorkArea 2, &sL, &sT, &sR, &sB
    txt := (
"monitor count    " monitorCount  "`r`n"
"primary montior  "  monitorPrimary  "`r`n"
"prmary work area " pL ", " pT ", " pR ", " pB  "`r`n" 
"2nd work area    " sL ", " sT ", " sR ", " sB  "`r`n" 
    )
    MsgBox(txt)
}



StackMap := Map(
    "WindowsTerminal.exe", 2,
    "explorer.exe", 2,
    "notepad++.exe", 3
)

ProcessPriority := Map(
    "Code.exe", 1,
    "EXCEL.EXE", 2,
    "chrome.exe", 3,
    "msedge.exe", 4,
    "WindowsTerminal.exe", 5,
    "explorer.exe", 6
)

process_ignore := Map()
process_ignore["Taskmgr.exe"] := true

title_ignore := Map()
title_ignore["Program Manager"] := true
title_ignore["Aqua Voice"] := true
title_ignore["Window Spy for AHKv2"] := true

class MyWin {
    __New(id, title, proc) {
        WinGetPos &X, &Y, &W, &H, id
        min_max := WinGetMinMax(id)

        this.id := "ahk_id " id
        this.rid := id
        this.title := StrSplit(title, " - ")[1]
        this.proc := proc
        this.x := X
        this.y := Y
        this.width := W
        this.height := H
        this.min_max := min_max

        try {
            this.exePath := ProcessGetPath(WinGetPID(id))
        } catch {
            this.exePath := ""
        }
    }
}

GetWinInfo() {
    windows := []
    for id in WinGetList() {
        title := WinGetTitle(id)

        if (title_ignore.Has(title) || title = "")
            continue

        proc := WinGetProcessName(id)

        if (process_ignore.Has(proc))
            continue

        windows.Push(MyWin(id, title, proc))
    }
    return windows
}

QuickSortWindows(arr, low, high, compareFunc) {
    if (low < high) {
        pi := PartitionWindows(arr, low, high, compareFunc)
        QuickSortWindows(arr, low, pi - 1, compareFunc)
        QuickSortWindows(arr, pi + 1, high, compareFunc)
    }
}

PartitionWindows(arr, low, high, compareFunc) {
    pivot := arr[high]
    i := low - 1

    for j in Range(low, high - 1) {
        comparison := compareFunc.Call(arr[j], pivot)
        if (comparison <= 0) {
            i++
            temp := arr[i]
            arr[i] := arr[j]
            arr[j] := temp
        }
    }

    temp := arr[i + 1]
    arr[i + 1] := arr[high]
    arr[high] := temp

    return i + 1
}

Range(start, end) {
    result := []
    loop end - start + 1 {
        result.Push(start + A_Index - 1)
    }
    return result
}

CompareByProcess(winA, winB) {
    proc_a := ProcessPriority.Get(WinA.proc, 999)
    proc_b := ProcessPriority.Get(WinB.proc, 999)

    if (proc_a < proc_b)
        return -1
    else if (proc_a > proc_b)
        return 1
    else
        return 0
}

