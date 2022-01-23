#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


; enable scroll lock at boot (or when this script is first enabled)
if not GetKeyState("ScrollLock", "T"){
    Send, {ScrollLock}
}

;hotkeys

;numlock on
Numpad5::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Play_Pause}
}
else{
    Send, 5
}
return
Numpad6::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Next}
}
else{
    Send, 6
}
return

Numpad4::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Prev}
}
else{
    Send, 4
}
return

Numpad8::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Up 1}
    Sleep, 69
    }
else{
    Send, 8
}
return

Numpad2::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Down 1}
    Sleep, 69
    }
else{
    Send, 2
}
return

NumpadDot::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Mute} 
    }
else{
    Send, .
}
return

Numpad0::
    if GetKeyState("ScrollLock", "T"){    
        toggle_mic_mute()
    }
    else{
    Send, 0
    }
    return

;numlock off
NumpadClear::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Play_Pause}
}
else{
    Send {NumpadClear}
}
return

NumpadRight::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Next}
}
else{
    Send {NumpadRight}
}
return

NumpadLeft::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Prev}
}
else{
    Send {NumpadLeft}
}
return

NumpadUp::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Up 1}
    Sleep, 69
    }
else{
    Send {NumpadUp}
}
return

NumpadDown::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Down 1}
    Sleep, 69
    }
else{
    Send {NumpadDown}
}
return

NumpadDel::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Mute} 
    }
else{
    Send {NumpadDel}
}
return

NumpadIns::
    if GetKeyState("ScrollLock", "T"){    
        toggle_mic_mute()
    }
    else{
    Send, 0
    }
    return

; functions
toggle_mic_mute(){
    ; mutes/unmutes the microphone for specific appliccations.
    ; supported applications: Discord and MS Teams
    
    ; source (for most of this functions code): https://github.com/stajp/Teams_mute_AHK/blob/main/mute_teams_standalone.ahk 
        doAltTab := 0
        SetTitleMatchMode, 2
        if WinExist("Discord") {
            WinGetTitle, active_title, A		 
            if active_title contains Discord		
            {
                Send ^+M		                  
            }
            else	   							 
            {
                SetTitleMatchMode, 2			  
                WinActivate, Discord	  
                Send ^+M		             
                Sleep 200
                Send, {AltDown}{Tab}{AltUp}
                Sleep, 200
            }
        }
        SetTitleMatchMode, 2
        if if WinExist("Microsoft Teams") {       ; Check if MS Teams is currently opened on this pc
            WinGetTitle, active_title, A		  ; looks at the active program on your screen and puts it into variable "active_title"
            if active_title contains Microsoft Teams		; if active program is Teams, move on
            {
                Send ^+M		                  ; send CTRL + SHIFT + M for Teams microphone mute on/off
            }
            else	   							  ; if active program is not Teams then go to Teams 
            {
                SetTitleMatchMode, 2			  ; set Title search to exact
                WinActivate, Microsoft Teams	  ; bring "Microsoft Teams" to the foreground so the shortcut will work
                Send ^+M		                  ; send CTRL + SHIFT + M for Teams microphone mute on/off
                Sleep 200
                Send, {AltDown}{Tab}{AltUp}
                Sleep, 200
            }
        }

        SetTitleMatchMode, 3 ; reset the match mode to an exact match
}