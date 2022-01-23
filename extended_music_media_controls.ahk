#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;this file is optimized specifically for the media keys on the logitech external keyboard 

;supportive functions
;set SSID variables
Runwait %comspec% /c netsh wlan show interface | clip,,hide
mySSID:=RegExReplace(clipboard, "s).*?\R\s+SSID\s+:(\V+).*", "$1")
SSIDlist:=  ["wifi_network1", " wifi_network2"]

hasValue(haystack, needle) {
    if(!isObject(haystack))
        return false
    if(haystack.Length()==0)
        return false
    for k,v in haystack
        if(v==needle)
            return true
    return false
}

; enable scroll lock at boot (or when this script is first enabled)
if not GetKeyState("ScrollLock", "T"){
    Send, {ScrollLock}
}


;hotkeys
;numlock on


;  Numpad1::
;     ;from: https://www.autohotkey.com/boards/viewtopic.php?t=49980
;     ; http://www.daveamenta.com/2011-05/programmatically-or-command-line-change-the-default-sound-playback-device-in-windows-7/
;     Devices := {}
;     IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

;     ; IMMDeviceEnumerator::EnumAudioEndpoints
;     ; eRender = 0, eCapture, eAll
;     ; 0x1 = DEVICE_STATE_ACTIVE
;     DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
;     ObjRelease(IMMDeviceEnumerator)

;     ; IMMDeviceCollection::GetCount
;     DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
;     Loop % (Count)
;     {
;         ; IMMDeviceCollection::Item
;         DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")

;         ; IMMDevice::GetId
;         DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
;         DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

;         ; IMMDevice::OpenPropertyStore
;         ; 0x0 = STGM_READ
;         DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
;         ObjRelease(IMMDevice)

;         ; IPropertyStore::GetValue
;         VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
;         VarSetCapacity(PROPERTYKEY, 20)
;         DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
;         NumPut(14, &PROPERTYKEY + 16, "UInt")
;         DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
;         DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
;         DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal
;         ObjRelease(IPropertyStore)

;         ObjRawSet(Devices, DeviceName, DeviceID)
;     }
;     ObjRelease(IMMDeviceCollection)


;     Devices2 := {}
;     For DeviceName, DeviceID in Devices
;         List .= "(" . A_Index . ") " . DeviceName . "`n", ObjRawSet(Devices2, A_Index, DeviceID)
;     InputBox n,, % List,,,,,,,, 1

;     MsgBox % Devices2[n]

;     ;IPolicyConfig::SetDefaultEndpoint
;     IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
;     R := DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "Str", Devices2[n], "UInt", 0, "UInt")
;     ObjRelease(IPolicyConfig)
;     MsgBox % Format("0x{:08X}", R)

;     return



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
    ; mic muting source: https://github.com/stajp/Teams_mute_AHK/blob/main/mute_teams_standalone.ahk 
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
    if if WinExist("Microsoft Teams") {
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
    ; mic muting source: https://github.com/stajp/Teams_mute_AHK/blob/main/mute_teams_standalone.ahk 
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
    if if WinExist("Microsoft Teams") {
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
    return

;hotkeys logitech ultraX
Browser_Favorites::
   if (hasValue(SSIDlist, mySSID))
       {
           		Send {Media_Prev}
       }
       else{
           Send {Browser_Favorites}
       }
return

Launch_Mail::
   if (hasValue(SSIDlist, mySSID))
       {
		Send {Media_Next}
       }
       else{
           Send {Launch_Mail}
       }
return

Launch_Media::
   if (hasValue(SSIDlist, mySSID))
       {
           while GetKeyState("Launch_Media")
           {
               Send {Volume_Down 1}
               Sleep, 69
           }
       }
       else{
           Send {Launch_Media}
       }
return

Browser_Home::
   if (hasValue(SSIDlist, mySSID))
   {
       while GetKeyState("Browser_Home")
       {
           Send {Volume_Up 1}
           Sleep, 69
       }
   }
   else{
       Send {Browser_Home}
   }
return




