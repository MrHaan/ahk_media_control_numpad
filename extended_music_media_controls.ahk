#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; global variables
output_audio_device_number := 1
output_device_id:= None

; enable scroll lock at boot (or when this script is first enabled)
if not GetKeyState("ScrollLock", "T"){
    Send, {ScrollLock}
}

;hotkeys

NumpadClear::   ; numlock off
Numpad5::       ;numlock on
if GetKeyState("ScrollLock", "T"){
    Send {Media_Play_Pause}
}
else{
    Send, 5
}
return

NumpadRight::
Numpad6::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Next}
}
else{
    Send, 6
}
return

NumpadLeft::
Numpad4::
if GetKeyState("ScrollLock", "T"){
    Send {Media_Prev}
}
else{
    Send, 4
}
return

NumpadUp::
Numpad8::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Up 1}
    Sleep, 69
    }
else{
    Send, 8
}
return

NumpadDown::
Numpad2::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Down 1}
    Sleep, 69
    }
else{
    Send, 2
}
return

NumpadDel::
NumpadDot::
if GetKeyState("ScrollLock", "T"){
    Send {Volume_Mute} 
    }
else{
    Send, .
}
return

NumpadIns::
Numpad0::
    if GetKeyState("ScrollLock", "T"){    
        toggle_mic_mute()
    }
    else{
    Send, 0
    }
    return

NumpadPgdn::
Numpad3::
if GetKeyState("ScrollLock", "T"){
    switch_audio_output_device()
}
else{
    Send {NumpadClear}
}
return

; functions
toggle_mic_mute(){
    ; mutes/unmutes the microphone for specific appliccations.
    ; supported applications: Discord and MS Teams
    
    ; This function is inspired by part of the work of Stipe Predanić: https://github.com/stajp/Teams_mute_AHK/blob/main/mute_teams_standalone.ahk 

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
            Send, {AltDown}{Tab}{AltUp}       ; Alt+Tab to re-activate the window that was active when pressing the mute button 
            Sleep, 200                         
        }
    }

    SetTitleMatchMode, 3 ; reset the match mode to an exact match
}

switch_audio_output_device(){
    ; This function is based on Flipeador, from: https://www.autohotkey.com/boards/viewtopic.php?t=49980
    ; similar code for autohotkey v2: https://www.autohotkey.com/boards/viewtopic.php?f=76&t=49980&p=387886#p387886 
    
    ; http://www.daveamenta.com/2011-05/programmatically-or-command-line-change-the-default-sound-playback-device-in-windows-7/

    ; initializing this function
    SetTimer, ProceedChangingSpeakerDevice, Off ;stop the timer if it is already running
    Devices := {}
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    ; IMMDeviceEnumerator::EnumAudioEndpoints ;(call for legacy autohotkey v1.0)
    ; eRender = 0, eCapture, eAll
    ; 0x1 = DEVICE_STATE_ACTIVE
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
    ObjRelease(IMMDeviceEnumerator)

    ; get properties (like name and ID) of currently active audio output devices
    ; IMMDeviceCollection::GetCount ;(call for legacy autohotkey v1.0)
    DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
    Loop % (Count)
    {
        ; IMMDeviceCollection::Item
        DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")

        ; IMMDevice::GetId
        DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
        DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

        ; IMMDevice::OpenPropertyStore
        ; 0x0 = STGM_READ
        DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
        ObjRelease(IMMDevice)

        ; IPropertyStore::GetValue
        VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
        VarSetCapacity(PROPERTYKEY, 20)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
        NumPut(14, &PROPERTYKEY + 16, "UInt")
        DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
        DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
        DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal
        ObjRelease(IPropertyStore)

        ObjRawSet(Devices, DeviceName, DeviceID)
    }
    ObjRelease(IMMDeviceCollection)

    ; put the available output devices in a list
    Devices2 := {}
    DevicesNames2:={}
    For DeviceName, DeviceID in Devices
        List .= "(" . A_Index . ") " . DeviceName . "`n", ObjRawSet(Devices2, A_Index, DeviceID), ObjRawSet(DevicesNames2, A_Index, DeviceName)

    ; print device names as a list to choose from
    ; InputBox n,, % List,,,,,,,, 1

    ; update the output audio device number
    ; goal is to iterate through all possible output audio devices.
    num_devices:= Devices2.MaxIndex() ;total number of audio output devices

    if (num_devices!=0){
        global output_audio_device_number
        if (output_audio_device_number< num_devices)
        {
            output_audio_device_number++
        }
        else{
            ; the highest device number is reached, so go back to the lowest number
            output_audio_device_number:=1
        }

        ; print ID of the device to change the output to 
        ; MsgBox % Devices2[n]
        
        ; print the name of the new output device
        ;MsgBox , ,,% DevicesNames2[output_audio_device_number],1
        ;TrayTip, New speaker device, % DevicesNames2[output_audio_device_number], 2, 

        ; Gui, +AlwaysOnTop +Disabled -SysMenu +Owner  ; +Owner avoids a taskbar button.
        ; Gui, Add, Text,, Some text to display.
        ; Gui, Show, NoActivate, Title of Window  ; NoActivate avoids deactivating the currently active window.

        ; Sleep, 1000
        ; ExitApp

        ; set the global variable parameter for 'SetAudioOutputDevice()'
        global output_device_id
        output_device_id:= Devices2[output_audio_device_number]

        ; display the name of the new output device on the screen
        WindowTitle:=% DevicesNames2[output_audio_device_number]

        SplashTextOn,,, % WindowTitle
        WinGetPos,,, Width, Height, % WindowTitle ; get the width and height of the popup window
        ; put the popup window in the bottom center of the screen
        WinMove, % WindowTitle,, (A_ScreenWidth/2)-(Width/2), A_ScreenHeight-Height-42

        ; set a timer to call a function with the next steps after 1 second
        SetTimer, ProceedChangingSpeakerDevice, 1000
    }
    else{
        MsgBox, 0,, No audio output devices found to switch between...
    }
}

ProceedChangingSpeakerDevice(){
    ; Turn timer off
    SetTimer, ProceedChangingSpeakerDevice, Off
    ; turn the splashtext message off
    SplashTextOff
    ; actual changing of the speaker device
    SetAudioOutputDevice()
}

SetAudioOutputDevice(){
    ; sets the output audio device to the given DeviceID

    global output_device_id

    ; IPolicyConfig::SetDefaultEndpoint ;(call for legacy autohotkey v1.0)
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
    R := DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "Str", output_device_id, "UInt", 0, "UInt")
    ObjRelease(IPolicyConfig)    

    ; print some hexadecimal number, an address or success/fail state (iguess)
    ; MsgBox % Format("0x{:08X}", R)
}

