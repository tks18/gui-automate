#Requires AutoHotkey v2.0
SendMode("Input")
SetWorkingDir(A_ScriptDir)
#SingleInstance
#Include "%A_ScriptDir%\init.ahk"
#Include "%A_ScriptDir%\keys.ahk"

; Other Hotkeys & Strings
#Include '%A_ScriptDir%\keys\helpers\index.ahk'
#Include '%A_ScriptDir%\keys\scripts\index.ahk'