#HotIf WinActive("ahk_exe excel.exe")
Alt & c::
{
    Send("{F2}")
    Send("^a")
    Send("{F9}")
    Send("^c")
    Send("^z")
    Send("{Esc}")
}

#HotIf WinActive("ahk_exe excel.exe")
Alt & n:: {
    KeyWait "Alt"
    Send("{Alt}")
    Send("n")
    Send("v")
    Send("t")
}

#HotIf WinActive("ahk_exe excel.exe")
Alt & ,:: {
    KeyWait "Alt"
    Send("{Alt}")
    Send("h")
    Send("d")
    Send("s")
}

#HotIf WinActive("ahk_exe excel.exe")
Alt & k:: {
    KeyWait "Alt"
    Send("{Alt}")
    Send("h")
    Send("f")
    Send("d")
    Send("s")
    Send("k")
    Send("{Enter}")
}

#HotIf WinActive("ahk_exe excel.exe")
Alt & [:: {
    KeyWait "Alt"
    Send("{Alt}")
    Send("h")
    Send("o")
    Send("i")
}

#HotIf WinActive("ahk_exe excel.exe")
Alt & ]:: {
    KeyWait "Alt"
    Send("{Alt}")
    Send("h")
    Send("o")
    Send("a")
}