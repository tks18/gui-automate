CopyandPreserveClipboardText() {
    preserveClipboard := ClipboardAll()
    Send("^c")
    Sleep(ClipWait())
    Cliptext := A_Clipboard
    A_Clipboard := preserveClipboard
    Sleep(ClipWait())
    preserveClipboard := ""
    Return Cliptext
}

GetClipboardTextNoFormats() {
    Cliptext := A_Clipboard
    Sleep(ClipWait())
    return Cliptext
}

GetClipboardTextFormats() {
    Cliptext := ClipboardAll()
    Sleep(ClipWait())
    return Cliptext
}

PasteandPreserveClipboardText(text) {
    preserveClipboard := ClipboardAll()
    A_Clipboard := text
    Send("^v")
    Sleep(ClipWait())
    A_Clipboard := preserveClipboard
    Sleep(ClipWait())
    preserveClipboard := ''
}