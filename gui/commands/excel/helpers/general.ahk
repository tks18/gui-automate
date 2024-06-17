formatTable() {
    Send("^a")
    Send("!6")
}

formatSheet() {
    KeyWait("Alt")
    Send("{Alt}")
    Send("0B")
}

deleteSheet() {
    KeyWait("Alt")
    Send("{Alt}")
    KeyWait("h")
    Send("h")
    KeyWait("d")
    Send("d")
    KeyWait("s")
    Send("s")
    KeyWait("Enter")
    Send("{Enter}")
}