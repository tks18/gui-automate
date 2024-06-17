; Format Pivot Values
formatPivotValues() {
    KeyWait("Alt")
    Send("{Alt}")
    Send(0)
    Send("I")
}

; Create Pivot Table
createPivotTable() {
    KeyWait("Alt")
    Send("{Alt}")
    KeyWait("n")
    Send("n")
    KeyWait("v")
    Send("v")
    KeyWait("t")
    Send("t")
    KeyWait("Enter")
    Send("{Enter}")
    formatPivotValues()
}

transposePivotFields() {
    KeyWait("Alt")
    Send("{Alt}")
    Send("Y1,")
    Send("YYG")
}

sortAscendingPivotTable() {
    KeyWait("Alt")
    Send("{Alt}")
    Send(1)
}

sortDescendingPivotTable() {
    KeyWait("Alt")
    Send("{Alt}")
    Send(2)
}

changePivotValueFieldOrientation() {
    KeyWait("Alt")
    Send("{Alt}")
    Send("Y1,")
    Send("YYH")
}
