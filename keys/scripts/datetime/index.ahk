; ==============================================
; 🕒 DateTime Expansions Library - v2
; ==============================================

; --- Standard Date & Time ---
::@date:: {
    SendInput FormatTime(, "MM/dd/yyyy")
}
::@time:: {
    SendInput FormatTime(, "HH:mm:ss")
}
::@datetime:: {
    SendInput FormatTime(, "MM/dd/yyyy HH:mm:ss")
}
::@datelong:: {
    SendInput FormatTime(, "dddd, MMMM dd, yyyy")
}
::@monthyear:: {
    SendInput FormatTime(, "MMM yyyy")
}
::@year:: {
    SendInput FormatTime(, "yyyy")
}

; --- ISO & File-safe Formats ---
::@iso:: {
    SendInput FormatTime(, "yyyy-MM-dd")
}
::@isotime:: {
    SendInput FormatTime(, "yyyy-MM-ddTHH:mm:ss")
}
::@ts:: { ; short timestamp
    SendInput FormatTime(, "yyyyMMdd_HHmmss")
}
::@tslog:: { ; log-style timestamp
    SendInput "[" FormatTime(, "yyyy-MM-dd HH:mm:ss") "]"
}
::@tsfile:: { ; filename-friendly timestamp
    SendInput FormatTime(, "yyyy-MM-dd_HH-mm-ss")
}

; --- Natural Language ---
::@today:: {
    SendInput "Today (" FormatTime(, "ddd, MMM dd") ")"
}
::@yesterday:: {
    SendInput "Yesterday (" FormatTime(A_Now - 86400, "ddd, MMM dd") ")"
}
::@tomorrow:: {
    SendInput "Tomorrow (" FormatTime(A_Now + 86400, "ddd, MMM dd") ")"
}

; --- Quarter / Fiscal ---
::@quarter:: {
    month := FormatTime(, "M")
    q := Ceil(month / 3)
    SendInput "Q" q " " FormatTime(, "yyyy")
}
::@fiscal:: {
    month := FormatTime(, "M")
    fiscalYear := (month >= 4) ? FormatTime(, "yy") "-" FormatTime(, "yy") + 1 : FormatTime(, "yy") - 1 "-" FormatTime(, "yy")
    SendInput "FY" fiscalYear
}

; ==============================================
; 🧩 End of DateTime Library
; ==============================================
