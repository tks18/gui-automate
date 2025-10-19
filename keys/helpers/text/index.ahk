; Common Helper Functions
transformText(text, transformType) {
    transformedText := ""
    if (transformType) {
        if (transformType = "upper") {
            transformedText := StrUpper(text)
        } else if (transformType = "lower") {
            transformedText := StrLower(text)
        } else if (transformType = "title") {
            transformedText := StrTitle(text)
        } else if (transformType = "sentence") {
            transformedText := RegExReplace(StrLower(text), "((?:^|[.!?]\s+)[a-z])", "$u1")
        } else if (transformType = "fix br") {
            transformedText := RegExReplace(text, "\R", "`r`n")
        } else if (transformType = "reverse") {
            transformedText := DllCall("msvcrt.dll\_wcsrev", "Str", text, "str")
        }
        return transformedText
    } else {
        return text
    }
}