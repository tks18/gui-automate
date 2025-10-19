CreateTextMenuFromStr(textOptions, displayOptions) {
    TextMenu := Menu()

    MenuAction(ItemName, ItemPos, *) {
        A_Clipboard := textOptions[ItemPos]
        Return
    }
    for option in displayOptions {
        TextMenu.Add(option, MenuAction)
    }
    TextMenu.Show()
    TextMenu.Delete()
}