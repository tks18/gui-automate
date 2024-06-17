#Include "../utils/index.ahk"

BACKGROUND := "2D2B55"
TEXT := "A599E9"
GUITITLESTYLE := "xm w200 -E0x200 c" . TEXT
GUIEDITBOXSTYLE := "xm w200 -E0x200 +Border Background" . BACKGROUND . " c" . TEXT
GUIBOTTOMTEXTSTYLE := "xm w200 +Center -E0x200 c" . TEXT


Class BaseUI {

    __New() {
        this.BaseGUI := Gui()
        this.BaseGUI.OnEvent("Escape", this.onEscape)
        this.uiDestroyed := false
        this.editBoxTracker := []
        this.BaseGUI.MarginX := "15"
        this.BaseGUI.MarginY := "15"
        this.BaseGUI.BackColor := "2D2B55"
        this.BaseGUI.SetFont("s9", "Verdana")
        this.BaseGUI.Title := "Shan.tk's Tools"
        this.BaseGUI.Add("Text", GUIBOTTOMTEXTSTYLE, "Developed by Shan.tk 💜")
        this.BaseGUI.Opt("+AlwaysOnTop -SysMenu -ToolWindow -caption +Border")
        this.BaseGUITitle := this.BaseGUI.Add("Text", GUITITLESTYLE, "Enter the Command:")
    }

    #WinActivateForce
    destroy() {
        this.BaseGUI.Destroy()
        this.uiDestroyed := true
        WinActivate("A")
    }

    disableAllEditBoxes() {
        for editBox in this.editBoxTracker {
            editBox.Enabled := false
        }
    }

    addEditBox(onChangeHandler, editTitle := "") {
        if (editTitle != "") {
            this.BaseGUI.Add("Text", GUITITLESTYLE, editTitle)
        }
        this.disableAllEditBoxes()
        uiEditBox := this.BaseGUI.Add("Edit", GUIEDITBOXSTYLE)
        uiEditBox.onEvent("Change", onChangeHandler)
        uiEditBox.Focus()
        this.editBoxTracker.Push(uiEditBox)
        return uiEditBox
    }

    addSearchBar(title, url) {
        this.BaseGUI.Add("Text", GUITITLESTYLE, title)
        this.BaseGUISearchBox := this.BaseGUI.Add("Edit", GUIEDITBOXSTYLE . " -WantReturn")

        onEnterPress(eventObject, item) {
            searchText := this.BaseGUISearchBox.value
            this.destroy()
            querySafe := uriEncode(searchText)
            finalQuery := StrReplace(url, "REPLACEME", querySafe, , , 1)
            Run(finalQuery)
        }

        this.BaseGUIDefaultButton := this.BaseGUI.Add("Button", "x-10 y-10 w1 h1 +default", "")
        this.BaseGUIDefaultButton.onEvent("Click", onEnterPress)
        this.disableAllEditBoxes()
        this.BaseGUI.Show("AutoSize")
        return
    }

    onEscape() {
        this.Destroy()
        return
    }
}