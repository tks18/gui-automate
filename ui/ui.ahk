#Include "%A_ScriptDir%/utils/index.ahk"

BACKGROUND := "2D2B55"
TEXT := "A599E9"
GUITITLESTYLE := "xm w200 -E0x200 c" . TEXT
GUIEDITBOXSTYLE := "xm w200 -E0x200 +Border Background" . BACKGROUND . " c" . TEXT
GUIBOTTOMTEXTSTYLE := "xm w200 +Center -E0x200 c" . TEXT


Class Interface {

    __New() {
        this.ui := Gui()
        this.ui.OnEvent("Escape", this.onEscape)
        this.uiDestroyed := false
        this.editBoxTracker := []
        this.btnTracker := []
        this.ui.MarginX := "15"
        this.ui.MarginY := "15"
        this.ui.BackColor := "2D2B55"
        this.ui.SetFont("s9", "Verdana")
        this.ui.Title := "Shan.tk's Tools"
        this.ui.Opt("+AlwaysOnTop -SysMenu -ToolWindow -caption +Border")
        this.ui.Add("Text", GUIBOTTOMTEXTSTYLE, "Developed by Shan.tk 💜")
        this.uiTitle := this.ui.Add("Text", GUITITLESTYLE, "Enter the Command:")
    }

    #WinActivateForce
    destroy() {
        this.ui.Destroy()
        this.uiDestroyed := true
        WinActivate("A")
    }

    disableAllEditBoxes() {
        for editBox in this.editBoxTracker {
            editBox.Enabled := false
        }
    }

    refreshUI() {
        this.ui.Show("AutoSize Center")
    }

    disableAllBtns() {
        for btn in this.btnTracker {
            btn.Enabled := false
        }
    }

    addEditBox(onChangeHandler, editTitle := "") {
        if (editTitle != "") {
            this.ui.Add("Text", GUITITLESTYLE, editTitle)
        }
        this.disableAllEditBoxes()
        uiEditBox := this.ui.Add("Edit", GUIEDITBOXSTYLE)
        uiEditBox.onEvent("Change", onChangeHandler)
        uiEditBox.Focus()
        this.editBoxTracker.Push(uiEditBox)
        return uiEditBox
    }

    addSearchBar(title, url) {
        this.disableAllEditBoxes()
        this.disableAllBtns()
        this.ui.Add("Text", GUITITLESTYLE, title)
        this.uiSearchBox := this.ui.Add("Edit", GUIEDITBOXSTYLE . " -WantReturn")

        onEnterPress(eventObject, item) {
            searchText := this.uiSearchBox.value
            this.destroy()
            querySafe := uriEncode(searchText)
            finalQuery := StrReplace(url, "REPLACEME", querySafe, , , 1)
            Run(finalQuery)
        }

        this.uiDefaultButton := this.ui.Add("Button", "x-10 y-10 w1 h1 +default", "")
        this.uiDefaultButton.onEvent("Click", onEnterPress)
        this.uiSearchBox.Focus()
        this.btnTracker.Push(this.uiDefaultButton)
        this.ui.Show("AutoSize")
        return this.uiSearchBox
    }

    addFreeUserInputBox(title, handler) {
        this.disableAllEditBoxes()
        this.disableAllBtns()
        this.ui.Add("Text", GUITITLESTYLE, title)
        this.uiUserInputBox := this.ui.Add("Edit", GUIEDITBOXSTYLE . " -WantReturn")
        this.editBoxTracker.Push(this.uiUserInputBox)

        this.uiDefaultButton := this.ui.Add("Button", "x-10 y-10 w1 h1 +default", "")
        this.uiDefaultButton.onEvent("Click", handler)
        this.uiUserInputBox.Focus()
        this.btnTracker.Push(this.uiDefaultButton)
        this.ui.Show("AutoSize")
        return this.uiUserInputBox
    }

    onEscape() {
        this.Destroy()
        return
    }
}