#Include "%A_ScriptDir%/utils/index.ahk"

BACKGROUND := "2D2B55"
TEXT := "A599E9"
GUITITLESTYLE := " w200 -E0x200 Wrap c" . TEXT
GUIEDITBOXSTYLE := "xm w200 -E0x200 +Border Background" . BACKGROUND . " c" . TEXT
GUIBOTTOMTEXTSTYLE := "xm w200 +Center -E0x200 c" . TEXT


Class Interface {

    __New() {
        this.ui := Gui()
        this.ui.OnEvent("Escape", this.onEscape)
        this.uiDestroyed := false
        this.editBoxTracker := []
        this.editBox := ""
        this.editBoxTitle := ""
        this.editBoxCallback := ""
        this.btnTracker := []
        this.ui.MarginX := "5"
        this.ui.MarginY := "5"
        this.ui.BackColor := "2D2B55"
        this.ui.SetFont("s9", "Verdana")
        this.ui.Title := "Shan.tk's Tools"
        this.ui.Opt("+AlwaysOnTop -SysMenu -ToolWindow -caption")
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
        this.ui.Show()
        xWindowPos := 0
        this.ui.GetPos(, , &xWindowPos,)
        xLocation := (A_ScreenWidth / 2) - (xWindowPos / 2) - (xWindowPos / 2)
        yLocation := 10
        this.ui.Move(xLocation, yLocation)
    }

    disableAllBtns() {
        for btn in this.btnTracker {
            btn.Enabled := false
        }
    }

    addEditBox(onChangeHandler, editTitle := "") {
        if (this.editBox = "") {
            this.editBoxTitle := this.ui.Add("Text", GUITITLESTYLE, "./")
            this.editBox := this.ui.Add("Edit", GUIEDITBOXSTYLE)
            this.editBoxCallback := onChangeHandler
            this.editBox.onEvent("Change", this.editBoxCallback)
            this.editBox.Focus()
            return this.editBox
        } else {
            newTitle := this.editBoxTitle.value
            if (editTitle != "") {
                newTitle := newTitle . this.editBox.value . "/"
            }
            this.editBoxTitle.value := newTitle
            this.editBox.value := ""
            this.editBox.onEvent("Change", this.editBoxCallback, 0)
            this.editBoxCallback := onChangeHandler
            this.editBox.onEvent("Change", onChangeHandler)
            this.editBox.Focus()
            return this.editBox
        }
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
        this.refreshUI()
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
        this.refreshUI()
        return this.uiUserInputBox
    }

    onEscape() {
        this.Destroy()
        return
    }
}