#Include "%A_ScriptDir%/utils/index.ahk"

BACKGROUND := "2D2B55"
TEXT := "A599E9"
GUITITLESTYLE := "xm w200 -E0x200 c" . TEXT
GUIEDITBOXSTYLE := "xm w200 -E0x200 +Border Background" . BACKGROUND . " c" . TEXT
GUIBOTTOMTEXTSTYLE := "xm w200 +Center -E0x200 c" . TEXT


Class Interface {

    __New() {
        this.gui := Gui()
        this.gui.OnEvent("Escape", this.onEscape)
        this.uiDestroyed := false
        this.editBoxTracker := []
        this.btnTracker := []
        this.gui.MarginX := "15"
        this.gui.MarginY := "15"
        this.gui.BackColor := "2D2B55"
        this.gui.SetFont("s9", "Verdana")
        this.gui.Title := "Shan.tk's Tools"
        this.gui.Opt("+AlwaysOnTop -SysMenu -ToolWindow -caption +Border")
        this.gui.Add("Text", GUIBOTTOMTEXTSTYLE, "Developed by Shan.tk 💜")
        this.guiCommandLabel := this.gui.Add("Text", GUITITLESTYLE, "Enter the Command:")
    }

    #WinActivateForce
    destroy() {
        this.gui.Destroy()
        this.uiDestroyed := true
        WinActivate("A")
    }

    disableAllEditBoxes() {
        for editBox in this.editBoxTracker {
            editBox.Enabled := false
        }
    }

    refreshUI() {
        this.gui.Show("AutoSize Center")
    }

    disableAllBtns() {
        for btn in this.btnTracker {
            btn.Enabled := false
        }
    }

    addEditBox(onChangeHandler, editTitle := "") {
        if (editTitle != "") {
            this.gui.Add("Text", GUITITLESTYLE, editTitle)
        }
        this.disableAllEditBoxes()
        uiEditBox := this.gui.Add("Edit", GUIEDITBOXSTYLE)
        uiEditBox.onEvent("Change", onChangeHandler)
        uiEditBox.Focus()
        this.editBoxTracker.Push(uiEditBox)
        return uiEditBox
    }

    addSearchBar(title, url) {
        this.disableAllEditBoxes()
        this.disableAllBtns()
        this.gui.Add("Text", GUITITLESTYLE, title)
        this.guiSearchBox := this.gui.Add("Edit", GUIEDITBOXSTYLE . " -WantReturn")

        onEnterPress(eventObject, item) {
            searchText := this.guiSearchBox.value
            this.destroy()
            querySafe := uriEncode(searchText)
            finalQuery := StrReplace(url, "REPLACEME", querySafe, , , 1)
            Run(finalQuery)
        }

        this.guiDefaultButton := this.gui.Add("Button", "x-10 y-10 w1 h1 +default", "")
        this.guiDefaultButton.onEvent("Click", onEnterPress)
        this.guiSearchBox.Focus()
        this.btnTracker.Push(this.guiDefaultButton)
        this.gui.Show("AutoSize")
        return this.guiSearchBox
    }

    addFreeUserInputBox(title, handler) {
        this.disableAllEditBoxes()
        this.disableAllBtns()
        this.gui.Add("Text", GUITITLESTYLE, title)
        this.guiUserInputBox := this.gui.Add("Edit", GUIEDITBOXSTYLE . " -WantReturn")
        this.editBoxTracker.Push(this.guiUserInputBox)

        this.guiDefaultButton := this.gui.Add("Button", "x-10 y-10 w1 h1 +default", "")
        this.guiDefaultButton.onEvent("Click", handler)
        this.guiUserInputBox.Focus()
        this.btnTracker.Push(this.guiDefaultButton)
        this.gui.Show("AutoSize")
        return this.guiUserInputBox
    }

    onEscape() {
        this.Destroy()
        return
    }

    handleCommands(commandConfig, eventObject, BaseUI) {
        for (config in commandConfig) {
            commandTitle := config.commandTitle
            commandCatchPhrase := config.phrase
            commandFunc := config.handleCommands
            commandFuncType := Type(commandFunc)
            if (commandCatchPhrase = eventObject.value) {
                if (commandFuncType = "Array") {
                    BaseUI.addEditBox(this.handleCommands(commandFunc, eventObject, BaseUI), commandTitle)
                    BaseUI.gui.Show("AutoSize")
                } else if (commandFuncType = "Func") {
                    commandFunc(eventObject, BaseUI)
                }
            }
        }
    }
}