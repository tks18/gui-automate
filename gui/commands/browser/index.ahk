getPersonalChromeProfile(url) {
    return '"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" --app "' url . '"'
}

getWorkChromeProfile(url) {
    return '"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 1" --app "' . url . '"'
}


handleBrowserFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "ap") {
        BaseUI.addSearchBar("Go to any Personal URL", getPersonalChromeProfile("REPLACEME"))
    }

    if (currentText = "aw") {
        BaseUI.addSearchBar("Go to any Work URL", getWorkChromeProfile("REPLACEME"))
    }

    if (currentText = "omnia") {
        BaseUI.destroy()
        Run(getWorkChromeProfile("https://deloitteomniaapa.deloitte.com/"))
    }

    if (currentText = "dc") {
        BaseUI.destroy()
        Run(getWorkChromeProfile("https://deloitteconnect.deloitte.com/"))
    }

    if (currentText = "music") {
        BaseUI.destroy()
        Run(getPersonalChromeProfile("https://open.spotify.com/"))
    }
}

addBrwoserEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleBrowserFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Application Shortcuts")
}