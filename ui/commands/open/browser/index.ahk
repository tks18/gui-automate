getPersonalChromeProfile(url) {
    return '"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Default" --app "' url . '"'
}

getWorkChromeProfile(url) {
    return '"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 1" --app "' . url . '"'
}


handleBrowserFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "ap") {
        Interface.addSearchBar("Go to any Personal URL", getPersonalChromeProfile("REPLACEME"))
    }

    if (currentText = "aw") {
        Interface.addSearchBar("Go to any Work URL", getWorkChromeProfile("REPLACEME"))
    }

    if (currentText = "omnia") {
        Interface.destroy()
        Run(getWorkChromeProfile("https://deloitteomniaapa.deloitte.com/"))
    }

    if (currentText = "dc") {
        Interface.destroy()
        Run(getWorkChromeProfile("https://deloitteconnect.deloitte.com/"))
    }

    if (currentText = "music") {
        Interface.destroy()
        Run(getPersonalChromeProfile("https://open.spotify.com/"))
    }
}

addBrwoserEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleBrowserFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Browser Actions")
}