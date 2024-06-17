handleSearchFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "g")
    {
        BaseUI.addSearchBar("Ask Google !", "https://www.google.com/search?num=50&safe=off&site=&source=hp&q=REPLACEME&btnG=Search&oq=&gs_l=")
    }

    if (currentText = "y")
    {
        BaseUI.addSearchBar("Search Youtube", "https://www.youtube.com/results?search_query=REPLACEME")
    }
}

addSearchEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleSearchFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Enter Search Engine")
}