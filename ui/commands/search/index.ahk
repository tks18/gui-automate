handleSearchFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "g")
    {
        Interface.addSearchBar("Ask Google !", "https://www.google.com/search?num=50&safe=off&site=&source=hp&q=REPLACEME&btnG=Search&oq=&gs_l=")
    }

    if (currentText = "y")
    {
        Interface.addSearchBar("Search Youtube", "https://www.youtube.com/results?search_query=REPLACEME")
    }
}

addSearchEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleSearchFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Enter Search Engine")
}