#Include chrome.ahk
#Include explorer.ahk
^F1::
global ChromeInst := new Chrome("C:\chromeahk\profile","--headless")
_File := Explorer_GetSelection() 
SplitPath, _File, , OutDir, , OutNameNoExt
Document := ComObjGet(_File)
Document.SaveAs2(OutDir "/" OutNameNoExt ".htm", 10) ;// wdFormatFilteredHTML:=10
Document.Close()
HTMLFilePath := "file:///" OutDir "/" OutNameNoExt ".htm"

PageInstance := ChromeInst.GetPage()
PageInstance.Call("Page.navigate", {"url": HTMLFilePath})
PageInstance.WaitForLoad()
html  := PageInstance.Evaluate("document.getElementsByClassName('WordSection1')[0].innerHTML").value
Sleep 500
Clipboard := html
WinKill, A	
return
