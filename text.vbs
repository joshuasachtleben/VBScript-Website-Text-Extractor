'https://www.frbservices.org/EPaymentsDirectory/FedACHdir.txt
'https://www.frbservices.org/EPaymentsDirectory/agreement.html

Set IE = CreateObject("InternetExplorer.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

With IE
  .Visible = True
  .Navigate "https://www.frbservices.org/EPaymentsDirectory/FedACHdir.txt"
  
  Do While .Busy
    WScript.Sleep 100
  Loop
  .Document.getElementById("agree_terms_use").Click
  
  Do While .Busy
    WScript.Sleep 100
  Loop
  
  Dim text
  Set elements = IE.Document.getElementsByTagName("pre")
  For Each element in elements
    text = element.innerHTML
  Next
  'msgBox(text)
  
  outFile = "c:\Users\joshuasachtleben\Desktop\data.txt"
  Set objFile = objFSO.CreateTextFile(outFile, True)
  objFile.Write text
  objFile.Close
  
  msgBox("Done")
  
End With