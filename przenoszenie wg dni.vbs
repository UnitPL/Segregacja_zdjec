Dim strVBSFolder 'As String
Dim strVBSFile 'AS String

strVBSFolder = Left(WScript.ScriptFullName, _
		    InstrRev(WScript.ScriptFullName, _
                            "\", _
                            Len(WScript.ScriptFullName)))
strVBSFile = Right(WScript.ScriptFullName, _
                   Len(WScript.ScriptFullName) - InstrRev(WScript.ScriptFullName, _
                                                          "\", _
                                                          Len(WScript.ScriptFullName)))

If MsgBox("Czy napewno chcech posortowaæ folder: " & Chr(10) & strVBSFolder, _
          vbQuestion + vbOKCancel, _
         "Nie bêdzie odwrotu!!!") = vbOK Then 

    Dim objShell 'As Object 'Shell32.Shell
    Dim objFolder 'As Object 'Shell32.Folder
    Dim oFile 'As Object 'Shell32.FolderItem
    Dim strTime 'As String

    Dim objFSO 'As Object 'Scripting.FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim strNewDir 'As String
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(strVBSFolder)

    Dim i 'As Integer
    Dim j 'As Long
    Dim k 'AS Long

    
    Const sHeading = 12 ' Date taken

    For Each oFile In objFolder.Items
        If Not oFile.IsFolder Then
            If oFile.Name <> strVBSFile Then
                strTime = Replace_RegExp(objFolder.GetDetailsOf(objFolder.Parsename(oFile.Name), _
                                                                sHeading), _
                                        "[^\d\-\: ]", "")
		If Len(strTime) > 0 then

                    strNewDir = objFolder.Self.Path & "\" & Left(strTime, 2) & "." & mid(strTime, 3,2) & "." & mid(strTime, 5,4) & " -" 
                    With objFSO
                        If Not .FolderExists(strNewDir) Then 
			    .CreateFolder strNewDir
			    i = i + 1
		        End If
    
                        .MoveFile oFile.Path, strNewDir & "\" & oFile.Name
		        j = j + 1
                    End With
		Else
		    k = k + 1
 		End If

            End If
        End If

    Next
    
    If i > 0 Then
	MsgbOx "przeniesionych plików: " & j & Chr(10) & _
               "utworzonych nowych folderów: " & i, vbInformation, "Koniec :-)"
    Else
        MsgbOx "Nie ma tu nic do roboty :-|", vbInformation
    End If

    If k > 0 then
	MsgbOx "Pozosta³e pliki (szt: " & k & ") maj¹ nieokreœlon¹ datê utworzenia pliku" & chr(10) & _
  	       "Ktoœ przy nich grzeba³!" & chr(10) & chr(10) & _
	       "Z tym nic nie zrobiê :-) sorka!", vbInformation
    End If

    Set objFSO = Nothing
    Set objShell = Nothing
    Set objFolder = Nothing
End IF

Function Replace_RegExp(vText, strFind, vReplace) 'As String
    
    Dim objRegExp 'As Object 'VBScript.RegExp
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = True
        .Pattern = strFind
        Replace_RegExp = .Replace(vText, vReplace)
    End With
    Set objRegExp = Nothing
    
End Function
