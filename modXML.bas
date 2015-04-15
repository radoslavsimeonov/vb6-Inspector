Attribute VB_Name = "modXML"
Option Explicit

Private sXML As String

Public Enum sTree
    Root = 0
    Element = 1
    SubElement = 2
    Entry = 3
    Entry1 = 4
    Entry2 = 5
    Entry3 = 6
End Enum

Public Sub ClearXMLFile()
    sXML = vbNullString
End Sub

Public Function SaveFileAsXML(Optional sFileName As String) As Boolean
On Error GoTo hErr
    
    Dim iFileNo As Integer
    Dim sHeader As String
    Dim sFolder As String
    
    SaveFileAsXML = False
    iFileNo = FreeFile
    
    If sFileName = vbNullString Then
        sFolder = App.Path & "\Reports\"
        sFileName = UCase$(COMPUTER_NAME & "_" & Format$(Now, "ddmmyyyyhhmm")) & ".xml"
    End If
    
    MakeSureDirectoryPathExists sFolder
        
    sHeader = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "windows-1251" & Chr(34) & "?>"
    
    Open sFolder & sFileName For Output As #iFileNo
    
    Print #iFileNo, sHeader
    
    Print #iFileNo, sXML
    
    Close #iFileNo
    
    MsgBox "Изходящия файл за " & COMPUTER_NAME & " е записан успешно в папка:" _
            & vbCrLf & vbCrLf & _
            sFolder & sFileName, vbInformation
    
    SaveFileAsXML = True
    
    Exit Function
hErr:
    MsgBox "Грешка при записване на изходящите данни. Проверете дали имате достъп до" & _
            vbCrLf & vbCrLf & _
            sFolder, vbCritical
End Function

Public Sub AddRow(iLevel As sTree, sField As String, Optional sValue As Variant, Optional bClose As Boolean = True)

    Dim sTab As String
    
    sTab = String$(iLevel, vbTab)
    
    If Not IsMissing(sValue) Then
        If sValue = vbNullString And bClose Then Exit Sub
        
        If VarType(sValue) = vbString Then _
            sValue = ReplaceSymbols(sValue)
    End If
    
    sXML = sXML & sTab
    sXML = sXML & "<" & IIf(bClose And IsMissing(sValue), "/", "") & sField & ">"
    
    If IsMissing(sValue) Then
        sXML = sXML & vbCrLf
    Else
        sXML = sXML & sValue
        
        If bClose Then
            sXML = sXML & "</" & sField & ">" & vbCrLf
        End If
    End If
    
End Sub

Private Function ReplaceSymbols(ByVal sValue As String) As String
    If InStr(sValue, "&") > 0 Then sValue = Replace(sValue, "&", "&amp;")
    If InStr(sValue, "<") > 0 Then sValue = Replace(sValue, "<", "&lt;")
    If InStr(sValue, ">") > 0 Then sValue = Replace(sValue, ">", "&gt;")
    If InStr(sValue, "'") > 0 Then sValue = Replace(sValue, "'", "&apos;")
    If InStr(sValue, Chr$(34)) > 0 Then sValue = Replace(sValue, Chr$(34), "&quot;")
    If InStr(sValue, Chr$(0)) > 0 Then sValue = Replace(sValue, Chr$(0), "")
    
    ReplaceSymbols = sValue
End Function

