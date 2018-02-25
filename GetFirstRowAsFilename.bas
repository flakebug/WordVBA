Attribute VB_Name = "NewMacros"
Option Explicit

Sub GetFirstRowAsFilename()
Attribute GetFirstRowAsFilename.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����2"
    On Error Resume Next
    
    Dim result As VbMsgBoxResult
    result = MsgBox("���~���{�����|�N�ɭP�ɮצW�ٷl�a�B�L�k�^�_�A�нT�{�A���{���ɮצ�m�w���T�s���s�ƥ����u�@��Ƨ���", vbExclamation + vbOKCancel)
    If result = vbCancel Then
        End
    End If
    result = MsgBox("�ЦA���T�{�A���{���ɮצ�m�w���T�s���s�ƥ����u�@��Ƨ���", vbExclamation + vbOKCancel)
    If result = vbCancel Then
        End
    End If
    result = MsgBox("OK�A�ǳư���", vbQuestion + vbOKCancel)
    If result = vbCancel Then
        End
    End If
   

    Dim VBAFullname As String
    VBAFullname = ActiveDocument.FullName
 
    Dim CurrentPath As String
    CurrentPath = ActiveDocument.Path
    
    Dim filename As String
    filename = Dir(CurrentPath & "\*.doc", vbDirectory)
    Dim wddoc As Document
    
    Application.Visible = False
    Application.ScreenUpdating = False
    
    Dim OriginalFullname As String
    Dim TargetFullname As String
    Dim RawDocTitle As String
    Dim TrimmedDocTitle As String
    
    Open CurrentPath & "\log.txt" For Append As #1
    Do While filename <> ""
        OriginalFullname = CurrentPath & "\" & filename
        If OriginalFullname <> VBAFullname Then
            Set wddoc = Application.Documents.Open(OriginalFullname)
            Selection.HomeKey Unit:=wdStory
            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
            RawDocTitle = Selection
            TrimmedDocTitle = RemoveSpecial(RawDocTitle)
            wddoc.Close
            TargetFullname = CurrentPath & "\" & TrimmedDocTitle & ".doc"
            If Trim(TrimmedDocTitle) <> "" Then
                Name OriginalFullname As TargetFullname
            Else
                Err.Raise 9999
            End If
          
            If Err Then
                Print #1, Now & ", �o�{�S��r���w���է�, " & filename
                If Err.Number <> 9999 Then
                    TryToRename filename, TrimmedDocTitle
                End If
                Err.Clear
            End If
        End If
        
        filename = Dir
    Loop
    Close #1

    Application.ScreenUpdating = True
    Application.Visible = True
    
    MsgBox "����", vbInformation
    
End Sub

Sub TryToRename(filename As String, TrimmedDocTitle As String)
    On Error Resume Next
    Dim OriginalFullname As String
    Dim TargetFullname As String
    Dim LengthOfTitle As Integer
    Dim CurrentPath As String
    CurrentPath = ActiveDocument.Path
    
    OriginalFullname = CurrentPath & "\" & filename
rename:
    LengthOfTitle = Len(TrimmedDocTitle)
    If LengthOfTitle = 0 Then
        Exit Sub
    End If
    TrimmedDocTitle = Left(TrimmedDocTitle, LengthOfTitle - 1)
    TargetFullname = CurrentPath & "\" & TrimmedDocTitle & ".doc"
    Name OriginalFullname As TargetFullname
    If Err Then
        GoTo rename
    End If
    Err.Clear
    
End Sub

Function RemoveSpecial(str As String) As String
    'updatebyExtendoffice 20160303
    Dim xChars As String
    Dim I As Long
    xChars = "�@"
    For I = 0 To 47
        xChars = xChars & Chr(I)
    Next
    For I = 58 To 64
        xChars = xChars & Chr(I)
    Next
    For I = 91 To 96
        xChars = xChars & Chr(I)
    Next
    For I = 123 To 127
        xChars = xChars & Chr(I)
    Next
    
    For I = 1 To Len(xChars)
        str = Replace$(str, Mid$(xChars, I, 1), "")
    Next
    RemoveSpecial = str
End Function
