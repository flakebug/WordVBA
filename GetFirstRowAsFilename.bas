Attribute VB_Name = "NewMacros"
Option Explicit

Sub GetFirstRowAsFilename()
Attribute GetFirstRowAsFilename.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����2"
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
    
    Dim FileName As String
    FileName = Dir(CurrentPath & "\*.doc", vbDirectory)
    Dim wddoc As Document
    
    Application.Visible = False
    Application.ScreenUpdating = False
    
    Dim OriginalFullname As String
    Dim TargetFullname As String
    Dim RawDocTitle As String
    Dim TrimmedDocTitle As String
    

    Do While FileName <> ""
        OriginalFullname = CurrentPath & "\" & FileName
        If OriginalFullname <> VBAFullname Then
            Set wddoc = Application.Documents.Open(OriginalFullname)
            Selection.HomeKey Unit:=wdStory
            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
            RawDocTitle = Selection
            TrimmedDocTitle = Replace(Replace(Replace(RawDocTitle, "�@�@�@", "-"), "�@", ""), vbCr, "")
            wddoc.Close
            TargetFullname = CurrentPath & "\" & TrimmedDocTitle & ".doc"
            Name OriginalFullname As TargetFullname
        End If
        
        FileName = Dir
    Loop
    

    Application.ScreenUpdating = True
    Application.Visible = True
    
    MsgBox "����", vbInformation
End Sub


