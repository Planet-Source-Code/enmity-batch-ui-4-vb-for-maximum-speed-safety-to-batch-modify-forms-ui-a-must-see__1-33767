VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************
'*
'*  Program : Batch UI 4 VB - Font
'*
'*  Purpose : for maximum speed and safety to batch modify forms' UI!
'*
'*  Version : v0.91
'*
'*  Author  : Unruled Boy
'*
'*  Last Modified : 2002.4.15
'*
'*  Contact : unruledboy@cmmail.com
'*
'*  class Name: cFileFunctions
'*
'*  Copyrights:
'*              Please do freely modify, transfer, redistrbute it
'*            as your wish!
'*              The author will not give any warranty for it.
'*
'*  Note : Any comment is welcome, otherwise, enjoy!
'*
'*  Please vote & commend for Batch UI 4 VB for a better speed and safety
'*  with custom borders, styles, texts, captions, tooltips, menus and more!!!!!!
'*
'*************************************************************

Option Explicit


Public Function FilePath(ByVal strFile As String) As String
    
    If DoesFileExist(strFile) Then
        If Len(strFile) > 3 Then 'x:\a.bbb
            If Mid(strFile, 2, 2) = ":\" Then
                FilePath = Left(strFile, InStrRev(strFile, "\") - 1)
            Else
                FilePath = ""
            End If
        Else
            FilePath = ""
        End If
    Else
        FilePath = ""
    End If
    
End Function


Public Function FileBaseName(ByVal strFileName As String) As String
        
    If DoesFileExist(strFileName) Then
        If Len(strFileName) > 3 Then 'x:\a.bbb
            If Mid(strFileName, 2, 2) = ":\" Then
                'a.bbb
                strFileName = Right(strFileName, Len(strFileName) - InStrRev(strFileName, "\"))
                If InStr(strFileName, ".") <> 0 Then
                    FileBaseName = Left(strFileName, InStrRev(strFileName, ".") - 1)
                Else
                    FileBaseName = strFileName
                End If
            Else
                FileBaseName = ""
            End If
        Else
            FileBaseName = ""
        End If
    Else
        FileBaseName = ""
    End If
    
End Function


Public Function GetFileName(ByVal strFile As String) As String
        
    If DoesFileExist(strFile) Then
        If Len(strFile) > 3 Then 'x:\a.bbb
            If Mid(strFile, 2, 2) = ":\" Then
                'a.bbb
                GetFileName = Right(strFile, Len(strFile) - InStrRev(strFile, "\"))
            Else
                GetFileName = ""
            End If
        Else
            GetFileName = ""
        End If
    Else
        GetFileName = ""
    End If
    
End Function


Public Function FileExtName(ByVal strFileName As String) As String
        
    If DoesFileExist(strFileName) Then
        If Len(strFileName) > 5 Then 'x:\a.bbb
            If Mid(strFileName, 2, 2) = ":\" Then
                If InStr(strFileName, ".") <> 0 Then
                    FileExtName = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") - 1)
                Else
                    FileExtName = ""
                End If
            Else
                FileExtName = vbNullString
            End If
        Else
            FileExtName = vbNullString
        End If
    Else
        FileExtName = vbNullString
    End If
    
End Function


Public Function DoesFileExist(ByVal strFile As String) As Boolean
    
On Error GoTo handleError
    
    Dim o_intFileNum As Integer
    
    o_intFileNum = FreeFile
    Open strFile For Input As #o_intFileNum
    Close #o_intFileNum
    DoesFileExist = True
    
    Exit Function
    
handleError:
    Close #o_intFileNum
    DoesFileExist = False
End Function



