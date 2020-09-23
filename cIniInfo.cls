VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniInfo"
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
'*  class Name: cIniInfo
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

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal keyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer

Private Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)

Private Const mc_lngZero As Long = 0
Private Const mc_intEntryValueSize As Integer = 1024
Private Const mc_strDefaultValue As String = "<unknown>"

Private m_intRet As Integer
Private m_intNullOffset As Integer

Private m_strRet As String

Private m_ucmFileFunctions As cFileFunctions



Public Function Rtv_Ini_Data(ByVal strINIFile As String, ByVal strSection As String, ByVal strEntry As String) As String
        
    On Error GoTo handleReadIniDataErr

        m_strRet = Space$(mc_intEntryValueSize)
        
        m_intRet = GetPrivateProfileString(strSection, strEntry, mc_strDefaultValue, m_strRet, mc_intEntryValueSize, strINIFile)
        
        m_intRet = Len(Trim$(m_strRet)) - 1
        
        m_strRet = Left$(m_strRet, m_intRet)
        
        Rtv_Ini_Data = m_strRet
        
    Exit Function
        
handleReadIniDataErr:
    Rtv_Ini_Data = "<unknown>"

End Function


Public Function Upd_Ini_Data(ByVal strINIFile As String, _
                             ByVal strSection As String, _
                             ByVal strEntry As String, _
                             ByVal strValue As String) As Boolean

    On Error GoTo handleUpdateIniDataErr
        
        m_strRet = vbNullString
        
        strValue = CStr(strValue)
        
        m_intRet = WritePrivateProfileString(strSection, strEntry, strValue, strINIFile)
        
        Upd_Ini_Data = m_intRet <> 0
        
    Exit Function
        
handleUpdateIniDataErr:
    Upd_Ini_Data = False
        
End Function

    
Public Function Delete_INI_Data(ByVal sPath As String, _
                               ByVal sSection As String, _
                               ByVal sKey As String) As Boolean
  
         Dim lStatus As Long
         
         sSection = FixSectionName(sSection) & vbNullChar 'remove [] if any
         If Len(sKey) > 0 Then
           sKey = sKey & vbNullChar
         End If
         sPath = sPath & vbNullChar
        
         lStatus = WritePrivateProfileString(sSection, _
                                             sKey, _
                                             mc_lngZero, _
                                             sPath)
         Delete_INI_Data = (lStatus <> 0)

End Function


Private Function FixSectionName(ByVal sOldName As String) As String
  
        Dim strNewName As String
        
        If Left$(sOldName, 1) = "[" Then
          strNewName = Mid$(sOldName, 2)
        Else
          strNewName = sOldName
        End If
        
        If Right$(sOldName, 1) = "]" Then
          strNewName = Left$(strNewName, Len(strNewName) - 1)
        End If
        
        FixSectionName = strNewName

End Function


Public Function EnumIniSections(ByVal strINIFile As String, ByRef vSections() As String) As Boolean

    On Error Resume Next
        
        If m_ucmFileFunctions.DoesFileExist(strINIFile) <> 0 Then
            
            Erase vSections()
            
            m_strRet = String$(128, 0)
            
            m_intRet = GetPrivateProfileStringSections(0, 0, "", m_strRet, 127, strINIFile)
            
            m_intRet = 0
            
            Do
                m_intNullOffset = InStr(m_strRet, Chr$(0))
                If m_intNullOffset > 1 Then
                   m_intRet = m_intRet + 1
                   ReDim Preserve vSections(1 To m_intRet) As String
                   vSections(m_intRet) = Mid$(m_strRet, 1, m_intNullOffset - 1)
                   m_strRet = Mid$(m_strRet, m_intNullOffset + 1)
                End If
            Loop While m_intNullOffset > 1
            EnumIniSections = True
        Else
            EnumIniSections = False
        End If
    
End Function


Public Function EnumIniEntries(strINIFile$, strSection$, ByRef vEntries() As String) As Boolean

On Error Resume Next
            
    If m_ucmFileFunctions.DoesFileExist(strINIFile$) <> 0 Then
        Erase vEntries()
        m_strRet = String$(10000, 0)
        m_intRet = GetPrivateProfileStringKeys(strSection$, 0, "", m_strRet, 9999, strINIFile$)     'get the entries from the section
        m_intRet = 0
        Do
            m_intNullOffset = InStr(m_strRet, Chr$(0))
            If m_intNullOffset > 1 Then
               m_intRet = m_intRet + 1
               ReDim Preserve vEntries(1 To m_intRet) As String
               vEntries(m_intRet) = Mid$(m_strRet, 1, m_intNullOffset - 1)
               m_strRet = Mid$(m_strRet, m_intNullOffset + 1)
            End If
        Loop While m_intNullOffset > 1
        EnumIniEntries = True
    Else
        EnumIniEntries = False
    End If
    
End Function


Private Sub Class_Initialize()

    Set m_ucmFileFunctions = New cFileFunctions
    
End Sub

Private Sub Class_Terminate()
    
    Set m_ucmFileFunctions = Nothing

End Sub
