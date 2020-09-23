VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBatchUI"
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
'*  Class Name: cBatchUI
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

Private Const LB_FINDSTRING = &H18F
Private Const LB_ERR = (-1)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type ControlType
    intIndex As Integer
    strName As String
    strID As String
    strFont As String
    blnSelected As Boolean
End Type
Private m_udtControlType(0 To 15) As ControlType

Private Const mc_strAllControlTypes = "1,1,1,1,1,1,1,1,1,1,1"

Private Const mc_strAppearanceID As String = "BeginProperty Font "
Private Const mc_strAppearanceIDEx As String = "BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"

Private Const mc_strTrueValue As String = "-1  'True"
Private Const mc_strFalseValue As String = "0   'False"

Private Const mc_strFontBoldValue As String = "700"
Private Const mc_strFontNonBoldValue As String = "400"

Private Const mc_strEndTag As String = "   End"
Private Const mc_strStandardWhiteSpace As String = "   "

Private m_ucmIniInfo As cIniInfo
'Private m_ucmFileLines As cFileLines
Private m_ucmFileFunctions As cFileFunctions

Public Event ChangeCompleted(ByVal blnSuccess As Boolean)



Public Property Get ControlTypesCount() As Integer
    
    ControlTypesCount = UBound(m_udtControlType) + 1
    
End Property


Public Property Get ControlTypeItems(ByVal intIndex As Integer) As Variant
    
    ControlTypeItems = m_udtControlType(intIndex)
    
End Property


Public Function Search4DuplicatedItemEx(ByVal lstTarget As Object, _
                                        ByVal strItem As String) As Long

    Dim o_lngRet As Long
    
    If Not (lstTarget Is Nothing) Then
        With lstTarget
            If .ListCount > 0 Then
                o_lngRet = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal strItem)
                If o_lngRet <> LB_ERR Then
                    Search4DuplicatedItemEx = o_lngRet
                Else
                    Search4DuplicatedItemEx = -1
                End If
            Else
                Search4DuplicatedItemEx = -1
            End If
        End With
    Else
        Search4DuplicatedItemEx = -1
    End If

End Function


Public Function Search4DuplicatedItem(ByVal lstTarget As Object, _
                                        ByVal strItem As String) As Boolean
    
    Dim o_intItems As Integer
    Dim o_blnItemIsDumplicated As Boolean
    
    o_blnItemIsDumplicated = False
    
    With lstTarget
        If .ListCount > 0 Then
            For o_intItems = 0 To .ListCount - 1
                If .List(o_intItems) = strItem Then
                    o_blnItemIsDumplicated = True
                    Exit For
                Else
                End If
            Next
            Search4DuplicatedItem = o_blnItemIsDumplicated
        Else
            Search4DuplicatedItem = False
        End If
    End With
    
End Function


'Public Function ChangeFormFileUIEx(ByVal strFormFile As String, _
'                                 ByVal objFont As StdFont, _
'                                 Optional ByVal strControlTypes As String = mc_strAllControlTypes) As Boolean
'
'On Error GoTo handleError
'
'    Dim o_intItems As Integer
'    Dim o_blnItemChanged As Boolean
'    Dim o_intFileNum As Integer
'    Dim o_intFileNum2 As Integer
'    Dim o_strLineData As String
'    Dim o_strWhiteSpaces As String
'    Dim o_strControlTypes() As String
'    Dim o_blnUIUpdated As Boolean
'
'    Dim o_lngLines As Long
'    Dim o_lngItems As Long
'
'    '*************** control types
'
'    o_strControlTypes() = Split(strControlTypes, ",")
'    For o_intItems = LBound(o_strControlTypes) To UBound(o_strControlTypes)
'        m_udtControlType(o_intItems).blnSelected = (o_strControlTypes(o_intItems) = "1")
'        'Debug.Print m_udtControlType(o_intItems).strID & m_udtControlType(o_intItems).blnSelected
'    Next
'
'    'o_intFileNum = FreeFile
'    'Open strFormFile For Input As #o_intFileNum
'
'    o_intFileNum2 = FreeFile
'    'Open m_ucmFileFunctions.FilePath(strFormFile) & "\" & m_ucmFileFunctions.FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile) For Output As #o_intFileNum2
'    Open strFormFile & ".new" For Output As #o_intFileNum2
'
'    With m_ucmFileLines
'        o_lngLines = .ReadFile(strFormFile)
'
'        o_lngItems = 0
'
'        Do While o_lngItems < o_lngLines
'            o_lngItems = o_lngItems + 1
'            o_strLineData = .Lines2String(o_lngItems)
'
'            o_blnItemChanged = False
'
'            For o_intItems = LBound(m_udtControlType) To UBound(m_udtControlType)
'                With m_udtControlType(o_intItems)
'                    If InStr(Trim(o_strLineData), Trim(.strID)) <> 0 Then
'                        If .blnSelected Then
'                            o_blnItemChanged = True
'                        Else
'
'                        End If
'
'                        Exit For
'                    Else
'
'                    End If
'                End With
'            Next
'
'            Print #o_intFileNum2, o_strLineData
'
'            If o_blnItemChanged Then
'
'                o_lngItems = o_lngItems + 1
'                o_strLineData = .Lines2String(o_lngItems)
'                'Debug.Print o_strLineData
'
'                o_blnUIUpdated = False
'
'                Do While Trim(o_strLineData) <> Trim(mc_strEndTag)
'
'                    If Trim(o_strLineData) = Trim(m_udtControlType(o_intItems).strFont) Then
'                        'ÒÑ×Ô¶¨Òå£¬ÐÞ¸Ä
'                        Print #o_intFileNum2, o_strLineData
'
'                        o_strWhiteSpaces = WhiteSpaces(o_strLineData) & _
'                                           mc_strStandardWhiteSpace
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Name =""" & _
'                                              objFont.Name & """"
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Size =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              CStr(objFont.Size)
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Charset =   0"
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Weight =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Bold, _
'                                                  mc_strFontBoldValue, _
'                                                  mc_strFontNonBoldValue)
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Underline =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Underline, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Italic =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Italic, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Strikethrough= " & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Strikethrough, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                        o_lngItems = o_lngItems + 1
'                        o_strLineData = .Lines2String(o_lngItems)
'                        Print #o_intFileNum2, o_strWhiteSpaces & "EndProperty"
'
'                        o_blnUIUpdated = True
'
'                        Exit Do
'                    Else
'                        '²»ÊÇÄ¿±ê£¬Ö±½ÓÐ´ÈëÎÄ¼þ
'                        Print #o_intFileNum2, o_strLineData
'                    End If
'
'                    o_lngItems = o_lngItems + 1
'                    o_strLineData = .Lines2String(o_lngItems)
'                    'Debug.Print o_strLineData
'
'                Loop
'
'                'Î´¶¨Òå£¬Ìí¼Ó
'                If Not o_blnUIUpdated Then
'                    o_strWhiteSpaces = WhiteSpaces(o_strLineData) & mc_strStandardWhiteSpace
'                    Print #o_intFileNum2, o_strWhiteSpaces & m_udtControlType(o_intItems).strFont
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Name =""" & _
'                                              objFont.Name & """"
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Size =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              CStr(objFont.Size)
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Charset =   0"
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Weight =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Bold, _
'                                                  mc_strFontBoldValue, _
'                                                  mc_strFontNonBoldValue)
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Underline =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Underline, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Italic =" & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Italic, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                        Print #o_intFileNum2, o_strWhiteSpaces & "   Strikethrough= " & _
'                                              mc_strStandardWhiteSpace & _
'                                              IIf(objFont.Strikethrough, _
'                                                  mc_strTrueValue, _
'                                                  mc_strFalseValue)
'
'                    Print #o_intFileNum2, o_strWhiteSpaces & "EndProperty"
'
'                    Print #o_intFileNum2, o_strLineData '"End"
'                 Else
'                 End If
'            Else
'            End If
'        Loop
'    End With
'
'    'Close #o_intFileNum
'    Close #o_intFileNum2
'
'    'Name FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile) As FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "3" & "." & FileExtName(strFormFile)
'    'Name strFormFile As FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile)
'    'Name FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "3" & "." & FileExtName(strFormFile) As strFormFile
'
'    SetAttr strFormFile, vbNormal
'
'    Name strFormFile As strFormFile & ".bak"
'    Name strFormFile & ".new" As strFormFile
'
'    ChangeFormFileUIEx = True
'
'    RaiseEvent ChangeCompleted(True)
'
'    Exit Function
'
'handleError:
'    ChangeFormFileUIEx = False
'
'    RaiseEvent ChangeCompleted(False)
'
'End Function
'


Public Function ChangeFormFileUI(ByVal strFormFile As String, _
                                 ByVal objFont As StdFont, _
                                 Optional ByVal strControlTypes As String = mc_strAllControlTypes) As Boolean
    
On Error GoTo handleError

    Dim o_intItems As Integer
    Dim o_blnItemChanged As Boolean
    Dim o_intFileNum As Integer
    Dim o_intFileNum2 As Integer
    Dim o_strLineData As String
    Dim o_strWhiteSpaces As String
    Dim o_strControlTypes() As String
    Dim o_blnUIUpdated As Boolean
            
    '*************** control types
    
    o_strControlTypes() = Split(strControlTypes, ",")
    For o_intItems = LBound(o_strControlTypes) To UBound(o_strControlTypes)
        m_udtControlType(o_intItems).blnSelected = (o_strControlTypes(o_intItems) = "1")
        'Debug.Print m_udtControlType(o_intItems).strID & m_udtControlType(o_intItems).blnSelected
    Next
    
    o_intFileNum = FreeFile
    Open strFormFile For Input As #o_intFileNum
    
    o_intFileNum2 = FreeFile
    'Open FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile) For Output As #o_intFileNum2
    Open strFormFile & ".new" For Output As #o_intFileNum2
        
    Do While Not EOF(o_intFileNum)
        Line Input #o_intFileNum, o_strLineData
        'Debug.Print o_strLineData
        
        o_blnItemChanged = False
        
        For o_intItems = LBound(m_udtControlType) To UBound(m_udtControlType)
            With m_udtControlType(o_intItems)
                If InStr(Trim(o_strLineData), Trim(.strID)) <> 0 Then
                    If .blnSelected Then
                        o_blnItemChanged = True
                    Else
                        
                    End If
                    
                    Exit For
                Else
                    
                End If
            End With
        Next
        
        Print #o_intFileNum2, o_strLineData
        
        If o_blnItemChanged Then
            
            Line Input #o_intFileNum, o_strLineData
            'Debug.Print o_strLineData
            
            o_blnUIUpdated = False
            
            Do While Trim(o_strLineData) <> Trim(mc_strEndTag)
            
                If Trim(o_strLineData) = Trim(m_udtControlType(o_intItems).strFont) Then
                    'ÒÑ×Ô¶¨Òå£¬ÐÞ¸Ä
                    Print #o_intFileNum2, o_strLineData
                    
                    o_strWhiteSpaces = WhiteSpaces(o_strLineData) & _
                                       mc_strStandardWhiteSpace
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Name =""" & _
                                          objFont.Name & """"
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Size =" & _
                                          mc_strStandardWhiteSpace & _
                                          CStr(objFont.Size)
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Charset =   0"
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Weight =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Bold, _
                                              mc_strFontBoldValue, _
                                              mc_strFontNonBoldValue)
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Underline =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Underline, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Italic =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Italic, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Strikethrough= " & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Strikethrough, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                    Line Input #o_intFileNum, o_strLineData
                    Print #o_intFileNum2, o_strWhiteSpaces & "EndProperty"
                    
                    o_blnUIUpdated = True
                    
                    Exit Do
                Else
                    '²»ÊÇÄ¿±ê£¬Ö±½ÓÐ´ÈëÎÄ¼þ
                    Print #o_intFileNum2, o_strLineData
                End If
                
                Line Input #o_intFileNum, o_strLineData
                'Debug.Print o_strLineData
            Loop
                        
            'Î´¶¨Òå£¬Ìí¼Ó
            If Not o_blnUIUpdated Then
                o_strWhiteSpaces = WhiteSpaces(o_strLineData) & mc_strStandardWhiteSpace
                Print #o_intFileNum2, o_strWhiteSpaces & m_udtControlType(o_intItems).strFont
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Name =""" & _
                                          objFont.Name & """"
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Size =" & _
                                          mc_strStandardWhiteSpace & _
                                          CStr(objFont.Size)
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Charset =   0"
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Weight =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Bold, _
                                              mc_strFontBoldValue, _
                                              mc_strFontNonBoldValue)
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Underline =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Underline, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Italic =" & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Italic, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                    Print #o_intFileNum2, o_strWhiteSpaces & "   Strikethrough= " & _
                                          mc_strStandardWhiteSpace & _
                                          IIf(objFont.Strikethrough, _
                                              mc_strTrueValue, _
                                              mc_strFalseValue)
                    
                Print #o_intFileNum2, o_strWhiteSpaces & "EndProperty"
                
                Print #o_intFileNum2, o_strLineData '"End"
             Else
             End If
        Else
            
        End If
        
    Loop
    
    Close #o_intFileNum
    Close #o_intFileNum2
    
    'Name FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile) As FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "3" & "." & FileExtName(strFormFile)
    'Name strFormFile As FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "2" & "." & FileExtName(strFormFile)
    'Name FilePath(strFormFile) & "\" & FileBaseName(strFormFile) & "3" & "." & FileExtName(strFormFile) As strFormFile
    SetAttr strFormFile, vbNormal
    Name strFormFile As strFormFile & ".bak"
    Name strFormFile & ".new" As strFormFile
    
    ChangeFormFileUI = True
    
    RaiseEvent ChangeCompleted(True)
    
    Exit Function
    
handleError:
    ChangeFormFileUI = False
    
    RaiseEvent ChangeCompleted(False)
    
End Function


Private Function WhiteSpaces(ByVal strData As String) As String
    
    Dim o_intItems As Integer
    
    If Len(strData) > 0 And InStr(strData, " ") <> 0 Then
        For o_intItems = 1 To Len(strData)
            If Mid(strData, o_intItems, 1) <> " " Then
                WhiteSpaces = Space(o_intItems - 1)
                Exit For
            Else
            End If
        Next
    Else
        strData = ""
    End If
    
End Function


Private Sub InitializeBatchUI()

    'ReDim m_udtControlType(0 To 15) As ControlType
    
    With m_udtControlType(0)
        .intIndex = 0
        .strName = "Label"
        .strID = "Begin VB.Label "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(1)
        .intIndex = 1
        .strName = "TextBox"
        .strID = "Begin VB.TextBox "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(2)
        .intIndex = 2
        .strName = "CommandButton"
        .strID = "Begin VB.CommandButton "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(3)
        .intIndex = 3
        .strName = "CheckBox"
        .strID = "Begin VB.CheckBox "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(4)
        .intIndex = 4
        .strName = "OptionButton"
        .strID = "Begin VB.OptionButton "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(5)
        .intIndex = 5
        .strName = "ListBox"
        .strID = "Begin VB.ListBox "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(6)
        .intIndex = 6
        .strName = "ComboBox"
        .strID = "Begin VB.ComboBox "
        .strFont = mc_strAppearanceID
        .blnSelected = True
    End With
    With m_udtControlType(7)
        .intIndex = 7
        .strName = "MSHFlexGrid"
        .strID = "Begin MSHierarchicalFlexGridLib.MSHFlexGrid "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(8)
        .intIndex = 8
        .strName = "MSFlexGrid"
        .strID = "Begin MSFlexGridLib.MSFlexGrid "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(9)
        .intIndex = 9
        .strName = "MaskEdBox"
        .strID = "Begin MSMask.MaskEdBox "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(10)
        .intIndex = 10
        .strName = "DataGrid"
        .strID = "Begin MSDataGridLib.DataGrid "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(11)
        .intIndex = 11
        .strName = "ImageCombo"
        .strID = "Begin MSComctlLib.ImageCombo "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(12)
        .intIndex = 12
        .strName = "ListView"
        .strID = "Begin MSComctlLib.ListView "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(13)
        .intIndex = 13
        .strName = "StatusBar"
        .strID = "Begin MSComctlLib.StatusBar "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(14)
        .intIndex = 14
        .strName = "TabStrip"
        .strID = "Begin MSComctlLib.TabStrip "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With
    With m_udtControlType(15)
        .intIndex = 15
        .strName = "TreeView"
        .strID = "Begin MSComctlLib.TreeView "
        .strFont = mc_strAppearanceIDEx
        .blnSelected = True
    End With

End Sub


Public Function RebuildControlsInfoFile() As Boolean

    Dim o_intItems As Integer

    For o_intItems = LBound(m_udtControlType) To UBound(m_udtControlType)
        With m_udtControlType(o_intItems)
            m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", "ControlsInfo", "Controls", CStr(o_intItems + 1)

            m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "Name", .strName
            m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "ID", .strID
            m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "Font", .strFont
            m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "Selected", IIf(.blnSelected, "1", "0")
        End With
    Next
    
    RebuildControlsInfoFile = True

End Function


Private Sub Class_Initialize()
    
    Set m_ucmIniInfo = New cIniInfo
        
    'Set m_ucmFileLines = New cFileLines
    
    Set m_ucmFileFunctions = New cFileFunctions
    
    InitializeBatchUI
            
End Sub


Private Sub Class_Terminate()
    
    'Erase m_udtControlType
    
    'Set m_ucmFileLines = Nothing
    
    Set m_ucmIniInfo = Nothing
    
    Set m_ucmFileFunctions = Nothing
    
End Sub
