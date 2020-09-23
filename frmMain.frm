VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch UI 4 VB"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9660
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.CommandButton cmdVote 
      Caption         =   "Please &Vote!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   6680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All..."
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectUI 
      Caption         =   "&Custom..."
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ListBox lstStatus 
      Height          =   780
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   7095
   End
   Begin VB.Frame fraControlUI 
      Caption         =   "Appearance Sample"
      Height          =   1575
      Left            =   7320
      TabIndex        =   11
      Top             =   4080
      Width           =   2175
      Begin VB.Label lblFontUI 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   810
         TabIndex        =   12
         Top             =   720
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdRemoveFiles 
      Caption         =   "&Remove Files"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddFiles 
      Caption         =   "&Add Files"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ListBox lstFiles2BeChanged 
      Height          =   2400
      Left            =   4800
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   4695
   End
   Begin VB.FileListBox filFormFiles 
      Height          =   2250
      Left            =   120
      MultiSelect     =   2  'Extended
      Pattern         =   "*.frm"
      TabIndex        =   7
      Top             =   3840
      Width           =   4580
   End
   Begin VB.DirListBox dirFormFiles 
      Height          =   2400
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.DriveListBox drvFormFiles 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4580
   End
   Begin VB.Frame fraControlTypes 
      Caption         =   "Control Types:"
      Height          =   1575
      Left            =   4800
      TabIndex        =   4
      Top             =   4080
      Width           =   2415
      Begin VB.ListBox lstControls 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdBatchChange 
      Caption         =   "&Batch Change!"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Frame fraTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Line linHR 
         BorderColor     =   &H8000000F&
         X1              =   960
         X2              =   6360
         Y1              =   400
         Y2              =   400
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Left            =   240
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For maximum speed and safety to batch modify forms' UI!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   5280
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batch UI 4 VB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   1
         Top             =   45
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
'*  Form Name: frmMain
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetInputState Lib "user32" () As Long

Private WithEvents m_ucmBatchUI As cBatchUI
Attribute m_ucmBatchUI.VB_VarHelpID = -1
Private m_ucmIniInfo As cIniInfo
Attribute m_ucmIniInfo.VB_VarHelpID = -1
Private m_ucmFileFunctions As cFileFunctions
Private m_ucmFileDialogs As cFileDialogs

Private m_blnItemChangedByCoder As Boolean



Private Sub SetHorizontalScrollBar(ByRef lstItem As Control, _
                                   ByVal strItem As String)

        Dim o_lngWidth As Long
        Static o_lngCurrentWidth As Long
                
        If o_lngCurrentWidth < TextWidth(strItem & "  ") Then
          o_lngCurrentWidth = TextWidth(strItem + " ")
          o_lngWidth = o_lngCurrentWidth
          
          If Me.ScaleMode = vbTwips Then o_lngWidth = o_lngWidth / Screen.TwipsPerPixelX
          
          SendMessageByNum lstItem.hwnd, LB_SETHORIZONTALEXTENT, o_lngWidth, 0
        End If

End Sub


Private Sub DoEventsEx()
    
    If GetInputState() <> 0 Then DoEvents
    
End Sub


Private Sub cmdAddFiles_Click()
        
    Dim o_intItems As Integer
    
    With filFormFiles
        If .ListCount > 0 Then
            For o_intItems = 0 To .ListCount - 1
                If .Selected(o_intItems) Then
                    If m_ucmBatchUI.Search4DuplicatedItemEx(lstFiles2BeChanged, _
                        .Path & "\" & .List(o_intItems)) <> -1 Then
                        'duplicated
                    Else
                        lstFiles2BeChanged.AddItem .Path & "\" & .List(o_intItems)
                        SetHorizontalScrollBar lstFiles2BeChanged, lstFiles2BeChanged.List(lstFiles2BeChanged.ListCount - 1)
                    End If
                Else
                End If
            Next
        End If
    End With
    
End Sub


Private Sub cmdBatchChange_Click()
    
    Dim o_intItems As Integer
    Dim o_strControlTypes As String
    Dim o_lngTimeElapsed As Long
    
    With lstFiles2BeChanged
        '.Visible = False
        
        If .ListCount > 0 Then
            
            o_strControlTypes = vbNullString
            
            With lstControls
                For o_intItems = 0 To .ListCount - 1
                    o_strControlTypes = o_strControlTypes & _
                                        IIf(.Selected(o_intItems), "1", "0") & _
                                        IIf(o_intItems < .ListCount - 1, ",", "")
                    m_ucmIniInfo.Upd_Ini_Data App.Path & "\controlsinfo.ini", _
                                              "ControlsInfo", _
                                              "Control" & CStr(o_intItems + 1) & "Selected", _
                                              IIf(.Selected(o_intItems), "1", "0")
                Next
            End With
            
            o_lngTimeElapsed = GetTickCount()
            
            For o_intItems = .ListCount - 1 To 0 Step -1
                lstStatus.AddItem "Changing #" & CStr(.ListCount - o_intItems) & ", totally " & _
                        CStr(.ListCount) & "files(" & GetFileName(.List(o_intItems)) & ")..."
                m_ucmBatchUI.ChangeFormFileUI .List(o_intItems), lblFontUI.Font, o_strControlTypes
                
                DoEventsEx
            Next
            
            MsgBox "Batch Changing Completed! Time Elaspsed:" & GetTickCount() - o_lngTimeElapsed & " ms.", vbInformation
        End If
        
        '.Visible = True
    End With
    
End Sub


Private Sub cmdRemoveFiles_Click()
    
    Dim o_intItems As Integer
    
    With lstFiles2BeChanged
        If .ListCount > 0 Then
            For o_intItems = .ListCount - 1 To 0 Step -1
                If .Selected(o_intItems) Then
                    .RemoveItem o_intItems
                Else
                End If
            Next
        Else
        End If
    End With
    
End Sub


Private Sub cmdSelectAll_Click()

    Dim o_intItems As Integer
    
    If Not m_blnItemChangedByCoder Then
        With lstControls
            For o_intItems = 0 To .ListCount - 1
                .Selected(o_intItems) = True
            Next
        End With
    Else
    End If

End Sub


Private Sub cmdSelectUI_Click()
    
    Dim o_fntSample As StdFont
    
    Set o_fntSample = lblFontUI.Font
    
    If m_ucmFileDialogs.VBChooseFont(o_fntSample, , Me.hwnd) Then
        With o_fntSample
            If .Size >= 8 And .Size <= 24 And LenB(.Name) Then
                Set lblFontUI.Font = o_fntSample
            End If
        End With
    Else
    End If
    
End Sub


Private Sub cmdVote_Click()

    ShellExecute 0, "Open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=33767&lngWId=1", vbNullString, vbNullString, vbNormalFocus

End Sub


Private Sub dirFormFiles_Change()
    
    filFormFiles.Path = dirFormFiles.Path
    
End Sub


Private Sub drvFormFiles_Change()
    
    dirFormFiles.Path = drvFormFiles.Drive
    
End Sub


Private Sub filFormFiles_DblClick()
    
    With lstFiles2BeChanged
        If m_ucmBatchUI.Search4DuplicatedItemEx(lstFiles2BeChanged, _
            filFormFiles.Path & "\" & filFormFiles.FileName) <> -1 Then
            'duplicated
        Else
            .AddItem filFormFiles.Path & "\" & filFormFiles.FileName
            SetHorizontalScrollBar lstFiles2BeChanged, .List(.ListCount - 1)
        End If
    End With
    
End Sub


Private Sub Form_Load()
        
    Set m_ucmBatchUI = New cBatchUI
    
    Set m_ucmIniInfo = New cIniInfo
    
    Set m_ucmFileFunctions = New cFileFunctions
    
    Set m_ucmFileDialogs = New cFileDialogs
    
    LoadControls
    
    Set imgLogo.Picture = Me.Icon
    
    lstStatus.AddItem "stand by..."
    
End Sub


Private Sub LoadControls()
    
    Dim o_intItems As Integer
    Dim o_intRet As Integer
    
    If m_ucmFileFunctions.DoesFileExist(App.Path & "\controlsinfo.ini") Then
        With m_ucmIniInfo
            o_intRet = Val(.Rtv_Ini_Data(App.Path & "\controlsinfo.ini", "ControlsInfo", "Controls"))
            If o_intRet > 0 Then
                For o_intItems = 0 To o_intRet - 1
                    lstControls.AddItem .Rtv_Ini_Data(App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "Name")
                    lstControls.ItemData(o_intItems) = o_intItems
                    lstControls.Selected(o_intItems) = IIf(.Rtv_Ini_Data(App.Path & "\controlsinfo.ini", "ControlsInfo", "Control" & CStr(o_intItems + 1) & "Selected") = "1", True, False)
                Next
                If lstControls.ListCount > 0 Then
                    lstControls.ListIndex = 0
                Else
                End If
            Else
            End If
        End With
    Else
    End If
    
    If lstControls.ListCount = 0 Then
        For o_intItems = 0 To m_ucmBatchUI.ControlTypesCount - 1
            With m_ucmBatchUI.ControlTypeItems(o_intItems)
                lstControls.AddItem .strName
                lstControls.ItemData(o_intItems) = .intIndex
                lstControls.Selected(o_intItems) = CBool(.blnSelected)
            End With
        Next
        lstControls.ListIndex = 0
    Else
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
        
    Set m_ucmBatchUI = Nothing
        
End Sub


Private Sub m_ucmBatchUI_ChangeCompleted(ByVal blnSuccess As Boolean)

    lstStatus.AddItem "Change " & IIf(blnSuccess, "successfully", " failed") & ")¡£"

End Sub
