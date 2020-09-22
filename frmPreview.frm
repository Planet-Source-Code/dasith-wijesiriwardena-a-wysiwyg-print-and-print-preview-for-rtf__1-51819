VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Preview [ Please Note: The Printed Page May Have Slight Variations. Use This As a Rough Guide. ]"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10470
   ControlBox      =   0   'False
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TabStrip tabPreview 
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   3560
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      HotTracking     =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 1"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picToolbar 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   10215
      TabIndex        =   5
      Top             =   120
      Width           =   10275
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3960
         ScaleHeight     =   375
         ScaleWidth      =   4695
         TabIndex        =   18
         Top             =   0
         Width           =   4695
         Begin VB.HScrollBar Slider1 
            Height          =   300
            Left            =   0
            Max             =   2
            Min             =   1
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Value           =   1
            Width           =   3000
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   315
            Left            =   4200
            Picture         =   "frmPreview.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   60
            Width           =   315
         End
         Begin VB.CommandButton cmdStart 
            Height          =   315
            Left            =   3120
            Picture         =   "frmPreview.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin VB.CommandButton cmdNext 
            Height          =   315
            Left            =   3840
            Picture         =   "frmPreview.frx":06D6
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   60
            Width           =   315
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   315
            Left            =   3480
            Picture         =   "frmPreview.frx":0820
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.CommandButton btnPreview 
         Appearance      =   0  'Flat
         Caption         =   "Zoom &Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton btnPreview 
         Appearance      =   0  'Flat
         Caption         =   "Zoom &In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   990
         TabIndex        =   9
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton btnPreview 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   795
      End
      Begin VB.ComboBox cboPercent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Done"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   6
         Top             =   80
         Width           =   855
      End
   End
   Begin VB.PictureBox picParent 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   120
      ScaleHeight     =   2805
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   720
      Width           =   4095
      Begin VB.PictureBox picPageNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   720
         ScaleHeight     =   240
         ScaleWidth      =   1335
         TabIndex        =   24
         Top             =   100
         Width           =   1335
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Page 1 Of 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.HScrollBar hscPreview 
         Height          =   240
         LargeChange     =   2000
         Left            =   300
         SmallChange     =   500
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2895
      End
      Begin VB.VScrollBar vscPreview 
         Height          =   2475
         LargeChange     =   2000
         Left            =   3720
         SmallChange     =   500
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox picH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   1
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   1935
         TabIndex        =   17
         Top             =   400
         Width           =   1935
         Begin VB.Line Line2 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   1
            X1              =   0
            X2              =   1800
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.PictureBox picH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   1935
         TabIndex        =   16
         Top             =   1000
         Width           =   1935
         Begin VB.Line Line2 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   0
            X2              =   1800
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.PictureBox picV1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   400
         ScaleHeight     =   1815
         ScaleWidth      =   15
         TabIndex        =   15
         Top             =   0
         Width           =   15
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   1
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   1680
         End
      End
      Begin VB.PictureBox picV1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   3480
         ScaleHeight     =   1815
         ScaleWidth      =   15
         TabIndex        =   14
         Top             =   360
         Width           =   15
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   1680
         End
      End
      Begin VB.PictureBox imgCorner 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   240
         Left            =   3360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   2340
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Index           =   0
         Left            =   60
         ScaleHeight     =   2055
         ScaleWidth      =   3075
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Image picChild 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000010&
         ForeColor       =   &H00C0C0C0&
         Height          =   135
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label picture1 
         BackColor       =   &H80000010&
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   3000
         TabIndex        =   11
         Top             =   80
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''
'   Credits (Please give credits if you use this on your own project.)
'   I found the print preview module at pscode.com
'   The User Interface and code for the Print Preview Window was done by me.
'   The WYSIWYG display example was found at MSDN.
'
'   - Dasith Wijesiriwardena. dasiths@hotmail.com
'
'   The WYSIWYG Display shows exactly what is going to be printed but
'   the Print Preview may have slight variations. Can't seem to find the
'   problem. Please email me if you can further improve this.
'
'''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private Const lBorder = 100
Private ScalePercent As Integer
Private bLoad As Boolean
Private intLeftMargin As Integer
Private intRightMargin As Integer
Private intTopMargin As Integer
Private intBottomMargin As Integer


Public Sub AddPage(PageNumber As Integer)

    If PageNumber > 1 Then
        Load picPreview(PageNumber - 1)
        Set picPreview(PageNumber - 1) = Nothing
        tabPreview.Tabs.Add PageNumber, , "Page " & PageNumber
    End If
    
End Sub

Private Sub FillCboPercent()

    Dim iCount As Integer
    Dim strSearch As String
    
    With cboPercent
        For iCount = 200 To 30 Step -10
            .AddItem CStr(iCount) & "%"
        Next
        strSearch = "100%"
        .ListIndex = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strSearch)
    End With
    
End Sub

Public Sub PictureShow()

On Error GoTo err1:
    
    Screen.MousePointer = vbHourglass
    With picChild
        .Height = (ScalePercent / 100) * picPreview(0).Height
        .Width = (ScalePercent / 100) * picPreview(0).Width
        ResizeScrollBars
        
        DoCenter
        
        picture1.Move .Left + .Width, .Top + (80 * ScalePercent / 100), (120 * ScalePercent / 100), .Height
        Label1.Move .Left + (80 * ScalePercent / 100), .Top + .Height, .Width, (80 * ScalePercent / 100)
    End With
err1:
    Screen.MousePointer = vbDefault
    Call Form_Resize

End Sub

Private Sub PreviewPrint()

    Dim iCount, iPicCount As Integer
    
    On Error GoTo ErrHandle
    For iCount = 0 To picPreview.Count - 1
        picPreview(iCount).Picture = picPreview(iCount).Image
    Next
                
    If Printer.Copies > 0 Then
        For iCount = 1 To Printer.Copies
            Printer.Print
            For iPicCount = 0 To picPreview.Count - 1
                Printer.PaintPicture picPreview(iPicCount).Picture, 0, 0
                If iPicCount < picPreview.Count - 1 Then _
                    Printer.NewPage
            Next
            Printer.EndDoc
        Next
    End If
    
    Exit Sub
    
ErrHandle:
    Select Case Err.Number
        Case 482
            MsgBox "Make sure that you have a printer installed.  If a " & _
                "printer is installed, go into your printer properties " & _
                "look under the Setup tab, and make sure the ICM checkbox " & _
                "is checked and try printing again.", , "Printer Error"
            Exit Sub
        Case 32755
            Exit Sub
        Case Else
            MsgBox Err.Number & " " & Err.Description, , "Preview - Printing"
            Resume Next
    End Select
    
End Sub

Private Sub PreviewZoomIn()
    
    With cboPercent
        If .ListIndex - 1 >= 0 Then
            ScalePercent = ScalePercent + 10
            .ListIndex = .ListIndex - 1
        End If
    End With
    
    Exit Sub
    
ErrHandle:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & " " & Err.Description, , "Preview - Printing"
            Resume Next
    End Select

End Sub

Private Sub PreviewZoomOut()

    With cboPercent
        If .ListIndex + 1 < .ListCount Then
            ScalePercent = ScalePercent - 10
            .ListIndex = .ListIndex + 1
        End If
    End With
    
    Exit Sub
    
ErrHandle:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & " " & Err.Description, , "Preview - Printing"
            Resume Next
    End Select

End Sub

Private Sub ResizeScrollBars()

    With vscPreview

        If picChild.Height > picParent.Height Then
            .Visible = True
            .Max = picChild.Height - picParent.ScaleHeight
            .Min = 0
            .LargeChange = picChild.Height - picParent.Height
            imgCorner.Visible = True
        Else
            .Visible = False
            imgCorner.Visible = False
        End If
    End With
    
    With hscPreview

        If picChild.Width > picParent.Width Then
            .Visible = True
            .Max = picChild.Width - picParent.ScaleWidth
            .Min = 0
            .LargeChange = picChild.Width - picParent.ScaleWidth
            imgCorner.Visible = True
        Else
            .Visible = False
            imgCorner.Visible = False
        End If
    End With

End Sub

Public Sub SizePreview(lWidth As Long, lHeight As Long)

    Dim iCount As Integer
    
    For iCount = 0 To picPreview.Count - 1
        With picPreview(iCount)
            .Left = 0
            .Top = 0
            .Width = lWidth
            .Height = lHeight
        End With
    Next
    picChild.Move 0, 0, lWidth, lHeight
    
    Form_Resize
    
End Sub

Private Sub btnPreview_Click(Index As Integer)

    Select Case Index
        Case 0  'Print
            'PreviewPrint
            Form1.Command1_Click
            Unload Me
        Case 1  'Zoom In
            PreviewZoomIn
        Case 2  'Zoom Out
            PreviewZoomOut
    End Select
    
End Sub

Private Sub cboPercent_Change()

    If bLoad = False Then
        vscPreview.Value = 0
        hscPreview.Value = 0
        With cboPercent
            ScalePercent = CInt(Left(.List(.ListIndex), Len(.List(.ListIndex)) - 1))
        End With
        PictureShow
    End If
        
End Sub

Private Sub cboPercent_Click()

    If bLoad = False Then
        vscPreview.Value = 0
        hscPreview.Value = 0
        With cboPercent
            ScalePercent = CInt(Left(.List(.ListIndex), Len(.List(.ListIndex)) - 1))
        End With
        PictureShow
    End If
    
    Call Form_Resize

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdEnd_Click()
Slider1.Value = Slider1.Max
End Sub

Private Sub cmdNext_Click()
If Not Slider1.Value = Slider1.Max Then
    Slider1.Value = Slider1.Value + 1
End If
End Sub

Private Sub cmdPrevious_Click()
If Not Slider1.Value = 1 Then
    Slider1.Value = Slider1.Value - 1
End If
End Sub

Private Sub cmdStart_Click()
Slider1.Value = 1
End Sub

Private Sub Form_Activate()

    With picPreview(0)
        .Picture = .Image
        picChild.Move 0, 0, .Width, .Height
        picChild.Picture = .Picture
        PictureShow

    End With
    
    Form_Resize
    
    On Error Resume Next
    Slider1.Max = tabPreview.Tabs.Count
    If tabPreview.Tabs.Count = 1 Then
        Slider1.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = False
        cmdStart.Enabled = False
        cmdEnd.Enabled = False
    End If
    
    Call Slider1_Change
    
    'If 1000 * (Slider1.Max - 1) > picToolbar.Width - 5160 Then
        'Slider1.Width = picToolbar.Width - 5160
    'Else
        'Slider1.Width = 1000 * (Slider1.Max - 1)
    'End If
    
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
Else
    If KeyCode = vbKeyAdd Then
        btnPreview_Click (1)
        DoCenter
    ElseIf KeyCode = vbKeySubtract Then
        btnPreview_Click (2)
        DoCenter
    End If
    
    If KeyCode = vbKeyPageUp Then
        vscPreview.Value = vscPreview.Min
    ElseIf KeyCode = vbKeyPageDown Then
         vscPreview.Value = vscPreview.Max
    End If
End If
End Sub

Private Sub Form_Load()

    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    bLoad = True
    FillCboPercent
    ScalePercent = 100
    WindowState = vbMaximized
    bLoad = False
    
    intLeftMargin = 320
    intRightMargin = 560
    intTopMargin = 440
    intBottomMargin = 600
    
End Sub

Private Sub Form_Resize()

Call DoCenter

If Me.Width < 10560 Then
    If Not Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
        Me.Width = 10560
    End If
End If

    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    
    picToolbar.Move 80, 40, Me.ScaleWidth - 100
    cmdClose.Left = picToolbar.Width - 980
 
    'If 1000 * (Slider1.Max - 1) > picToolbar.Width - 5160 Then
        'Slider1.Width = picToolbar.Width - 5160
    'Else
        'Slider1.Width = 1000 * (Slider1.Max - 1)
    'End If
    
    picPageNum.Left = picChild.Left + (picChild.Width / 2) - (picPageNum.Width / 2)
    
    picV1(0).Left = picChild.Left + picChild.Width - (intRightMargin * (ScalePercent / 100))
    picV1(0).Top = picChild.Top
    picV1(0).Height = picChild.Height
    Line1(0).Y2 = picV1(0).Height
        
    picV1(1).Left = picChild.Left + intLeftMargin * (ScalePercent / 100)
    picV1(1).Top = picChild.Top
    picV1(1).Height = picChild.Height
    Line1(1).Y2 = picV1(1).Height
    
    picH(0).Left = picChild.Left + 10
    picH(0).Top = picChild.Top + (intTopMargin * (ScalePercent / 100))
    picH(0).Width = picChild.Width - 20
    Line2(0).X2 = picH(0).Width
        
    picH(1).Left = picChild.Left + 10
    picH(1).Top = picChild.Top + picChild.Height - (intBottomMargin * (ScalePercent / 100))
    picH(1).Width = picChild.Width - 20
    Line2(1).X2 = picH(1).Width

    
    With tabPreview
        .Move lBorder - 20, ScaleHeight - .Height - lBorder - 350, ScaleWidth - (2 * lBorder) + 80
        picParent.Move lBorder - 20, lBorder + picToolbar.Height, ScaleWidth - (2 * lBorder) + 80, ScaleHeight - .Height - picToolbar.Height - (2 * lBorder) - 350
        '.Move picParent.Left + picParent.Width, picParent.Top
    End With
    
End Sub

Private Sub hscPreview_Change()
    picChild.Left = (-hscPreview.Value)
    Label1.Left = (-hscPreview.Value) + (80 * (ScalePercent / 100))
    picV1(1).Left = picChild.Left + (intLeftMargin * (ScalePercent / 100))
    picV1(0).Left = picChild.Left + picChild.Width - (intRightMargin * ((ScalePercent / 100)))
    
    picH(0).Left = picChild.Left
    picH(1).Left = picChild.Left
    
    picPageNum.Left = picChild.Left + (picChild.Width / 2) - (picPageNum.Width / 2)
End Sub

Private Sub hscPreview_Scroll()
    Call hscPreview_Change
End Sub

Private Sub picParent_Resize()
On Error Resume Next

If Me.Width < 12000 Then
    Me.Width = 12000
End If

If Me.Height < 8000 Then
    Me.Height = 8000
End If
    
    Dim iCount As Integer
    
    With picParent
    
        vscPreview.Move .ScaleLeft + .ScaleWidth - vscPreview.Width, 20 + .ScaleTop, vscPreview.Width, .ScaleHeight - hscPreview.Height


        hscPreview.Move 20, .ScaleHeight - hscPreview.Height, .ScaleWidth - vscPreview.Width

        imgCorner.Move vscPreview.Left, hscPreview.Top
    End With
    ResizeScrollBars
    
End Sub

Private Sub Slider1_Change()
If Not tabPreview.Tabs(Slider1.Value).Selected = True Then
    tabPreview.Tabs(Slider1.Value).Selected = True
    DoCenter
    'Slider1.ToolTipText = "Page " & Slider1.Value
End If

    Label2.Caption = "Page " & Slider1.Value & " of " & Slider1.Max
    Select Case Slider1.Value
        Case Slider1.Min
            cmdPrevious.Enabled = False
            cmdStart.Enabled = False
            cmdNext.Enabled = True
            cmdEnd.Enabled = True
        Case Slider1.Max
            cmdPrevious.Enabled = True
            cmdStart.Enabled = True
            cmdNext.Enabled = False
            cmdEnd.Enabled = False
        Case Else
            cmdPrevious.Enabled = True
            cmdStart.Enabled = True
            cmdNext.Enabled = True
            cmdEnd.Enabled = True
    End Select

End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Slider1.ToolTipText = "Page " & Slider1.Value
End Sub

Private Sub Slider1_Scroll()
Call Slider1_Change
End Sub

Private Sub tabPreview_Click()

'LockWindowUpdate Me.hWnd

Call Form_Resize

    vscPreview.Value = 0
    hscPreview.Value = 0
    Slider1.Value = tabPreview.SelectedItem.Index

    With picPreview(tabPreview.SelectedItem.Index - 1)
        .Picture = .Image
        picChild.Picture = .Picture
        PictureShow
    End With
    
'LockWindowUpdate 0

End Sub

Private Sub vscPreview_Change()
    picChild.Top = (-vscPreview.Value)
    picture1.Top = picChild.Top + (80 * ((ScalePercent / 100)))
    picH(0).Top = picChild.Top + (intTopMargin * ((ScalePercent / 100)))
    picH(1).Top = picChild.Top + picChild.Height - (intBottomMargin * ((ScalePercent / 100)))
    
    picPageNum.Top = picChild.Top + 100
    
    picV1(0).Top = picChild.Top
    picV1(1).Top = picChild.Top
End Sub

Private Sub vscPreview_Scroll()
Call vscPreview_Change
End Sub

Private Function DoCenter()
    
    Dim i As Integer

    For i = 0 To picPreview.Count - 1
        If Not picChild.Width > Me.Width Then
            picPreview(i).Left = (Me.Width / 2) - (picChild.Width / 2)
        Else
            picPreview(i).Left = 0
        End If
        If Not picChild.Height > Me.Height - 1000 Then
            picPreview(i).Top = (Me.Height / 2) - (picChild.Height / 2) - 910
        Else
            picPreview(i).Top = 0
        End If
    Next i
    
    picChild.Left = picPreview(0).Left
    picChild.Top = picPreview(0).Top
    
    With picChild
        picture1.Move .Left + .Width - 20, .Top + (80 * ScalePercent / 100), (120 * ScalePercent / 100), .Height
        Label1.Move .Left + (80 * ScalePercent / 100), .Top + .Height, .Width, ((80 + 10) * ScalePercent / 100)
   End With

End Function

