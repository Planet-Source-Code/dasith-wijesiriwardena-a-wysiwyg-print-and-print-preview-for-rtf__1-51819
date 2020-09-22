VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "Form1.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Turn Off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Turn On"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11456
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0272
   End
   Begin VB.Label Label1 
      Caption         =   "WYSIWYG Display:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
LineWidth = WYSIWYG_RTF(RichTextBox1, 1440, 1440, True)
End Sub

Private Sub Command3_Click()
Dim LineWidth As Long
LineWidth = WYSIWYG_RTF(RichTextBox1, 1440, 1440, False)
End Sub

Private Sub Command4_Click()
PrintPreview RichTextBox1, 1400, 1400, 1400, 1400, Printer.Orientation
End Sub

Private Sub Command5_Click()
MsgBox Text1.Text, vbInformation
End Sub

   Private Sub Form_Load()
      Dim LineWidth As Long
      
      RichTextBox1.LoadFile App.Path & "\file1.rtf"

      ' Initialize Form and Command button
      Me.Caption = "Rich Text Box WYSIWYG Printing Example by Dasith Wijesiriwardena. dasiths@hotmail.com"

      ' Set the font of the RTF to a TrueType font for best results

      ' Tell the RTF to base it's display off of the printer
      LineWidth = WYSIWYG_RTF(RichTextBox1, 1440, 1440, True)
      '1440 Twips=1 Inch

      ' Set the form width to match the line width
      On Error Resume Next
      Me.Width = LineWidth + 200
      Me.WindowState = vbMaximized
   End Sub

   Private Sub Form_Resize()
      ' Position the RTF on form
      If Not Me.WindowState = vbMinimized Then
        RichTextBox1.Move 100, 500, Me.ScaleWidth - 200, _
            Me.ScaleHeight - 600
      End If
   End Sub

   Public Sub Command1_Click()
      ' Print the contents of the RichTextBox with a one inch margin
      On Error GoTo err1
      lngPrinterWidth = Printer.Width
      PrintRTF RichTextBox1, 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch
      Exit Sub
err1:
    Select Case Err.Number
        Case 482
            MsgBox "Make sure that you have a printer installed.  If a " & _
                "printer is installed, go into your printer properties " & _
                "look under the Setup tab, and make sure the ICM checkbox " & _
                "is checked and try printing again.", , "Printer Error"
            Exit Sub
        Case Else
            MsgBox Err.Number & " " & Err.Description
    End Select
End Sub


