Attribute VB_Name = "modPrinting"

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

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long
    cpMax As Long
End Type

Private Type FormatRange
    hdc As Long
    hdcTarget As Long
    rc As Rect
    rcPage As Rect
    chrg As CharRange
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, ByVal nIndex As Long) As Long


Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
    lp As Any) As Long

Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
    (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
    ByVal lpOutput As Long, ByVal lpInitData As Long) As Long



Public Sub PrintPreview(RTF As RichTextBox, LeftMarginWidth As Currency, _
    TopMarginHeight As Currency, RightMarginWidth As Currency, BottomMarginHeight As Currency, _
    pgOrientation As Integer)
      
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    Dim iCount As Integer

    On Error GoTo ErrHandle
    
'Set the orientation of the printer
    Printer.Orientation = pgOrientation
    Printer.ScaleMode = vbTwips

' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = CLng(LeftMarginWidth - LeftOffset)
    TopMargin = CLng(TopMarginHeight - TopOffset)
    RightMargin = CLng((Printer.Width - RightMarginWidth) - LeftOffset)
    BottomMargin = CLng((Printer.Height - BottomMarginHeight) - TopOffset)

' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin


    frmPreview.SizePreview Printer.Width, Printer.Height

    fr.hdc = frmPreview.picPreview(0).hdc
    fr.hdcTarget = frmPreview.picPreview(0).hdc
    fr.rc = rcDrawTo
    fr.rcPage = rcPage
    fr.chrg.cpMin = 0
    fr.chrg.cpMax = -1


    TextLength = Len(RTF.Text)

    Dim iPage As Integer
    
    iPage = 1
    
    Do
        With frmPreview
            If iPage > 1 Then
                .AddPage iPage
                fr.hdc = .picPreview(iPage - 1).hdc
                fr.hdcTarget = .picPreview(iPage - 1).hdc
            End If
            .picPreview(iPage - 1).Print
        End With

        NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        
        iPage = iPage + 1
    Loop

    r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))

    Printer.KillDoc
    Printer.EndDoc
    
    frmPreview.Show

    Exit Sub
    
ErrHandle:
    Select Case Err.Number
        Case 482
            MsgBox "Make sure that you have a printer installed.  If a " & _
                "printer is installed, go into your printer properties " & _
                "look under the Setup tab, and make sure the ICM checkbox " & _
                "is checked and try printing again.", , "Printer Error"
            Exit Sub
        Case Else
            MsgBox Err.Number & " " & Err.Description
            Resume Next
    End Select
    
End Sub


