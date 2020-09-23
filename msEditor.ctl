VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl msEditor 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   ScaleHeight     =   4020
   ScaleWidth      =   7260
   ToolboxBitmap   =   "msEditor.ctx":0000
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   6180
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0312
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0424
            Key             =   "Font"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0536
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0648
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":075A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":086C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":097E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0A90
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0BA2
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0CB4
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0DC6
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0ED8
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":0FEA
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":10FC
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":120E
            Key             =   "StrikeThru"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1320
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1642
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1754
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1866
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1978
            Key             =   "DecreaseIndent"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor.ctx":1D82
            Key             =   "IncreaseIndent"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5940
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox msEdit 
      Height          =   2205
      Left            =   390
      TabIndex        =   0
      Top             =   870
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3889
      _Version        =   393217
      BackColor       =   12648447
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"msEditor.ctx":218C
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open a file"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThru"
            Object.ToolTipText     =   "StrikeThru"
            ImageKey        =   "StrikeThru"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "Color"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font type and Size"
            ImageKey        =   "Font"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Left"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Right"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DecreaseIndent"
            Object.ToolTipText     =   "Decrease Indent"
            ImageKey        =   "DecreaseIndent"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IncreaseIndent"
            Object.ToolTipText     =   "Increase Indent"
            ImageKey        =   "IncreaseIndent"
         EndProperty
      EndProperty
   End
   Begin VB.Menu sad 
      Caption         =   "gsdfgsdfg"
   End
   Begin VB.Menu ss 
      Caption         =   "ee"
   End
End
Attribute VB_Name = "msEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_PASTE = &H302

Private mmDataSource As DataSource
Private mnuBold As Boolean
Private mnuItalic As Boolean
Private mnuUnderline As Boolean
Private mnuStrikeThru As Boolean

Private DocumentName As String
Private DefaultFont As String
Private DefaultFontSize As Integer
Private DefaultTextColor As String
Private DefaultBackgroundColor As String
Private DefaultBold As Boolean
Private DefaultItalic As Boolean
Private DefaultUnderline As Boolean
Private DefaultStrikeThru As Boolean
Private DefaultAlignment As String
Private IndentSize As Integer

Private mDocumentName As String
Private mDefaultFont As String
Private mDefaultFontSize As String
Private mDefaultTextColor As String
Private mDefaultBackgroundColor As String
Private mDefaultBold As String
Private mDefaultItalic As String
Private mDefaultUnderline As String
Private mDefaultStrikeThru As String
Private mDefaultAlignment As String
Private mItentSize As String

Private Cancelled As Boolean
Private Saved As Boolean

Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
'Default Property Values:
Const m_def_mDataField = "0"
Const m_def_Enabled = True
Const m_def_ForeColor = 0
Const m_def_FontUnderline = 0
Const m_def_FontStrikethru = 0
Const m_def_FontSize = 0
Const m_def_FontName = ""
Const m_def_FontItalic = 0
Const m_def_FontBold = 0
Const m_def_TextRTF = "0"
Const m_def_Open_Visible = True
Const m_def_Save_Visible = True
Const m_def_New_Visible = True

'Property Variables:
Dim m_mDataField As String
Dim m_Enabled As Boolean
Dim m_ForeColor As Long
Dim m_FontUnderline As Boolean
Dim m_FontStrikethru As Boolean
Dim m_FontSize As Single
Dim m_FontName As String
Dim m_FontItalic As Boolean
Dim m_FontBold As Boolean
Dim m_TextRTF As String

Enum VisEn
        mSave = 0
        mNew = 1
        mOpen = 2
        mCut = 3
        mCopy = 4
        mPaste = 5
        mUndo = 6
        mRedo = 7
        mColor = 8
        mFont = 9
        mStrikeThru = 10
        mUnderline = 11
        mItalic = 12
        mBold = 13
        mLeft = 14
        mCenter = 15
        mRight = 16
        mIncreaseIndent = 17
        mDecreaseIndent = 18
End Enum

Private Sub UserControl_Initialize()
        
    mDocumentName = "NewPage"
    mDefaultFont = "Tahoma"
    mDefaultFontSize = 10
    mDefaultTextColor = 0
    mDefaultBackgroundColor = &HC0FFFF
    mDefaultBold = False
    mDefaultItalic = False
    mDefaultUnderline = False
    mDefaultStrikeThru = False
    mDefaultAlignment = "Left"
    mItentSize = 500
    
    DocumentName = mDocumentName
    DefaultFont = mDefaultFont
    DefaultFontSize = mDefaultFontSize
    DefaultTextColor = mDefaultTextColor
    DefaultBackgroundColor = mDefaultBackgroundColor
    DefaultBold = mDefaultBold
    DefaultItalic = mDefaultItalic
    DefaultUnderline = mDefaultUnderline
    DefaultStrikeThru = mDefaultStrikeThru
    DefaultAlignment = mDefaultAlignment
    IndentSize = mItentSize
        
    Saved = False
    Cancelled = False
    
    With tlbar
    
        .Buttons("Cut").Enabled = False
        .Buttons("Copy").Enabled = False
        .Buttons("Undo").Enabled = False
        .Buttons("Redo").Enabled = False
        .Buttons("Paste").Enabled = False
        .Buttons("Save").Enabled = False
        
    End With
 
    With msEdit
       
        .SelFontSize = DefaultFontSize
        .SelFontName = DefaultFont
        .SelColor = DefaultTextColor
        .BackColor = DefaultBackgroundColor
        .SelBold = DefaultBold
        .SelItalic = DefaultItalic
        .SelUnderline = DefaultUnderline
        .SelStrikeThru = DefaultStrikeThru
        
    End With

End Sub

Private Sub mnuBackgroundColor_Click()
    CommonDialog1.ShowColor
    msEdit.BackColor = CommonDialog1.Color
End Sub

Private Sub msBold()
    If mnuBold Then
        msEdit.SelBold = False
        mnuBold = False
        tlbar.Buttons("Bold").Value = tbrUnpressed
    Else
        msEdit.SelBold = True
        mnuBold = True
        tlbar.Buttons("Bold").Value = tbrPressed
    End If
End Sub

Private Sub msCenter()
    
    msEdit.SelAlignment = 2
    
    tlbar.Buttons("Left").Value = tbrUnpressed
    tlbar.Buttons("Center").Value = tbrPressed
    tlbar.Buttons("Right").Value = tbrUnpressed

End Sub


Private Sub msCopy()
    Clipboard.Clear
    Clipboard.SetText msEdit.SelText, 1
    
    tlbar.Buttons("Paste").Enabled = True
    mnuPaste.Enabled = True
End Sub

Private Sub msCut()
    Clipboard.Clear
    Clipboard.SetText msEdit.SelText, 1
    msEdit.SelText = ""
    
    tlbar.Buttons("Paste").Enabled = True

End Sub

Private Sub msDecreaseIndent()
    msEdit.SelIndent = msEdit.SelIndent - IndentSize
End Sub

Sub SaveNow()
    
    If Saved = False Then
        Dim blnAnnuleren As Boolean
    
        On Error GoTo handlelit
        blnAnnuleren = False
        With CommonDialog1
            .filter = "Editor Documents (*.rtf)|*.rtf"
            .ShowSave
        End With
        
        If Not blnAnnuleren Then
            msEdit.SaveFile CommonDialog1.FileName
            Saved = True
        End If
     End If
     
     Exit Sub
    
handlelit:
    If Err.Number = cdlCancel Then
        Saved = False
        Resume Next
    Else
            MsgBox Err.Description
    End If
    
    
End Sub


Private Sub msFontColor()
    On Error Resume Next
    CommonDialog1.ShowColor
    msEdit.SelColor = CommonDialog1.Color
End Sub
Private Sub msFont()
    On Error Resume Next
    
    CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cdlCFEffects
                                                      
    With msEdit
        CommonDialog1.FontName = .SelFontName
        CommonDialog1.FontSize = .SelFontSize
        CommonDialog1.FontBold = .SelBold
        CommonDialog1.FontStrikethru = .SelStrikeThru
        CommonDialog1.FontUnderline = .SelUnderline
        CommonDialog1.FontItalic = .SelItalic
        CommonDialog1.Color = .SelColor
    End With
    
    CommonDialog1.ShowFont
    
    With msEdit
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
        .SelStrikeThru = CommonDialog1.FontStrikethru
        .SelUnderline = CommonDialog1.FontUnderline
        .SelColor = CommonDialog1.Color
    End With
     
End Sub

Private Sub msIncreaseIndent()
    msEdit.SelIndent = msEdit.SelIndent + IndentSize
End Sub

Private Sub msItalic()
    If mnuItalic Then
        msEdit.SelItalic = False
        mnuItalic = False
        tlbar.Buttons("Italic").Value = tbrUnpressed
    Else
        msEdit.SelItalic = True
        mnuItalic = True
        tlbar.Buttons("Italic").Value = tbrPressed
    End If
End Sub

Private Sub msLeft()
    msEdit.SelAlignment = 0
    
    tlbar.Buttons("Left").Value = tbrPressed
    tlbar.Buttons("Center").Value = tbrUnpressed
    tlbar.Buttons("Right").Value = tbrUnpressed

End Sub

Private Sub msNew()
    
    With tlbar
        .Buttons("Cut").Enabled = True
        .Buttons("Copy").Enabled = True
        .Buttons("Undo").Enabled = True
        .Buttons("Redo").Enabled = True
        .Buttons("Paste").Enabled = True
        .Buttons("Save").Enabled = True
    End With
    
    msEdit.Text = ""
    
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub msPaste()
    
    msEdit.SelText = Clipboard.GetText(1)
    tlbar.Buttons("Save").Enabled = True
    
End Sub

Private Sub msRedo()
    
    Redo
    tlbar.Buttons("Undo").Enabled = True
    
End Sub

Private Sub msRight()
    msEdit.SelAlignment = 1
    
    tlbar.Buttons("Left").Value = tbrUnpressed
    tlbar.Buttons("Center").Value = tbrUnpressed
    tlbar.Buttons("Right").Value = tbrPressed
    
End Sub

Private Sub mnuSelectAll_Click()
    
    msEdit.SelStart = 0
    msEdit.SelLength = Len(msEdit.Text)
    msEdit.SetFocus
    
End Sub
Private Sub msStrikeThrough()
    If mnuStrikeThru Then
        msEdit.SelStrikeThru = False
        mnuStrikeThru = False
        tlbar.Buttons("StrikeThru").Value = tbrUnpressed
    Else
        msEdit.SelStrikeThru = True
        mnuStrikeThru = True
        tlbar.Buttons("StrikeThru").Value = tbrPressed
    End If
End Sub


Private Sub msUnderline()
    If mnuUnderline Then
        msEdit.SelUnderline = False
        mnuUnderline = False
        tlbar.Buttons("Underline").Value = tbrUnpressed
    Else
        msEdit.SelUnderline = True
        mnuUnderline = True
        tlbar.Buttons("Underline").Value = tbrPressed
    End If
End Sub
Sub SetButtons()
    
    If msEdit.SelBold = True Then
        mnuBold = True
        tlbar.Buttons("Bold").Value = tbrPressed
    Else
        mnuBold = False
        tlbar.Buttons("Bold").Value = tbrUnpressed
    End If
    
    If msEdit.SelItalic = True Then
        mnuItalic = True
        tlbar.Buttons("Italic").Value = tbrPressed
    Else
        mnuItalic = False
        tlbar.Buttons("Italic").Value = tbrUnpressed
    End If
    
    If msEdit.SelUnderline = True Then
        mnuUnderline = True
        tlbar.Buttons("Underline").Value = tbrPressed
    Else
        mnuUnderline = False
        tlbar.Buttons("Underline").Value = tbrUnpressed
    End If
    
    If msEdit.SelStrikeThru = True Then
        mnuStrikeThru = True
        tlbar.Buttons("StrikeThru").Value = tbrPressed
    Else
        mnuStrikeThru = False
        tlbar.Buttons("StrikeThru").Value = tbrUnpressed
    End If
    
    If msEdit.SelAlignment = 0 Then
        tlbar.Buttons("Left").Value = tbrPressed
        tlbar.Buttons("Center").Value = tbrUnpressed
        tlbar.Buttons("Right").Value = tbrUnpressed
    Else
        If msEdit.SelAlignment = 1 Then
            tlbar.Buttons("Left").Value = tbrUnpressed
            tlbar.Buttons("Center").Value = tbrUnpressed
            tlbar.Buttons("Right").Value = tbrPressed
        Else
            tlbar.Buttons("Left").Value = tbrUnpressed
            tlbar.Buttons("Center").Value = tbrPressed
            tlbar.Buttons("Right").Value = tbrUnpressed
        End If
        
    End If
    
End Sub
Private Sub msUndo()
    Undo
    tlbar.Buttons("Redo").Enabled = True
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Save"
             SaveNow
        Case "New"
            msNew
        Case "Open"
            msOpen
        Case "Cut"
            msCut
        Case "Copy"
            msCopy
        Case "Paste"
            msPaste
        Case "Undo"
            msUndo
        Case "Redo"
            msRedo
        Case "Color"
            msFontColor
        Case "Font"
            msFont
        Case "StrikeThru"
            msStrikeThrough
        Case "Underline"
            msUnderline
        Case "Italic"
            msItalic
        Case "Bold"
            msBold
        Case "Left"
            msLeft
        Case "Center"
            msCenter
        Case "Right"
            msRight
        Case "IncreaseIndent"
            msIncreaseIndent
        Case "DecreaseIndent"
            msDecreaseIndent
        End Select
End Sub

Private Sub tlbar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuView
    End If
End Sub

Public Sub Undo()
    'if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    msEdit.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub
Public Sub Redo()
    'This is the basic redo
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    msEdit.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub msEdit_Change()
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = msEdit.TextRTF
    End If
    Saved = False
    tlbar.Buttons("Save").Enabled = True
 
End Sub

Private Sub msEdit_Click()
    SetButtons
End Sub

Private Sub msEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    If msEdit.SelText <> "" Then
        tlbar.Buttons("Copy").Enabled = True
        tlbar.Buttons("Cut").Enabled = True
    Else
        tlbar.Buttons("Copy").Enabled = False
        tlbar.Buttons("Cut").Enabled = False
    End If
    
    SetButtons
End Sub

Private Sub msEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If msEdit.SelText <> "" Then
        tlbar.Buttons("Copy").Enabled = True
        tlbar.Buttons("Cut").Enabled = True
    Else
        tlbar.Buttons("Copy").Enabled = False
        tlbar.Buttons("Cut").Enabled = False
    End If
    
    SetButtons
    
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    msEdit.Move 0, 0 + tlbar.Height, UserControl.ScaleWidth, UserControl.ScaleHeight - tlbar.Height
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
    BackColor = msEdit.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    msEdit.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = msEdit.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    On Error Resume Next
    msEdit.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = m_FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    m_FontStrikethru = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = m_FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
    PropertyChanged "FontItalic"
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = msEdit.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set msEdit.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
    Text = msEdit.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    msEdit.Text() = New_Text
    PropertyChanged "Text"
End Property
Public Property Get TextRTF() As String
    TextRTF = msEdit.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    msEdit.TextRTF = TextRTF
    PropertyChanged "TextRTF"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_ForeColor = m_def_ForeColor
    m_FontUnderline = m_def_FontUnderline
    m_FontStrikethru = m_def_FontStrikethru
    m_FontSize = m_def_FontSize
    m_FontName = m_def_FontName
    m_FontItalic = m_def_FontItalic
    m_FontBold = m_def_FontBold
    m_TextRTF = m_def_TextRTF
    m_mDataField = m_def_mDataField
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    msEdit.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    msEdit.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_FontUnderline = PropBag.ReadProperty("FontUnderline", m_def_FontUnderline)
    m_FontStrikethru = PropBag.ReadProperty("FontStrikethru", m_def_FontStrikethru)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontItalic = PropBag.ReadProperty("FontItalic", m_def_FontItalic)
    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    Set msEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)
    msEdit.Text = PropBag.ReadProperty("Text", "")
    m_TextRTF = PropBag.ReadProperty("TextRTF", m_def_TextRTF)
    Set mmDataSource = PropBag.ReadProperty("mDataSource", Nothing)
    
    m_mDataField = PropBag.ReadProperty("mDataField", m_def_mDataField)
    msEdit.MaxLength = PropBag.ReadProperty("MaxLength", 0)
        
    tlbar.Buttons("Open").Visible = PropBag.ReadProperty("Open_Visible", m_def_Open_Visible)
    tlbar.Buttons("Save").Visible = PropBag.ReadProperty("Save_Visible", m_def_Save_Visible)
    tlbar.Buttons("New").Visible = PropBag.ReadProperty("New_Visible", m_def_New_Visible)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", msEdit.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", msEdit.BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("FontUnderline", m_FontUnderline, m_def_FontUnderline)
    Call PropBag.WriteProperty("FontStrikethru", m_FontStrikethru, m_def_FontStrikethru)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    Call PropBag.WriteProperty("FontItalic", m_FontItalic, m_def_FontItalic)
    Call PropBag.WriteProperty("FontBold", m_FontBold, m_def_FontBold)
    Call PropBag.WriteProperty("Font", msEdit.Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", msEdit.Text, "")
    Call PropBag.WriteProperty("TextRTF", m_TextRTF, m_def_TextRTF)
    Call PropBag.WriteProperty("mDataSource", mmDataSource, Nothing)
    
    Call PropBag.WriteProperty("mDataField", m_mDataField, m_def_mDataField)
    Call PropBag.WriteProperty("MaxLength", msEdit.MaxLength, 0)
    
    Call PropBag.WriteProperty("Open_Visible", tlbar.Buttons("Open").Visible, m_def_Open_Visible)
    Call PropBag.WriteProperty("Save_Visible", tlbar.Buttons("Save").Visible, m_def_Save_Visible)
    Call PropBag.WriteProperty("New_Visible", tlbar.Buttons("New").Visible, m_def_New_Visible)
    
End Sub
Public Property Get mDataSource() As DataSource
Attribute mDataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set mDataSource = msEdit.DataSource
End Property

Public Property Set mDataSource(ByVal New_mDataSource As DataSource)
    Set msEdit.DataSource = New_mDataSource
    PropertyChanged "mDataSource"
End Property
Public Property Get mDataField() As String
    mDataField = msEdit.DataField
End Property

Public Property Let mDataField(ByVal New_mDataField As String)
    m_mDataField = New_mDataField
    msEdit.DataField = m_mDataField
    PropertyChanged "mDataField"
End Property
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets a value indicating whether there is a maximum number of characters a RichTextBox control can hold and, if so, specifies the maximum number of characters."
    MaxLength = msEdit.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    msEdit.MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Sub msOpen()
    
    On Error GoTo errHandling

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .filter = "Text Documents|*.txt;*.rtf"
        .ShowOpen
    End With
    
    msEdit.LoadFile CommonDialog1.FileName
    
    With tlbar
    
        .Buttons("Cut").Enabled = False
        .Buttons("Copy").Enabled = False
        .Buttons("Undo").Enabled = False
        .Buttons("Redo").Enabled = False
        .Buttons("Paste").Enabled = False
        .Buttons("Save").Enabled = False
        
    End With
   
    Exit Sub

errHandling:

    If Err.Number <> cdlCancel Then
        MsgBox Err.Description
    End If
End Sub
Public Property Get Open_Visible() As Boolean
    Open_Visible = tlbar.Buttons("Open").Visible
End Property

Public Property Let Open_Visible(ByVal New_Open_Visible As Boolean)
    tlbar.Buttons("Open").Visible = New_Open_Visible
    PropertyChanged "Open_Visible"
End Property
Public Property Get Save_Visible() As Boolean
    Save_Visible = tlbar.Buttons("Save").Visible
End Property

Public Property Let Save_Visible(ByVal New_Save_Visible As Boolean)
    tlbar.Buttons("Save").Visible = New_Save_Visible
    PropertyChanged "Save_Visible"
End Property
Public Property Get New_Visible() As Boolean
    New_Visible = tlbar.Buttons("New").Visible
End Property
Public Property Let New_Visible(ByVal New_New_Visible As Boolean)
    tlbar.Buttons("New").Visible = New_New_Visible
    PropertyChanged "New_Visible"
End Property

