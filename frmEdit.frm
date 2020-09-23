VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VERY simple demo of msEditor"
   ClientHeight    =   6060
   ClientLeft      =   2160
   ClientTop       =   2265
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOpen 
      Caption         =   "Open Visible"
      Height          =   225
      Left            =   2670
      TabIndex        =   4
      Top             =   60
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Save Visible"
      Height          =   225
      Left            =   1380
      TabIndex        =   3
      Top             =   60
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox chkNew 
      Caption         =   "New Visible"
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   5580
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   480
      Left            =   60
      Top             =   5550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   1
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Project1.msEditor msEditor1 
      Height          =   5175
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   10095
      _extentx        =   17806
      _extenty        =   8229
      backcolor       =   12648447
      enabled         =   0   'False
      font            =   "frmEdit.frx":0000
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim adoNotes  As ADODB.Recordset       ' optional if you use code to open and connect to a database
'Dim dbMain As ADODB.Connection         ' same here

Private Sub chkNew_Click()
   msEditor1.New_Visible = chkNew.Value
End Sub

Private Sub chkOpen_Click()
   msEditor1.Open_Visible = chkOpen.Value
End Sub

Private Sub chkSave_Click()
   msEditor1.Save_Visible = chkSave.Value
End Sub

Private Sub Command1_Click()
  ' Dirty programming to demonstrate the control
  Adodc1.Recordset.AddNew
End Sub

Private Sub Form_Load()
    
   ' Use can use code to connect to a database..........
   
   ' Set dbMain = New ADODB.Connection
   ' dbMain.CursorLocation = adUseClient
   ' dbMain.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbNotes.mdb" & ";Persist Security Info=False"
       
   '...... or (for convenience here) use the Adocontrol
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbNotes.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "tblNotes"
   
   ' The same for the recordset. You can use code as the commented code below and forget the Adocontrol
   
   ' Set adoNotes = New ADODB.Recordset
   ' adoNotes.Open "SELECT * FROM tblNotes", dbMain, adOpenDynamic, adLockOptimistic
    
    Set msEditor1.mDataSource = Adodc1  ' for code you have to use the 'adoNotes' instead of the Adodc1
    msEditor1.MaxLength = 2000          ' Optional maximum length.
    
    msEditor1.mDataField = "nNotes"     ' The field of your text
    msEditor1.Enabled = True            ' I have it default false so set it to true
    
End Sub
