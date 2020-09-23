VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   2010
   ClientTop       =   3765
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12165
   Begin MSComctlLib.ListView lvHeader 
      Height          =   2955
      Left            =   6120
      TabIndex        =   16
      Top             =   5310
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Element"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LIST"
      Height          =   345
      Left            =   5040
      TabIndex        =   14
      Top             =   780
      Width           =   825
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UIDL"
      Height          =   345
      Left            =   5910
      TabIndex        =   13
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HEADER"
      Height          =   345
      Left            =   6870
      TabIndex        =   12
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOP"
      Height          =   345
      Left            =   5040
      TabIndex        =   11
      Top             =   1140
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETR"
      Height          =   345
      Left            =   5910
      TabIndex        =   10
      Top             =   780
      Width           =   915
   End
   Begin VB.CheckBox chkAPOP 
      Caption         =   "Use APOP"
      Height          =   225
      Left            =   2520
      TabIndex        =   9
      Top             =   270
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdUIDL 
      Caption         =   "UIDL"
      Height          =   435
      Left            =   2640
      TabIndex        =   8
      Top             =   1050
      Width           =   2355
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   8400
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   6870
      TabIndex        =   6
      Text            =   "1"
      Top             =   1170
      Width           =   975
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1050
      Width           =   2355
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H00FFC0C0&
      Height          =   3315
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5040
      Width           =   5745
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2490
      TabIndex        =   3
      Text            =   "mypasswod"
      Top             =   525
      Width           =   2475
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Text            =   "mrenigma"
      Top             =   510
      Width           =   2325
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "myemail"
      Top             =   150
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Connect"
      Height          =   375
      Left            =   5010
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
   Begin MSComctlLib.ListView lvMessages 
      Height          =   3465
      Left            =   180
      TabIndex        =   15
      Top             =   1530
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   6112
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6150
      TabIndex        =   17
      Top             =   5070
      Width           =   420
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cPOPTest As clsPOP3
Attribute cPOPTest.VB_VarHelpID = -1
Dim bCollecting As Boolean

Private Sub cmdGo_Click()
Dim sReturn As String
Dim Records As Long
Dim TotalSize As Double

      cPOPTest.TimeOut = 10
      cPOPTest.QUIT
      If cPOPTest.Connect(Me.chkAPOP.Value, Me.txtServer, Me.txtUsername, Me.txtPassword, 110) = True Then
         ' Connected Get Stat
         sReturn = cPOPTest.STAT
         If sReturn <> "" Then
            Records = Split(sReturn, " ")(0)
            TotalSize = Split(sReturn, " ")(1)
         
         End If
         PrintStatus vbCrLf & "Total Emails:" & Records
         PrintStatus "Total Mailbox Size:" & Format(TotalSize / 1024, 0) & " KB"

      End If

End Sub

Private Sub cmdList_Click()
Dim sResult As String
Dim asMessages() As String
Dim liD As Long
Dim lSize As Long
Dim sSize As String
Dim lMessageCount As Long
Dim i As Long
Dim HeaderItem As Email_Header
Dim xNode As Node

      If cPOPTest.IsConnected = True Then
         bCollecting = True
         sResult = cPOPTest.STAT
         lMessageCount = Split(sResult, " ")(0)
         sResult = cPOPTest.List

         If lMessageCount > 0 Then
            ' Messages
            Me.lvMessages.ListItems.Clear
            Me.PB.Max = lMessageCount
         
            asMessages = Split(sResult, vbCrLf)
            For i = 1 To lMessageCount
               Me.PB.Value = i
            
               Set HeaderItem = cPOPTest.header(i)
               lSize = Split(asMessages(i - 1), " ")(1)
               
               If lSize > 1024 Then
                  sSize = Format(lSize / 1024, ".##") & " kb"
               Else
                  sSize = lSize & " b"
               End If
               
               Me.lvMessages.ListItems.Add i, "M" & i, i
               Me.lvMessages.ListItems(i).SubItems(1) = HeaderItem.ElementValue("SUBJECT")
               Me.lvMessages.ListItems(i).SubItems(2) = HeaderItem.ElementValue("From")
               Me.lvMessages.ListItems(i).SubItems(3) = HeaderItem.ElementValue("date")
               Me.lvMessages.ListItems(i).SubItems(4) = sSize
               
               ' Stop
            Next
         End If
         ' Me.vsMessages.AutoSize 0, Me.vsMessages.Cols - 1
      
         bCollecting = False
      End If
End Sub

Private Sub cmdUIDL_Click()
      Debug.Print cPOPTest.UIDL
End Sub

Private Sub Command1_Click()
Dim sMessage As String

      sMessage = cPOPTest.RETR(Me.txtID.Text)
      ' Me.PB.Value = Me.PB.Max
      Me.txtResult = Left$(sMessage, 2000) ' Output first 2000 bytes
      ' Stop
End Sub

Private Sub Command2_Click()
      Debug.Print cPOPTest.Top(Me.txtID.Text, 99)
End Sub

Private Sub Command3_Click()
Dim HeaderItem As Email_Header
Dim i As Long


      Set HeaderItem = cPOPTest.header(Me.txtID.Text)
      Me.lvHeader.ListItems.Clear
      ' Stop
      If cPOPTest.GetLastError = "" Then
         For i = 1 To HeaderItem.ElementCount
            Me.lvHeader.ListItems.Add i, , HeaderItem.ElementNameFromIndex(i)
            Me.lvHeader.ListItems.Item(i).SubItems(1) = HeaderItem.ElementValueFromIndex(i)
         Next
         Me.lblFrom = "From:" & HeaderItem.ElementValue("From")
      Else
         Stop
      End If
      ' Stop
End Sub

Private Sub Command4_Click()
      Debug.Print cPOPTest.UIDL(Me.txtID.Text)

End Sub

Private Sub Command5_Click()
      Debug.Print cPOPTest.List(Me.txtID.Text)
End Sub

Private Sub cPOPTest_DataArrival(sData As String)
      ' If Not (bCollecting) Then
      PrintStatus "<" & sData
      ' End If
End Sub

Private Sub cPOPTest_DataSent(sData As String)
      ' If Not (bCollecting) Then
      PrintStatus ">" & sData
      ' End If
End Sub

Private Sub cPOPTest_Progress(dCurrectBytes As Double, dTotalSize As Double)
      On Error Resume Next
      Me.PB.Max = dTotalSize
      Me.PB.Value = dCurrectBytes
End Sub

Private Sub cPOPTest_POP3Error(ByVal sData As String)
      MsgBox ("Error occured - " & cPOPTest.GetLastError)
End Sub

Private Sub Form_Load()
      Set cPOPTest = New clsPOP3
      
      Me.lvMessages.ColumnHeaders(1).Width = "299"
      Me.lvMessages.ColumnHeaders(2).Width = "3839"
      Me.lvMessages.ColumnHeaders(3).Width = "4169"
      Me.lvMessages.ColumnHeaders(4).Width = "1844"
      Me.lvMessages.ColumnHeaders(5).Width = "1604"
      Me.txtServer = GetSetting("enPOP3", "Settings", "Server", "myemail")
      Me.txtUsername = GetSetting("enPOP3", "Settings", "Username", "mrenigma")
      Me.txtPassword = GetSetting("enPOP3", "Settings", "Password", "mypassword")
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
      If cPOPTest.IsConnected = True Then
         cPOPTest.QUIT
      End If
      Set cPOPTest = Nothing
      
      Call SaveSetting("enPOP3", "Settings", "Server", Me.txtServer)
      Call SaveSetting("enPOP3", "Settings", "Username", Me.txtUsername)
      Call SaveSetting("enPOP3", "Settings", "Password", Me.txtPassword)
      
End Sub
Private Sub PrintStatus(sText As String)
      Me.txtResult.Text = txtResult.Text & sText & vbCrLf
      Me.txtResult.SelStart = Len(Me.txtResult.Text)
End Sub


