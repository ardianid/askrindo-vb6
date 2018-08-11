VERSION 5.00
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form Frm_New_Seting 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_New_Seting.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1035
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   3915
         Begin VB.TextBox txtPasswd 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   600
            Width           =   2355
         End
         Begin VB.TextBox txtUsername 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Text            =   "sa"
            Top             =   240
            Width           =   2355
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            Enabled         =   0   'False
            Height          =   210
            Left            =   240
            TabIndex        =   12
            Top             =   300
            Width           =   870
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Enabled         =   0   'False
            Height          =   210
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   825
         End
      End
      Begin VB.CheckBox chkAuthentication 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Use NT Authentication"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.ComboBox Cbo_Nama_Server 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   480
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   135
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   5415
      End
      Begin IsButton_Ard.isButton Cmd_Simpan 
         Height          =   450
         Left            =   4440
         TabIndex        =   2
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   794
         Icon            =   "Frm_New_Seting.frx":8104
         Style           =   10
         Caption         =   "Connect"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin IsButton_Ard.isButton Cmd_Keluar 
         Height          =   450
         Left            =   4440
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   794
         Icon            =   "Frm_New_Seting.frx":8120
         Style           =   10
         Caption         =   "Cancel"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SETING NAMA SERVER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Server"
         Height          =   210
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Frm_New_Seting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private oSQLServer As SQLDMO.SQLServer

Private Sub Cmd_Keluar_Click()

'  If Not oSQLServer Is Nothing Then
'    oSQLServer.Disconnect
'  End If
'
'  Set oSQLServerActive = Nothing
'  Set oSQLServer = Nothing
'  Set oSQLServerDMOApp = Nothing

    Unload Me
    
End Sub

Private Sub Cmd_Simpan_Click()

On Error GoTo ErrorHandler

'Set oSQLServer = New SQLDMO.SQLServer
'
'oSQLServer.LoginTimeout = -1 '-1 is the ODBC default (60) seconds
''Connect to the Server
'If chkAuthentication Then
'  With oSQLServer
'  'Use NT Authentication
'    .LoginSecure = True
'  'Do not reconnect automatically
'    .AutoReConnect = False
'  'Now connect
'    .Connect Cbo_Nama_Server.Text
'  End With
'Else
'  With oSQLServer
'  'Use SQL Server Authentication
'    .LoginSecure = False
'  'Do not reconnect automatically
'    .AutoReConnect = False
'  'Use SQL Security
'    .Connect Cbo_Nama_Server.Text, txtUsername.Text, txtPasswd.Text
'  End With
'End If
''
'Set oSQLServerActive = oSQLServer
''MsgBox "Your Login: " & oSQLServer.Login
''Show next form

If Set_Lokasi_Database(Trim(Cbo_Nama_Server.Text)) = True Then
    
    If MsgBox("Server telah terkoneksi, anda ingin melanjutkan program", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
       Unload Me
       Utama.Show
    Else
        Cmd_Keluar_Click
    End If
End If
Exit Sub

ErrorHandler:
MsgBox "Error: " & Err.Number & " " & Err.Description, vbOKOnly, "Login Error"

End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Simpan.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Cbo_Nama_Server.SetFocus
End Sub

Private Sub Form_Load()

'    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 2, _
'                        Me.Top / 2, Me.Width / 2, _
'                      Me.Height / 2, SWP_NOACTIVATE Or SWP_SHOWWINDOW


'  Dim i As Integer
''  gbShowCommandEvents = True
'  Set oSQLServerDMOApp = New SQLDMO.Application
'
'  Dim namX As NameList
'  Set namX = oSQLServerDMOApp.ListAvailableSQLServers
'  For i = 1 To namX.Count
'   Cbo_Nama_Server.AddItem namX.Item(i)
'  Next

 ' Cbo_Nama_Server.ListIndex = 0
  
End Sub
