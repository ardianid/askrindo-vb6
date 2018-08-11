VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_penggajian2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Detail Penggajian"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
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
   ScaleHeight     =   3135
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Selesai"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Rubah"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox tjab 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox tnama 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox tkode 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_gapok 
      Height          =   345
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   609
      Calculator      =   "frm_penggajian2.frx":0000
      Caption         =   "frm_penggajian2.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian2.frx":008C
      Keys            =   "frm_penggajian2.frx":00AA
      Spin            =   "frm_penggajian2.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   12582912
      Format          =   "###,###,###"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   1028849669
      MinValueVT      =   1598423045
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_lembur 
      Height          =   345
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   609
      Calculator      =   "frm_penggajian2.frx":011C
      Caption         =   "frm_penggajian2.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian2.frx":01A8
      Keys            =   "frm_penggajian2.frx":01C6
      Spin            =   "frm_penggajian2.frx":0210
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   12582912
      Format          =   "###,###,###"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   1028849669
      MinValueVT      =   1598423045
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_pot 
      Height          =   345
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   609
      Calculator      =   "frm_penggajian2.frx":0238
      Caption         =   "frm_penggajian2.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian2.frx":02C4
      Keys            =   "frm_penggajian2.frx":02E2
      Spin            =   "frm_penggajian2.frx":032C
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   12582912
      Format          =   "###,###,###"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   1028849669
      MinValueVT      =   1598423045
   End
   Begin TDBNumber6Ctl.TDBNumber tdb_bersih 
      Height          =   345
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   609
      Calculator      =   "frm_penggajian2.frx":0354
      Caption         =   "frm_penggajian2.frx":0374
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian2.frx":03E0
      Keys            =   "frm_penggajian2.frx":03FE
      Spin            =   "frm_penggajian2.frx":0448
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###;;0"
      EditMode        =   0
      Enabled         =   0
      ErrorBeep       =   0
      ForeColor       =   12582912
      Format          =   "###,###,###"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   1028849669
      MinValueVT      =   1598423045
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Gaji Bersih :"
      Height          =   210
      Left            =   960
      TabIndex        =   13
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Potongan :"
      Height          =   210
      Left            =   960
      TabIndex        =   11
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lembur :"
      Height          =   210
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Gaji Pokok :"
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jabatan :"
      Height          =   210
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Karyawan :"
      Height          =   210
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode Karyawan :"
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "frm_penggajian2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub setting_awal(ByVal kodekary As String, ByVal namakary As String, ByVal jab As String, _
            ByVal gapok As Double, ByVal lembur As Double, ByVal pot As Double)
    
    tkode.Text = kodekary
    tnama.Text = namakary
    tjab.Text = jab
    tdb_gapok.Value = gapok
    tdb_lembur.Value = lembur
    tdb_pot.Value = pot
    
    
End Sub

Private Sub hitunggaji()
    
    Dim gapok, lembur, pot As Double
    
    If Not IsNull(tdb_gapok.Value) Then
        gapok = 0
    Else
        gapok = tdb_gapok.Value
    End If
    
    If Not IsNull(tdb_lembur.Value) Then
        lembur = 0
    Else
        lembur = tdb_lembur.Value
    End If
    
    If Not IsNull(tdb_pot.Value) Then
        pot = 0
    Else
        pot = tdb_pot.Value
    End If
    
    Dim gajibersih As Double
        gajibersih = (gapok + lembur) - pot
    
    tdb_bersih.Value = gajibersih
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    tdb_gapok.SetFocus
End Sub

Private Sub Form_Load()

Me.Left = Utama.Width / 2 - Me.Width / 2
Me.Top = (Utama.Height / 2 - Me.Height / 2) - 1500

End Sub

Private Sub tdb_gapok_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_lembur.SetFocus
End Sub

Private Sub tdb_lembur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tdb_pot.SetFocus
End Sub

Private Sub tdb_pot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command1.SetFocus
End Sub
