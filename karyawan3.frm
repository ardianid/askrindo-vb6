VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form karyawan3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Non Aktifkan..."
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox tket 
         Height          =   795
         Left            =   1440
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
      End
      Begin VB.ComboBox cbalasan 
         Height          =   315
         ItemData        =   "karyawan3.frx":0000
         Left            =   1440
         List            =   "karyawan3.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin TDBDate6Ctl.TDBDate ttgl 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         Calendar        =   "karyawan3.frx":0068
         Caption         =   "karyawan3.frx":0180
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "karyawan3.frx":01EC
         Keys            =   "karyawan3.frx":020A
         Spin            =   "karyawan3.frx":0268
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   1863103
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "05/01/2007"
         ValidateMode    =   0
         ValueVT         =   2010382343
         Value           =   39087
         CenturyMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   6
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Alasan :"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   3
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Non-Aktif :"
         Height          =   195
         Index           =   34
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "karyawan3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode As String

Private Sub cbalasan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then tket.SetFocus
End Sub

Private Sub cmdbatal_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   
 On Error GoTo err_handler
   
    If MsgBox("Yakin akan dinonaktifkan ???", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then On Error GoTo 0: Exit Sub
    
    kon.BeginTrans
    Dim sql As String
    Dim rs As Recordset
    
    sql = "update Tb_Karyawan set tgl_keluar='" & Format(ttgl.Value, "yyyy/mm/dd") & "',alasankeluar='" & Trim(cbalasan.Text) & "',ket_keluar='" & Trim(tket.Text) & "' where Kode_Karyawan='" & kode & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    kon.CommitTrans
    
    MsgBox "Pegawai dinonaktifkan", vbOKOnly + vbInformation, "Informasi"
    
    Karyawan.set_setelah_edit_status
    
    Unload Me
    
    
On Error GoTo 0
Exit Sub

err_handler:
    
    kon.RollbackTrans
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear
End Sub

Private Sub Form_Activate()
    ttgl.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Left = Utama.Width / 2 - Me.Width / 2
    Me.Top = (Utama.Height / 2 - Me.Height / 2)
    
    Me.Caption = Karyawan.Txt_Nama.Text
    kode = Karyawan.Txt_Kode.Text
    
    cbalasan.ListIndex = 0
End Sub

Private Sub tket_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdok.SetFocus
End Sub

Private Sub ttgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cbalasan.SetFocus
End Sub
