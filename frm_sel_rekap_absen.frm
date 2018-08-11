VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_sel_rekap_absen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi rekap absen"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "Detail"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rekap"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cbbulan 
         Height          =   315
         ItemData        =   "frm_sel_rekap_absen.frx":0000
         Left            =   960
         List            =   "frm_sel_rekap_absen.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox tnama 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox tkode 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_Lihat 
         Caption         =   "&Tampil"
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin TDBNumber6Ctl.TDBNumber tdbthn 
         Height          =   320
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   564
         Calculator      =   "frm_sel_rekap_absen.frx":0004
         Caption         =   "frm_sel_rekap_absen.frx":0024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_sel_rekap_absen.frx":0089
         Keys            =   "frm_sel_rekap_absen.frx":00A7
         Spin            =   "frm_sel_rekap_absen.frx":00F1
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama :"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kode :"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tahun :"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bulan :"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_sel_rekap_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbbulan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tkode.SetFocus
End Sub

Private Sub Cmd_Keluar_Click()
    Unload Me
End Sub

Private Sub Cmd_Lihat_Click()

    Dim sql As String
    
    If Option1.Value = True Then
        
        sql = "select * from View_r_absen where Tahun=" & tdbthn.Text & " and Bulan=" & cbbulan.ListIndex + 1
        
        
        If tkode.Text <> "" Then
            sql = sql & " and Kode_Karyawan like '%" & Trim(tkode.Text) & "%'"
        End If
        
        If tnama.Text <> "" Then
            sql = sql & " and Nama_Karyawan like '%" & Trim(tnama.Text) & "%'"
        End If
        
        sql = sql & " order by Nama_Karyawan asc"
        
    End If

    If Option2.Value = True Then
    
'        sql = "SELECT Periode.Bulan,tbAbsen.tanggal,tbAbsen.jam1,tbAbsen.jam2,tbAbsen.bulan,tbAbsen.tahun,Tb_Karyawan.Nama_Karyawan,Tb_Karyawan.Kode_Karyawan"
'            sql = sql & " From "
'        sql = sql & "(ASKRINDO.dbo.Periode Periode INNER JOIN ASKRINDO.dbo.tbAbsen tbAbsen ON"
'        sql = sql & " Periode.Tanggal = tbAbsen.tanggal)"
'        sql = sql & " INNER JOIN ASKRINDO.dbo.Tb_Karyawan Tb_Karyawan ON"
'        sql = sql & " tbAbsen.kode_absen = Tb_Karyawan.kd_absen"
'        sql = sql & " where Periode.tahun=" & tdbthn.Text & " and Periode.bulan=" & cbbulan.ListIndex + 1
'
'
'
'        sql = sql & " Order By"
'        sql = sql & " Periode.Bulan ASC,Tb_Karyawan.Nama_Karyawan Asc"
        
        sql = "select * from View_absendetail"
        
        sql = sql & " where tahun=" & tdbthn.Text & " and bulan=" & cbbulan.ListIndex + 1
        
        If tkode.Text <> "" Then
            sql = sql & " and Kode_Karyawan like '%" & Trim(tkode.Text) & "%'"
        End If

        If tnama.Text <> "" Then
            sql = sql & " and Nama_Karyawan like '%" & Trim(tnama.Text) & "%'"
        End If
        
        
    End If
    
    Mysq = sql
    
    If Option1.Value = True Then
    
        Load frm_lap_rek_absen2
        frm_lap_rek_absen2.Show
    
    Else
    
        Load frm_lap_absen_detail
        frm_lap_absen_detail.Show
    
    End If
    



End Sub

Private Sub Form_Activate()
    cbbulan.SetFocus
End Sub

Private Sub Form_Load()

Dim status As String
status = Buka_Koneksi
If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
        
        End
        Exit Sub
End If
    
    Option1.Value = True
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 250
    End With
    
        With cbbulan
        .AddItem "Januari"
        .AddItem "Februari"
        .AddItem "Maret"
        .AddItem "April"
        .AddItem "Mei"
        .AddItem "Juni"
        .AddItem "Juli"
        .AddItem "Agustus"
        .AddItem "September"
        .AddItem "Oktober"
        .AddItem "November"
        .AddItem "Desember"
    End With
    
    tdbthn.Text = Year(Now)
    cbbulan.ListIndex = 0
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If

End Sub

Private Sub tdbthn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cbbulan.SetFocus
End Sub

Private Sub tkode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tnama.SetFocus
End Sub

Private Sub tnama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Lihat.SetFocus
End Sub
