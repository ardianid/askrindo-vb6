VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_sel_rkapgaji 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi Laporan Gaji"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "&Detail"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Rekap"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Keluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Lihat 
         Caption         =   "&Tampil"
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox tkode 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox tnama 
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox cbbulan 
         Height          =   330
         ItemData        =   "frm_sel_rkapgaji.frx":0000
         Left            =   960
         List            =   "frm_sel_rkapgaji.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
      Begin TDBNumber6Ctl.TDBNumber tdbthn 
         Height          =   320
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   564
         Calculator      =   "frm_sel_rkapgaji.frx":0004
         Caption         =   "frm_sel_rkapgaji.frx":0024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_sel_rkapgaji.frx":0089
         Keys            =   "frm_sel_rkapgaji.frx":00A7
         Spin            =   "frm_sel_rkapgaji.frx":00F1
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bulan :"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tahun :"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kode :"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama :"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   510
      End
   End
End
Attribute VB_Name = "frm_sel_rkapgaji"
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
    
'    If Opt_Semua.Value = True Then
'
'    sql = "select * from VIEW_Karyawan where tgl_keluar is null order by Nama_Karyawan asc"
'
'    Else
    
'    If Txt_Kode.Text <> "" Or Txt_Nama.Text <> "" Or Tgl_Masuk1.Text <> "__/__/____" Or Tgl_Masuk2.Text <> "__/__/____" Then
        
        sql = "select * from View_gaji where Tahun=" & tdbthn.Text & " and Bulan=" & cbbulan.ListIndex + 1
        
        
        If tkode.Text <> "" Then
            sql = sql & " and KodeKaryawan like '%" & Trim(tkode.Text) & "%'"
        End If
        
        If tnama.Text <> "" Then
            sql = sql & " and Nama_Karyawan like '%" & Trim(tnama.Text) & "%'"
        End If
        
        sql = sql & " order by Nama_Karyawan asc"
        
        
'    Else
'
'        Dim konfirm As Integer
'            konfirm = CInt(MsgBox("Kriteria pencarian harus diisi", vbOKOnly + vbInformation, "Informasi"))
'
'        Exit Sub
'    End If
    
'    End If
    
'    khusus_user = Mid(Utama.StatusBar1.Panels(5).Text, 7, Len(Utama.StatusBar1.Panels(5).Text))
    
    Mysq = sql
    
    If Option1.Value = True Then
    
    Load frm_lap_rekap_gaji
        frm_lap_rekap_gaji.Show
    
    Else
    
    Load frm_lapgajidetail
        frm_lapgajidetail.Show
    
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
    Option1.Value = True
    
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

