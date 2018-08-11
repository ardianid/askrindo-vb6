VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Utama 
   BackColor       =   &H00404040&
   Caption         =   "ASKRINDO"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10440
   Icon            =   "Utama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":27C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2856C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":28E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":29720
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":29FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2A8D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2B1AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2BA88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Utama.frx":2C362
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7620
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "23:21"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   975
         _cx             =   1720
         _cy             =   1085
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "login2"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "logof2"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rubahpwd_t"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "download_mesin_T"
            Object.ToolTipText     =   "Download data (Mesin Absensi)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "penggajian_T"
            Object.ToolTipText     =   "Penggajian"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Menu fL 
      Caption         =   "&File"
      Begin VB.Menu login 
         Caption         =   "&Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu logof 
         Caption         =   "Log &Off"
         Shortcut        =   ^O
      End
      Begin VB.Menu grs1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu user 
      Caption         =   "&User"
      Begin VB.Menu User_Baru_M 
         Caption         =   "&Tambah User"
      End
      Begin VB.Menu Form_Hak_Akses_M 
         Caption         =   "&Seting Hak Akses"
      End
      Begin VB.Menu grspw 
         Caption         =   "-"
      End
      Begin VB.Menu rubahpwd_M 
         Caption         =   "&Rubah Password"
      End
   End
   Begin VB.Menu mast 
      Caption         =   "&Master"
      Begin VB.Menu Frm_Mast_Bagian_M 
         Caption         =   "Bagian Pega&wai"
      End
      Begin VB.Menu Frm_Mast_Jabatan_M 
         Caption         =   "Ja&batan Pegawai"
      End
      Begin VB.Menu Karyawan_M 
         Caption         =   "Pe&gawai"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "&Transaksi"
      Begin VB.Menu download_mesin_M 
         Caption         =   "&Download Data (Mesin Absensi)"
      End
      Begin VB.Menu penggajian_M 
         Caption         =   "&Penggajian"
      End
   End
   Begin VB.Menu lap 
      Caption         =   "&Laporan"
      Begin VB.Menu Frm_sel_Karyawan_M 
         Caption         =   "&Karyawan"
      End
      Begin VB.Menu frm_sel_sisapensiun_M 
         Caption         =   "&Sisa Masa Kerja"
      End
      Begin VB.Menu frm_sel_jabterakhir_M 
         Caption         =   "&Jabatan Terakhir Pegawai"
      End
      Begin VB.Menu glaptrans 
         Caption         =   "-"
      End
      Begin VB.Menu frm_sel_histkary_M 
         Caption         =   "&Histori Karyawan"
      End
      Begin VB.Menu glap2 
         Caption         =   "-"
      End
      Begin VB.Menu frm_sel_rekap_absen_M 
         Caption         =   "&Absensi"
      End
      Begin VB.Menu frm_sel_rkapgaji_M 
         Caption         =   "&Gaji"
      End
   End
End
Attribute VB_Name = "Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim status As String

Public Sub SetAktifMenu(ByVal sql As String)
    
    Dim obj As Object
     Dim a As Long
            
    Dim rec As Recordset
        Set rec = New ADODB.Recordset
            rec.Open sql, kon, adOpenKeyset

    With rec
        If Not .EOF Then
        Do While Not .EOF

               Dim nama_f
               Dim namatol
                    nama_f = !nama_form
                    namatol = nama_f
                    nama_f = nama_f & "_M"
                    namatol = namatol & "_T"
                    
               For Each obj In Me
               
                If TypeOf obj Is Toolbar Then
                Else
                If obj.Name = nama_f Then
                    obj.Enabled = True
                    Exit For
                End If
                End If
                
               Next
                
               
               For a = 1 To 9
                    If UCase(Toolbar1.Buttons.Item(a).Key) = UCase(namatol) Then
                        Toolbar1.Buttons.Item(a).Enabled = True
                        Exit For
                    End If
               Next
                
        .MoveNext
        Loop
        End If

    End With

    rubahpwd_M.Enabled = True
    Toolbar1.Buttons.Item(2).Enabled = True
    Toolbar1.Buttons.Item(1).Enabled = False
    Toolbar1.Buttons.Item(4).Enabled = True
    
End Sub

Private Sub download_mesin_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("download_mesin") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_download_absen
        Frm.Show
        
    Else
        
        If Cek_akses_Form("download_mesin") = False Then Exit Sub
        
        Set Frm = frm_download_absen
        Frm.Show
    End If

End Sub

Private Sub exit_Click()
    End
End Sub

Public Sub enable_menu(ByVal sett As Boolean)
   
   Dim a As Object
   Dim X As Long
        For Each a In Me
        
            If TypeOf a Is Toolbar Then
            Else
            If (UCase(Right(a.Name, 1)) = UCase("M")) Then
                a.Enabled = sett
            End If
            End If
            
        Next
   
        For X = 1 To 9
            If UCase(Right(Toolbar1.Buttons.Item(X).Key, 1)) = UCase("t") Then
                 Toolbar1.Buttons.Item(X).Enabled = sett
            End If
        Next
        
'   adduser_S.Enabled = sett
'   setingakses_S.Enabled = sett
'   rubahpwd_S.Enabled = sett
'
'   kary_S.Enabled = sett
'   anggota_S.Enabled = sett
'   hargaperkilo_S.Enabled = sett
'   simpananwajib_S.Enabled = sett
'
'   timbang_S.Enabled = sett
'   timbang_btl.Enabled = sett
'
'
'   lapkary_S.Enabled = sett
'   lapanggota_S.Enabled = sett
'   laptimbang.Enabled = sett
    
End Sub

Private Sub Form_Hak_Akses_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Form_Hak_Akses
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Form_Hak_Akses") = False Then Exit Sub
        
        Set Frm = Form_Hak_Akses
        Frm.Show
    End If

End Sub







Private Sub Frm_Mast_Bagian_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Mast_Bagian") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Mast_Bagian
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Mast_Bagian") = False Then Exit Sub
        
        Set Frm = Frm_Mast_Bagian
        Frm.Show
    End If


End Sub

'Private Sub Frm_Mast_Bhn_Bkr_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Bhn_Bkr") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Bhn_Bkr
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Bhn_Bkr") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Bhn_Bkr
'        Frm.Show
'    End If
'
'End Sub

'Private Sub Frm_Mast_Biaya_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Biaya") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Biaya
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Biaya") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Biaya
'        Frm.Show
'    End If
'
'
'
'End Sub

Private Sub Frm_Mast_Jabatan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_Mast_Jabatan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_Mast_Jabatan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_Mast_Jabatan") = False Then Exit Sub
        
        Set Frm = Frm_Mast_Jabatan
        Frm.Show
    End If


End Sub

'Private Sub Frm_Mast_Kend_SJKB_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Kend_SJKB") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Kend_SJKB
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Kend_SJKB") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Kend_SJKB
'        Frm.Show
'    End If
'
'End Sub

'Private Sub Frm_Mast_Peny_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Peny") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Peny
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Peny") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Peny
'        Frm.Show
'    End If
'
'End Sub

'Private Sub Frm_Mast_Tol_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Tol") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Tol
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Tol") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Tol
'        Frm.Show
'    End If
'
'
'End Sub

'Private Sub Frm_Mast_Uang_Makan_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Mast_Uang_Makan") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Mast_Uang_Makan
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Mast_Uang_Makan") = False Then Exit Sub
'
'        Set Frm = Frm_Mast_Uang_Makan
'        Frm.Show
'    End If
'
'End Sub

'Private Sub Frm_Master_asmen_M_Click()
'
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Master_asmen") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Master_asmen
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Master_asmen") = False Then Exit Sub
'
'        Set Frm = Frm_Master_asmen
'        Frm.Show
'    End If
'
'End Sub

'Private Sub Frm_Sel_Biaya_SJKB_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SJKB") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Biaya_SJKB
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SJKB") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Biaya_SJKB
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_Biaya_SPB_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SPB") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Biaya_SPB
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SPB") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Biaya_SPB
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_Biaya_SPPD_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SPPD") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Biaya_SPPD
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Biaya_SPPD") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Biaya_SPPD
'        Frm.Show
'    End If
'End Sub

Private Sub frm_sel_histkary_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_histkary") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_histkary
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_histkary") = False Then Exit Sub
        
        Set Frm = frm_sel_histkary
        Frm.Show
    End If

End Sub

Private Sub frm_sel_jabterakhir_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_jabterakhir") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_jabterakhir
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_jabterakhir") = False Then Exit Sub
        
        Set Frm = frm_sel_jabterakhir
        Frm.Show
    End If

End Sub

Private Sub Frm_sel_Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Frm_sel_Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Frm_sel_Karyawan") = False Then Exit Sub
        
        Set Frm = Frm_sel_Karyawan
        Frm.Show
    End If

End Sub

'Private Sub Frm_Sel_Perbulan_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Perbulan") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Perbulan
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Perbulan") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Perbulan
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_Pertahun_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Pertahun") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Pertahun
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Pertahun") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Pertahun
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_Pertanggal1_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_Pertanggal1") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_Pertanggal1
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_Pertanggal1") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_Pertanggal1
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_SJKB2_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_SJKB2") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_SJKB2
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_SJKB2") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_SJKB2
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_SPB1_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_SPB1") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_SPB1
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_SPB1") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_SPB1
'        Frm.Show
'    End If
'End Sub

'Private Sub Frm_Sel_SPPD_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Sel_SPPD") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Sel_SPPD
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Sel_SPPD") = False Then Exit Sub
'
'        Set Frm = Frm_Sel_SPPD
'        Frm.Show
'    End If
'
'End Sub

Private Sub Frm_Trans_SJKB_M_Click()
'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Trans_SJKB") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Trans_SJKB
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Trans_SJKB") = False Then Exit Sub
'
'        Set Frm = Frm_Trans_SJKB
'        Frm.Show
'    End If

End Sub

Private Sub Frm_Trans_SPB_M_Click()

'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Trans_SPB") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Trans_SPB
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Trans_SPB") = False Then Exit Sub
'
'        Set Frm = Frm_Trans_SPB
'        Frm.Show
'    End If

End Sub

Private Sub Frm_Trans_SPPD_M_Click()

'    If Not (Frm Is Nothing) Then
'        Unload Frm
'
'        If Cek_akses_Form("Frm_Trans_SPPD") = False Then Exit Sub
'
'        Set Frm = Nothing
'        Set Frm = Frm_Trans_SPPD
'        Frm.Show
'
'    Else
'
'        If Cek_akses_Form("Frm_Trans_SPPD") = False Then Exit Sub
'
'        Set Frm = Frm_Trans_SPPD
'        Frm.Show
'    End If


End Sub

Private Sub frm_sel_rekap_absen_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_rekap_absen") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_rekap_absen
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_rekap_absen") = False Then Exit Sub
        
        Set Frm = frm_sel_rekap_absen
        Frm.Show
    End If

End Sub

Private Sub frm_sel_rkapgaji_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_rkapgaji") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_rkapgaji
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_rkapgaji") = False Then Exit Sub
        
        Set Frm = frm_sel_rkapgaji
        Frm.Show
    End If

End Sub

Private Sub frm_sel_sisapensiun_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("frm_sel_sisapensiun") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_sel_sisapensiun
        Frm.Show
        
    Else
        
        If Cek_akses_Form("frm_sel_sisapensiun") = False Then Exit Sub
        
        Set Frm = frm_sel_sisapensiun
        Frm.Show
    End If

End Sub

Private Sub Karyawan_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = Karyawan
        Frm.Show
        
    Else
        
        If Cek_akses_Form("Karyawan") = False Then Exit Sub
        
        Set Frm = Karyawan
        Frm.Show
    End If


End Sub

Private Sub login_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show

End Sub

Private Sub logof_Click()
    
    If kon.State = adStateClosed Then
            Buka_Koneksi
    End If
    
    If Not (Frm Is Nothing) Then
        Unload Frm
    End If
    
    enable_menu False
    StatusBar1.Panels(1).Text = "User Actived :"
    U_Masuk.Show
    
End Sub
Private Sub MDIForm_Load()
    
    StatusBar1.Panels(2).Text = Format(Date, "dd mmmm yyyy")
      
    enable_menu False

 status = Buka_Koneksi
 If status = "-2147467259" Then
    
            Dim konfirm As Integer
            Dim Informasi As String
                Informasi = "Koneksi terhadap server tidak berhasil :"
                Informasi = Informasi & vbCrLf & "1. Pastikan server telah hidup dan SQL Server telah dijalankan pada server,atau"
                Informasi = Informasi & vbCrLf & "2. Apabila masih terjadi kegagalan koneksi,periksa nama komputer server,Pastikan nama komputer server tidak berubah"
                Informasi = Informasi & vbCrLf & vbCrLf & "apakah anda ingin menyeting ulang koneksi nama komputer server ?"
        
                konfirm = CInt(MsgBox(Informasi, vbYesNo + vbQuestion, "Konfimasi"))
        
                If konfirm = vbYes Then
        
                    Load Frm_New_Seting
                    Frm_New_Seting.Show
        
                    Unload Me
                    Exit Sub
                Else
                    Unload Me
                    End
                    Exit Sub
                End If

 End If

'    Dim btas As Double
'        btas = Batas
'
'    If btas = 100 Then
'        End
'        Exit Sub
'    Else
'        btas = btas + 1
'        SaveSetting "bts", "bts", "bts", btas
'    End If


    logof.Enabled = False
    Toolbar1.Buttons.Item(2).Enabled = False
    
     ShockwaveFlash2.Movie = App.Path & "\MOVIE2.swf"
    
    login_Click
    
End Sub



Private Sub MDIForm_Resize()
    
    With ShockwaveFlash2
        .Left = 0
        .Width = Me.Width
    End With
    
End Sub

Private Sub penggajian_M_Click()

If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("penggajian") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = frm_penggajian
        Frm.Show
        
    Else
        
        If Cek_akses_Form("penggajian") = False Then Exit Sub
        
        Set Frm = frm_penggajian
        Frm.Show
    End If

End Sub

Private Sub rubahpwd_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
                
        Set Frm = Nothing
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
        
    Else
                
        Set Frm = Frm_Rubah_Pwd
        Frm.Show
    End If

ShockwaveFlash2.Movie = App.Path & "\MOVIE2.swf"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1
            login_Click
        Case 2
            logof_Click
        Case 4
            rubahpwd_M_Click
        Case 6
            download_mesin_M_Click
        Case 7
            penggajian_M_Click
        Case 9
            exit_Click
    End Select
    
End Sub

Private Sub User_Baru_M_Click()

    If Not (Frm Is Nothing) Then
        Unload Frm
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = Nothing
        Set Frm = User_Baru
        Frm.Show
        
    Else
        
        If Cek_akses_Form("User_Baru") = False Then Exit Sub
        
        Set Frm = User_Baru
        Frm.Show
    End If
    
End Sub
