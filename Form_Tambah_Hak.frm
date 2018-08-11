VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Form_Tambah_Hak 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tambah Hak Akses ..."
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Tambah_Hak.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   13590
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13575
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id User :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lbl_id_user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl_nama_user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   7080
      TabIndex        =   7
      Top             =   840
      Width           =   6495
      Begin VB.CommandButton cmd_batal 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Selesai"
         Height          =   495
         Left            =   5520
         TabIndex        =   10
         Top             =   6360
         Width           =   855
      End
      Begin VB.CommandButton cmd_simpan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4560
         TabIndex        =   9
         Top             =   6360
         Width           =   855
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_hak_akses 
         Height          =   6135
         Left            =   120
         OleObjectBlob   =   "Form_Tambah_Hak.frx":27C92
         TabIndex        =   8
         Top             =   120
         Width           =   6255
      End
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   5880
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox cek_akses_tambah 
            Caption         =   "&Tambah"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox cek_akses_rubah 
            Caption         =   "&Rubah"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox cek_akses_hapus 
            Caption         =   "&Hapus"
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox cek_akses_laporan 
            Caption         =   "&Laporan/Bukti"
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   6480
      TabIndex        =   2
      Top             =   1560
      Width           =   495
      Begin VB.CommandButton cmd_keluar_satu 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton cmd_keluar_semua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton cmd_masuk_semua 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmd_masuk_satu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6375
      Begin TrueOleDBGrid60.TDBGrid Grid_aplikasi 
         Height          =   6015
         Left            =   120
         OleObjectBlob   =   "Form_Tambah_Hak.frx":2C489
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox cek_laporan 
            Caption         =   "&Laporan/Bukti"
            Height          =   255
            Left            =   1920
            TabIndex        =   20
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox cek_hapus 
            Caption         =   "&Hapus"
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox cek_rubah 
            Caption         =   "&Rubah"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox cek_tambah 
            Caption         =   "&Tambah"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "Form_Tambah_Hak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Arr_Aplikasi As New XArrayDB
Dim Arr_Akses As New XArrayDB

Private Sub cek_akses_percommand(ByVal id As Integer)

Dim sql As String
Dim rs As Recordset
    
    sql = "select Tambah,Rubah,Hapus,Cetak_Laporan from Tb_Aplikasi where Id=" & id
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon, adOpenKeyset

With rs

    If Not rs.EOF Then
    
    ''' tambah '''
    
    Dim tambah As Integer
        tambah = IIf(Not IsNull(!tambah), !tambah, 0)
    
    If tambah <> 0 Then
        
        cek_tambah.Enabled = True
        cek_tambah.Value = vbChecked
    
    Else
        
        cek_tambah.Enabled = False
        cek_tambah.Value = vbUnchecked
    
    End If
    
    ''' --- '''
    
    ''' rubah '''
    
    Dim rubah As Integer
        rubah = IIf(Not IsNull(!rubah), !rubah, 0)
    
    If rubah <> 0 Then
        
        cek_rubah.Enabled = True
        cek_rubah.Value = vbChecked
    
    Else
        
        cek_rubah.Enabled = False
        cek_rubah.Value = vbUnchecked
    
    End If
    
    
    ''' --- '''
    
    ''' hapus '''
    
    Dim hapus As Integer
        hapus = IIf(Not IsNull(!hapus), !hapus, 0)
    
    If hapus <> 0 Then
        
        cek_hapus.Enabled = True
        cek_hapus.Value = vbChecked
    
    Else
        
        cek_hapus.Enabled = False
        cek_hapus.Value = vbUnchecked
    
    End If
    
    
    ''' --- '''
    
    ''' laporan '''
    
    Dim laporan As Integer
        laporan = IIf(Not IsNull(!Cetak_Laporan), !Cetak_Laporan, 0)
    
    If laporan <> 0 Then
        
        cek_laporan.Enabled = True
        cek_laporan.Value = vbChecked
    
    Else
        
        cek_laporan.Enabled = False
        cek_laporan.Value = vbUnchecked
    
    End If
    
    
    ''' --- '''
    
    Else
    
    cek_tambah.Enabled = False
    cek_rubah.Enabled = False
    cek_hapus.Enabled = False
    cek_laporan.Enabled = False
    
    cek_tambah.Value = vbUnchecked
    cek_rubah.Value = vbUnchecked
    cek_hapus.Value = vbUnchecked
    cek_laporan.Value = vbUnchecked
    
    End If
    
End With


End Sub

Public Sub isi_grid_hak()
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select Id,Nama_Aplikasi,Ket from Tb_Aplikasi where Id not in (select Id_aplikasi from VIEW_Hak_Akses where Id_User=" & Trim(lbl_id_user.Caption) & ") order by Nama_Aplikasi asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
    
    Dim a As Long
    Dim id_a, NAMA, ket As String
       
            a = 1
            Arr_Aplikasi.ReDim 0, 0, 0, 0
            Arr_Aplikasi.ReDim 1, 1, 1, 1
                Grid_aplikasi.ReBind
                Grid_aplikasi.Refresh
            
            With rs
                
                Do While Not .EOF
                    Arr_Aplikasi.ReDim 1, a, 0, Grid_aplikasi.Columns.Count
                        Grid_aplikasi.ReBind
                        Grid_aplikasi.Refresh
                    
                    id_a = IIf(Not IsNull(!id), !id, "")
                    NAMA = IIf(Not IsNull(!Nama_APlikasi), !Nama_APlikasi, "")
                    ket = IIf(Not IsNull(!ket), !ket, "")
                    
                    Arr_Aplikasi(a, 0) = id_a
                    Arr_Aplikasi(a, 1) = NAMA
                    Arr_Aplikasi(a, 2) = ket
                
                a = a + 1
                .MoveNext
                Loop
                
                Grid_aplikasi.ReBind
                Grid_aplikasi.Refresh
                
                Grid_aplikasi.MoveFirst
                
            End With
       
        
End Sub

Private Sub Cmd_Batal_Click()

    Unload Me
    Form_Hak_Akses.Enabled = True
    Form_Hak_Akses.Show
    
End Sub

Private Sub cmd_keluar_satu_Click()

    Dim baris As Long
    
    If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then Exit Sub
    
    If Arr_Aplikasi.UpperBound(1) = 1 And Arr_Aplikasi(1, 1) = Empty Then
        baris = 1
    Else
        baris = Arr_Aplikasi.UpperBound(1) + 1
    End If
    
    Arr_Aplikasi.ReDim 1, baris, 0, Grid_aplikasi.Columns.Count
        Grid_aplikasi.ReBind
        Grid_aplikasi.Refresh
    
    Arr_Aplikasi(baris, 0) = Arr_Akses(Grid_hak_akses.Bookmark, 0)
    Arr_Aplikasi(baris, 1) = Arr_Akses(Grid_hak_akses.Bookmark, 1)
    Arr_Aplikasi(baris, 2) = Arr_Akses(Grid_hak_akses.Bookmark, 2)

    Grid_aplikasi.ReBind
    Grid_aplikasi.Refresh
    
    If Arr_Akses.UpperBound(1) = 1 Then
        Arr_Akses.ReDim 0, 0, 0, 0
        Arr_Akses.ReDim 1, 1, 1, 1
    Else
        Grid_hak_akses.Delete
    End If
    
    Grid_hak_akses.ReBind
    Grid_hak_akses.Refresh
    
End Sub

Private Sub cmd_keluar_semua_Click()

Dim a As Long
    
    If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then Exit Sub
        
    Dim b As Long
    If Arr_Aplikasi.UpperBound(1) = 1 And Arr_Aplikasi(1, 1) = Empty Then
        b = 1
    Else
        b = Arr_Aplikasi.UpperBound(1) + 1
    End If
        
    For a = Arr_Akses.LowerBound(1) To Arr_Akses.UpperBound(1)
    
        Arr_Aplikasi.ReDim 1, b, 0, Grid_aplikasi.Columns.Count
            Grid_aplikasi.ReBind
            Grid_aplikasi.Refresh
        
        Arr_Aplikasi(b, 0) = Arr_Akses(a, 0)
        Arr_Aplikasi(b, 1) = Arr_Akses(a, 1)
        Arr_Aplikasi(b, 2) = Arr_Akses(a, 2)
        
        Grid_aplikasi.ReBind
        Grid_aplikasi.Refresh
        
        b = b + 1
        
    Next
    
    Arr_Akses.ReDim 0, 0, 0, 0
    Arr_Akses.ReDim 1, 1, 1, 1
        Grid_hak_akses.ReBind
        Grid_hak_akses.Refresh

End Sub

Private Sub cmd_masuk_satu_Click()
    
    Dim baris As Long
    
    If Arr_Aplikasi.UpperBound(1) = 1 And Arr_Aplikasi(1, 1) = Empty Then Exit Sub
    
    If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then
        baris = 1
    Else
        baris = Arr_Akses.UpperBound(1) + 1
    End If
    
    Arr_Akses.ReDim 1, baris, 0, Grid_hak_akses.Columns.Count
        Grid_hak_akses.ReBind
        Grid_hak_akses.Refresh
    
    ' tambah
    Dim tambah As Integer
    If cek_tambah.Value = vbChecked Then
        tambah = 1
    Else
        tambah = 0
    End If
    
    ' sampe sini
    
    ' rubah
    Dim rubah As Integer
    If cek_rubah.Value = vbChecked Then
        rubah = 1
    Else
        rubah = 0
    End If
    
    ' sampe sini
    
    ' rubah
    Dim hapus As Integer
    If cek_hapus.Value = vbChecked Then
        hapus = 1
    Else
        hapus = 0
    End If
    
    ' sampe sini
    
    ' rubah
    Dim laporan As Integer
    If cek_laporan.Value = vbChecked Then
        laporan = 1
    Else
        laporan = 0
    End If
    
    ' sampe sini
    
    Arr_Akses(baris, 0) = Arr_Aplikasi(Grid_aplikasi.Bookmark, 0)
    Arr_Akses(baris, 1) = Arr_Aplikasi(Grid_aplikasi.Bookmark, 1)
    Arr_Akses(baris, 2) = Arr_Aplikasi(Grid_aplikasi.Bookmark, 2)
    Arr_Akses(baris, 3) = tambah
    Arr_Akses(baris, 4) = rubah
    Arr_Akses(baris, 5) = hapus
    Arr_Akses(baris, 6) = laporan
    
    Grid_hak_akses.ReBind
    Grid_hak_akses.Refresh
    
    If Arr_Aplikasi.UpperBound(1) = 1 Then
        Arr_Aplikasi.ReDim 0, 0, 0, 0
        Arr_Aplikasi.ReDim 1, 1, 1, 1
        
        cek_tambah.Enabled = False
        cek_rubah.Enabled = False
        cek_hapus.Enabled = False
        cek_laporan.Enabled = False
        
        cek_tambah.Value = vbUnchecked
        cek_rubah.Value = vbUnchecked
        cek_hapus.Value = vbUnchecked
        cek_laporan.Value = vbUnchecked
        
    Else
        Grid_aplikasi.Delete
    End If
    
    Grid_aplikasi.ReBind
    Grid_aplikasi.Refresh
        
    End Sub

Private Sub cmd_masuk_semua_Click()

Dim a As Long
Dim b As Long
    
    If Arr_Aplikasi.UpperBound(1) = 1 And Arr_Aplikasi(1, 1) = Empty Then Exit Sub
    
    If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then
        b = 1
    Else
        b = Arr_Akses.UpperBound(1) + 1
    End If
    
    
    For a = Arr_Aplikasi.LowerBound(1) To Arr_Aplikasi.UpperBound(1)
        
        cek_akses_percommand Arr_Aplikasi(a, 0)
        
        ' tambah
        Dim tambah As Integer
        If cek_tambah.Value = vbChecked Then
            tambah = 1
        Else
            tambah = 0
        End If
        
        ' sampe sini
        
        ' rubah
        Dim rubah As Integer
        If cek_rubah.Value = vbChecked Then
            rubah = 1
        Else
            rubah = 0
        End If
        
        ' sampe sini
        
        ' rubah
        Dim hapus As Integer
        If cek_hapus.Value = vbChecked Then
            hapus = 1
        Else
            hapus = 0
        End If
        
        ' sampe sini
        
        ' rubah
        Dim laporan As Integer
        If cek_laporan.Value = vbChecked Then
            laporan = 1
        Else
            laporan = 0
        End If
        
        ' sampe sini
        
        Arr_Akses.ReDim 1, b, 0, Grid_hak_akses.Columns.Count
            Grid_hak_akses.ReBind
            Grid_hak_akses.Refresh
        
        Arr_Akses(b, 0) = Arr_Aplikasi(a, 0)
        Arr_Akses(b, 1) = Arr_Aplikasi(a, 1)
        Arr_Akses(b, 2) = Arr_Aplikasi(a, 2)
        Arr_Akses(b, 3) = tambah
        Arr_Akses(b, 4) = rubah
        Arr_Akses(b, 5) = hapus
        Arr_Akses(b, 6) = laporan
        
        Grid_hak_akses.ReBind
        Grid_hak_akses.Refresh
        
        b = b + 1
        
    Next
    
    Arr_Aplikasi.ReDim 0, 0, 0, 0
    Arr_Aplikasi.ReDim 1, 1, 1, 1
        Grid_aplikasi.ReBind
        Grid_aplikasi.Refresh
    
    cek_tambah.Enabled = False
    cek_rubah.Enabled = False
    cek_hapus.Enabled = False
    cek_laporan.Enabled = False
    
    cek_tambah.Value = vbUnchecked
    cek_rubah.Value = vbUnchecked
    cek_hapus.Value = vbUnchecked
    cek_laporan.Value = vbUnchecked
    
End Sub

Private Sub Cmd_Simpan_Click()
On Error GoTo err_handler

    Dim a As Long
    Dim sql As String
    Dim rs As Recordset
        
        
        If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Tidak ada data akses yang akan disimpan", vbOKOnly + vbInformation, "Informasi"))
                
                On Error GoTo 0
                Exit Sub
        End If
        
        If MsgBox("Yakin semua akses user yang diberikan sudah benar", vbYesNo + vbInformation, "Informasi") = vbNo Then
            
            On Error GoTo 0
            Exit Sub
        End If
        
        kon.BeginTrans
        
        For a = Arr_Akses.LowerBound(1) To Arr_Akses.UpperBound(1)
            
            sql = "insert into Tb_Hak_Akses (Id_User,Id_Aplikasi,Tambah,Rubah,Hapus,Cetak_Laporan) values(" & Trim(lbl_id_user.Caption) & "," & Arr_Akses(a, 0) & "," & Arr_Akses(a, 3) & "," & Arr_Akses(a, 4) & "," & Arr_Akses(a, 5) & "," & Arr_Akses(a, 6) & " )"
                Set rs = New ADODB.Recordset
                    rs.Open sql, kon
            
        Next
        
        kon.CommitTrans
        
            konfirm = CInt(MsgBox("Data hak akses user telah disimpan", vbOKOnly + vbInformation, "Informasi"))
        
        With Form_Hak_Akses
            .Isi_Hak_Akses_PerUser Trim(lbl_id_user.Caption)
        End With
        
        Cmd_Batal_Click
        
        On Error GoTo 0
        Exit Sub
        
err_handler:
        
        kon.RollbackTrans
        
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Informasi"))
            Err.Clear
        
End Sub

Private Sub Form_Load()
    
'SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 2, _
'                Me.Top / 2, Me.Width / 2, _
'              Me.Height / 2, SWP_NOACTIVATE Or SWP_SHOWWINDOW

    
    
    Grid_aplikasi.Array = Arr_Aplikasi
    Grid_hak_akses.Array = Arr_Akses
        
    If lbl_id_user.Caption <> "Id User" And lbl_nama_user.Caption <> "Id User" Then
        isi_grid_hak
    End If
        
    Arr_Akses.ReDim 0, 0, 0, 0
    Arr_Akses.ReDim 1, 1, 1, 1
        Grid_hak_akses.ReBind
        Grid_hak_akses.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Hak_Akses.Enabled = True
    Form_Hak_Akses.Show
End Sub


Private Sub Grid_aplikasi_Click()

If Arr_Aplikasi.UpperBound(1) = 1 And Arr_Aplikasi(1, 1) = Empty Then
    
    cek_tambah.Enabled = False
    cek_rubah.Enabled = False
    cek_hapus.Enabled = False
    cek_laporan.Enabled = False
    
    cek_tambah.Value = vbUnchecked
    cek_rubah.Value = vbUnchecked
    cek_hapus.Value = vbUnchecked
    cek_laporan.Value = vbUnchecked
    
    Exit Sub
        
    End If
    
    cek_akses_percommand Arr_Aplikasi(Grid_aplikasi.Bookmark, 0)

End Sub

Private Sub Grid_aplikasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Grid_aplikasi_Click
End Sub

Private Sub Grid_hak_akses_Click()

    If Arr_Akses.UpperBound(1) = 1 And Arr_Akses(1, 1) = Empty Then
        
        cek_akses_tambah.Value = vbUnchecked
        cek_akses_rubah.Value = vbUnchecked
        cek_akses_hapus.Value = vbUnchecked
        cek_akses_laporan.Value = vbUnchecked
    
    Else
        
        If Arr_Akses(Grid_hak_akses.Bookmark, 3) = 0 Then
            cek_akses_tambah.Value = vbUnchecked
        Else
            cek_akses_tambah.Value = vbChecked
        End If
        
        If Arr_Akses(Grid_hak_akses.Bookmark, 4) = 0 Then
            cek_akses_rubah.Value = vbUnchecked
        Else
            cek_akses_rubah.Value = vbChecked
        End If
        
        If Arr_Akses(Grid_hak_akses.Bookmark, 5) = 0 Then
            cek_akses_hapus.Value = vbUnchecked
        Else
            cek_akses_hapus.Value = vbChecked
        End If
        
        If Arr_Akses(Grid_hak_akses.Bookmark, 6) = 0 Then
            cek_akses_laporan.Value = vbUnchecked
        Else
            cek_akses_laporan.Value = vbChecked
        End If
    End If

End Sub

Private Sub Grid_hak_akses_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Grid_hak_akses_Click
End Sub
