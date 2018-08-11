VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form Form_Hak_Akses 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hak Akses User ..."
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13875
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Hak_Akses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   13935
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.Frame Frame1 
         Caption         =   "User"
         Height          =   8055
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   5055
         Begin TrueOleDBGrid60.TDBGrid Grid_User 
            Height          =   7575
            Left            =   120
            OleObjectBlob   =   "Form_Hak_Akses.frx":08CA
            TabIndex        =   7
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hak Akses"
         Height          =   8055
         Left            =   5280
         TabIndex        =   1
         Top             =   0
         Width           =   8535
         Begin VB.CommandButton cmd_kurang 
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   7440
            Width           =   615
         End
         Begin VB.CommandButton cmd_tambah 
            BackColor       =   &H00FFFFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   7440
            Width           =   615
         End
         Begin VB.CommandButton cmd_keluar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   7440
            TabIndex        =   2
            Top             =   7440
            Width           =   975
         End
         Begin TrueOleDBGrid60.TDBGrid Grid_Hak 
            Height          =   7095
            Left            =   120
            OleObjectBlob   =   "Form_Hak_Akses.frx":3834
            TabIndex        =   5
            Top             =   240
            Width           =   8295
         End
         Begin VB.Frame Frame6 
            Enabled         =   0   'False
            Height          =   855
            Left            =   2400
            TabIndex        =   8
            Top             =   7320
            Visible         =   0   'False
            Width           =   3615
            Begin VB.CheckBox cek_akses_laporan 
               Caption         =   "&Laporan/Bukti"
               Height          =   255
               Left            =   1920
               TabIndex        =   12
               Top             =   480
               Width           =   1575
            End
            Begin VB.CheckBox cek_akses_hapus 
               Caption         =   "&Hapus"
               Height          =   255
               Left            =   1920
               TabIndex        =   11
               Top             =   120
               Width           =   975
            End
            Begin VB.CheckBox cek_akses_rubah 
               Caption         =   "&Rubah"
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox cek_akses_tambah 
               Caption         =   "&Tambah"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   120
               Width           =   975
            End
         End
      End
   End
End
Attribute VB_Name = "Form_Hak_Akses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Arr_User As New XArrayDB
Dim Arr_Hak As New XArrayDB

Private Sub Cmd_Keluar_Click()
    
    Unload Me
    
End Sub

Private Sub cmd_kurang_Click()
On Error GoTo err_handler

If Arr_Hak.UpperBound(1) = 1 And Arr_Hak(1, 1) = Empty Then Exit Sub

If MsgBox("Yakin akan membatalkan form/aplikasi ini dari akses user", vbYesNo + vbInformation, "Informasi") = vbNo Then

On Error GoTo 0
Exit Sub

End If

kon.BeginTrans

Dim sql As String
Dim rs As Recordset
    sql = "delete from Tb_Hak_Akses where Id_Hak= " & Arr_Hak(Grid_Hak.Bookmark, 0)
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
            
            kon.CommitTrans
            
            Dim konfirm As Integer
                konfirm = CInt(MsgBox("Aplikasi berhasil dibatalkan", vbOKOnly + vbInformation, "Informasi"))
                
                Grid_User_DblClick
        
On Error GoTo 0
Exit Sub


err_handler:
    
    kon.RollbackTrans
    
    konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear

End Sub

Private Sub Cmd_Tambah_Click()
    
    If Arr_User.UpperBound(1) = 1 And Arr_User(1, 1) = Empty Then Exit Sub
    
    With Form_Tambah_Hak
        
        .lbl_id_user.Caption = Arr_User(Grid_User.Bookmark, 1)
        .lbl_nama_user.Caption = Arr_User(Grid_User.Bookmark, 2)
        .isi_grid_hak
        .Show
        
        Me.Enabled = False
        
    End With
    
End Sub

Private Sub Form_Load()

With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = 10
End With

Dim status As String
status = Buka_Koneksi
If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Koneksi terputus ....", vbOKOnly + vbInformation, "Informasi"))
        
        End
        Exit Sub
End If

''' akses command ''
'
'    hak_akses_percommand CStr(Me.Name)
'
'    Cmd_Tambah.Enabled = c_tambah
'    cmd_kurang.Enabled = c_hapus
'
''' stop here ''

'Me.PaintPicture Utama.Picture, 0, 0
'Picture1.PaintPicture Utama.Picture, 0, 0
'Frame1.BackColor = RGB(154, 162, 209)
'Frame2.BackColor = RGB(154, 162, 209)
'Frame6.BackColor = RGB(154, 162, 209)
'cek_akses_tambah.BackColor = RGB(154, 162, 209)
'cek_akses_rubah.BackColor = RGB(154, 162, 209)
'cek_akses_hapus.BackColor = RGB(154, 162, 209)
'cek_akses_laporan.BackColor = RGB(154, 162, 209)

Grid_User.Array = Arr_User
Grid_Hak.Array = Arr_Hak

Isi_User

End Sub

Private Sub Isi_User()
    
    Dim sql As String
    Dim rs As Recordset
        
        sql = "select Id_User,Nama_Karyawan from VIEW_User order by Id_User asc"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
        
        Isi_Grid_user rs
        
End Sub

Private Sub Isi_Grid_user(ByVal rec As Recordset)

Dim a As Long
Dim id_u, NAMA As String
    
    a = 1
    
    Arr_User.ReDim 0, 0, 0, 0
    Arr_User.ReDim 1, 1, 1, 1
        Grid_User.ReBind
        Grid_User.Refresh
    
    With rec
        
        Do While Not .EOF
            Arr_User.ReDim 1, a, 0, Grid_User.Columns.Count
                Grid_User.ReBind
                Grid_User.Refresh
        
        id_u = IIf(Not IsNull(!Id_User), !Id_User, "")
        NAMA = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
        
        Arr_User(a, 0) = a
        Arr_User(a, 1) = id_u
        Arr_User(a, 2) = NAMA
        
        a = a + 1
        .MoveNext
        Loop
        
        Grid_User.ReBind
        Grid_User.Refresh
        
        Grid_User.MoveFirst
        
    End With

End Sub

Public Sub Isi_Hak_Akses_PerUser(ByVal Id_User As String)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select Id_Hak,Nama_Aplikasi,Ket,Tambah,Rubah,Hapus,Cetak_Laporan from VIEW_Hak_Akses where Id_User=" & Id_User
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
            
            Dim a As Long
            Dim id_hak, NAMA, ket As String
            Dim tambah, rubah, hapus, laporan As Integer
            
            a = 1
            Arr_Hak.ReDim 0, 0, 0, 0
            Arr_Hak.ReDim 1, 1, 1, 1
                Grid_Hak.ReBind
                Grid_Hak.Refresh
            
            With rs
                
                Do While Not .EOF
                    Arr_Hak.ReDim 1, a, 0, Grid_Hak.Columns.Count
                        Grid_Hak.ReBind
                        Grid_Hak.Refresh
                    
                    id_hak = IIf(Not IsNull(!id_hak), !id_hak, "")
                    NAMA = IIf(Not IsNull(!Nama_APlikasi), !Nama_APlikasi, "")
                    ket = IIf(Not IsNull(!ket), !ket, "")
                    tambah = IIf(Not IsNull(!tambah), !tambah, 0)
                    rubah = IIf(Not IsNull(!rubah), !rubah, 0)
                    hapus = IIf(Not IsNull(!hapus), !hapus, 0)
                    laporan = IIf(Not IsNull(!Cetak_Laporan), !Cetak_Laporan, 0)
                    
                    Arr_Hak(a, 0) = id_hak
                    Arr_Hak(a, 1) = a
                    Arr_Hak(a, 2) = NAMA
                    Arr_Hak(a, 3) = ket
                    Arr_Hak(a, 4) = IIf((tambah = True), 1, 0)
                    Arr_Hak(a, 5) = IIf((rubah = True), 1, 0)
                    Arr_Hak(a, 6) = IIf((hapus = True), 1, 0)
                    Arr_Hak(a, 7) = IIf((laporan = True), 1, 0)
                    
                a = a + 1
                .MoveNext
                Loop
                
                Grid_Hak.ReBind
                Grid_Hak.Refresh
                
                Grid_Hak.MoveFirst
                
            End With
        
End Sub

Private Sub Form_Resize()
    
'    With Picture1
'        .Top = 50
'        .Left = Me.Width / 2 - .Width / 2
'    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

If kon.State = adStateOpen Then
    
    kon.Close
    Set kon = Nothing
End If

'If kon1.State = adStateOpen Then
'
'    kon1.Close
'    Set kon1 = Nothing
'End If
    

End Sub

Private Sub Grid_Hak_Click()
    
    If Arr_Hak.UpperBound(1) = 1 And Arr_Hak(1, 1) = Empty Then
        
        cek_akses_tambah.Value = vbUnchecked
        cek_akses_rubah.Value = vbUnchecked
        cek_akses_hapus.Value = vbUnchecked
        cek_akses_laporan.Value = vbUnchecked
        
    Else
        
        cek_akses_tambah.Value = IIf((Arr_Hak(Grid_Hak.Bookmark, 4) = 1), vbChecked, vbUnchecked)
        cek_akses_rubah.Value = IIf((Arr_Hak(Grid_Hak.Bookmark, 5) = 1), vbChecked, vbUnchecked)
        cek_akses_hapus.Value = IIf((Arr_Hak(Grid_Hak.Bookmark, 6) = 1), vbChecked, vbUnchecked)
        cek_akses_laporan.Value = IIf((Arr_Hak(Grid_Hak.Bookmark, 7) = 1), vbChecked, vbUnchecked)
        
    End If
    
End Sub

Private Sub Grid_Hak_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Grid_Hak_Click
End Sub

Private Sub Grid_User_DblClick()
    
    If Arr_User.UpperBound(1) = 1 And Arr_User(1, 1) = Empty Then Exit Sub
    
    Isi_Hak_Akses_PerUser Arr_User(Grid_User.Bookmark, 1)
    
End Sub

Private Sub Grid_User_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    Grid_User_DblClick
    
End Sub
