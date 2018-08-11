VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form frm_penggajian 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penggajian"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
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
   ScaleHeight     =   7440
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   12480
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox tkodeslip 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdkalk 
      Caption         =   "&Kalkulasi"
      Height          =   615
      Left            =   12360
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "frm_penggajian.frx":0000
      Left            =   960
      List            =   "frm_penggajian.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      Calculator      =   "frm_penggajian.frx":0004
      Caption         =   "frm_penggajian.frx":0024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian.frx":0089
      Keys            =   "frm_penggajian.frx":00A7
      Spin            =   "frm_penggajian.frx":00F1
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
   Begin TrueOleDBGrid60.TDBGrid grid1 
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "frm_penggajian.frx":0119
      TabIndex        =   5
      Top             =   960
      Width           =   13695
   End
   Begin IsButton_Ard.isButton cmdedit 
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
      Icon            =   "frm_penggajian.frx":5433
      Style           =   1
      Caption         =   "-"
      IconAlign       =   0
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Informasi"
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin TDBDate6Ctl.TDBDate ttgl 
      Height          =   315
      Left            =   7320
      TabIndex        =   11
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frm_penggajian.frx":5B05
      Caption         =   "frm_penggajian.frx":5C1D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_penggajian.frx":5C89
      Keys            =   "frm_penggajian.frx":5CA7
      Spin            =   "frm_penggajian.frx":5D05
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Penggajian :"
      Height          =   210
      Left            =   5880
      TabIndex        =   10
      Top             =   480
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Transaksi :"
      Height          =   210
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bulan :"
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tahun :"
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frm_penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr_kalk As New XArrayDB
Dim totpot, totlembur As Double

Private Sub kosongkan_grid()
    arr_kalk.ReDim 0, 0, 0, 0
    arr_kalk.ReDim 1, 1, 1, grid1.Columns.Count
    grid1.ReBind
    grid1.Refresh
End Sub
Private Sub isi_bkti()
    
    Dim sql As String
    Dim rs As Recordset
    Dim noutur As Integer
        noutur = 0
    
        sql = "select count(KodeSlip) as jml from tbGajiHeader where tahun=" & TDBNumber1.Value
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon
        
        With rs
            If Not .EOF Then
                noutur = IIf(Not IsNull(!jml), !jml, 0)
            End If
        End With
    
        rs.Close
    
        noutur = noutur + 1
        
        Dim notrans As String
        notrans = noutur
        
        If Len(notrans) = 1 Then
            notrans = "00" & notrans
        ElseIf Len(notrans) = 2 Then
            notrans = "0" & notrans
        End If
        
        Dim bulan As String
            bulan = Combo1.ListIndex + 1
            If Len(bulan) = 1 Then bulan = "0" & bulan
        
        notrans = notrans & bulan & TDBNumber1.Value
        
        tkodeslip.Text = notrans
        
End Sub

Private Sub hapus_sebelumnya()
    
    Dim sql As String
    Dim sql2 As String
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    
    Dim sql3 As String
    Dim rs3 As Recordset
    
    sql = "select kodeslip from tbGajiHeader where bulan=" & Combo1.ListIndex + 1 & " and tahun=" & TDBNumber1.Value
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    If Not rs.EOF Then
        
        MsgBox "Terdapat kalkulasi gaji terdahulu,data akan dihapus dan dikalkulasi ulang dengan data saat ini", vbInformation + vbOKOnly, "Konfirmasi"
        
        sql2 = "delete from tbGajiDetail where kodeslip='" & rs!kodeslip & "'"
        
        Set rs2 = New ADODB.Recordset
            rs2.Open sql2, kon
            
        sql3 = "delete from tbGajiHeader where kodeslip='" & rs!kodeslip & "'"
        
        Set rs3 = New ADODB.Recordset
            rs3.Open sql3, kon
        
    End If
    
End Sub

Private Sub simpan()
On Error GoTo er_data
    
    Dim sql As String
    Dim rs As Recordset
    
    kon.BeginTrans
    
    hapus_sebelumnya
    
    isi_bkti
    
    simpan_detail
    
    sql = "insert into tbGajiHeader (kodeslip,bulan,tahun,tanggal) values ('" & Trim(tkodeslip.Text) & "'," & Combo1.ListIndex + 1 & "," & TDBNumber1.Value & ",'" & Format(ttgl.Value, "yyyy/mm/dd") & "')"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
    
    kon.CommitTrans
    
    MsgBox "Data disimpan", vbOKOnly + vbInformation, "Informasi"
    
    kosongkan_grid
    
    On Error GoTo 0
    Exit Sub

er_data:
    
    kon.RollbackTrans
    
    Dim bukan_tipe
        bukan_tipe = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub simpan_detail()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim lembur As Double
    Dim pot As Double
    Dim gapok As Double
    Dim gajibersih As Double
    
    Dim a As Integer
        a = 1
        
        For a = 1 To arr_kalk.UpperBound(1)
            
            If IsEmpty(arr_kalk(a, 4)) Then
                lembur = 0
            Else
                lembur = arr_kalk(a, 4)
            End If
            
            If IsEmpty(arr_kalk(a, 5)) Then
                pot = 0
            Else
                pot = arr_kalk(a, 5)
            End If
            
            If IsEmpty(arr_kalk(a, 3)) Then
                gapok = 0
            Else
                gapok = arr_kalk(a, 3)
            End If
            
            If IsEmpty(arr_kalk(a, 6)) Then
                gajibersih = 0
            Else
                gajibersih = arr_kalk(a, 6)
            End If
            
            
            sql = "insert into tbGajiDetail (kodeslip,kodekaryawan,gajipokok,lembur,potongan,gajibersih,kode_jab) values("
            sql = sql & "'" & Trim(tkodeslip.Text) & "','" & arr_kalk(a, 0) & "'"
            sql = sql & "," & gapok & "," & lembur & "," & pot & "," & gajibersih & ",'" & arr_kalk(a, 7) & "')"
            
            Set rs = New ADODB.Recordset
                rs.Open sql, kon
            
        Next a
    
End Sub

Private Sub isi_grid_kalkulasi()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim kodekary As String
    Dim namakary As String
    Dim jab As String
    Dim gapok As Double
    Dim lembur As Double
    Dim pot As Double
    Dim totakhir As Double
    Dim kdabsen As String
    Dim kodejab As String
    Dim kodeabsen As String
    Dim gajibersih As Double
    
    Dim a As Integer
        a = 1
    
        sql = "select * from View_jab_terakhir order by kode_karyawan"
        
    MousePointer = vbHourglass
        
    Set rs = New ADODB.Recordset
        rs.Open sql, kon
        
        With rs
            Do While Not .EOF
            
                kodekary = IIf(Not IsNull(!kode_karyawan), !kode_karyawan, "")
                namakary = IIf(Not IsNull(!Nama_Karyawan), !Nama_Karyawan, "")
                jab = IIf(Not IsNull(!nama_jab), !nama_jab, "")
                gapok = IIf(Not IsNull(!gapok), !gapok, 0)
                kodejab = IIf(Not IsNull(!kode_jab), !kode_jab, "")
                kodeabsen = IIf(Not IsNull(!kd_absen), !kd_absen, "")
                
                If kodeabsen = "" Then
                    totpot = 0: totlembur = 0
                Else
                    cek_kalk (kodeabsen)
                End If
                
                gajibersih = (gapok + totlembur) - totpot
                
                arr_kalk.ReDim 1, a, 0, grid1.Columns.Count
                grid1.ReBind
                grid1.Refresh
                
                arr_kalk(a, 0) = kodekary
                arr_kalk(a, 1) = namakary
                arr_kalk(a, 2) = jab
                arr_kalk(a, 3) = gapok
                arr_kalk(a, 4) = IIf(IsEmpty(totlembur), 0, totlembur)
                arr_kalk(a, 5) = IIf(IsEmpty(totpot), 0, totpot)
                arr_kalk(a, 6) = gajibersih
                arr_kalk(a, 7) = kodejab
                
                a = a + 1
            .MoveNext
            Loop
                
            grid1.ReBind
            grid1.Refresh
                
              grid1.MoveLast
                
            MousePointer = vbDefault
            
        End With
    
End Sub

Private Sub cek_kalk(ByVal kdabsen As String)
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select jmlpot,jmllembur from tbAbsen2 where kd_absen='" & kdabsen & "' and bulan = " & Combo1.ListIndex + 1 & " and tahun=" & TDBNumber1.Value
    
    Set rs = New ADODB.Recordset
    rs.Open sql, kon
    
    With rs
        If Not .EOF Then
            totpot = IIf(Not IsNull(!jmlpot), !jmlpot, 0)
            totlembur = IIf(Not IsNull(!jmllembur), !jmllembur, 0)
        End If
    End With
    
End Sub

Private Sub cmdedit_Click()
    
    If arr_kalk.UpperBound(1) = 1 And arr_kalk(1, 1) = Empty Then Exit Sub
        
        Dim kodekary As String
        Dim namakary As String
        Dim jab As String
        Dim gapok As Double
        Dim lembur As Double
        Dim pot As Double
        
        
        kodekary = arr_kalk(grid1.Bookmark, 0)
        namakary = arr_kalk(grid1.Bookmark, 1)
        jab = arr_kalk(grid1.Bookmark, 2)
        gapok = arr_kalk(grid1.Bookmark, 3)
        lembur = arr_kalk(grid1.Bookmark, 4)
        pot = arr_kalk(grid1.Bookmark, 5)
        
        With frm_penggajian2
            .setting_awal kodekary, namakary, jab, gapok, lembur, pot
        End With
        frm_penggajian2.Show 1
        
End Sub

Private Sub cmdkalk_Click()
    
  '  tkodeslip.Text = Combo1.Text & "-" & TDBNumber1.Value
    
    kosongkan_grid

    isi_grid_kalkulasi
    
End Sub

Private Sub cmdsimpan_Click()

    If TDBNumber1.Value = 0 Then
        MsgBox "Tahun harus diisi", vbInformation + vbOKOnly, "Informasi"
        TDBNumber1.SetFocus
        Exit Sub
    End If

    If arr_kalk.UpperBound(1) = 1 And arr_kalk(1, 1) = Empty Then
        MsgBox "Tidak ada data yang akan disimpan", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If

    simpan
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    TDBNumber1.SetFocus
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

Me.Left = Utama.Width / 2 - Me.Width / 2
    Me.Top = (Utama.Height / 2 - Me.Height / 2) - 1500

ttgl.Value = Date
TDBNumber1.Text = Year(Now)

'TDBNumber1.Value = "2012"
    
        With Combo1
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
    
    Combo1.ListIndex = 0
    
    grid1.Array = arr_kalk
    kosongkan_grid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If kon.State = adStateOpen Then
        
        kon.Close
        Set kon = Nothing
    End If

End Sub

Private Sub grid1_AfterColUpdate(ByVal ColIndex As Integer)

On Error GoTo er_data

If ColIndex = 4 Then
    
        arr_kalk(grid1.Bookmark, ColIndex) = grid1.Columns(ColIndex).Text
                
    End If
    
    If ColIndex = 5 Then
    
         arr_kalk(grid1.Bookmark, ColIndex) = grid1.Columns(ColIndex).Text
        
    End If
   
   
   If ColIndex = 6 Then
        
         arr_kalk(grid1.Bookmark, ColIndex) = grid1.Columns(ColIndex).Text
        
   End If
   
   On Error GoTo 0
   Exit Sub
   
er_data:
    
    Dim bukan_tipe
        bukan_tipe = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub grid1_Change()

On Error GoTo er_data

    Dim pot As Double
    Dim lbr As Double
    Dim gapok As Double
    Dim gajibersih As Double
        
        
        If Not IsNumeric(grid1.Columns(4).Text) Then
            lbr = 0
        Else
        
            If grid1.Columns(4).Text = 0 Then
                lbr = 0
            Else
                lbr = grid1.Columns(4).Text
            End If
        
        End If
        
        
        If Not IsNumeric(grid1.Columns(5).Text) Then
            pot = 0
        Else
        
            If grid1.Columns(5).Text = 0 Then
                pot = 0
            Else
                pot = grid1.Columns(5).Text
            End If
        
        End If
    
        
               
    
     gapok = grid1.Columns(3).Text
    
     gajibersih = (gapok + lbr) - pot
        
        
     grid1.Columns(4).Text = lbr
     arr_kalk(grid1.Bookmark, 4) = grid1.Columns(4).Text
     
     grid1.Columns(5).Text = pot
     arr_kalk(grid1.Bookmark, 5) = grid1.Columns(5).Text
     
     grid1.Columns(6).Text = gajibersih
     arr_kalk(grid1.Bookmark, 6) = grid1.Columns(6).Text
            
        
'    grid1.ReBind
'    grid1.Refresh
        
     On Error GoTo 0
     Exit Sub
        
    
er_data:
    
    Dim bukan_tipe
        bukan_tipe = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

