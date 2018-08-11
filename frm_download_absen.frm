VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_download_absen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Download Data Absensi"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
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
   ScaleHeight     =   1725
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      Calculator      =   "frm_download_absen.frx":0000
      Caption         =   "frm_download_absen.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_download_absen.frx":0085
      Keys            =   "frm_download_absen.frx":00A3
      Spin            =   "frm_download_absen.frx":00ED
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "frm_download_absen.frx":0115
      Left            =   840
      List            =   "frm_download_absen.frx":0117
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tahun :"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bulan :"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frm_download_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jammasuk As String
Dim potperjam As Double
Dim mxpot As Double
Dim jampulang As String

Private Sub buka_util()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim jmlpermenit As Double
    Dim jmldetik As Single
    
    
    sql = "select * from mas_util"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, kon
    
    With rs
        jammasuk = !jam_masuk
        potperjam = !pot_perjam
        
        jmlpermenit = potperjam / 5
       ' jmldetik = jmlpermenit / 60
        
        
        potperjam = jmlpermenit
        
        mxpot = !mak_potongan
        jampulang = !jam_pulang
    End With
    
End Sub

Private Sub downloaddata()
On Error GoTo salah:
    
    Dim p As Integer
    
    MousePointer = vbHourglass
    
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\att2000.mdb" & ";Persist Security Info=False"
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim rs3 As Recordset
    
    Dim id, id2 As String
    Dim tanggal As String
    Dim jam1 As String
    Dim jam2 As String
    Dim stat As String
    Dim ket As String
    Dim a As Integer
        a = 0
        
    Dim totpot, totlembur As Double
        totpot = 0
        totlembur = 0
        
    Dim tglk1, tglk2 As Date
    
        
    Dim sql As String
    Dim sql2 As String
    Dim sql3 As String
    
        sql = "select * from qinout where bulan=" & Combo1.ListIndex + 1 & " and tahun=" & TDBNumber1.Value
        sql = sql & " order by userid asc"
        
        kon.BeginTrans
        
    Set rs = New ADODB.Recordset
    rs.Open sql, cn
    With rs
        
'        If .re <= 0 Then
'            p = CInt(MsgBox("Tidak ada data absensi yang akan diproses", vbInformation + vbOKOnly, "Informasi"))
'            MousePointer = vbDefault
'            Exit Sub
'        End If
        
        Do While Not .EOF
            
            id = !userid
            
            If id <> id2 And a > 0 Then
                set_total_absen id2, totlembur, totpot
                totpot = 0
                totlembur = 0
            End If
            
            tanggal = !tanggal
            jam1 = !jammasuk
            jam2 = !jamkeluar
            stat = !stat
            ket = !ket
            
            sql3 = "delete from tbabsen where kode_absen='" & id & "' and tanggal='" & Format(tanggal, "yyyy/mm/dd") & "'"
            sql3 = sql3 & " and jam1='" & jam1 & "' and jam2='" & jam2 & "' and bulan=" & Combo1.ListIndex + 1 & " and tahun=" & TDBNumber1.Value
            
            Set rs3 = New ADODB.Recordset
                rs3.Open sql3, kon
            
            sql2 = "insert into tbAbsen (kode_absen,tanggal,jam1,jam2,bulan,tahun,stat,ket) values("
            sql2 = sql2 & "'" & id & "','" & Format(tanggal, "yyyy/mm/dd") & "',"
            sql2 = sql2 & "'" & jam1 & "','" & jam2 & "'," & Combo1.ListIndex + 1 & "," & TDBNumber1.Value & ",'" & stat & "','" & ket & "')"
            
            Set rs2 = New ADODB.Recordset
            rs2.Open sql2, kon
            
            tglk1 = jammasuk
            tglk2 = Format(jammasuk, "dd/mm/yyyy") & " " & jam1
                Dim selisihjam As Integer
                    selisihjam = DateDiff("n", tglk1, tglk2)
                
                Dim jmlpot As Double
                    jmlpot = 0
                If selisihjam > 0 Then
                    jmlpot = selisihjam * potperjam
                    
                    If jmlpot > 100000 Then
                        jmlpot = mxpot
                    Else
                        jmlpot = jmlpot
                    End If
                    
                End If
                
                totpot = totpot + jmlpot
                
                tglk1 = jampulang
                tglk2 = Format(jampulang, "dd/mm/yyyy") & " " & jam2
                selisihjam = DateDiff("h", tglk1, tglk2)
                
                If selisihjam >= 1 Then
                    totlembur = totlembur + setlembur(id)
                End If
                
                id2 = !userid
                
            a = a + 1
            
        .MoveNext
        
            If .EOF Then
                set_total_absen id2, totlembur, totpot
                totpot = 0
                totlembur = 0
            End If
        
        Loop
    
    End With
    
    rs.Close
    cn.Close
    
    MousePointer = vbDefault
    
    kon.CommitTrans
    
    If a = 0 Then
        p = CInt(MsgBox("Tidak ada data absensi yang akan diproses", vbInformation + vbOKOnly, "Informasi"))
    Else
        p = CInt(MsgBox("Download Data Selesai....", vbInformation + vbOKOnly, "Informasi"))
    End If
    
    
    
    On Error GoTo 0
    Exit Sub
    
salah:
    
    MousePointer = vbDefault
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    
    
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
    
End Sub

Public Function setlembur(ByVal kodeabsen As String) As Double
    
    Dim lem, mak, trans As Double
    
    Dim tot As Double
    
    Dim sql As String
    Dim rs As Recordset
    
    sql = "select lembur,makan,transport from View_jab_terakhir where kd_absen='" & kodeabsen & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, kon
    
    With rs
        If Not .EOF Then
            lem = IIf(Not IsNull(!lembur), !lembur, 0)
            mak = IIf(Not IsNull(!makan), !makan, 0)
            trans = IIf(Not IsNull(!transport), !transport, 0)
            
            tot = lem + mak + trans
            
        End If
    End With
    
    setlembur = tot
    
End Function

Private Sub set_total_absen(ByVal kdabsen As String, ByVal totlembur As Double, ByVal totpot As Double)
    
    Dim sql1 As String
    Dim sql2 As String
    
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    
    sql1 = "delete from tbAbsen2 where kd_absen='" & kdabsen & "' and bulan=" & Combo1.ListIndex + 1 & " and tahun=" & TDBNumber1.Value
    
    Set rs1 = New ADODB.Recordset
        rs1.Open sql1, kon
        
    sql2 = "insert into tbAbsen2 (kd_absen,bulan,tahun,jmlpot,jmllembur) values('" & kdabsen & "'"
    sql2 = sql2 & "," & Combo1.ListIndex + 1 & "," & TDBNumber1.Value
    sql2 = sql2 & "," & totpot & "," & totlembur & ")"
    
    Set rs2 = New ADODB.Recordset
        rs2.Open sql2, kon
    
End Sub

Private Sub Command1_Click()
    buka_util
    downloaddata
End Sub

Private Sub Form_Activate()
    Combo1.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Left = Utama.Width / 2 - Me.Width / 2
    Me.Top = (Utama.Height / 2 - Me.Height / 2) - 1500
    
    TDBNumber1.Value = "2012"
    
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
    
End Sub
