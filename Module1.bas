Attribute VB_Name = "Module1"
Option Explicit
Public kon As New ADODB.Connection
Public X As String
Public M_Id_User As String
Public M_Id_Aplikasi As String
Public Frm As Form
Public Rs_Nav As ADODB.Recordset
Public Mysq As String
Public macem2 As String
Public macem2_lagi As String
Public macem2_ja As String
Public Flag_tempat As String
Public khusus_user As String
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public nama_otoris As String
Public jabatan_otoris As String

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const HKCU = &H80000001

Public tahun_u, bulan_u, hari_u As String

Public Sub kalkulasi_umur(ByVal tanggallahir As String)

Dim HARI As Double
Dim bulan As Double
Dim tahun As Double
Dim MAXIMALHARI As Double

Dim umurtahun As String
Dim umurbulan As String
Dim umurhari As String

Dim ketemu As String

ketemu = Format(tanggallahir, "dd-mm-yyyy")
HARI = DatePart("d", Now) - DatePart("d", ketemu)
bulan = DatePart("m", Now) - DatePart("m", ketemu)
tahun = DatePart("yyyy", Now) - DatePart("yyyy", ketemu)
MAXIMALHARI = Day((DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year(Format(Now, "dd/mm/yy")), Month(Format(Now, "dd/mm/yy")), 1)))))
If bulan < 0 And HARI >= 0 Then
umurtahun = tahun - 1
umurbulan = 12 + bulan
umurhari = HARI
ElseIf bulan < 0 And HARI <= 0 Then
umurtahun = tahun - 1
umurbulan = bulan + 11
umurhari = MAXIMALHARI + (HARI)
ElseIf bulan = 0 And HARI < 0 Then
umurtahun = tahun - 1
umurbulan = bulan + 11
umurhari = MAXIMALHARI + (HARI)
ElseIf bulan > 0 And HARI < 0 Then
umurtahun = tahun
umurbulan = bulan - 1
umurhari = MAXIMALHARI + HARI
Else
umurtahun = tahun
umurbulan = bulan
umurhari = HARI
End If

tahun_u = tahun
bulan_u = bulan
hari_u = HARI

End Sub

Public Function Cek_akses_Form(ByVal nama_form As String) As Boolean
    
    If kon.State = adStateClosed Then
        
        Buka_Koneksi
    
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select id_hak from VIEW_Hak_Akses where nama_form ='" & nama_form & "' and id_user=" & Flag_tempat
        
        Set rs = New ADODB.Recordset
            rs.Open sql, kon, adOpenKeyset
            
            If Not rs.EOF Then
                Cek_akses_Form = True
            Else
                Cek_akses_Form = False
            End If
        
    
End Function

Public Function FormatTgl(ByVal tgl As String) As String
    
    Dim dday, dmonth, dyear As String
        dday = DatePart("d", tgl)
        dmonth = DatePart("m", tgl)
        dyear = DatePart("yyyy", Date)
    
    Dim hasil As String
    hasil = dyear & "/" & dmonth & "/" & dday
    FormatTgl = Format(hasil, "yyyy/mm/dd")
    
End Function

Public Sub hak_akses_percommand(ByVal nama_form As String)

'c_tambah = 0
'c_rubah = 0
'c_hapus = 0
'c_laporan = 0

Dim comd As Command
Set comd = New ADODB.Command
With comd
    .ActiveConnection = kon
    .CommandText = "cek_id_aplikasi"
    .CommandType = adCmdStoredProc
    
    .Parameters("@nama_form").Value = nama_form
    .Execute
End With

Dim rs As Recordset
    Set rs = New ADODB.Recordset
    rs.Open comd
    
    With rs
        
        If Not .EOF Then
        
        Dim comd1 As Command
        Set comd1 = New ADODB.Command
        
        comd1.ActiveConnection = kon
        comd1.CommandText = "cek_hak_akses_percommand"
        comd1.CommandType = adCmdStoredProc
        comd1.Parameters("@Id_User").Value = Flag_tempat
        comd1.Parameters("@id_aplikasi").Value = !id
        
        comd1.Execute
        
'        c_tambah = IIf(Not IsNull(comd1.Parameters("@tambah")), comd1.Parameters("@tambah"), False)
'        c_rubah = IIf(Not IsNull(comd1.Parameters("@rubah")), comd1.Parameters("@rubah"), False)
'        c_hapus = IIf(Not IsNull(comd1.Parameters("@hapus")), comd1.Parameters("@hapus"), False)
'        c_laporan = IIf(Not IsNull(comd1.Parameters("@cetak_laporan")), comd1.Parameters("@cetak_laporan"), False)
        
        Set comd1.ActiveConnection = Nothing
        
        End If
        
        
    End With

Set comd.ActiveConnection = Nothing
Set rs = Nothing
End Sub

Public Function encrypt_pwd(ByVal pwd As String) As String

Dim hasil As String
    hasil = crypt("E", RTrim$("ARGIE"), RTrim$(pwd))
    
    encrypt_pwd = hasil
    
End Function

Public Function decrypt_pwd(ByVal pwd As String) As String

Dim hasil As String
    hasil = crypt("D", RTrim$("ARGIE"), RTrim$(pwd))
    
    decrypt_pwd = hasil
    
End Function

Public Function Buka_Koneksi() As String
    
    On Error GoTo salah:
    
    If kon.State = adStateOpen Then kon.Close
        
       ' kon.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=pegasus;Initial Catalog=ASKRINDO;Data Source=ardian\sql2008"  ' & Lokasi_Database & ""
        
        kon.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;PWD=123admin;Initial Catalog=ASKRINDO;Data Source=USER-06-PC\SQL2008" ' " & Lokasi_Database & ""
        
        Buka_Koneksi = Err.Number
    Exit Function
        
salah:

        If Buka_Lagi = 0 Then
            Exit Function
            Buka_Koneksi = 0
        Else
            Buka_Koneksi = "-2147467259"
            Exit Function
        End If
                
        Buka_Koneksi = Err.Number
                
End Function

Public Function Buka_Lagi() As String
    
    On Error GoTo salah:
    
    If kon.State = adStateOpen Then kon.Close
    
        kon.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=123admin;Initial Catalog=ASKRINDO;Data Source=" & Lokasi_Database & ""
        
        Buka_Lagi = Err.Number
    Exit Function
        
salah:
                
        Buka_Lagi = Err.Number
                
End Function



Public Function Batas() As Double
    
    Batas = GetSetting("bts", "bts", "bts", 0)
    
End Function

Public Function Lokasi_Database() As Variant
    
    Lokasi_Database = GetSetting("ValvKop", "ValvKop.v", "ValvKop.v01", 0)
    
End Function

Public Function Set_Lokasi_Database(ByVal Letak As String) As Boolean
    
On Error GoTo err_handler

    SaveSetting "ValvKop", "ValvKop.v", "ValvKop.v01", Letak
    SaveSetting "bts", "bts", "bts", "0"
    
    Set_Lokasi_Database = True
    
    On Error GoTo 0
    Exit Function

err_handler:
    
    Set_Lokasi_Database = False
    
    Dim p As Integer
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
       
End Function

Public Sub Focus_(ByVal obj As Object)
On Error Resume Next
    
    With obj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Public Function periksa_tanggal(ByVal tgl As String) As Boolean
    
    On Error GoTo err_tu
    
    Dim periksa As String
        periksa = CStr(CDate(tgl))
        
        periksa_tanggal = True
    
    On Error GoTo 0
    Exit Function
    
err_tu:
        periksa_tanggal = False
    
End Function
