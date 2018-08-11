Attribute VB_Name = "Module2"
Option Explicit

Function Konv1(n)
Dim vNilai, vHasil As Long
Dim kiri, kanan, kanan1, kanan2, kanan3, tengah, angka As String
vNilai = n
Select Case vNilai
Case 0 To 999
   Konv1 = n
     Select Case Len(Konv1)
         Case 1
              Konv1 = "              " & Konv1
         Case 2
              Konv1 = "             " & Konv1
         Case 3
              Konv1 = "            " & Konv1
     End Select
Case 1000 To 999999
   kanan = Trim(Right(Str(n), 3))
   kiri = Trim(Left(Str(n), Len(Str(n)) - 3))
   Konv1 = kiri & "." & kanan
     Select Case Len(Konv1)
         Case 5
              Konv1 = "          " & Konv1
         Case 6
              Konv1 = "         " & Konv1
         Case 7
              Konv1 = "        " & Konv1
     End Select

Case 1000000 To 999999999
   kanan = Trim(Right(Str(n), 6))
   kanan1 = Trim(Right(Str(n), 3))
   tengah = Trim(Left(kanan, 3))
   kiri = Trim(Left(Str(n), Len(Str(n)) - 6))
   Konv1 = Trim(kiri & "." & tengah & "." & kanan1)
     Select Case Len(Konv1)
         Case 9
              Konv1 = "      " & Konv1
         Case 10
              Konv1 = "     " & Konv1
         Case 11
              Konv1 = "    " & Konv1
     End Select


Case Is > 999999999
   kanan = Trim(Right(Str(n), 9))
   kanan1 = Trim(Right(kanan, 3))
   kanan2 = Trim(Mid(kanan, 4, 3))
   kanan3 = Trim(Left(kanan, 3))
   kiri = Trim(Left(Str(n), Len(Str(n)) - 9))
   Konv1 = Trim(kiri & "." & kanan3 & "." & kanan2 & "." & kanan1)
     Select Case Len(Konv1)
         Case 13
              Konv1 = "  " & Konv1
         Case 14
              Konv1 = " " & Konv1
         Case 15
              Konv1 = Konv1
              
     End Select
     
End Select
End Function


Function Konv2(n)
Dim vNilai, vHasil As Long
Dim kiri, kanan, kanan1, tengah, angka As String
vNilai = n
Select Case vNilai
Case 0 To 999
   Konv2 = n
     Select Case Len(Konv2)
         Case 1
              Konv2 = "           " & Konv2
         Case 2
              Konv2 = "          " & Konv2
         Case 3
              Konv2 = "         " & Konv2
     End Select
Case 1000 To 999999
   kanan = Trim(Right(Str(n), 3))
   kiri = Trim(Left(Str(n), Len(Str(n)) - 3))
   Konv2 = kiri & "." & kanan
     Select Case Len(Konv2)
         Case 5
              Konv2 = "       " & Konv2
         Case 6
              Konv2 = "      " & Konv2
         Case 7
              Konv2 = "     " & Konv2
     End Select
   
Case Is > 999999
   kanan = Trim(Right(Str(n), 6))
   kanan1 = Trim(Right(Str(n), 3))
   tengah = Trim(Left(kanan, 3))
   kiri = Trim(Left(Str(n), Len(Str(n)) - 6))
   Konv2 = Trim(kiri & "." & tengah & "." & kanan1)
     Select Case Len(Konv2)
         Case 9
              Konv2 = "   " & Konv2
         Case 10
              Konv2 = "  " & Konv2
         Case 11
              Konv2 = " " & Konv2
     End Select
End Select
End Function
Private Function Lokasiker(n)
Dim vloker As String
Select Case n
   Case 1
     Lokasiker = "Tanjung Enim"
   Case 2
     Lokasiker = "Kertapati"
   Case 3
     Lokasiker = "Tarahan"
   Case 4
     Lokasiker = "Jakarta"
   Case 5
     Lokasiker = "Ombilin"
   Case 6
     Lokasiker = "Samarinda"
   Case 7
     Lokasiker = "Banjarmasin"
   Case 9
     Lokasiker = "Briket T.Enim"
   Case 10
     Lokasiker = "Briket Lampung"
   Case 11
     Lokasiker = "Briket Serang"
   Case 12
     Lokasiker = "Briket jakarta"
   Case 13
     Lokasiker = "Briket Cilacap"
   Case 14
     Lokasiker = "Briket Semarang"
   Case 15
     Lokasiker = "Briket Gresik"
   Case 16
     Lokasiker = "Bukit Kendi"
   Case Else
     Lokasiker = "-"
End Select
End Function
Function terbilang(NILAI)
Dim mil, mily, vnil, juta, ribu, satu, jut, rib, sat, ucap As String
vnil = Format(Right(NILAI, 12), "000000000000")
mily = Mid(vnil, 1, 3)
juta = Mid(vnil, 4, 3)
ribu = Mid(vnil, 7, 3)
satu = Mid(vnil, 10, 3)

If mily = "000" Then
   mil = ""
    Else
   ucap = ucapan(mily)
   mil = ucap & " " & "Milyar"
   End If
If juta = "000" Then
 jut = ""
  Else
   ucap = ucapan(juta)
   jut = ucap & " " & "Juta"
End If
If ribu = "000" Then
   rib = ""
 Else
   ucap = ucapan(ribu)
   rib = ucap & " " & "Ribu"
   End If
 If satu = "000" Then
   sat = ""
 Else
   ucap = ucapan(satu)
   sat = ucap
   End If
   terbilang = mil & " " & jut & " " & rib & " " & sat & " " & "Rupiah"
     
End Function
Function ucapan(bilang)

Dim ratusan, puluhan, satuan, sratus As String
Dim spuluh, ssatu As String

ratusan = Left(bilang, 1)
puluhan = Mid(bilang, 2, 1)
satuan = Right(bilang, 1)
Select Case ratusan
  Case "0"
   sratus = ""
   Case "1"
   sratus = "Seratus"
  Case "2"
   sratus = "Dua Ratus"
  Case "3"
   sratus = "Tiga Ratus"
  Case "4"
   sratus = "Empat Ratus"
  Case "5"
   sratus = "Lima Ratus"
  Case "6"
   sratus = "Enam Ratus"
  Case "7"
   sratus = "Tujuh Ratus"
  Case "8"
   sratus = "Delapan Ratus"
  Case "9"
   sratus = "Sembilan Ratus"
End Select
Select Case puluhan
  Case "0"
   spuluh = ""
  Case "1"
   spuluh = ""
  Case "2"
   spuluh = "Dua Puluh"
  Case "3"
   spuluh = "Tiga Puluh"
  Case "4"
   spuluh = "Empat Puluh"
  Case "5"
   spuluh = "Lima Puluh"
  Case "6"
   spuluh = "Enam Puluh"
  Case "7"
   spuluh = "Tujuh Puluh"
  Case "8"
   spuluh = "Delapan Puluh"
  Case "9"
   spuluh = "Sembilan Puluh"
  End Select
    If puluhan = "1" Then
      Select Case satuan
         Case "0"
       ssatu = "Sepuluh"
         Case "1"
       ssatu = "Sebelas"
        Case "2"
       ssatu = "Duabelas"
        Case "3"
       ssatu = "Tigabelas"
        Case "4"
       ssatu = "Empatbelas"
        Case "5"
       ssatu = "Limabelas"
        Case "6"
       ssatu = "Enambelas"
        Case "7"
       ssatu = "Tujuhbelas"
        Case "8"
       ssatu = "Delapanbelas"
        Case "9"
       ssatu = "Sembilanbelas"
       End Select
    Else
     Select Case satuan
    Case "0"
      ssatu = ""
    Case "1"
      ssatu = "Satu"
    Case "2"
      ssatu = "Dua"
    Case "3"
      ssatu = "Tiga"
    Case "4"
      ssatu = "Empat"
    Case "5"
      ssatu = "Lima"
    Case "6"
      ssatu = "Enam"
    Case "7"
      ssatu = "Tujuh"
    Case "8"
      ssatu = "Delapan"
    Case "9"
      ssatu = "Sembilan"
   End Select
   End If
    ucapan = sratus & " " & spuluh & " " & ssatu
  
  End Function
Function Terbil(NILAI)
Dim mil, mily, vnil, juta, ribu, satu, jut, rib, sat, ucap As String
vnil = Format(Right(NILAI, 12), "000000000000")
mily = Mid(vnil, 1, 3)
juta = Mid(vnil, 4, 3)
ribu = Mid(vnil, 7, 3)
satu = Mid(vnil, 10, 3)

If mily = "000" Then
   mil = ""
    Else
   ucap = Kata(mily)
   mil = ucap & " " & "Milyar"
   End If
If juta = "000" Then
 jut = ""
  Else
   ucap = Kata(juta)
   jut = ucap & " " & "Juta"
End If
If ribu = "000" Then
   rib = ""
 Else
   ucap = Kata(ribu)
   rib = ucap & " " & "Ribu"
   End If
 If satu = "000" Then
   sat = ""
 Else
   ucap = Kata(satu)
   sat = ucap
   End If
   Terbil = mil & " " & jut & " " & rib & " " & sat
     
End Function
Function Kata(bilang)

Dim ssatu, spuluh, sratus, ratusan, puluhan, satuan As String

ratusan = Left(bilang, 1)
puluhan = Mid(bilang, 2, 1)
satuan = Right(bilang, 1)
Select Case ratusan
  Case "0"
   sratus = ""
   Case "1"
   sratus = "Seratus"
  Case "2"
   sratus = "Dua Ratus"
  Case "3"
   sratus = "Tiga Ratus"
  Case "4"
   sratus = "Empat Ratus"
  Case "5"
   sratus = "Lima Ratus"
  Case "6"
   sratus = "Enam Ratus"
  Case "7"
   sratus = "Tujuh Ratus"
  Case "8"
   sratus = "Delapan Ratus"
  Case "9"
   sratus = "Sembilan Ratus"
End Select
Select Case puluhan
  Case "0"
   spuluh = ""
  Case "1"
   spuluh = ""
  Case "2"
   spuluh = "Dua Puluh"
  Case "3"
   spuluh = "Tiga Puluh"
  Case "4"
   spuluh = "Empat Puluh"
  Case "5"
   spuluh = "Lima Puluh"
  Case "6"
   spuluh = "Enam Puluh"
  Case "7"
   spuluh = "Tujuh Puluh"
  Case "8"
   spuluh = "Delapan Puluh"
  Case "9"
   spuluh = "Sembilan Puluh"
  End Select
    If puluhan = "1" Then
      Select Case satuan
         Case "0"
       ssatu = "Sepuluh"
         Case "1"
       ssatu = "Sebelas"
        Case "2"
       ssatu = "Duabelas"
        Case "3"
       ssatu = "Tigabelas"
        Case "4"
       ssatu = "Empatbelas"
        Case "5"
       ssatu = "Limabelas"
        Case "6"
       ssatu = "Enambelas"
        Case "7"
       ssatu = "Tujuhbelas"
        Case "8"
       ssatu = "Delapanbelas"
        Case "9"
       ssatu = "Sembilanbelas"
       End Select
    Else
     Select Case satuan
    Case "0"
      ssatu = ""
    Case "1"
      ssatu = "Satu"
    Case "2"
      ssatu = "Dua"
    Case "3"
      ssatu = "Tiga"
    Case "4"
      ssatu = "Empat"
    Case "5"
      ssatu = "Lima"
    Case "6"
      ssatu = "Enam"
    Case "7"
      ssatu = "Tujuh"
    Case "8"
      ssatu = "Delapan"
    Case "9"
      ssatu = "Sembilan"
   End Select
   End If
    Kata = sratus & " " & spuluh & " " & ssatu
  
  End Function
Function konbul(NN)
Dim cmonth As Integer
cmonth = Month(Date)
Select Case cmonth
  Case 1: konbul = "I"
  Case 2: konbul = "II"
  Case 3: konbul = "III"
  Case 4: konbul = "IV"
  Case 5: konbul = "V"
  Case 6: konbul = "VI"
  Case 7: konbul = "VII"
  Case 8: konbul = "VIII"
  Case 9: konbul = "IX"
  Case 10: konbul = "X"
  Case 11: konbul = "XI"
  Case 12: konbul = "XII"
End Select

End Function

