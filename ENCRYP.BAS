Attribute VB_Name = "Encrypt"
' Problem noted when offset = a single digit hex number
' decode routine is looking for 2 digits. simple fix noted below
'fix by Michael W. Smith (74001,457) of NBS
' routine look good!
'
Function crypt(Action As String, Key As String, Src As String) As String
Dim Count As Integer, KeyPos As Integer, KeyLen As Integer, SrcAsc As Integer, dest As String, offset As Integer, TmpSrcAsc, SrcPos
KeyLen = Len(Key)

If Action = "E" Then
    Randomize
    offset = (Rnd * 10000 Mod 255) + 1

    dest = Hex$(offset)  ' problem with "offset" of single digit hex numbers
    ' when decodeing, decode procedure is looking for 2 digits, whereis lower number produce single digit hex numbers (ie 1,2,3,12...)
    If Len(dest) = 1 Then   ' adds 0 in front of single digit hex numbers
        dest = "0" + dest
    End If

    For SrcPos = 1 To Len(Src)
        SrcAsc = (Asc(Mid$(Src, SrcPos, 1)) + offset) Mod 255
        If KeyPos < KeyLen Then KeyPos = KeyPos + 1 Else KeyPos = 1
        'Fill Dest$ with HEX representation of Encrypted field
        'Hex used to keep nasties such as eof or lf from mangling stream
        'Use format$ to make Hex$ return " 0" instead of "0" when the same
        'values are Xor'ed together (Null) - keeps placeholder for decrypt
        SrcAsc = SrcAsc Xor Asc(Mid$(Key, KeyPos, 1))
        dest = dest + Format$(Hex$(SrcAsc), "@@")
        offset = SrcAsc

    Next

ElseIf Action = "D" Then
    offset = Val("&H" + Left$(Src, 2))
    For SrcPos = 3 To Len(Src) Step 2
        SrcAsc = Val("&H" + Trim(Mid$(Src, SrcPos, 2)))
        If KeyPos < KeyLen Then KeyPos = KeyPos + 1 Else KeyPos = 1
        TmpSrcAsc = SrcAsc Xor Asc(Mid$(Key, KeyPos, 1))
        If TmpSrcAsc <= offset Then
            TmpSrcAsc = 255 + TmpSrcAsc - offset
        Else
            TmpSrcAsc = TmpSrcAsc - offset
        End If
        dest = dest + Chr(TmpSrcAsc)
        offset = SrcAsc
    Next

End If
crypt = dest
End Function

