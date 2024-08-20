Attribute VB_Name = "DocSoTien"
Function VND(ByVal So As Long) As String
    VND = DocSo(So) & " " & ChrW(273) & ChrW(7891) & "ng"
End Function
Function DocSo(ByVal So As Long) As String
    'Mã hóa hiễn thị Tiếng Việt
    Dim HangDonVi As String, HangChuc As String, HangTram As String, HangNgan As String, HangTrieu As String, HangTy As String
    Dim SoDu As Long, SoHang As Long
    Dim DonVi() As Variant
    DonVi = Array("", "m" & ChrW(7897) & "t", "hai", "ba", "b" & ChrW(7889) & "n", "n" & ChrW(259) & "m", "s" & ChrW(225) & "u", "b" & ChrW(7843) & "y", "t" & ChrW(225) & "m", "ch" & ChrW(237) & "n")
    Dim DonViChuc() As Variant
    DonViChuc = Array("", "m" & ChrW(432) & ChrW(7901) & "i", "hai" & " m" & ChrW(432) & ChrW(417) & "i", "ba" & " m" & ChrW(432) & ChrW(417) & "i", "b" & ChrW(7889) & "n" & " m" & ChrW(432) & ChrW(417) & "i", "n" & ChrW(259) & "m" & " m" & ChrW(432) & ChrW(417) & "i", "s" & ChrW(225) & "u" & " m" & ChrW(432) & ChrW(417) & "i", "b" & ChrW(7843) & "y" & " m" & ChrW(432) & ChrW(417) & "i", "t" & ChrW(225) & "m" & " m" & ChrW(432) & ChrW(417) & "i", "ch" & ChrW(237) & "n" & " m" & ChrW(432) & ChrW(417) & "i")
    Dim DonViTram() As Variant
    DonViTram = Array("", "m" & ChrW(7897) & "t" & " " & "tr" & ChrW(259) & "m", "hai" & " " & "tr" & ChrW(259) & "m", "ba" & " " & "tr" & ChrW(259) & "m", "b" & ChrW(7889) & "n" & " " & "tr" & ChrW(259) & "m", "n" & ChrW(259) & "m" & " " & "tr" & ChrW(259) & "m", "s" & ChrW(225) & "u" & " " & "tr" & ChrW(259) & "m", "b" & ChrW(7843) & "y" & " " & "tr" & ChrW(259) & "m", "t" & ChrW(225) & "m" & " " & "tr" & ChrW(259) & "m", "ch" & ChrW(237) & "n" & " " & "tr" & ChrW(259) & "m")

    'Xử lý từng hàng đơn vị
    If So = 0 Then
        DocSo = "kh" & ChrW(244) & "ng"
        Exit Function
    End If
    
    If So Mod 10 > 0 Then
        HangDonVi = ""
        If So Mod 100 > 11 And So Mod 10 = 1 Then
            HangDonVi = "m" & ChrW(7889) & "t"
        Else
            HangDonVi = DonVi(So Mod 10)
        End If
    Else
        HangDonVi = ""
    End If
    
    If So > 9 Then
        HangChuc = ""
        SoHang = Int((So Mod 100) / 10)
        If SoHang > 0 Then
            HangChuc = DonViChuc(SoHang) & " "
        ElseIf SoHang = 0 Then
            If (So Mod 100) = 0 Then
            HangChuc = ""
            Else
            HangChuc = "l" & ChrW(7867) & " "
            End If
        End If
    Else
        DocSo = Trim(HangDonVi)
        Exit Function
    End If
    
    If So > 99 Then
        HangTram = ""
        SoHang = Int((So Mod 1000) / 100)
        If (So Mod 1000) > 0 Then
            If SoHang = 0 Then
                HangTram = "kh" & ChrW(244) & "ng" & " " & "tr" & ChrW(259) & "m" & " "
            Else
                HangTram = DonViTram(SoHang) & " "
            End If
        End If
    Else
        DocSo = Trim(HangChuc & HangDonVi)
        Exit Function
    End If
    
    If So > 999 Then
        HangNgan = ""
        SoHang = Int((So Mod 1000000) / 1000)
        If SoHang > 0 Then
            HangNgan = DocSo(SoHang) & " " & "ng" & ChrW(224) & "n" & " "
        End If
    Else
        DocSo = Trim(HangTram & HangChuc & HangDonVi)
        Exit Function
    End If
    
    If So > 999999 Then
        HangTrieu = ""
        SoHang = Int((So Mod 1000000000) / 1000000)
        If SoHang > 0 Then
            HangTrieu = DocSo(SoHang) & " " & "tri" & ChrW(7879) & "u" & " "
        End If
    Else
        DocSo = Trim(HangNgan & HangTram & HangChuc & HangDonVi)
        Exit Function
    End If
    
    If So > 999999999 Then
        HangTy = ""
        SoDu = So
        SoHang = Int(So / 1000000000)
        If SoHang > 0 Then
            HangTy = DocSo(SoHang) & " " & "t" & ChrW(7927) & " "
        End If
    Else
        DocSo = Trim(HangTrieu & HangNgan & HangTram & HangChuc & HangDonVi)
        Exit Function
    End If
    DocSo = Trim(HangTy & HangTrieu & HangNgan & HangTram & HangChuc & HangDonVi)
End Function

