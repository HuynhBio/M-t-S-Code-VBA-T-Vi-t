Attribute VB_Name = "Module1"
Function VND(ByVal So As Long) As String
    VND = DocSo(So) & " " & ChrW(273) & ChrW(7891) & "ng"
End Function
Function DocSo(ByVal So As Long) As String
    'Mã hóa hiễn thị Tiếng Việt
    Dim HangDonVi As String, HangChuc As String, HangTram As String, HangNgan As String, HangTrieu As String, HangTy As String
    Dim SoDu As Long, SoHang As Long, Ntest As Long
    Dim khong As String, mot As String, mott As String, hai As String, ba As String, bon As String, nam As String, sau As String, bay As String, tam As String, chin As String
    Dim muoi As String, muoii As String, tram As String, ngan As String, trieu As String, ty As String
    Dim le As String
    khong = "kh" & ChrW(244) & "ng"
    mot = "m" & ChrW(7897) & "t"
    mott = "m" & ChrW(7889) & "t"
    hai = "hai"
    ba = "ba"
    bon = "b" & ChrW(7889) & "n"
    nam = "n" & ChrW(259) & "m"
    sau = "s" & ChrW(225) & "u"
    bay = "b" & ChrW(7843) & "y"
    tam = "t" & ChrW(225) & "m"
    chin = "ch" & ChrW(237) & "n"
    muoi = "m" & ChrW(432) & ChrW(7901) & "i"
    muoii = " m" & ChrW(432) & ChrW(417) & "i"
    tram = " tr" & ChrW(259) & "m"
    ngan = " ng" & ChrW(224) & "n"
    trieu = " tri" & ChrW(7879) & "u"
    ty = " t" & ChrW(7927)
    le = "l" & ChrW(7867)
    Dim DonVi() As Variant
    DonVi = Array("", mot, hai, ba, bon, nam, sau, bay, tam, chin)
    Dim DonViChuc() As Variant
    DonViChuc = Array("", muoi, hai & muoii, ba & muoii, bon & muoii, nam & muoii, sau & muoii, bay & muoii, tam & muoii, chin & muoii)
    Dim DonViTram() As Variant
    DonViTram = Array("", mot & tram, hai & tram, ba & tram, bon & tram, nam & tram, sau & tram, bay & tram, tam & tram, chin & tram)

    'Xử lý từng hàng đơn vị
    HangTy = ""
    SoDu = So
    SoHang = Int(SoDu / 1000000000)
    If SoHang > 0 Then
        HangTy = DocSo(SoHang) & ty & " "
        SoDu = SoDu Mod 1000000000
    End If

    HangTrieu = ""
    SoHang = Int(SoDu / 1000000)
    If SoHang > 0 Then
        HangTrieu = DocSo(SoHang) & trieu & " "
        SoDu = SoDu Mod 1000000
    End If

    HangNgan = ""
    SoHang = Int(SoDu / 1000)
    If SoHang > 0 Then
        HangNgan = DocSo(SoHang) & ngan & " "
        SoDu = SoDu Mod 1000
    End If

    HangTram = ""
    If So > 99 Then
        SoHang = Int(SoDu / 100)
        If SoHang = 0 Then
            HangTram = khong & tram & " "
        Else
            HangTram = DonViTram(SoHang) & " "
            SoDu = SoDu Mod 100
        End If
    End If

    HangChuc = ""
    If So > 9 Then
        SoHang = Int(SoDu / 10)
        If SoHang > 0 Then
            HangChuc = DonViChuc(SoHang) & " "
        ElseIf SoHang = 0 Then
            If SoDu Mod 10 = 0 Then
            HangChuc = ""
            Else
            HangChuc = le & " "
            End If
        End If
    End If
    If SoDu > 11 And SoDu Mod 10 = 1 Then
        HangDonVi = mott
    Else
        HangDonVi = DonVi(SoDu Mod 10)
    End If
    DocSo = Trim(HangTy & HangTrieu & HangNgan & HangTram & HangChuc & HangDonVi)
    If DocSo = "" Then DocSo = khong
End Function

