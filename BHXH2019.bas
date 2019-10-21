Attribute VB_Name = "BHXH2019"
Public KETQUA
Sub TINHBAOHIEM(GIATRI, GiaTriDoiChieu)
	If GIATRI < GiaTriDoiChieu Then
		KETQUA = GIATRI
	Else
		KETQUA = GiaTriDoiChieu
	End If
End Sub
'TINH BHXH PHAI DONG
Function BHXH(TongLuong)
	Dim XH, TD
	XH = TongLuong * 0.08 'BHXH BANG 8% LUONG, TOI DA 2384000 VND
	TD = 2384000
	TINHBAOHIEM XH, TD
	BHXH = KETQUA
End Function
'TINH BHYT PHAI DONG
Function BHYT(TongLuong)
	Dim YT, TD
	YT = TongLuong * 0.015 'BHYT BANG 1.5% LUONG TOI DA 447000 VND
	TD = 447000
	TINHBAOHIEM YT, TD
	BHYT = KETQUA
End Function
'TINH BHTN PHAI DONG
Function BHTN(TongLuong, VungLamViec)
	Dim TN, TD
	TN = TongLuong * 0.01 'BHTN BANG 1% LUONG TOI DA THEO VUNG
	If VungLamViec = 1 Or VungLamViec = "I" Then 'VUNG I TOI DA 836000 VND
		TD = 836000
		TINHBAOHIEM TN, TD
		BHTN = KETQUA
	ElseIf VungLamViec = 2 Or VungLamViec = "II" Then 'VUNG II TOI DA 742000 VND
		TD = 742000
		TINHBAOHIEM TN, TD
		BHTN = KETQUA
	ElseIf VungLamViec = 3 Or VungLamViec = "III" Then 'VUNG III TOI DA 650000 VND
		TD = 650000
		TINHBAOHIEM TN, TD
		BHTN = KETQUA
	ElseIf VungLamViec = 4 Or VungLamViec = "IV" Then 'VUNG IV TOI DA 584000 VND
		TD = 584000
		TINHBAOHIEM TN, TD
		BHTN = KETQUA
	Else
		BHTN = "NHAP VUNG LAM VIEC !"
	End If
End Function
'TINH TONG BAO HIEM PHAI DONG
Function BAOHIEM(TongLuong, VungLamViec) 'CAC GHI CHU TUONG TU HAM TREN
	'##### TINH BHXH PHAI DONG #####
	Dim BaoHiemXaHoi
	BaoHiemXaHoi = BHXH(TongLuong)
	'##### TINH BHYT PHAI DONG #####
	Dim BaoHiemYTe
	BaoHiemYTe = BHYT(TongLuong)
	'##### TINH BHTN PHAI DONG #####
	Dim BaoHiemTaiNan
	BaoHiemTaiNan = BHTN(TongLuong, VungLamViec)
	BAOHIEM = BaoHiemXaHoi + BaoHiemYTe + BaoHiemTaiNan
End Function
