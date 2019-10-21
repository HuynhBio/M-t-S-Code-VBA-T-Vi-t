Attribute VB_Name = "Tinh_Thue"
Public THUE
Sub TINHTHUE(GIATRIMUCTHUE, GIATRITOIDA, PHANTRAM)
	If GIATRIMUCTHUE < 0 Then
		THUE = 0
	ElseIf GIATRIMUCTHUE < GIATRITOIDA Then
		THUE = GIATRIMUCTHUE * PHANTRAM
	Else
		THUE = GIATRITOIDA * PHANTRAM
	End If
End Sub
Function ThueTNCN2019(ThuNhap, SoNguoiPhuThuoc)
	Dim MUC1, MUC2, MUC3, MUC4, MUC5, Muc6, Muc7, MucKhongThue, KhauTruPhuThuoc
	Dim THUEMUC1, THUEMUC2, THUEMUC3, THUEMUC4, THUEMUC5, THUEMUC6, THUEMUC7
	MUC1 = ThuNhap - 9000000 - SoNguoiPhuThuoc * 3600000 'MUC KHONG THUE LA 9000000, MOI NGUOI PHU THUOC GIAM 3600000
	MUC2 = MUC1 - 5000000
	MUC3 = MUC2 - 5000000
	MUC4 = MUC3 - 8000000
	MUC5 = MUC4 - 14000000
	Muc6 = MUC5 - 20000000
	Muc7 = Muc6 - 28000000
	Call TINHTHUE(MUC1, 5000000, 0.05)
	THUEMUC1 = THUE
	Call TINHTHUE(MUC2, 5000000, 0.1)
	THUEMUC2 = THUE
	Call TINHTHUE(MUC3, 8000000, 0.15)
	THUEMUC3 = THUE
	Call TINHTHUE(MUC4, 14000000, 0.2)
	THUEMUC4 = THUE
	Call TINHTHUE(MUC5, 20000000, 0.25)
	THUEMUC5 = THUE
	Call TINHTHUE(Muc6, 28000000, 0.3)
	THUEMUC6 = THUE
	If Muc7 < 0 Then
		THUEMUC7 = 0
	Else
		THUEMUC7 = Muc7 * 0.35
	End If
	ThueTNCN2019 = THUEMUC1 + THUEMUC2 + THUEMUC3 + THUEMUC4 + THUEMUC5 + THUEMUC6 + THUEMUC7
End Function