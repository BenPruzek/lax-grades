' *File / Get Customer-ID and Host-ID
' !!! Do not change the line above !!!

Sub Main () 
	MsgBox  "Customer-ID = "     + GetLicenseCustomerNumber + vbCrLf + _
			"Host-ID         = " + GetLicenseHostId
End Sub
