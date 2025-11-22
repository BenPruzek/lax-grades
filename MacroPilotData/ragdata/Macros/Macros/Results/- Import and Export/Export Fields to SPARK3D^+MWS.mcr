' ExportToSpark3D

Sub Main ()
Dim res As Object          'Create an empty object
Set res = Result3D("")
MsgBox("The .f3e file to be imported into SPARK 3D can be found here:" + vbCrLf + vbCrLf + res.ExportFieldsToSPARK3D + vbCrLf + vbCrLf + "Please note, that SPARK3D version has to be 1.6.1 or higher.",vbInformation,"Export to SPARK finished")
End Sub
