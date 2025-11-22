' *Calculate / Calculate Velocity
' !!! Do not change the line above !!!

'
' ================================================================================================
' Copyright 2007-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 25-May-2007 mbk: initial version
'------------------------------------------------------------------------------------

Option Explicit

Sub Main () 

	Dim cst_eng As Double
	Dim cst_velocity
	Dim cst_ruheeng As Double

	cst_eng = 1
	cst_ruheeng = 0.51e6 'Ruheenergie des Elektrons

	While True

	cst_velocity = 2.998e8*Sqr(1-1/(cst_eng/cst_ruheeng+1)^2)
	
	Begin Dialog UserDialog 340,154,"Calculate Velocity from Energy" ' %GRID:10,7,1,1
		Text 30,28,100,14,"Energy  (eV)",.Text1
		TextBox 150,28,170,21,.eng
		OKButton 110,91,90,21
		CancelButton 220,91,90,21
		Text 30,63,290,14,"Velocity: " + Format(cst_velocity,"Standard") + " m/s",.Text2
	End Dialog
		Dim dlg As UserDialog

		dlg.eng = CStr(cst_eng)

		If (Dialog(dlg) = 0) Then Exit All
		
		cst_eng = Evaluate(dlg.eng)
	Wend


End Sub

