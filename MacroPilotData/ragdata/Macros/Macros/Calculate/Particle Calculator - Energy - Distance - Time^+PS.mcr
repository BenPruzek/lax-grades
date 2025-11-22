' ================================================================================================
' Copyright 2008-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 21-Jul-2009 ube: temperature-kelvin added
' 21-Feb-2008 ube: first version
'------------------------------------------------------------------------------------

Option Explicit

Const kboltzman = 1.3806504E-23
Const masselectron = 9.10938215E-31

Sub Main ()

	Dim i_energy As Integer

	Dim rev As Double
	Dim rgamma As Double
	Dim rbeta As Double
	Dim rvelo As Double
	Dim rmomt As Double
	Dim rkelvin As Double

	Dim i_timedist As Integer

	Dim dist_m As Double
	Dim rtime As Double

	i_energy = 1
	rgamma = 2.0

	i_timedist = 1
	dist_m = 0.1

	While True
	
	Begin Dialog UserDialog 730,224,"Particle Calculator (ONLY ELECTRONS !!)" ' %GRID:10,7,1,1
		Text 20,7,120,14,"check input value",.Text4
		OptionGroup .Group_energy
			OptionButton 30,28,20,14,"eV",.OptionButton1
			OptionButton 30,56,20,14,"OptionButton2",.OptionButton2
			OptionButton 30,84,20,14,"OptionButton3",.OptionButton3
			OptionButton 30,112,120,14,"velocity (m/s)",.OptionButton4
			OptionButton 30,140,130,14,"norm.momentum",.OptionButton5
			OptionButton 30,168,130,14,"temper. (Kelvin)",.OptionButton6
		Text 60,28,50,14,"eV",.Text1
		Text 60,56,60,14,"gamma",.Text2
		Text 60,84,80,14,"beta",.Text3
		TextBox 170,28,170,21,.ev
		TextBox 170,56,170,21,.ga
		TextBox 170,84,170,21,.beta
		TextBox 170,112,170,21,.v
		TextBox 170,140,170,21,.u
		TextBox 170,168,170,21,.k
		OKButton 20,196,90,21
		CancelButton 120,196,90,21
		OptionGroup .Group_timedist
			OptionButton 410,63,90,14,"time (s)",.OptionButton7
			OptionButton 410,91,100,14,"distance (m)",.OptionButton8
		TextBox 530,63,180,21,.t
		TextBox 530,91,180,21,.dist_m
		Text 410,35,130,14,"check input value",.Text5
	End Dialog
		Dim dlg As UserDialog

		dlg.Group_energy = i_energy
		Select Case i_energy
			Case 0 ' given ev
				rgamma= 1+Abs(rev)/511e3
				rbeta = Sqr(1-1/(rgamma*rgamma))
				rvelo= clight*rbeta
				rmomt= Sqr((rgamma*rgamma)-1)
				rkelvin = masselectron*rvelo*rvelo/(3*kboltzman)

			Case 1 ' given rgamma
				rev   = (rgamma-1)*511e3
				rbeta = Sqr(1-1/(rgamma*rgamma))
				rvelo = clight*rbeta
				rmomt = Sqr((rgamma*rgamma)-1)
				rkelvin = masselectron*rvelo*rvelo/(3*kboltzman)

			Case 2 ' given rbeta
				rgamma= 1/Sqr(1-(rbeta*rbeta))
				rev=(rgamma-1)*511e3
				rvelo =clight*rbeta
				rmomt =Sqr((rgamma*rgamma)-1)
				rkelvin = masselectron*rvelo*rvelo/(3*kboltzman)

			Case 3 ' given rvelo
				rbeta  = rvelo/clight
				rgamma = 1/Sqr(1-(rbeta*rbeta))
				rev = (rgamma-1)*511e3
				rmomt= Sqr((rgamma*rgamma)-1)
				rkelvin = masselectron*rvelo*rvelo/(3*kboltzman)

			Case 4 ' given rmomt
				rgamma = Sqr(1+(rmomt*rmomt))
				rev = (rgamma-1)*511e3
				rbeta = Sqr(1-1/(rgamma*rgamma))
				rvelo = clight*rbeta
				rkelvin = masselectron*rvelo*rvelo/(3*kboltzman)

			Case 5 ' given kelvin
				rvelo = Sqr(3*kboltzman*rkelvin/masselectron)
				rbeta = rvelo/clight
				rgamma = 1/Sqr(1-(rbeta*rbeta))
				rev = (rgamma-1)*511e3
				rmomt= Sqr((rgamma*rgamma)-1)
		End Select

		dlg.ev=CStr(rev)
		dlg.ga=CStr(rgamma)
		dlg.beta=CStr(rbeta)
		dlg.v=CStr(rvelo)
		dlg.u=CStr(rmomt)
		dlg.k=CStr(rkelvin)

		dlg.Group_timedist = i_timedist
		Select Case i_timedist
			Case 0 ' given time
				dist_m = rvelo * rtime
			Case 1 ' given distance
				rtime = dist_m / rvelo
		End Select

		dlg.t=CStr(rtime)
		dlg.dist_m=CStr(dist_m)

		If (Dialog(dlg) = 0) Then Exit All
		
		rev = Evaluate(dlg.ev)
		rgamma = Evaluate(dlg.ga)
		rbeta = Evaluate(dlg.beta)
		rvelo = Evaluate(dlg.v)
		rmomt = Evaluate(dlg.u)
		rkelvin = Evaluate(dlg.k)

		i_energy = dlg.Group_energy
		i_timedist = dlg.Group_timedist

		rtime = Evaluate(dlg.t)
		dist_m = Evaluate(dlg.dist_m)
	Wend


End Sub

