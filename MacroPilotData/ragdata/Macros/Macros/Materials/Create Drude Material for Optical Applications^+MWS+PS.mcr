' *Calculate / Calculate Drude Parameter for Optical Applications
' !!! Do not change the line above !!!

' macro.546
' ================================================================================================
' Copyright 2006-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
'------------------------------------------------------------------------------------
' 10-Feb-2015 ckr: replace 3e8 by Clight, which is more accurate
' 20-Dec-2006 fde: first version
'------------------------------------------------------------------------------------

Option Explicit
'#include "vba_globals_all.lib"

	Public cst_frq As Double
	Public cst_lambda As Double
	Public cst_epss As Double
	Public cst_wavelength As Double
	Public cst_n As Double
	Public cst_k As Double
	Public cst_eps1 As Double
	Public cst_eps2 As Double
	Public cst_wp As Double
	Public cst_vc As Double
	Public cst_type As String
	Public cst_d1g_Parameter1 As String
	Public cst_d1g_Parameter2 As String
	Public cst_d1g_Parameter3 As String
	Public cst_d1g_Parameter4 As String
	Public cst_material_name As String
	Public sCommand As String


Sub Main ()
    Dim eps_str As String

	cst_lambda=633.0
	cst_frq = Clight/(633*1e-9)
	cst_epss=1.0
	cst_type = "n_k"
	cst_k = .001
	cst_n = 1
	cst_material_name = "Drude Metal"

	While True
	
		' wavelength = clight / frq

	Begin Dialog UserDialog 640,287,"Calculate Drude Parameter",.DialogFunc ' %GRID:10,7,1,1
		Text 340,70,40,14,"k",.Text_im
		Text 30,63,120,14,"eps infinity",.Text4
		Text 30,28,120,14,"Lambda  [nm]",.Text1
		Text 340,28,40,14,"n",.Text_re
		TextBox 170,28,110,21,.lambda_p
		TextBox 400,28,110,21,.const_re
		TextBox 170,63,110,21,.epss_p
		TextBox 400,63,110,21,.const_im
		PushButton 60,189,90,21, "Quit",.stopit
		GroupBox 30,105,290,70,"Select Input Parameter (n/k or eps1/eps2) ",.GroupBox1
		OptionGroup .options
			OptionButton 60,126,90,14,"n/k",.Option1
			OptionButton 60,147,120,14,"eps1/eps2",.Option2
		PushButton 180,189,90,21,"Calculate",.Calculate
		Text 340,119,120,14,"frequency [Thz]",.Text6
		Text 340,140,130,14,"plasma frequency",.Text3
		Text 340,161,120,14,"colision frequency",.Text5
		Text 60,231,150,14,"Material Name",.Text_material
		TextBox 180,224,180,21,.Materialname
		PushButton 60,259,200,21,"Create/Change Material",.CreateMaterial
		Text 500,119,90,14,"frequency",.Text_freq
		Text 500,140,90,14,"wp",.Text_wp
		Text 500,161,90,14,"vp",.Text_vc
	End Dialog
		Dim dlg As UserDialog
		dlg.lambda_p=CStr(cst_lambda)
		dlg.epss_p=CStr(cst_epss)
        dlg.const_re=CStr(cst_n)
        dlg.const_im=CStr(cst_k)
        dlg.Materialname = CStr(cst_material_name)

		
		If (Dialog(dlg) = 0) Then Exit All
		
		cst_lambda = Evaluate(dlg.lambda_p)
		cst_epss = Evaluate(dlg.epss_p)
		
	Wend
	End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)

    Select Case Action%  '(1 open)
    Case 1 ' Dialog box initialization
            DlgText "Text_re","n:"
			DlgText "Text_im","k:"
		   	DlgText "Text_freq", Left$(CStr(cst_frq/1e12),7)
		   	DlgText "const_re", Format$(cst_n,"0.0####")
		    DlgText "const_im", Format$(cst_k,"0.0####")
		    DlgEnable "CreateMaterial",False

			cst_type = "n_k"
    Case 2 ' Value changing or button pressed
        	Select Case DlgItem$ '(2 open )
            'Case "Help"
			'StartHelp "common_preloadedmacro_calculate_calculate_analytical_line_impedance"
			'DialogFunc = True
			Case "options"
			 Select Case SuppValue% '(3 open)
              Case 0
            	DlgText "Text_re","n:"
				DlgText "Text_im","k:"
				cst_n= Sqr((cst_eps1+Sqr(cst_eps1^2+cst_eps2^2))/(2))
				cst_k=  Sqr((-cst_eps1+Sqr(cst_eps1^2+cst_eps2^2))/(2))
				DlgText "const_re", Format$(cst_n,"0.0####")
				DlgText "const_im", Format$(cst_k,"0.0####")

				cst_type = "n_k"
				DialogFunc% = True
              Case 1
            	DlgText "Text_re","eps1:"
				DlgText "Text_im","eps2:"
            	cst_eps1 = cst_n^2-cst_k^2
            	cst_eps2 = 2*cst_n*cst_k

				DlgText "const_re", Format$(cst_eps1,"0.0####")
				DlgText "const_im", Format$(cst_eps2,"0.0####")

				cst_type = "eps1_eps2"

				DialogFunc% = True
               End Select '(3 close)
			 Case "Calculate"
             DlgEnable "CreateMaterial",True

              Select Case cst_type '(4 open)
             Case "n_k"
             cst_vc= -(2*cst_n*cst_k*2*pi*cst_frq)/(cst_n^2-cst_k^2-cst_epss)
             cst_wp=Sqr(-(cst_n^2-cst_k^2-cst_epss)*((2*pi*cst_frq)^2+cst_vc^2))
             DlgText "Text_wp", Format$(cst_wp,"0.00E+00")
             DlgText "Text_vc", Format$(cst_vc,"0.00E+00")

             DialogFunc% = True

             Case "eps1_eps2"
             cst_n= Sqr((cst_eps1+Sqr(cst_eps1^2+cst_eps2^2))/(2))
             cst_k=  Sqr((-cst_eps1+Sqr(cst_eps1^2+cst_eps2^2))/(2))

             cst_vc=- (2*cst_n*cst_k*2*pi*cst_frq)/(cst_n^2-cst_k^2-cst_epss)
             cst_wp=Sqr(-(cst_n^2-cst_k^2-cst_epss)*((2*pi*cst_frq)^2+cst_vc^2))
             DlgText "Text_wp", Format$(cst_wp,"0.00E+00")
             DlgText "Text_vc", Format$(cst_vc,"0.00E+00")

             DialogFunc% = True
               End Select '(4 close)

             Case "CreateMaterial"
                sCommand = ""
    			sCommand = sCommand + "With Material" + vbLf
    			sCommand = sCommand + "     .Reset" + vbLf
   				sCommand = sCommand + "     .Name """ + CStr(cst_material_name) + """" + vbLf
    			sCommand = sCommand + "     .FrqType """ + "All" + """" + vbLf
   				sCommand = sCommand + "     .Type """ + "Normal" +"""" + vbLf
   				sCommand = sCommand + "     .Epsilon """ + "1" +"""" + vbLf
   				sCommand = sCommand + "     .Mu """ + "1" +"""" + vbLf
   				sCommand = sCommand + "     .Kappa """ + "0" +"""" + vbLf
    			sCommand = sCommand + "     .SetMaterialUnit """ + "GHz" + """, """ + "mm" + """" + vbLf
   				sCommand = sCommand + "     .TanD """ + "0.0" +"""" + vbLf
    			sCommand = sCommand + "     .TanDFreq """ + "0.0" +"""" + vbLf
    			sCommand = sCommand + "     .TanDGiven """ + "False" +"""" + vbLf
    			sCommand = sCommand + "     .TanDModel """ + "ConstTanD" +"""" + vbLf
   			 	sCommand = sCommand + "     .KappaM """ + "0" +"""" + vbLf
   			 	sCommand = sCommand + "     .TanDM """ + "0.0" +"""" + vbLf
   			 	sCommand = sCommand + "     .TanDMFreq """ + "0.0" +"""" + vbLf
   			 	sCommand = sCommand + "     .TanDMGiven """ + "False" +"""" + vbLf
   			 	sCommand = sCommand + "     .TanDMModel  """ + "ConstTanD" +"""" + vbLf
   				sCommand = sCommand + "     .DispModelEps """ + "Drude" +"""" + vbLf
   				sCommand = sCommand + "     .EpsInfinity """ + CStr(cst_epss) +"""" + vbLf
   			 	sCommand = sCommand + "     .DispCoeff1Eps """ + CStr(cst_wp) +"""" + vbLf
   			 	sCommand = sCommand + "     .DispCoeff2Eps """ + CStr(cst_vc) +"""" + vbLf
   			 	sCommand = sCommand + "     .DispModelMu """ + "None" +"""" + vbLf
   			 	sCommand = sCommand + "     .DispersiveFittingSchemeEps """ + "General 1st" +"""" + vbLf
   			 	sCommand = sCommand + "     .DispersiveFittingSchemeMu """ + "General 1st" +"""" + vbLf
   			 	sCommand = sCommand + "     .UseGeneralDispersionEps """ + "False" +"""" + vbLf
   			 	sCommand = sCommand + "     .UseGeneralDispersionMu """ + "False" +"""" + vbLf
   			 	sCommand = sCommand + "     .Rho """ + "0" +"""" + vbLf
   			 	sCommand = sCommand + "     .ThermalType """ + "Normal" +"""" + vbLf
   			 	sCommand = sCommand + "     .ThermalConductivity """ + "0" +"""" + vbLf
   			 	sCommand = sCommand + "     .Colour """ + "0" + """,""" + "1" +""",""" + "1" + """" +  vbLf
    			sCommand = sCommand + "     .Transparency """ + "0" +"""" + vbLf
       			sCommand = sCommand + "     .Create" + vbLf
       			sCommand = sCommand + "End With" + vbLf


			AddToHistory "define Material: " + CStr(cst_material_name), sCommand


             DialogFunc% = True

            Case "stopit"
 			Exit All


          End Select '(2 close)

       Case 3   ' Text box changed
         Select Case DlgItem '(5 open)
            Case "lambda_p"
      			cst_d1g_Parameter1 = CStr(Abs(RealVal(CStr(DlgText("lambda_p")))))
	        	DlgText "lambda_p" ,cst_d1g_Parameter1
	        	cst_frq =Clight/Abs(RealVal(CStr(DlgText("lambda_p"))))*1e9
	        	DlgText "Text_freq", Left$(CStr(cst_frq/1e12),7)

	        Case "epss_p"
      			cst_d1g_Parameter2 = CStr(Abs(RealVal(CStr(DlgText("epss_p")))))
                cst_epss = Abs(RealVal(CStr(DlgText("epss_p"))))
	        	DlgText "epss_p" ,cst_d1g_Parameter2
	        Case "const_re"
	            If (cst_type = "n_k") Then
	             cst_n =Abs(RealVal(CStr(DlgText("const_re"))))
	             cst_d1g_Parameter3 = CStr(Abs(RealVal(CStr(DlgText("const_re")))))

	            Else
	             cst_eps1 = RealVal(CStr(DlgText("const_re")))
   	             cst_d1g_Parameter3 = CStr(RealVal(CStr(DlgText("const_re"))))


	             End If
	        	DlgText "const_re" ,cst_d1g_Parameter3
	        Case "const_im"
	             If (cst_type = "n_k") Then
	             cst_k =(Abs(RealVal(CStr(DlgText("const_im"))))) 'value alway positiv
	             cst_d1g_Parameter4 = CStr(Abs(RealVal(CStr(DlgText("const_im")))))
	             Else
	             cst_eps2 = (Abs(RealVal(CStr(DlgText("const_im"))))) 'value alway positive
	           	 cst_d1g_Parameter4 = CStr(Abs(RealVal(CStr(DlgText("const_im")))))
	             End If
	        	DlgText "const_im" ,cst_d1g_Parameter4

	        Case "Materialname"
	           cst_material_name = NoForbiddenFilenameCharacters(CStr(DlgText "Materialname"))
               DlgText "Materialname",cst_material_name


      	   End Select '(5 close)


	End Select '(1 close)
 	End Function


