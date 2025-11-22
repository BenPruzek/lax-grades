' Export results for EMIT

' ================================================================================================
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 13-Mar-2012 ube: initial version  Jwa
' 28-Jun-2012 ube: support no pattern. set the WG ports' location to 0,0,0 if tetra/surface mesh used JWA

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "mws_ports.lib"

Sub Main ()
Dim CST_FN As Long, cst_iff As Long, cst_np As Long,cst_pName As String
Dim cst_EMIT As String, cst_folder As String
Dim cst_filename As String,cst_farfieldname As String,CST_FFN As Long

	'=============================================================================
	Dim cst_tree() As String, cst_tmpstr As String
	Dim cst_iloop As Long, cst_iloop2 As Long
	Dim cst_nff As Long, cst_nom As Long
	Dim cst_ffq(3) As String
	Dim b_ffs As Boolean


	Dim cst_theta_start As Double, cst_theta_step As Double, cst_theta_stop As Double, cst_ntheta As Long
	Dim cst_phi_start As Double, cst_phi_step As Double, cst_phi_stop As Double, cst_nphi As Long
	Dim cst_phi As Double, cst_theta As Double, cst_theta_calc As Double, cst_phi_calc As Double
	Dim cst_origin_type As String
	Dim cst_origin_x As Double, cst_origin_y As Double, cst_origin_z As Double
	Dim cst_dt As Double, cst_dp As Double

	Dim cst_icomp As Integer, cst_ncomp As Integer
	Dim cst_floop_start As Long, cst_floop_end As Long, cst_ifloop As Long

	'Dim cst_title As String, cst_title_ini As String, cst_ffname As String, cst_filename As String
	Dim cst_radpow As Double, cst_ffam_renorm As Double
	'Dim CST_FN As Long, cst_iff As Long
	Dim cst_theta_ph_renorm As Double, cst_phi_ph_renorm As Double, cst_phase_renorm As Double
	Dim cst_theta_re As Double, cst_theta_im As Double, cst_theta_am As Double, cst_theta_ph As Double
	Dim cst_phi_re As Double, cst_phi_im As Double, cst_phi_am As Double, cst_phi_ph As Double
	Dim cst_th_re As Double, cst_th_im As Double, cst_ph_re As Double, cst_ph_im As Double
	Dim cst_ffpol(3) As String
	Dim cst_fforigin As String
	Dim cst_xval As Double, cst_yval As Double, cst_zval As Double
	Dim cst_rad As Double, cst_rad_re As Double, cst_rad_im As Double
	Dim cst_time As Double
	'=============================================================================

	cst_time = Timer
	'On Error GoTo ERROR_NO_FARFIELDS
	ReDim cst_tree(1)
	cst_tree(0) = Resulttree.GetFirstChildName ("Farfields")
	cst_iloop = 0
	'If cst_tree(0)="" Then
	'	GoTo ERROR_NO_FARFIELDS
	'End If
	Do
		cst_tmpstr = Resulttree.GetNextItemName(cst_tree(cst_iloop))
		If cst_tmpstr <> "" Then
			cst_iloop = cst_iloop+1
			ReDim Preserve cst_tree(cst_iloop)
			cst_tree(cst_iloop)=cst_tmpstr
		End If
	Loop Until cst_tmpstr = ""
	cst_nff = cst_iloop+1

	For cst_iloop = 1 To cst_nff
		cst_tmpstr = Replace(cst_tree(cst_iloop-1),"Farfields\","")
		cst_tree(cst_iloop-1)=cst_tmpstr
	Next cst_iloop

	ReDim cst_freq(cst_nff)
	cst_nom = Monitor.GetNumberOfMonitors
	For cst_iloop = 1 To cst_nom
		If Monitor.GetMonitorTypeFromIndex(cst_iloop-1) = "Farfield" Then
			For cst_iloop2 = 1 To cst_nff
				If InStr(cst_tree(cst_iloop2-1),Monitor.GetMonitorNameFromIndex(cst_iloop-1)) > 0 Then
					cst_freq(cst_iloop2-1) = Monitor.GetMonitorFrequencyFromIndex(cst_iloop-1)
				End If
			Next cst_iloop2
		End If
	Next cst_iloop

	'Check if Directory where List and files should be stored exists otherwise create

	cst_EMIT = "\_exported files for EMIT"
	cst_folder = GetProjectPath("Project")+cst_EMIT

    On Error GoTo Folder_already_existing
    	MkDir cst_folder
	Folder_already_existing:
	On Error GoTo 0

'-----export Antenna Attribute--------
    Dim portlist() As String
	Dim np As Long
	Dim iii As Long
	Dim tmpstr As String
    Dim nptot As Long

	np = Solver.GetNumberOfPorts

	If np=0 Then
		MsgBox "No Ports defined."
	End If

	nptot = getNoPortsModes (np)

'----------------------------------------
' redim portlist()
'---------------------------------------
	Dim jjj As Long, kkk As Long
	Dim ALable As String,ALocation As String, ARotation As String
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, z0 As Double, z1 As Double
    Dim orie_cst As Long

  ' "np=nptot" For All only one mode ports,"np<nptot" With high order mode ports.

	ReDim portlist(nptot)
	kkk = 0
	For iii = 0 To np-1
		If Port.gettype(PortNumberArray(iii)) = "Waveguide" Then
			For jjj = 1 To Port.GetNumberOfModes (PortNumberArray(iii))
				kkk=kkk+1
				If Port.GetNumberOfModes (PortNumberArray(iii)) >1 Then
					portlist(kkk) = CStr(PortNumberArray(iii))+ "("+CStr(jjj)+")
				Else
					portlist(kkk) = CStr(PortNumberArray(iii))+"(1)"   'IIf(np=nptot,"","(1)" )
				End If
			Next jjj
		Else
			kkk=kkk+1
			portlist(kkk) = CStr(PortNumberArray(iii))+"(1)"  'IIf(np=nptot,"","(1)" )
		End If
	Next iii

         CST_FN = FreeFile
         cst_filename=cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+".att"

        Open cst_filename For Output As #CST_FN

        		For iii = 0 To np-1

                ALocation="0, 0, 0"
                If Port.gettype(PortNumberArray(iii)) = "Discrete" Then

                 DiscretePort.GetCoordinates( PortNumberArray(iii), x0, y0, z0, x1, y1,z1)
                 ALocation=Str((x0+x1)/2*Units.GetGeometryUnitToSI )+ ", " + Str((y0+y1)/2*Units.GetGeometryUnitToSI)+  ", " +Str((z0+z1)/2*Units.GetGeometryUnitToSI)
                End If

                If Port.gettype(PortNumberArray(iii)) = "Waveguide" Then

                 	If Mesh.GetMeshType =  "PBA" Then
                    Port.GetPortMeshCoordinates( PortNumberArray(iii),orie_cst, x0, x1, y0,y1, z0, z1)
                    ALocation=Str((x0+x1)/2*Units.GetGeometryUnitToSI )+ ", " + Str((y0+y1)/2*Units.GetGeometryUnitToSI)+  ", " +Str((z0+z1)/2*Units.GetGeometryUnitToSI)
                    Else
                     x0=0:    x1=0 :   y0=0 :  y1=0 : z0=0 : z1=0      ' set the location to 0,0,0 for tetra mesh and surface mesh
                    End If
                 End If

                 ARotation="0, 0, 0"

                'ALable=IIf(Port. GetLabel(PortNumberArray(iii))="","Antenna-"+PortNumberArray(iii),Port. GetLabel(PortNumberArray(iii)))
                'Print #CST_FN, ALable + ", "+ PortNumberArray(iii) + ", " + ALocation + ", "+ ARotation

                 Print #CST_FN, "Antenna-"+PortNumberArray(iii) + ", "+ PortNumberArray(iii) + ", " + ALocation + ", "+ ARotation

'-------------export farfield for each port---------------------------------------------

	If cst_tree(0)="" Then
		GoTo NO_FARFIELDS
	End If


   CST_FFN = FreeFile
   cst_farfieldname=cst_folder+"\"+"Antenna-"+PortNumberArray(iii)+".ffs"


   With FarfieldPlot
     .Plottype "3D"
     .Vary "angle1"
     .Theta "0"
     .Phi "0"
     .Step "5"
     .Step2 "5"
     .SetLockSteps "True"
     .SetPlotRangeOnly "False"
     .SetThetaStart "0"
     .SetThetaEnd "180"
     .SetPhiStart "0"
     .SetPhiEnd "360"
     .SetTheta360 "False"
     .SymmetricRange "False"
     .SetTimeDomainFF "False"
     .SetFrequency "150"
     .SetTime "0"
     .SetColorByValue "True"
     .DrawStepLines "False"
     .DrawIsoLongitudeLatitudeLines "False"
     .ShowStructure "True"
     .SetStructureTransparent "False"
     .SetFarfieldTransparent "False"
     .SetSpecials "enablepolarextralines"
     .SetPlotMode "Directivity"
     .Distance "1"
     .UseFarfieldApproximation "True"
     .SetScaleLinear "True"
     .SetLogRange "40"
     .SetLogNorm "0"
     .DBUnit "0"
     .EnableFixPlotMaximum "False"
     .SetFixPlotMaximumValue "1.0"
     .SetInverseAxialRatio "False"
     .SetAxesType "xyz"
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000"
     .Thetastart "0.000000e+000", "1.000000e+000", "1.000000e+000"
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000"
     .SetCoordinateSystemType "spherical"
     .SetPolarizationType "Linear"
     .SlantAngle 0.000000e+000
     .Origin "bbox"
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000"
     .SetUserDecouplingPlane "False"
     .UseDecouplingPlane "False"
     .DecouplingPlaneAxis "X"
     .DecouplingPlanePosition "0.000000e+000"
     .EnablePhaseCenterCalculation "False"
     .SetPhaseCenterAngularLimit "3.000000e+001"
     .SetPhaseCenterComponent "boresight"
     .SetPhaseCenterPlane "both"
     .ShowPhaseCenter "True"
     .StoreSettings
End With

    b_ffs=False
    cst_iloop = -1
	Do
    cst_iloop=cst_iloop+1

    b_ffs=(Right(cst_tree(cst_iloop),Len(cst_tree(cst_iloop))-InStr(cst_tree(cst_iloop),"[")+1)="["+Trim(PortNumberArray(iii))+"]")
    b_ffs=b_ffs Or(Right(cst_tree(cst_iloop),Len(cst_tree(cst_iloop))-InStr(cst_tree(cst_iloop),"[")+1)= "["+Trim(PortNumberArray(iii))+"(1)]")


    Loop Until b_ffs


   SelectTreeItem("Farfields\"+cst_tree(cst_iloop))


   FarfieldPlot.ASCIIExportAsBroadbandSource (cst_farfieldname )


   NO_FARFIELDS:

'-------------end of export farfield for each port--------------------------------------

	            Next iii

   Close #CST_FN

'-----end of export Antenna Attribute--------
'-------export STL file (combine each solid files)-----------

Dim cst_index As Long, cst_iii As Long, Component_cst As String, Shape_cst As String

cst_index=Solid.GetNumberOfShapes

For cst_iii=0 To cst_index-1

Component_cst=Left(Solid.GetNameOfShapeFromIndex(cst_iii),InStr(Solid.GetNameOfShapeFromIndex(cst_iii),":")-1)
Shape_cst=Replace(Solid.GetNameOfShapeFromIndex(cst_iii),Component_cst+":", "")


With STL
    .Reset
    .FileName (cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_iii) +".stl")
    .name(Shape_cst)
    .Component (Component_cst)
    .Write
End With

  CST_FN = FreeFile
         cst_filename=cst_folder + Replace(GetProjectPath("Project"),GetProjectPath("Root"),"") + ".stl"

         If cst_iii=0 Then

             STL_string0$=""
         Else
             Open cst_filename For Input As #CST_FN
             Input  #CST_FN, STL_string0$
             Close #CST_FN
         End If

         Open cst_filename For Output As #CST_FN

         CST_FN2 = FreeFile
         Open cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_iii) +".stl" For Input As #CST_FN2
         Input  #CST_FN2, STL_string$
         STL_string$ = Replace(STL_string$,cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_iii) +".stl",Solid.GetNameOfShapeFromIndex(cst_iii))
         Print  #CST_FN , STL_string0$+STL_string$
         Close #CST_FN2

         Kill cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_iii) +".stl"

   Close #CST_FN

Next cst_iii

  CST_FN = FreeFile
  cst_filename=cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+"_STLgeometry unit.txt"
  Open cst_filename For Output As #CST_FN
  Print  #CST_FN , "Exported STL unit is " & Units.GetUnit("Length")
  Close #CST_FN

'-----export S parameters

With TOUCHSTONE
    .Reset
    .FileName (cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),""))
    .Impedance (50)
    .FrequencyRange ("Full")
    .Renormalize (True)
    .UseARResults (False)
    .SetNSamples (100)
    .Write
End With

	'Exit All

    'ERROR_NO_FARFIELDS:
	'MsgBox("No farfield results available !!",vbOkOnly+vbCritical,"Macro Execution Stopped")
	'Exit Sub

End Sub


Function getNoPortsModes(np As Long) As Long
	Dim nptot As Long, iii As Long
	nptot = np
	FillPortNumberArray
	For iii = 0 To np-1
		If Port.gettype(PortNumberArray(iii)) = "Waveguide" Then	nptot = nptot + Port.GetNumberOfModes (PortNumberArray(iii))-1
		If Port.GetNumberOfModes(PortNumberArray(iii)) > 1 Then multimode = True
	Next iii
	getNoPortsModes = nptot
End Function
