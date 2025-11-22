
' ================================================================================================
' All the selected files are exported in the result folder specified by the user <.\userdefined_name>.dlxcst"
' It converts the Field Source Monitor (.fsm) file to .NFS folder (24 .dat + 24 .xml files).
' It exports the selected farfield results as .ffs files.
' It exports the 3D CAD structure to STL (units: meter) and the wires (if present) to IGES (units: meter).


' ================================================================================================
' Copyright 2013-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------
' 28-Mar-2018 ube: Excluded farfield cuts from selection list
' 28-Jul-2015 fsr: Replaced obsolete GetFileFromItemName with GetFileFromTreeItem
' 15-Dec-2014 dta: corrected index problem when writing complex impdedance data
' 01-Jul-2014 dta: Rearranged dialog box field related to Nearfield results
' 01-Jul-2014 dta: Removed macro execution stop when no FSM monitors are available. If no FSM are available or is FSM export is not selected CAD & Wire exports are disabled.
' 13-Jun-2014 dta: Reworked CAD export to STL (VBA:ExportFileUnits)
' 23-May-2014 dta: Corrected problem with function bSeparateFarfieldFrq (e.g. f=f0).
' 22-Apr-2014 dta: Fixed problem when exporting Sim.Excitation results (No Z-Parameter available)
' 07-Mar-2014 dta: Scaling /CAD export performed always using Global Coordinate System
' 20-Feb-2014 dta: Export folder fully userdefined
' 28-Jan-2014 dta: Added function DRound to make equidistant farfield check more robust
' 27-Jan-2014 dta: Exit the macro if there are no near field source results available
'				   Added check to allow the export of excitations with equidistant farfield sampling
'				   Added function to retrieve frequency when bSeparateFarfieldFrq is False
' 22-Jan-2014 dta: Corrected problem when dealing with broadband farfield source extraction named without "automatic labeling"
'				   Farfield is not exported by default even if available
' 07-Nov-2013 dta: Added automatic extension ".dlxcst" to the selectable export folder
'				   Moved ff settings to Properties button
' 24-Sep-2013 dta: Corrected bug to get correct scaling factor
' 02-Sep-2013 dta: Index file containing NF & FF file names
' 19-Aug-2013 dta: Boundary conditions info exported to ASCII file (if NFS export is performed)
'                  Complex Impedance export for near field source monitor frequency points
' 23-Jul-2013 dta: Added farfield export capability (broadband/single frequency,
'				   All .fsm files in the result folder automatically converted To NFS and copied To SAVANT folder
' 03-Jul-2013 dta: Automatic Wires export to IGES in meter
' 12-Jun-2013 dta: Automatic CAD export to STL in meter
' 04-Jun-2013 dta: Added CAD export capability
' 05-Apr-2013 dta: Initial version
' ================================================================================================

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"
'#include "mws_evaluate-results.lib"


Const bDebug = False 'Debug flag
Dim GlobalDataFileNamePath As String

'-------SAVANT export folder creation-----------

Dim cst_SAVANT_extension As String, cst_folder As String
Dim iversion As Integer
Dim cst_filename As String, cst_projectpath As String
Dim cst_project_name As String
Dim sfilename As String
Dim cst_tree() As String, cst_ff_monitor_names() As String, cst_ff_monitor_index() As Long
Dim cst_n_of_excitations As Long, cst_temp_n_of_excitations As Long, cst_n_of_equi_ff_excitations As Long,cst_n_of_non_equi_ff_excitations As Long
Dim cst_nff_monitors As Long
Dim FSM_Exist As Boolean, dummy_CST_project As Boolean
Dim NFS_Excitation As Variant
Dim CST_NF_Folder As String

Dim CST_FN As Long, CST_FN2 As Long


Dim cst_theta_step As Double, cst_phi_step As Double
Dim cst_origin_type As String, cst_fforigin As String
Dim cst_origin_x As Double, cst_origin_y As Double, cst_origin_z As Double


Sub Main


Dim iii As Long, counter As Long
Dim Nmoni_Fieldsource As Long, N_FSM_files As Long
Dim Fieldsource_name()  As String, FSM_file_name() As String
Dim CAD_STL_filenames(2) As String, temp_sfilename As String

Dim cst_tmpstr As String

Dim cst_temp_project_name As String

Dim ExportNFS As Boolean, ExportCAD As Boolean, ExportWires As Boolean, ExportFarfield As Boolean

Dim Message_Counter As Integer
Dim Message_String () As String
Dim Message_Type () As Boolean     'True--> Info ; False--> Warning

Dim cst_iloop As Long, cst_iloop2 As Long, index As Long

Dim cst_nff As Long


Dim cst_excit_names() As String, cst_temp_excit_names() As String, cst_ff_equi_excit_names() As String, cst_ff_non_equi_excit_names () As String
Dim cst_excit_equi_ff () As Boolean   'True= equidistant ; False= Non Equidistant
Dim ff_index As Long, excit_counter As Long
Dim ExcitAlreadyExist As Boolean
Dim FFS_Single_Freq As Boolean

Dim cst_excit_loop_start As Long, cst_excit_loop_end As Long, cst_i_excit_loop As Long, cst_iloop_start As Long, cst_iloop_end As Long
Dim cst_floop_start As Long, cst_floop_end As Long, cst_ifloop As Long
Dim ff_file_name As String



Dim InfoMessage As String

'===============Save .cst file in case the project is Untitled============
Save

'===============Get cst project name and project path============

cst_projectpath=GetProjectPath("Project")+".cst"
index=InStrRev (GetProjectPath("Project"), "\")			'Returns the index of the last "\" in the datafilename path
cst_project_name= Mid$(GetProjectPath("Project"),index+1)		'Trim the project path starting from the given index+1



'Dialog Box properties inizialization
'==========================================================================
	cst_theta_step=5.0
	cst_phi_step=5.0

	cst_origin_x = 0.0
	cst_origin_y = 0.0
	cst_origin_z = 0.0

	cst_origin_type="0"

'==========================================================================
'=======check if any .FSM file is available in the result folder===========

InfoMessage=""



FSM_Exist=True

    Nmoni_Fieldsource=0

    For iii= 0 To Monitor.GetNumberOfMonitors-1
            If (Monitor.GetMonitorTypeFromIndex(iii) = "Fieldsource") Then
                    Nmoni_Fieldsource = Nmoni_Fieldsource + 1

            End If
    Next iii



If (Nmoni_Fieldsource=0) And FindFirstFile(GetProjectPath("Result"),"*.fsm*", False)="" Then  ' check if any .fsm file is present in the project result folder
        'MsgBox "The current project hasn't got any ""Field source"" Monitor defined."
        InfoMessage=InfoMessage +"NEARFIELD RESULTS:"+vbCrLf+"The current project hasn't got any ""Field source"" Monitor defined."+vbCrLf+"The CAD & Wires Export options will be disabled."+ vbCrLf + vbCrLf
        FSM_Exist=False


	ElseIf (Nmoni_Fieldsource=0) And FindFirstFile(GetProjectPath("Result"),"*.fsm*", False)<>"" Then
		InfoMessage=InfoMessage +"NEARFIELD RESULTS:"+vbCrLf+"The current project hasn't got any ""Field source"" Monitor defined but"+vbCrLf+"there are .FSM files in the project result folder."+ vbCrLf + vbCrLf

	ElseIf  (Nmoni_Fieldsource<>0) And FindFirstFile(GetProjectPath("Result"),"*.fsm*", False)="" Then
		'MsgBox "Launch the simulation to get the field source monitor results"
		InfoMessage=InfoMessage +"NEARFIELD RESULTS:"+vbCrLf+"Launch the simulation to get the field source monitor results."+ vbCrLf + vbCrLf
		FSM_Exist=False

End If

If FSM_Exist=True Then
	If Resulttree.GetFirstChildName("Wires")<>"" Then
	InfoMessage="WIRES EXPORT POTENTIAL ISSUES:"+vbCrLf+"1) Please ensure that all the wires belong to a folder before executing the export."+ vbCrLf +"E.g. (Wires/<FolderName>/<WireName>)."+ vbCrLf +"2) Uncheck the WIRES Export if you don't have the ""3D Import Token"" feature in your license."+ vbCrLf + vbCrLf
	End If
End If

cst_nff_monitors=0

'=======calculate number of farfield monitors in the tree===========

	cst_iloop=0

	For iii= 0 To Monitor.GetNumberOfMonitors-1
            If (Monitor.GetMonitorTypeFromIndex(iii) = "Farfield") Then
                    cst_nff_monitors = cst_nff_monitors + 1

                    ReDim Preserve cst_ff_monitor_names(cst_nff_monitors-1)
					cst_ff_monitor_names(cst_iloop)=Monitor.GetMonitorNameFromIndex(iii)      'record in the array the ff monitor names
					ReDim Preserve cst_ff_monitor_index((cst_nff_monitors-1))
					cst_ff_monitor_index(cst_iloop)=iii										 'record in the array the ff monitor index
					cst_iloop=cst_iloop+1

            End If
    Next iii


'=======calculate number of farfield results in the tree===========


	ReDim cst_tree(1)
	cst_tree(0) = Resulttree.GetFirstChildName ("Farfields")
	cst_iloop = 0
	If cst_tree(0)="Farfields\Farfield Cuts" Then
	   cst_tree(0) = Resulttree.GetNextItemName ("Farfields\Farfield Cuts")
	End If
	If cst_tree(0)="" Then
		GoTo ERROR_NO_FARFIELDS
	End If
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

'=============================================================================
'=======calculate number of different excitations=====================


		ReDim cst_temp_excit_names(cst_nff)

			For cst_iloop2 = 0 To cst_nff-1    ' loop over farfield results to get the excitation strings

				index=InStr(cst_tree(cst_iloop2), "[")
				cst_temp_excit_names(cst_iloop2)= Mid$(cst_tree(cst_iloop2),index)

			Next cst_iloop2



		ReDim cst_excit_names(1)

		cst_excit_names(0)=cst_temp_excit_names(0)			'assign automatically the first excitation
		cst_n_of_excitations=1

		For cst_iloop2 = 0 To cst_nff-1

			ExcitAlreadyExist=False

				For cst_iloop=0 To cst_n_of_excitations-1    ' loop over current array of excitations

					If cst_temp_excit_names(cst_iloop2)=cst_excit_names(cst_iloop) Then
						ExcitAlreadyExist=True
					End If
				Next cst_iloop

					If ExcitAlreadyExist=False Then
						cst_n_of_excitations=cst_n_of_excitations+1
						ReDim Preserve cst_excit_names(cst_n_of_excitations)
						cst_excit_names(cst_n_of_excitations-1)=cst_temp_excit_names(cst_iloop2)
					End If

		Next cst_iloop2

	ReDim cst_excit_equi_ff(cst_n_of_excitations-1)

	ReDim nff_per_excit(cst_n_of_excitations-1)    	 	 'Writes in this array how many farfield results are available in the Tree for each excitation

	ReDim index_ff_excit_list(cst_nff-1)				 'Writes in this array the farfield result indexes for each excitation  e.g. [1st excit. 0,3,7...2nd excit 1,4,8, ecc]

'======= 1)load farfield result indexes for each excitation==============================
'======= 2)check if excitation contains farfield with equidistant frequency sampling=====

Dim dfrq_1st As Double, dfrq_2nd As Double, sname_without_frq_orig As String
Dim dfrq_ff_nth As Double, dfrq_ff_n_plus_1_th As Double, dfrq_ff_sampling As Double, dfrq_ff_difference As Double
Dim Is_FF_Equidistant As Boolean


	ff_index=0

	cst_n_of_equi_ff_excitations=0
	cst_n_of_non_equi_ff_excitations=0

For cst_iloop = 0 To cst_n_of_excitations-1
		excit_counter=0

'Variables inizialization
Is_FF_Equidistant=True
dfrq_1st=0.0
dfrq_2nd=0.0
dfrq_ff_nth=0.0
dfrq_ff_n_plus_1_th=0.0


		For cst_iloop2=0 To cst_nff-1

			If cst_temp_excit_names(cst_iloop2)=cst_excit_names(cst_iloop) Then
				index_ff_excit_list(ff_index)=cst_iloop2
				ff_index=ff_index+1
				excit_counter=excit_counter+1


'================================================================================================
'check equidistant farfield sampling for the current excitation
'================================================================================================

				If excit_counter=1 Then

					If Not bSeparateFarfieldFrq(cst_tree(cst_iloop2), dfrq_1st, sname_without_frq_orig) Then   'Farfield Result Name does not contain frequency term: (f=...)

							NoAutomaticLabelingFarfFreq(cst_tree(cst_iloop2), dfrq_1st)
					End If

					dfrq_1st=DRound(dfrq_1st,4)

				ElseIf excit_counter=2 Then

					If Not bSeparateFarfieldFrq(cst_tree(cst_iloop2), dfrq_2nd, sname_without_frq_orig) Then   'Farfield Result Name does not contain frequency term: (f=...)

							NoAutomaticLabelingFarfFreq(cst_tree(cst_iloop2), dfrq_2nd)
					End If

				dfrq_ff_sampling = DRound(dfrq_2nd-dfrq_1st,4) 		'calculate farfield sampling between the first two frequencies
				dfrq_ff_nth=DRound(dfrq_2nd,4)

				ElseIf excit_counter>2 And Is_FF_Equidistant Then

					If Not bSeparateFarfieldFrq(cst_tree(cst_iloop2), dfrq_ff_n_plus_1_th, sname_without_frq_orig) Then   'Farfield Result Name does not contain frequency term: (f=...)

						NoAutomaticLabelingFarfFreq(cst_tree(cst_iloop2), dfrq_ff_n_plus_1_th)
					End If

				dfrq_ff_n_plus_1_th=DRound(dfrq_ff_n_plus_1_th,4)
				dfrq_ff_nth=DRound(dfrq_ff_nth,4)

				dfrq_ff_difference=DRound (dfrq_ff_n_plus_1_th-dfrq_ff_nth,4)

				If 	dfrq_ff_difference=dfrq_ff_sampling Then

					dfrq_ff_nth=dfrq_ff_n_plus_1_th

					Else
						Is_FF_Equidistant=False
				End If

				End If

'================================================================================================
'================================================================================================
			End If

		Next cst_iloop2

		If cst_iloop=0 Then
		nff_per_excit(cst_iloop)=excit_counter  							 'assign n° of farfield results for the 1st excitation
		Else
		nff_per_excit(cst_iloop)=excit_counter+nff_per_excit(cst_iloop-1)    'assign for the following excitation the sum with previous excitation  [e.g. 3,6,9,10,11] 2nd excitation starts at index given by 1st excitation and ends at index given by actual excitation
		End If

	If Is_FF_Equidistant Then

		'excitation with equidistant farfield sampling
		cst_n_of_equi_ff_excitations = cst_n_of_equi_ff_excitations+1
		ReDim Preserve cst_ff_equi_excit_names(cst_n_of_equi_ff_excitations)
		cst_ff_equi_excit_names(cst_n_of_equi_ff_excitations-1)=cst_excit_names(cst_iloop)

		cst_excit_equi_ff(cst_iloop)=True
	Else
		'excitation with non equidistant farfield sampling
		cst_n_of_non_equi_ff_excitations = cst_n_of_non_equi_ff_excitations+1
		ReDim Preserve cst_ff_non_equi_excit_names(cst_n_of_non_equi_ff_excitations)
		cst_ff_non_equi_excit_names(cst_n_of_non_equi_ff_excitations-1)=cst_excit_names(cst_iloop)

		cst_excit_equi_ff(cst_iloop)=False

	End If

	Next cst_iloop

	'Write all the excitations with non equidistant farfield sampling in a single string
	'==========================================================
	Dim cst_ff_non_equi_excit_names_String As String

	cst_ff_non_equi_excit_names_String=""

	For iii=0 To cst_n_of_non_equi_ff_excitations-1
		If iii=0 Then
			cst_ff_non_equi_excit_names_String=cst_ff_non_equi_excit_names_String+cst_ff_non_equi_excit_names(iii)
		ElseIf iii>0 And iii<cst_n_of_non_equi_ff_excitations-1 Then
			cst_ff_non_equi_excit_names_String=cst_ff_non_equi_excit_names_String+", "+cst_ff_non_equi_excit_names(iii)
		Else
			cst_ff_non_equi_excit_names_String=cst_ff_non_equi_excit_names_String+" and "+cst_ff_non_equi_excit_names(iii)
		End If
	Next iii

	'==========================================================

	If cst_n_of_non_equi_ff_excitations=1 Then
		InfoMessage=InfoMessage + "FARFIELD RESULTS:"+vbCrLf+"The excitation """+cst_ff_non_equi_excit_names_String+""" has got non equidistant farfield sampling."+vbCrLf+"This excitation will not be considered for the broadband export."
	ElseIf cst_n_of_non_equi_ff_excitations>1 Then
		InfoMessage=InfoMessage + "FARFIELD RESULTS:"+vbCrLf+"The excitations """+cst_ff_non_equi_excit_names_String+""" have got non equidistant farfield sampling."+vbCrLf+"These excitations will not be considered for the broadband export."
	End If

	If cst_n_of_non_equi_ff_excitations=cst_n_of_excitations Then
		InfoMessage=InfoMessage+vbCrLf+"Please select ""Single Frequency"" if you want to perform the farfield export."
	End If
	'=============================================================================
	If cst_tree(0)="" Then
	ERROR_NO_FARFIELDS:
			InfoMessage=InfoMessage + "FARFIELD RESULTS:"+vbCrLf+"No farfield results available!!"
			'MsgBox("No farfield results available!!")
	End If


	If InfoMessage<>"" Then
		MsgBox(InfoMessage,"INFO")
	End If

	ERROR_NO_OPTIONS:
	ERROR_NO_BROADBAND:

	Begin Dialog UserDialog 460,469,"Export antenna data for SAVANT",.DialogFunc ' %GRID:10,7,1,1
		OKButton 50,441,90,21
		GroupBox 20,91,430,336,"Antenna Data Export",.GroupBox1
		GroupBox 30,133,390,115,"Nearfield Results (mandatory information)",.GroupBox3
		GroupBox 50,175,350,60,"3D Structure",.GroupBox4
		CheckBox 50,154,320,14,"Export all Field source monitors (.FSM) to NFS",.ExportFSMtoNFS
		GroupBox 30,259,390,161,"Farfield Results",.GroupBox5
		'GroupBox 430,217,370,112,"Farfield origin",.GroupBox6
		'GroupBox 430,154,370,56,"Angular resolution in degree",.GroupBox2
		'Text 450,185,90,14,"Theta Step",.Text1
		'Text 640,185,90,14,"Phi Step",.Text2
		'TextBox 500,301,50,21,.xcenter
		'TextBox 580,301,50,21,.ycenter
		'TextBox 670,301,50,21,.zcenter
		'TextBox 540,182,80,21,.th_step
		'TextBox 710,182,80,21,.ph_step
		'OptionGroup .GroupCenterBox
		'	OptionButton 460,238,290,14,"Center of bounding box",.OptionButton3
		'	OptionButton 460,259,290,14,"Origin of coordinate system",.OptionButton4
		'	OptionButton 460,280,290,14,"Free",.OptionButton5
		CheckBox 80,196,250,14," Export CAD Structure to STL [m] ",.ExportCADtoSTL
		CheckBox 80,217,220,14," Export Wires to IGES [m] ",.ExportWirestoIGES
		CancelButton 150,441,90,21
		PushButton 300,280,100,21,"Properties...",.Properties
		CheckBox 40,110,180,21," Export all selected data",.ExportAll
		'GroupBox 430,154,370,56,"Angular resolution in degree",.GroupBox2
		CheckBox 50,287,250,14," Export all farfields as source (.ffs)",.ExportFF
		GroupBox 40,315,370,98,"Type of export",.GroupBox7
		CheckBox 240,364,140,14," Export all",.FFall
		CheckBox 50,364,170,14," Export all",.ExcitAll
		DropListBox 240,385,160,121,cst_tree(),.FarfieldSelect
		DropListBox 50,385,160,121,cst_ff_equi_excit_names(),.ExcitationSelect
		'Text 450,185,90,14,"Theta Step",.Text1
		'Text 640,185,90,14,"Phi Step",.Text2
		OptionGroup .Export_Type
			OptionButton 240,336,160,14,"Single Frequency",.OptionButton1
			OptionButton 50,336,160,14,"Broadband",.OptionButton2
		GroupBox 20,7,430,77,"Specify export directory and folder name",.GroupBox2
		'OptionGroup .GroupCenterBox
		'	OptionButton 460,238,290,14,"Center of bounding box",.OptionButton3
		'	OptionButton 460,259,290,14,"Origin of coordinate system",.OptionButton4
		'	OptionButton 460,280,290,14,"Free",.OptionButton5
		'TextBox 500,301,50,21,.xcenter
		'TextBox 580,301,50,21,.ycenter
		'TextBox 670,301,50,21,.zcenter
		'TextBox 540,182,80,21,.th_step
		'TextBox 710,182,80,21,.ph_step
		Text 290,63,50,14,".dlxcst",.Text3
		TextBox 40,28,240,21,.dlg_savant_folder
		TextBox 40,56,240,21,.name
		PushButton 310,28,120,21,"Browse...",.Browse

	End Dialog
	Dim dlg As UserDialog

	dlg.name=cst_project_name
	dlg.dlg_savant_folder = GetProjectPath("Root")

	If (Dialog(dlg) = 0) Then Exit All

	cst_SAVANT_extension = ".dlxcst"

'	If dlg_savant_folder
	cst_folder = dlg.dlg_savant_folder+"\"+dlg.name+cst_SAVANT_extension


	CST_RmDir_NotEmpty (cst_folder)    'If folder already existing will be deleted without asking
	Wait 0.5
	CST_MkDir (cst_folder)				' creating folder for export

 'Store export options
'===============================================================================================
	If dlg.ExportAll=1 Then
		If FSM_Exist Then
			ExportNFS=True
			ExportCAD=True
			ExportWires=True
		Else
			ExportNFS=False
			ExportCAD=False
			ExportWires=False
		End If

		'ExportCAD=True
		'ExportWires=True

		If cst_tree(0)="" Then
			ExportFarfield=False
		Else
			ExportFarfield=dlg.ExportFF
			'ExportFarfield=True
		End If
	Else
		ExportNFS=dlg.ExportFSMtoNFS
		ExportCAD=dlg.ExportCADtoSTL
		ExportWires=dlg.ExportWirestoIGES
		ExportFarfield=dlg.ExportFF
	End If

	'===============check if no export option has been selected===================================

	If (ExportNFS=False And ExportCAD=False And ExportWires=False And ExportFarfield=False) Then
        MsgBox "Select at least one option in order to perform the export"
        GoTo ERROR_NO_OPTIONS
 	'Exit All
 	End If

	If dlg.Export_Type = 1 Then
		FFS_Single_Freq= False
	Else
		FFS_Single_Freq= True
	End If


	If FFS_Single_Freq=False And ExportFarfield  And cst_n_of_non_equi_ff_excitations=cst_n_of_excitations Then
		MsgBox "Please select ""Single frequency"" in order to perform the farfield export"
        GoTo ERROR_NO_BROADBAND
	End If

	'--- get registry settings
	'dlg.th_step  = GetString("CST STUDIO SUITE", "SavantExport", "theta_step", "5.0")
	'dlg.ph_step = GetString("CST STUDIO SUITE", "SavantExport", "phi_step", "5.0")


	'--- write back registry settings
	'SaveString  "CST STUDIO SUITE", "SavantExport", "theta_step", dlg.th_step
	'SaveString  "CST STUDIO SUITE", "SavantExport", "phi_step", dlg.ph_step



	'cst_theta_step=CDbl(dlg.th_step)
	'cst_phi_step=CDbl(dlg.ph_step)

	'cst_origin_type = CInt(dlg.GroupCenterBox)
	'cst_origin_x = RealVal(dlg.xcenter)
	'cst_origin_y = RealVal(dlg.ycenter)
	'cst_origin_z = RealVal(dlg.zcenter)
'================================================================================================


'===============Export FSM to NFS==========================================

'-------exectute the listed operations only if there is a field source monitor defined and the .fsm file has been written to the result folder (results are present) -----------

	If ExportNFS Then

	N_FSM_files=1
	ReDim FSM_file_name(N_FSM_files)
	FSM_file_name(N_FSM_files-1)=FindFirstFile(GetProjectPath("Result"),"*.fsm*", False)

	While (FSM_file_name(N_FSM_files-1) <> "")
		N_FSM_files=N_FSM_files+1
		ReDim Preserve FSM_file_name(N_FSM_files)
		FSM_file_name(N_FSM_files-1) = FindNextFile()
	Wend

	'Solver.CalculateZandYMatrices

	Message_Counter=0

	'Export NF_file_Names
	'================================================================================================
	CST_FN2 = FreeFile
	Open cst_folder+"\NFS_index.txt" For Output As #CST_FN2
	Print #CST_FN2, "// Number near field source"
	Print #CST_FN2, CStr(N_FSM_files-1)


	For iii=0 To N_FSM_files-2

		GlobalDataFileNamePath=GetProjectPath("Result")+FSM_file_name(iii)

		'--execute FSM to NFS conversion creating a subfolder in the .dlxcst folder
		ExportFSMtoNFS(GlobalDataFileNamePath)

	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="NFS-write: Field source monitor was successfuly exported to directory """+CST_NF_Folder+""""+"."
	Message_Type(Message_Counter-1)=True     'Info message
	'=========================================================================================

	If SelectTreeItem("1D Results\Z Matrix") And IsNumeric(NFS_Excitation)  Then     'Check if Z-Parameter folder is available and if the result comes from a standard single port excitation
		WriteComplexImpedanceNFSFrequencySamples (GlobalDataFileNamePath)
		ReportInformation ("Z-Matrix @ selected NFS frequencies has been successfully exported.")

	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="Z-Matrix @ selected NFS frequencies has been successfully exported."
	Message_Type(Message_Counter-1)=True     'Info message
	'=========================================================================================

	Else
	ReportWarning ("Z-Matrix is not available for the excitation:"""+CStr(NFS_Excitation)+""". Therefore the corresponding data have not been exported.")

	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="Z-Matrix is not available for the excitation:"""+CStr(NFS_Excitation)+""". Therefore the corresponding data have not been exported."
	Message_Type(Message_Counter-1)=False     'Warning message
	'=========================================================================================

	End If

		index=InStrRev(FSM_file_name(iii), ".")					'Returns the index of the last "." in the filename
		Print #CST_FN2, Left$(FSM_file_name(iii),index-1)		'Trim the project path starting from the given index+1
	Next iii

	Close #CST_FN2


	'Export boundary conditions info to .txt file
	'================================================================================================
	CST_FN = FreeFile
	Open cst_folder+"\NFS_Boundaries.txt" For Output As #CST_FN

	'check is symmetry plane is enabled --> if yes Xmax is taken
	If Boundary.GetXSymmetry="electric" Or Boundary.GetXSymmetry="magnetic" Then

	If Boundary.GetXmax="expanded open" Then
		Print #CST_FN, "Xmin=open"
	Else
		Print #CST_FN, "Xmin="+Boundary.GetXmax
	End If

	ElseIf Boundary.GetXmin="expanded open" Then
		Print #CST_FN, "Xmin=open"
	Else
		Print #CST_FN, "Xmin="+Boundary.GetXmin
	End If


	If Boundary.GetXmax="expanded open" Then       'Xmax
		Print #CST_FN, "Xmax=open"
	Else
		Print #CST_FN, "Xmax="+Boundary.GetXmax
	End If


	If Boundary.GetYSymmetry="electric" Or Boundary.GetYSymmetry="magnetic" Then
	If Boundary.GetYmax="expanded open" Then
		Print #CST_FN, "Ymin=open"
	Else
		Print #CST_FN, "Ymin="+Boundary.GetYmax
	End If

	ElseIf Boundary.GetYmin="expanded open" Then
		Print #CST_FN, "Ymin=open"
	Else
		Print #CST_FN, "Ymin="+Boundary.GetYmin
	End If


	If Boundary.GetYmax="expanded open" Then       'Ymax
		Print #CST_FN, "Ymax=open"
	Else
		Print #CST_FN, "Ymax="+Boundary.GetYmax
	End If

	If Boundary.GetZSymmetry="electric" Or Boundary.GetZSymmetry="magnetic" Then
	If Boundary.GetZmax="expanded open" Then
		Print #CST_FN, "Zmin=open"
	Else
		Print #CST_FN, "Zmin="+Boundary.GetZmax
	End If

	ElseIf Boundary.GetZmin="expanded open" Then
		Print #CST_FN, "Zmin=open"
	Else
		Print #CST_FN, "Zmin="+Boundary.GetZmin
	End If


	If Boundary.GetZmax="expanded open" Then		'Zmax
		Print #CST_FN, "Zmax=open"
	Else
		Print #CST_FN, "Zmax="+Boundary.GetZmax
	End If

	Close #CST_FN
	'================================================================================================


	End If

'================================================================================================

'===============Export Farfield results (if available)==========================================


Dim sResultFile As String

If ExportFarfield  Then

CST_FN2 = FreeFile
Open cst_folder+"\FFS_index.txt" For Output As #CST_FN2
Print #CST_FN2, "// Number far field source"

'---farfield plot init
	With FarfieldPlot
		.Reset
		Select Case cst_origin_type
			Case 0
				.Origin "bbox"
				cst_fforigin = "center of bounding box"
			Case 1
				.Origin "zero"
				cst_fforigin = "(x=0, y=0, z=0)"
			Case 2
				.Origin "free"
				'.Freeorigin cst_origin_x, cst_origin_y, cst_origin_z
				.Userorigin cst_origin_x, cst_origin_y, cst_origin_z
				cst_fforigin = "(x=" + Replace(Format(cst_origin_x,"Scientific"),",",".") + ", y=" + Replace(Format(cst_origin_y,"Scientific"),",",".") + ", z=" + Replace(Format(cst_origin_z,"Scientific"),",",".") + ")"
		End Select

	End With

'===================broadband farfield source export=======================================

	If FFS_Single_Freq=False Then   'broadband source

		If dlg.ExcitAll = 0 Then  'single excitation
			cst_excit_loop_start = CInt(dlg.ExcitationSelect)
			cst_excit_loop_end = CInt(dlg.ExcitationSelect)

		Print #CST_FN2, CStr(1)

		Else
			cst_excit_loop_start = 0 'all excitations
			cst_excit_loop_end = cst_n_of_excitations-1

		Print #CST_FN2, CStr(cst_n_of_equi_ff_excitations)
		End If

		For cst_i_excit_loop = cst_excit_loop_start To cst_excit_loop_end

				If cst_i_excit_loop=0 Then
					cst_iloop_start=0
					cst_iloop_end =nff_per_excit(cst_excit_loop_start)-1
				Else
					cst_iloop_start=nff_per_excit(cst_i_excit_loop-1)
					cst_iloop_end =nff_per_excit(cst_i_excit_loop)-1
				End If

	If cst_excit_equi_ff(cst_i_excit_loop) Then      'export is performed only if the excitation has equidistant farfield sampling

		Dim dfrq_orig As Double, dfrq_end As Double

		If Not bSeparateFarfieldFrq(cst_tree(index_ff_excit_list(cst_iloop_start)), dfrq_orig, sname_without_frq_orig) Then   'Farfield Result Name does not contain frequency term: (f=...)

				NoAutomaticLabelingFarfFreq(cst_tree(index_ff_excit_list(cst_iloop_start)), dfrq_orig)

		End If

		If Not bSeparateFarfieldFrq(cst_tree(index_ff_excit_list(cst_iloop_end)), dfrq_end, sname_without_frq_orig) Then		'Farfield Result Name does not contain frequency term: (f=...)

				NoAutomaticLabelingFarfFreq(cst_tree(index_ff_excit_list(cst_iloop_end)), dfrq_end)

		End If

				SelectTreeItem "Farfields\"+cst_tree(index_ff_excit_list(cst_iloop_start))
				FarfieldPlot.Reset
				FarfieldPlot.Plottype ("3d")
				FarfieldPlot.SetLockSteps(False)
				FarfieldPlot.Step(cst_theta_step)
				FarfieldPlot.Step2(cst_phi_step)
				FarfieldPlot.Plot

			sResultFile = cst_folder+"\farfield (f="+CStr(dfrq_orig)+".."+CStr(dfrq_end)+" ("+CStr(cst_iloop_end-cst_iloop_start+1)+")) " + cst_excit_names(cst_i_excit_loop)+".ffs"
			FarfieldPlot.ASCIIExportAsBroadbandSource sResultFile

			index=InStrRev(sResultFile, "\")
			ff_file_name=Mid$(sResultFile,index+1)
			index=InStrRev(ff_file_name, ".")															'Returns the index of the last "." in the filename
			Print #CST_FN2, Left$(ff_file_name,index-1)													'write farfield name without file extension

		End If

		Next cst_i_excit_loop

'===================single freqquency farfield source export=======================================

	Else 'single frequency
		If dlg.FFall = 0 Then  '--- All Farfields or just one
			cst_floop_start = CInt(dlg.FarfieldSelect)
			cst_floop_end = CInt(dlg.FarfieldSelect)

			Print #CST_FN2, CStr(1)       'Write number of farfield
		Else
			cst_floop_start = 0
			cst_floop_end = cst_nff - 1

			Print #CST_FN2, CStr(cst_nff)       'Write number of farfield
		End If

		For cst_ifloop = cst_floop_start To cst_floop_end


			SelectTreeItem "Farfields\"+cst_tree(cst_ifloop)
				FarfieldPlot.Reset
				FarfieldPlot.Plottype ("3d")
				FarfieldPlot.SetLockSteps(False)
				FarfieldPlot.Step(cst_theta_step)
				FarfieldPlot.Step2(cst_phi_step)
				FarfieldPlot.Plot

			sResultFile = cst_folder+"\"+cst_tree(cst_ifloop)+ ".ffs"
			FarfieldPlot.ASCIIExportAsSource sResultFile
			Print #CST_FN2, cst_tree(cst_ifloop)

		Next cst_ifloop

	End If

	Close #CST_FN2

	ReportInformation ("Selected farfields have been successfully exported as sources.")

	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="Selected farfields have been successfully exported as sources."
	Message_Type(Message_Counter-1)=True     'Info message
	'=========================================================================================
End If



'-------------------------------------------------------------------


'----	If option export enabled execute it------------------------------
'-------Check model unit. If not meter automatically rescales the model-----------


dummy_CST_project=False   'if new dummy .cst model used for wires scaling will be openend switch to True

'------------------export SOLIDS to STL---------------------------------

	If ExportCAD Then

	ExportCADtoSTL ()

	ReportInformation ("CAD structure has been successfully exported to STL.")
	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export Case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="CAD structure has been successfully exported to STL."
	Message_Type(Message_Counter-1)=True     'Info message
	'=========================================================================================
	End If




	If ExportWires And Resulttree.GetFirstChildName("Wires")<>""  Then

	'------------------rescales Wires to m --------------------

	Dim OldUnits As String, NewUnits As String
	Dim ComponentToScale As String, ScaleFactor As Double

	Save    'Save current project

	'--------------------Save the file with a different name to a temporary cst project in order to scale the model to m-----------------------------

	SaveAs(GetProjectPath("Project")+"_dummy"+".cst",False)

	dummy_CST_project=True

	WCS.ActivateWCS "global"  'disable local WCS if enabled

	cst_temp_project_name= GetProjectPath("Project")+".cst"

	ScaleFactor = Units.GetGeometryUnittoSI

	OldUnits = Units.GetUnit("Length")

	NewUnits="m"
	Units.SetUnit("Length",NewUnits)

	If OldUnits<>"m" Then     'perform wire scaling

	BeginHide

		' Scale wires
		ComponentToScale = Resulttree.GetFirstChildName("Wires")
		If (ComponentToScale <> "") Then
			AddToHistory("transform: scale "+Split(ComponentToScale,"\")(1), _
			"With Transform" + vbNewLine _
		     +".Reset" + vbNewLine _
		     +".Name "+Quote(Split(ComponentToScale,"\")(1)) + vbNewLine _
		     +".Origin "+Quote("Free") + vbNewLine _
		     +".Center "+Quote("0")+", "+Quote("0")+", "+Quote("0") + vbNewLine _
		     +".ScaleFactor "+Quote(CStr(ScaleFactor))+", "+Quote(CStr(ScaleFactor))+", "+Quote(CStr(ScaleFactor)) + vbNewLine _
		     +".Repetitions "+Quote("1") + vbNewLine _
		     +".Transform "+Quote("Wire")+", "+Quote("Scale") + vbNewLine _
			+"End With" + vbNewLine)
			' Scale all following components
			ComponentToScale = Resulttree.GetNextItemName(ComponentToScale)
			While(ComponentToScale <> "")
				AddToHistory("transform: scale "+Split(ComponentToScale,"\")(1), _
				"With Transform" + vbNewLine _
			     +".Reset" + vbNewLine _
			     +".Name "+Quote(Split(ComponentToScale,"\")(1)) + vbNewLine _
			     +".Origin "+Quote("Free") + vbNewLine _
			     +".Center "+Quote("0")+", "+Quote("0")+", "+Quote("0") + vbNewLine _
			     +".ScaleFactor "+Quote(CStr(ScaleFactor))+", "+Quote(CStr(ScaleFactor))+", "+Quote(CStr(ScaleFactor)) + vbNewLine _
			     +".Repetitions "+Quote("1") + vbNewLine _
			     +".Transform "+Quote("Wire")+", "+Quote("Scale") + vbNewLine _
				+"End With" + vbNewLine)
				ComponentToScale = Resulttree.GetNextItemName(ComponentToScale)
			Wend
		End If

	EndHide

	End If

	'------------------export WIRES to IGES--------------------


	ExportWIREStoIGES ()
	ReportInformation ("Wires have been successfully exported to IGES.")
	'========================================================================================
	' Store messages in order to display them again after reopening the file (Wire Export case)
	'========================================================================================
	Message_Counter=Message_Counter+1
	ReDim Preserve Message_String(Message_Counter)
	ReDim Preserve Message_Type(Message_Counter)

	Message_String(Message_Counter-1)="Wires have been successfully exported to IGES."
	Message_Type(Message_Counter-1)=True     'Info message
	'=========================================================================================

	Save
	OpenFile (cst_projectpath)   'Open original cst file


	End If


'================plot messages on window only if the dummy project needed for wires scaling has been created=======================
'================delete dummy cst project==========================================================================================

	If dummy_CST_project=True Then


	'--------------Delete temporary project directory + .cst file used to scale the structure
	CST_RmDir_NotEmpty (GetProjectPath("Project")+"_dummy")
	Kill(cst_temp_project_name)
	'----------------------------------------------------------------------------

	'----------Rename IGES files if WIRES are present in the project----------

	sfilename = FindFirstFile(cst_folder,"*.igs*", False)
    Name cst_folder+"\"+sfilename As cst_folder+"\"+cst_project_name+"_wires.igs"

    sfilename = FindFirstFile(cst_folder,"*.hlg*", False)
    Name cst_folder+"\"+sfilename As cst_folder+"\"+cst_project_name+"_wires.hlg"

	For iii=0 To Message_Counter-1
		If Message_Type(iii) Then
			ReportInformation(Message_String(iii))
		Else
			ReportWarning(Message_String(iii))
		End If
	Next iii


	End If



Save

	Begin Dialog UserDialog 720,126,"SAVANT folder" ' %GRID:10,7,1,1
		GroupBox 10,14,700,77,"The selected data have been exported to the following folder",.GroupBox1
		Text 30,70,300,14,"Click OK to open it.",.Text2
		OKButton 80,98,90,21
		CancelButton 200,98,90,21
		TextBox 20,42,670,21,.SAVANTFolderpath
	End Dialog
	Dim dlg2 As UserDialog

	dlg2.SAVANTFolderpath=cst_folder

	'Dialog dlg2
	If (Dialog(dlg2) = 0) Then Exit All


Shell "explorer " & Chr$(34)+cst_folder+Chr$(34), 1


End Sub

Function dialogfunc(DlgItem$, Action%, SuppValue%) As Boolean


	Dim filename As String, index As Integer

    Select Case Action%
	    Case 1 ' Dialog box initialization
				DlgValue "ExportAll",1

			If FSM_Exist Then
				DlgValue "ExportFSMtoNFS",1
				DlgValue "ExportCADtoSTL",1
				DlgValue "ExportWIREStoIGES",1
			Else
				DlgValue "ExportFSMtoNFS",0
				DlgValue "ExportCADtoSTL",0
				DlgValue "ExportWIREStoIGES",0
			End If

				'DlgValue "ExportCADtoSTL",1
				'DlgValue "ExportWIREStoIGES",1
				DlgEnable "ExportFSMtoNFS",False
				DlgEnable "ExportCADtoSTL",False
				DlgEnable "ExportWIREStoIGES",False
				DlgEnable "ExportFF", False


				DlgValue "ExportFF",0
				DlgEnable "Properties",False
				DlgEnable "Export_Type", False

				DlgValue  "Export_Type", 1
				DlgValue  "FFAll",1
				DlgValue  "ExcitAll",1
				DlgEnable "FFAll",False
				DlgEnable "ExcitAll",False
				DlgEnable "FarfieldSelect",False
				DlgEnable "ExcitationSelect",False



	    Case 2 ' Value changing or button pressed
	    	Select Case DlgItem$
	    		Case "Browse"
	    								filename = DlgText("dlg_savant_folder") + "\" + "Use this directory"
                                        filename = GetFilePath(filename, "", "", "Choose Root-directory", 2)
                                        If (filename <> "") Then
                                                DlgText "dlg_savant_folder", DirName(filename)
												iversion  = 1
												' double dirname, because of "Use this directory"
												sfilename = FindFirstFile(DirName(DirName(filename)), ShortName(DirName(filename))+ "_compare_##", False)
												While (sfilename <> "")
												        iversion  = CInt(Right$(ShortName(sfilename), 2)) + 1
												        sfilename = FindNextFile
												Wend
                                        End If
                                        dialogfunc = True
	    		Case "Properties"
				dialogfunc = True       ' Don't close the dialog box.
				PushProperties()


	    		Case "ExportAll"
		    		If SuppValue = 1 Then
						DlgEnable "ExportFSMtoNFS",False
						DlgEnable "ExportCADtoSTL",False
						DlgEnable "ExportWIREStoIGES",False
						DlgEnable "ExportFF",False
						DlgEnable "Export_Type", False
						'DlgValue "ExportFF",1
						'DlgEnable "th_step",True
						'DlgEnable "ph_step",True
						dialogfunc = True
		    		Else
						DlgEnable "ExportFSMtoNFS",True
						DlgEnable "ExportCADtoSTL",True
						DlgEnable "ExportWIREStoIGES",True
						DlgEnable "ExportFF",True
						DlgEnable "Export_Type", True
						dialogfunc = True
					End If
				Case "ExportFSMtoNFS"
					If SuppValue = 1 Then
						DlgEnable "ExportWIREStoIGES",True
						DlgEnable "ExportCADtoSTL",True
						'DlgValue "ExportCADtoSTL",1
						'DlgValue "ExportWIREStoIGES",1
					Else
						DlgEnable "ExportCADtoSTL",False
						DlgEnable "ExportWIREStoIGES",False
						DlgValue "ExportCADtoSTL",0
						DlgValue "ExportWIREStoIGES",0
					End If
				Case"ExportFF"
					If SuppValue = 1 Then
						'DlgEnable "th_step",True
						'DlgEnable "ph_step",True
						DlgEnable "Export_Type", True
						'DlgEnable "GroupCenterBox", True
						'DlgEnable "FFall",True
						'DlgEnable "Excitall",True
						'DlgEnable "FarfieldSelect",True
						'DlgEnable "ExcitationSelect",True
						dialogfunc = True
					Else
						'DlgEnable "th_step",False
						'DlgEnable "ph_step",False
						DlgEnable "Export_Type", False
						'DlgEnable "GroupCenterBox", False
						'DlgEnable "xcenter",False
						'DlgEnable "ycenter",False
						'DlgEnable "zcenter", False
						'DlgEnable "FFall",False
						'DlgEnable "Excitall",False
						'DlgEnable "FarfieldSelect",False
						'DlgEnable "ExcitationSelect",False
						dialogfunc = True
					End If
				Case "Export_Type"
					If SuppValue = 0 Then
					DlgEnable "FFAll",True
					DlgEnable "ExcitAll",False
					'DlgValue  "FFall",1
					'DlgEnable "ExcitationSelect", False
					dialogfunc = True
					Else
					DlgEnable "FFAll",False
					DlgEnable "ExcitAll",True
					'DlgEnable "FarfieldSelect",False
					'DlgValue  "Excitall",1
					'DlgEnable "ExcitationSelect", False
					dialogfunc = True
					End If
				Case "ExcitAll"
					If SuppValue = 1 Then
						DlgValue  "ExcitAll",1
						DlgEnable "ExcitationSelect",False
						dialogfunc = True
					Else
						DlgValue  "ExcitAll",0
						DlgEnable "ExcitationSelect",True
						dialogfunc = True
					End If
				Case "FFall"
					If SuppValue = 1 Then
						DlgValue  "FFAll",1
						DlgEnable "FarfieldSelect",False
						dialogfunc = True
					Else
						DlgValue  "FFAll",0
						DlgEnable "FarfieldSelect",True
						dialogfunc = True
					End If
				Case "th_step"
                	DlgValue "th_step",SuppValue
                	dialogfunc = True
                Case "ph_step"
                      DlgValue "ph_step",SuppValue
                      dialogfunc = True
	    		End Select
		Case 3 ' ComboBox or TextBox Value changed
	    Case 4 ' Focus changed
	    Case 5 ' Idle
	    Case 6 ' Function key
    End Select

	If (Action%=1) Or (Action%=2) Then

	If FSM_Exist=False Then
				DlgEnable "ExportFSMtoNFS",False
				DlgEnable "ExportCADtoSTL",False
				DlgEnable "ExportWIREStoIGES",False

	ElseIf (DlgValue("ExportAll")=0) And (DlgValue("ExportFSMtoNFS")=0) Then
				DlgEnable "ExportCADtoSTL",False
				DlgEnable "ExportWIREStoIGES",False
	End If


	If cst_tree(0)="" Then   ' no FF results available
			DlgEnable "ExportFF",False
			DlgEnable "Export_Type", False
			DlgEnable "FFAll",False
			DlgEnable "ExcitAll",False
			DlgEnable "FarfieldSelect",False
			DlgEnable "ExcitationSelect",False

	Else


	'("ExportAll")=0)

		If (DlgValue("ExportAll")=0) And (DlgValue("ExportFF")=0) Then
			DlgEnable "Export_Type", False
			'DlgEnable("th_step", False)
			'DlgEnable("ph_step", False)
			DlgEnable "FFAll",False
			DlgEnable "ExcitAll",False
			DlgEnable "FarfieldSelect",False
			DlgEnable "ExcitationSelect",False
			DlgEnable "Properties", False
			'DlgEnable("GroupCenterBox", False)
			'DlgEnable "xcenter",False
			'DlgEnable "ycenter",False
			'DlgEnable "zcenter", False
		ElseIf (DlgValue("ExportAll")=0) And(DlgValue("ExportFF")=1) Then
			DlgEnable "Export_Type", True
			DlgEnable "Properties", True
			'If (DlgValue("GroupCenterBox")=2) Then   'Origin free
			'	DlgEnable "xcenter",True
			'	DlgEnable "ycenter",True
			'	DlgEnable "zcenter", True
			'End If

			If (DlgValue("Export_Type")=1) Then
				DlgEnable "ExcitAll",True
				DlgEnable "FarfieldSelect",False
				 If (DlgValue("ExcitAll")=1) Then
					DlgEnable "ExcitationSelect",False
				 Else
					DlgEnable "ExcitationSelect",True
				 End If
			ElseIf (DlgValue("Export_Type")=0) Then
				DlgEnable "FFAll",True
				DlgEnable "ExcitationSelect",False
				 If (DlgValue("FFAll")=1) Then
					DlgEnable "FarfieldSelect",False
				 Else
					DlgEnable "FarfieldSelect",True
				 End If
			End If
			'("ExportAll")=1)
		ElseIf (DlgValue("ExportAll")=1) And (DlgValue("ExportFF")=1) Then
				DlgEnable("Export_Type", True)
			If (DlgValue("Export_Type")=1) Then
				DlgEnable "ExcitAll",True
				DlgEnable "FarfieldSelect",False
				 If (DlgValue("ExcitAll")=1) Then
					DlgEnable "ExcitationSelect",False
				 Else
					DlgEnable "ExcitationSelect",True
				 End If
			ElseIf (DlgValue("Export_Type")=0) Then
				DlgEnable "FFAll",True
				DlgEnable "ExcitationSelect",False
				 If (DlgValue("FFAll")=1) Then
					DlgEnable "FarfieldSelect",False
				 Else
					DlgEnable "FarfieldSelect",True
				 End If
			End If
		ElseIf (DlgValue("ExportAll")=1) And (DlgValue("ExportFF")=0) Then
				DlgEnable"Export_Type", False
				DlgEnable "ExcitAll", False

			If (DlgValue("Export_Type")=1) Then
				'DlgEnable "ExcitAll",True
				DlgEnable "FarfieldSelect",False
				 If (DlgValue("ExcitAll")=1) Then
					DlgEnable "ExcitationSelect",False
				 Else
					DlgEnable "ExcitationSelect",True
				 End If
			ElseIf (DlgValue("Export_Type")=0) Then
				DlgEnable "FFAll",True
				DlgEnable "ExcitationSelect",False
				 If (DlgValue("FFAll")=1) Then
					DlgEnable "FarfieldSelect",False
				 Else
					DlgEnable "FarfieldSelect",True
				 End If
			End If

		End If
		End If
	End If

End Function

Private Function DialogFunctionProperties(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------
Select Case Action%
	    Case 1 ' Dialog box initialization
			'DlgText "th_step","5.0"
			'DlgText "ph_step","5.0"

			DlgEnable "xcenter",False
			DlgEnable "ycenter",False
			DlgEnable "zcenter",False
			'DlgText  "xcenter","0"
			'DlgText  "ycenter","0"
			'DlgText  "zcenter","0"
		If cst_origin_type="2" Then
			DlgEnable "xcenter",True
			DlgEnable "ycenter",True
			DlgEnable "zcenter",True
		End If

		Case 2 ' Value changing or button pressed
			Select Case DlgItem$
			Case "GroupCenterBox"
		    		If SuppValue = 2 Then
						DlgEnable "xcenter",True
						DlgEnable "ycenter",True
						DlgEnable "zcenter",True
						DialogFunctionProperties = True
		    		Else
						DlgEnable "xcenter",False
						DlgEnable "ycenter",False
						DlgEnable "zcenter",False
						DialogFunctionProperties = True
					End If
		End Select
		Case 3 ' ComboBox or TextBox Value changed
	    Case 4 ' Focus changed
	    Case 5 ' Idle
	    Case 6 ' Function key

End Select

'If (Action%=1) Or (Action%=2)  Or (Action%=3)  Or (Action%=4) Then

'End If

End Function

Sub PushProperties()

	Begin Dialog UserDialog 400,225,"Properties",.DialogFunctionProperties ' %GRID:10,5,1,1
		OKButton 20,200,90,20
		CancelButton 120,200,90,20
		GroupBox 20,80,370,110,"Farfield origin",.GroupBox6
		GroupBox 20,15,370,55,"Angular resolution in degree",.GroupBox2
		Text 40,40,80,15,"Theta Step",.Text1
		Text 230,40,80,15,"Phi Step",.Text2
		TextBox 90,160,50,20,.xcenter
		TextBox 160,160,50,20,.ycenter
		TextBox 230,160,50,20,.zcenter
		TextBox 125,35,80,20,.th_step
		TextBox 300,35,80,20,.ph_step
		OptionGroup .GroupCenterBox
			OptionButton 40,100,290,15,"Center of bounding box",.OptionButton3
			OptionButton 40,120,290,15,"Origin of coordinate system",.OptionButton4
			OptionButton 40,140,290,15,"Free",.OptionButton5

	End Dialog
	Dim dlg As UserDialog

	dlg.th_step=CStr(cst_theta_step)
	dlg.ph_step=CStr(cst_phi_step)

	dlg.GroupCenterBox=CInt(cst_origin_type)

	dlg.xcenter=CStr(cst_origin_x)
	dlg.ycenter=CStr(cst_origin_y)
	dlg.zcenter=CStr(cst_origin_z)

	If (Dialog(dlg) <> 0) Then
	cst_theta_step=CDbl(dlg.th_step)
	cst_phi_step=CDbl(dlg.ph_step)

	cst_origin_type = CInt(dlg.GroupCenterBox)
	cst_origin_x = RealVal(dlg.xcenter)
	cst_origin_y = RealVal(dlg.ycenter)
	cst_origin_z = RealVal(dlg.zcenter)

	End If

End Sub

Sub ExportFSMtoNFS(DataFileName As String)

Dim NFS_Name As String
Dim index As Integer



index=InStrRev (DataFileName, "\")			'Returns the index of the last "\" in the datafilename path
NFS_Name= Mid$(DataFileName,index+1)		'Trim the project path starting from the given index+1
index=InStrRev (NFS_Name, ".")			'Returns the index of the last "." in the NFS name
NFS_Name=Left$(NFS_Name,index-1)

index=InStrRev (NFS_Name,"_")
NFS_Excitation=Mid$(NFS_Name, index+1)

CST_NF_Folder = cst_folder+"\"+NFS_Name


With NFSFile
    .Reset
    .Write(DataFileName, CST_NF_Folder)
End With

End Sub




Sub WriteComplexImpedanceNFSFrequencySamples (DataFileName As String)

Dim NFS_Name As String, NFS_Excitation As String
Dim index As Integer, iii As Integer, Freq_samples As Integer
Dim CST_NF_Folder As String, LineString As String, Temp_String As String, Temp_String_Frequency_List As String
Dim Frequency_List() As String


index=InStrRev (DataFileName, "\")			'Returns the index of the last "\" in the datafilename path
NFS_Name= Mid$(DataFileName,index+1)		'Trim the project path starting from the given index+1
index=InStrRev (NFS_Name, ".")			'Returns the index of the last "." in the NFS name
NFS_Name=Left$(NFS_Name,index-1)

index=InStrRev (NFS_Name,"_")
NFS_Excitation=Mid$(NFS_Name, index+1)

CST_NF_Folder = cst_folder+"\"+NFS_Name

Open CST_NF_Folder+"\Ex_ymax.xml" For Input As #1

	While Not EOF(1)
		Line Input #1, LineString
	Wend

Close #1

Temp_String=LineString

For iii=0 To 28
index=InStr(Temp_String, "<")			'Returns the index of the first occurence of "<" in the string
Temp_String= Mid$(Temp_String,index+1)		'Trim the project path starting from the given index+1
Next iii

index=InStr(Temp_String, ">")
Temp_String= Mid$(Temp_String,index+1)

index=InStr(Temp_String, "<")
Temp_String=Left$(Temp_String,index-2)


Temp_String_Frequency_List=Temp_String
index=InStr(Temp_String_Frequency_List, " ")

If index=0 Then

	Freq_samples=1 'only one frequency value is contained in the NFS data

Else
	Freq_samples=0

End If


	While index<>0
		index=InStr(Temp_String_Frequency_List, " ")
		Temp_String_Frequency_List= Mid$(Temp_String_Frequency_List,index+1)
		Freq_samples= Freq_samples+1
	Wend

If Freq_samples=1 Then

	ReDim Frequency_List(Freq_samples)
Else
	ReDim Frequency_List(Freq_samples-1)

End If

CSTSplit(Temp_String, Frequency_List)

Dim ZMatrix As Object, Re_Z As Object, Im_Z As Object
Dim temp_frequency As Double
Dim sTreeItem, sFile As String

Dim Zfilename As String

If InStr(NFS_Excitation,"+")= 0 Then     'don't write any file if the excitation string contains a + : measn sim. excitation/combine results

index=InStr(NFS_Excitation,"[")

If index<> 0 Then
		Zfilename=CST_NF_Folder+"_Z.txt"
		sTreeItem ="1D Results\Z Matrix\Z"+Left$(NFS_Excitation,index-1)+","+Left$(NFS_Excitation,index-1)
Else
		Zfilename=CST_NF_Folder+"_Z.txt"
		sTreeItem ="1D Results\Z Matrix\Z"+NFS_Excitation+","+NFS_Excitation

End If

sFile = Resulttree.GetFileFromTreeItem(sTreeItem)

Set ZMatrix= Result1DComplex(sFile)

Set Re_Z=ZMatrix.Real
Set Im_Z=ZMatrix.Imaginary

CST_FN = FreeFile
Open Zfilename For Output As #CST_FN

For iii=0 To Freq_samples-1

temp_frequency=CDBl(Frequency_List(iii))/Units.GetFrequencyUnitToSI
index=Re_Z.GetClosestIndexFromX(temp_frequency)

Print #CST_FN, CStr(temp_frequency)+"	"+CStr(Re_Z.GetY(index))+"	"+CStr(Im_Z.GetY(index))

Next iii

Close #CST_FN

Else
' return an info about Z complex data not exported

End If

End Sub

Sub ExportCADtoSTL ()
        '-------export STL file (combine each solid files)-----------

Dim CST_FN As Long, CST_FN2 As Long 'Metal export
Dim CST_FN3 As Long, CST_FN4 As Long 'Dielectric export
Dim cst_index As Long, cst_iii As Long, cst_count_metal As Long, cst_count_diel As Long, Component_cst As String, Shape_cst As String, Material_Type_cst As String

Dim STL_string0 As String, STL_string As String
Dim STL_string0_diel As String, STL_string_diel As String


cst_index=Solid.GetNumberOfShapes


cst_count_metal=0
cst_count_diel=0

For cst_iii=0 To cst_index-1

Component_cst=Left(Solid.GetNameOfShapeFromIndex(cst_iii),InStr(Solid.GetNameOfShapeFromIndex(cst_iii),":")-1)
Shape_cst=Replace(Solid.GetNameOfShapeFromIndex(cst_iii),Component_cst+":", "")

'check the material type of the selected shape

Material_Type_cst= Material.GetTypeOfMaterial(Solid.GetMaterialNameForShape(Component_cst+":"+Shape_cst))

If (Material_Type_cst = "PEC") Or (Material_Type_cst= "Lossy Metal") Then


With STL
    .Reset
    .FileName (cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_metal) +"_metal.stl")
    .name(Shape_cst)
    .Component (Component_cst)
    .ExportFileUnits("m")      'export automatically to m
    .Write
End With


  CST_FN = FreeFile
         cst_filename=cst_folder + Replace(GetProjectPath("Project"),GetProjectPath("Root"),"") + "_metal.stl"

         If cst_count_metal=0 Then

             STL_string0$=""
         Else
             Open cst_filename For Input As #CST_FN
             Input  #CST_FN, STL_string0$
             Close #CST_FN
         End If

         Open cst_filename For Output As #CST_FN

         CST_FN2 = FreeFile
         Open cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_metal) +"_metal.stl" For Input As #CST_FN2
         Input  #CST_FN2, STL_string$
         STL_string$ = Replace(STL_string$,cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_metal) +"_metal.stl",Solid.GetNameOfShapeFromIndex(cst_iii))
         Print  #CST_FN , STL_string0$+STL_string$
         Close #CST_FN2

         Kill cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_metal) +"_metal.stl"

	Close #CST_FN

	cst_count_metal=cst_count_metal+1

Else 'normal material (dieletric)

	With STL
    .Reset
    .FileName (cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_diel) +"_dielectric.stl")
    .name(Shape_cst)
    .Component (Component_cst)
    .ExportFileUnits("m")
    .Write
	End With


  	CST_FN3 = FreeFile
         cst_filename=cst_folder + Replace(GetProjectPath("Project"),GetProjectPath("Root"),"") + "_dielectric.stl"

         If cst_count_diel=0 Then

             STL_string0_diel$=""
         Else
             Open cst_filename For Input As #CST_FN3
             Input  #CST_FN3, STL_string0_diel$
             Close #CST_FN3
         End If

         Open cst_filename For Output As #CST_FN3

         CST_FN4 = FreeFile
         Open cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_diel) +"_dielectric.stl" For Input As #CST_FN4
         Input  #CST_FN4, STL_string_diel$
         STL_string_diel$ = Replace(STL_string_diel$,cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_diel) +"_dielectric.stl",Solid.GetNameOfShapeFromIndex(cst_iii))
         Print  #CST_FN3 , STL_string0_diel$+STL_string_diel$
         Close #CST_FN4

         Kill cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+ Trim(cst_count_diel) +"_dielectric.stl"

   Close #CST_FN3

   cst_count_diel=cst_count_diel+1

End If

Next cst_iii

'  CST_FN = FreeFile
'  cst_filename=cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"")+"_SOLIDS_STLgeometry unit.txt"
'  Open cst_filename For Output As #CST_FN
'  Print  #CST_FN , "Exported STL unit is " & Units.GetUnit("Length")
'  Close #CST_FN

'DlgText("OutputT", "Done!")
'DlgEnable("OK", True)

End Sub

Sub ExportWIREStoIGES ()


Dim Curve_cst As String, cst_wire_folder_name As String
Dim index As Long


		Curve_cst = Resulttree.GetFirstChildName("Wires")
		If (Curve_cst <> "") Then

			Wire.NewFolder "Dummy_Folder"

			If (Curve_cst <> "Wires\Dummy_Folder") Then

				index=InStrRev (Curve_cst, "\")
				cst_wire_folder_name=Mid$(Curve_cst, index+1)

				Curve_cst = Resulttree.GetNextItemName(Curve_cst)

				Wire.RenameFolder cst_wire_folder_name, "Dummy_Folder/"+cst_wire_folder_name

			Else

			Curve_cst = Resulttree.GetNextItemName(Curve_cst)

			End If


				While(Curve_cst <> "")

				If (Curve_cst <> "Wires\Dummy_Folder") Then

				index=InStrRev (Curve_cst, "\")
				cst_wire_folder_name=Mid$(Curve_cst, index+1)

				Curve_cst = Resulttree.GetNextItemName(Curve_cst)

				Wire.RenameFolder cst_wire_folder_name, "Dummy_Folder/"+cst_wire_folder_name

				Else

			Curve_cst = Resulttree.GetNextItemName(Curve_cst)

			End If

			Wend





		With IGES
				.Filename(cst_folder+Replace(GetProjectPath("Project"),GetProjectPath("Root"),"") +"_wires.igs")
				.Write("Dummy_Folder")
				End With

		End If



End Sub




Function bSeparateFarfieldFrq(ffname As String, dfrq As Double, sname_without_frq As String) As Boolean
	' This function checks if a meaningful expression "(f=x.xxx)" can be found in ffname, where x.xxx represents a frequency
	' Input: ffname
	' Output: dfrg: the frequency entry found, sname_without_frq: ffname with the frequency entry removed
	' Returns "True" if a frequency has been identified, "False" otherwise
	Dim i2, i3 As Integer

	bSeparateFarfieldFrq = False
	sname_without_frq = ""

	i2 = InStr(ffname,"(f=")
	If i2>0 Then
		' frequency start found
		i3 = InStr(Mid(ffname,i2+3,),")")

		If i3>1 Then

			If  IsNumeric(Mid(ffname,i2+3,i3-1)) Then
			' meaningful frequency end found

			bSeparateFarfieldFrq = True
			dfrq = Evaluate(Mid(ffname,i2+3,i3-1))
			sname_without_frq = Left(ffname,i2+2) + Mid(ffname,i2+i3+2)
		End If
		End If

	End If

End Function

Function NoAutomaticLabelingFarfFreq (ffname As String, dfreq As Double) As Boolean

'This function retrieves the farfield result frequency when the result of the function bSeparateFarfieldFrq is False.
'Generally it's executed when the farfield automatic labeling is not used.

Dim index As Long, iii As Long
Dim cst_ff_monitor_name_w_o_excit As String

index=InStr(ffname,"[")
cst_ff_monitor_name_w_o_excit= Left(ffname,index-2)

For iii=0 To cst_nff_monitors

	If StrComp(cst_ff_monitor_name_w_o_excit,cst_ff_monitor_names(iii),1)=0 Then   ' ff monitor string comparison 0-->check is fulfilled


	     Exit For  'Exit the For execution
	End If

	Next iii

	dfreq=Monitor.GetMonitorFrequencyFromIndex(cst_ff_monitor_index(iii))

	NoAutomaticLabelingFarfFreq=True


End Function
