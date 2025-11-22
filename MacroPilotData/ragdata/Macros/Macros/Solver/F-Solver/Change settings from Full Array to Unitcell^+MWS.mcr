' ================================================================================================
' Copyright 2020-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 06-Oct-2020 ube: First version
' ================================================================================================

Sub Main ()

Dim sCommand As String

sCommand = ""
AddToHistory "_____Start change Full Array to Unit Cell_____", sCommand

sCommand = ""
sCommand = sCommand + "ChangeSolverType ""HF Frequency Domain""" + vbLf
AddToHistory "change solver type", sCommand

sCommand = ""
sCommand = sCommand + "With Boundary" + vbLf
sCommand = sCommand + ".Xmin ""unit cell""" + vbLf
sCommand = sCommand + ".Xmax ""unit cell""" + vbLf
sCommand = sCommand + ".Ymin ""unit cell""" + vbLf
sCommand = sCommand + ".Ymax ""unit cell""" + vbLf
sCommand = sCommand + ".Zmin ""expanded open""" + vbLf
sCommand = sCommand + ".Zmax ""expanded open""" + vbLf
sCommand = sCommand + ".Xsymmetry ""none""" + vbLf
sCommand = sCommand + ".Ysymmetry ""none""" + vbLf
sCommand = sCommand + ".Zsymmetry ""none""" + vbLf
sCommand = sCommand + ".ApplyInAllDirections ""False""" + vbLf
sCommand = sCommand + ".OpenAddSpaceFactor ""0.5""" + vbLf
sCommand = sCommand + ".XPeriodicShift ""0.0""" + vbLf
sCommand = sCommand + ".YPeriodicShift ""0.0""" + vbLf
sCommand = sCommand + ".ZPeriodicShift ""0.0""" + vbLf
sCommand = sCommand + ".PeriodicUseConstantAngles ""False""" + vbLf
sCommand = sCommand + ".SetPeriodicBoundaryAngles ""PAA_UC_THETA"", ""PAA_UC_PHI""" + vbLf
sCommand = sCommand + ".SetPeriodicBoundaryAnglesDirection ""outward""" + vbLf
sCommand = sCommand + ".UnitCellFitToBoundingBox ""True""" + vbLf
sCommand = sCommand + ".UnitCellDs1 ""0.0""" + vbLf
sCommand = sCommand + ".UnitCellDs2 ""0.0""" + vbLf
sCommand = sCommand + ".UnitCellAngle ""90.0""" + vbLf
sCommand = sCommand + "End With"
AddToHistory "define boundaries", sCommand

sCommand = ""
sCommand = sCommand + "With FloquetPort" + vbLf
sCommand = sCommand + ".Reset" + vbLf
sCommand = sCommand + ".SetDialogTheta ""0""" + vbLf
sCommand = sCommand + ".SetDialogPhi ""0""" + vbLf 
sCommand = sCommand + ".SetPolarizationIndependentOfScanAnglePhi ""0.0"", ""False""" + vbLf  
sCommand = sCommand + ".SetSortCode ""+beta/pw""" + vbLf 
sCommand = sCommand + ".SetCustomizedListFlag ""False""" + vbLf 
sCommand = sCommand + ".Port ""Zmin""" + vbLf
sCommand = sCommand + ".SetNumberOfModesConsidered ""18""" + vbLf 
sCommand = sCommand + ".SetDistanceToReferencePlane ""0.0""" + vbLf 
sCommand = sCommand + ".SetUseCircularPolarization ""True""" + vbLf 
sCommand = sCommand + ".Port ""Zmax""" + vbLf 
sCommand = sCommand + ".SetNumberOfModesConsidered ""18""" + vbLf 
sCommand = sCommand + ".SetDistanceToReferencePlane ""0.0""" + vbLf 
sCommand = sCommand + ".SetUseCircularPolarization ""True""" + vbLf 
sCommand = sCommand + "End With"
AddToHistory "define Floquet port boundaries", sCommand

sCommand = ""
sCommand = sCommand + "With FDSolver" + vbLf
sCommand = sCommand + ".SetMethod ""Tetrahedral"", ""General purpose""" + vbLf
sCommand = sCommand + ".SetRecordUnitCellScanFarfield ""Auto""" + vbLf
sCommand = sCommand + ".SetDisableResultTemplatesDuringUnitCellScanAngleSweep ""True""" + vbLf 
sCommand = sCommand + "End With"
AddToHistory "AEP", sCommand

sCommand = ""
AddToHistory "_____End change Full Array to Unit Cell_____", sCommand

With ParameterSweep
     .ResetOptions
     .SetOptionResetParameterValuesAfterRun "False"
     .SetOptionSkipExistingCombinations "False"
     .SaveOptions
     .SetSimulationType "Frequency"
     .DeleteAllSequences
     .AddSequence "Scanning"
     .AddParameter_Stepwidth "Scanning", "PAA_UC_PHI", "-90", "90", "10"
     .AddParameter_Stepwidth "Scanning", "PAA_UC_THETA", "-60", "60", "10"
End With

MsgBox "Switch to unitcell successful. Additional history lines and Parameter Sweep for scanning defined.", vbInformation, "Change settings from Full Array to Unitcell"

End Sub
