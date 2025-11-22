Option Explicit
'#include "vba_globals_all.lib"

' DumpedParticlesNew
' ------------------------------------------------------------------------------
' This macro exports the first and last stored (=typically close to dump) positions
'   of every stored particle, contained within TRK trajectory plot
' ================================================================================================
' Copyright 2017-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------
' 11-Sep-2017 mbk: modernized Version of old Macro using the Particle Reader Object
' ------------------------------------------------------------------------------

Type Coordinates
	x As Single
	y As Single
	z As Single
End Type


Sub Main ()

	With ParticleTrajectoryReader
		.Reset
		.LoadTrajectoryData

		Dim N_traj As Long
		Dim lstPosX() As Single, lstPosY() As Single, lstPosZ() As Single
		Dim first_point As Coordinates, last_point As Coordinates

		N_traj = .GetNTrajectories

		Dim sfilename As String
		sfilename = GetProjectPath("Root")+"\"+"dumped_positions.txt"

		Open sfilename For Output As #1
		' Print Header
		Print #1,  "Particle ID" + "    "+  "u_start" + "    "+  "v_start" +"    "+  "w_start" +"    "+  "u_end" +"    "+  "v_end" +"    "+  "w_end"
		Print #1,  "------------------------------------------------------------------------------------------------"

		Dim i As Integer
		For i=0 To N_traj-1
			.SelectTrajectory(i)

			lstPosX = .GetQuantityValues("Position", "X")
			lstPosY = .GetQuantityValues("Position", "Y")
			lstPosZ = .GetQuantityValues("Position", "Z")

			first_point.x = lstPosX(LBound(lstPosX))*Units.GetGeometryUnitToSI
			first_point.y = lstPosY(LBound(lstPosY))*Units.GetGeometryUnitToSI
			first_point.z = lstPosZ(LBound(lstPosZ))*Units.GetGeometryUnitToSI

			last_point.x = lstPosX(UBound(lstPosX))*Units.GetGeometryUnitToSI
			last_point.y = lstPosY(UBound(lstPosY))*Units.GetGeometryUnitToSI
			last_point.z = lstPosZ(UBound(lstPosZ))*Units.GetGeometryUnitToSI

			' PP10 and PP15 are contained in included vba_globals_all.lib (stored in Library\Includes)
			Print #1,  PP10(i+1) + PP15(first_point.x) + PP15(first_point.y) + PP15(first_point.z) + PP15(last_point.x) + PP15(last_point.y) + PP15(last_point.z)

		Next i

		Close #1

		MsgBox "File "+sfilename+" has been successfully written.",vbInformation+vbOkOnly

	End With

End Sub
