' *Results / Evaluate Fields on Curves / Line-Integration of Electric Fields (incl. Transit Time Factor)
' !!! Do not change the line above !!!
' macro.905
'-----------------------------------------------------------------------------------------------------------------------------
' Copyright 2012-2023 Dassault Systemes Deutschland GmbH
' ================================================================================================
' History of Changes
' ------------------------------------------------------------------------------------------------
' 08-May-2012 ube: First version
' ================================================================================================

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

Public nx As Long
Public ny As Long
Public nz As Long
'-----------------------------------------------------------------------------------------------------------------------------
Const CLight= 299792448
'-----------------------------------------------------------------------------------------------------------------------------
Sub Main ()

        Dim Freq As Double
        Dim lok As Boolean

        Dim xp As Double
        Dim yp As Double
        Dim zp As Double

        Dim xyzstring$(3)
        xyzstring$(1)        = "x"
        xyzstring$(2)        = "y"
        xyzstring$(3)        = "z"

        nx        = Mesh.GetNx
        ny        = Mesh.GetNy
        nz        = Mesh.GetNz

        Begin Dialog UserDialog 420,252,.DialogFunc
                Text 100,7,210,14,"Line-Integration of electric fields",.Text1
                GroupBox 10,29,190,98,"(picked) point (on int.path)",.GroupBox1
                PushButton 40,224,160,21,"Start Integration",.PushButton1
                CancelButton 240,224,90,21
                Text 30,50,30,14,"x =",.Text2
                Text 30,71,30,14,"y =",.Text3
                Text 30,92,30,14,"z =",.Text4
                TextBox 60,50,120,21,.xp
                TextBox 60,71,120,21,.yp
                TextBox 60,92,120,21,.zp
                GroupBox 220,29,190,98,"Integration path",.GroupBox2
                Text 260,75,40,14,"from",.Text5
                Text 260,97,30,14,"to",.Text6
                Text 250,50,60,14,"direction",.Text7
                DropListBox 320,43,50,192,xyzstring(),.xyzDropListBox
                TextBox 300,75,90,21,.wlow
                TextBox 300,97,90,21,.whigh
                GroupBox 10,140,190,70,"Integrated symbol",.GroupBox3
                Text 20,168,70,14,"esymbol =",.Text8
                TextBox 90,168,100,21,.esymbol
                GroupBox 220,140,190,70,"Transit Time Factor",.GroupBox4
                Text 240,182,50,14,"beta =",.Text9
                TextBox 290,182,100,21,.beta
                CheckBox 240,161,160,14,"consider part.velocity",.Check_ttf
        End Dialog
        Dim dlg As UserDialog

        ' default-settings

        If (Pick.GetNumberOfPickedPoints >0) Then
                ' take first pickpoint
                lok        = Pick.GetPickpointCoordinates(1,xp,yp,zp)
                dlg.xp        = CStr(xp)
                dlg.yp        = CStr(yp)
                dlg.zp        = CStr(zp)
        Else
                dlg.xp        = "0"
                dlg.yp        = "0"
                dlg.zp        = "0"
        End If

        dlg.xyzDropListBox   = 0
        dlg.wlow        = CStr(Mesh.GetX(0))
        dlg.whigh        = CStr(Mesh.GetX(nx-1))

        dlg.esymbol = "Mode 1"

        dlg.Check_ttf        = 0
        dlg.beta                 = "not used"

        If (Dialog(dlg) = 0) Then Exit All

        Dim idir As Long
        idir        = 1 + CLng(dlg.xyzDropListBox)

        ' switch to electric field plot

        Dim esymbol As String
        esymbol        = dlg.esymbol

        If InStr(esymbol,"Mode ")=0 Then
                 SelectTreeItem "2D/3D Results\E-Field\" + esymbol
        Else
                 SelectTreeItem "2D/3D Results\Modes\" + esymbol + "\e"
        End If

        Freq        = GetFieldFrequency() * Units.GetFrequencyUnitToSI()

        Dim beta As Double
        If (dlg.Check_ttf = 1) Then
                beta        = RealVal(dlg.beta)
                If (beta <= 0 Or beta >1) Then
                        MsgBox "beta out of range   (0 < beta <= 1)"
                        Exit All
                End If
        Else
                ' for safety, although not used
                beta        = 0
        End If

        xp        = RealVal(dlg.xp)
        yp        = RealVal(dlg.yp)
        zp        = RealVal(dlg.zp)

        Dim wp1 As Double
        Dim wp2 As Double
        Dim wpd As Double
        Dim dw  As Double

        Dim iw     As Long
        Dim iwlow  As Long
        Dim iwhigh As Long

        Dim alfa  As Double
        Dim cosa  As Double
        Dim sina  As Double

        If idir=1 Then
                iwlow        = Mesh.GetClosestXIndex(dlg.wlow)
                iwhigh        = Mesh.GetClosestXIndex(dlg.whigh)
        ElseIf idir=2 Then
                iwlow        = Mesh.GetClosestYIndex(dlg.wlow)
                iwhigh        = Mesh.GetClosestYIndex(dlg.whigh)
        ElseIf idir=3 Then
                iwlow        = Mesh.GetClosestZIndex(dlg.wlow)
                iwhigh        = Mesh.GetClosestZIndex(dlg.whigh)
        End If

        Dim datafile As String

        datafile        = "debug.txt"
        Open datafile For Output As #1
        Print #1, ""
        Print #1, " iw        dw                Ew_real                Ew_imag                cosa                sina                SUM_VRE                SUM_VIM"
        Print #1, ""

        Dim vwsumre As Double, vwsumim As Double

        vwsumre        = 0
        vwsumim        = 0

        Dim vxre As Double, vxim As Double
        Dim vyre As Double, vyim As Double
        Dim vzre As Double, vzim As Double

        For iw=iwlow To iwhigh-1

                If idir=1 Then
                        wp1        = Mesh.GetX(iw)
                        wp2        = Mesh.GetX(iw+1)
                        xp        = 0.5 * (wp1+wp2)
                        wpd        = xp
                ElseIf idir=2 Then
                        wp1        = Mesh.GetY(iw)
                        wp2        = Mesh.GetY(iw+1)
                        yp        = 0.5 * (wp1+wp2)
                        wpd        = yp
                ElseIf idir=3 Then
                        wp1        = Mesh.GetZ(iw)
                        wp2        = Mesh.GetZ(iw+1)
                        zp        = 0.5 * (wp1+wp2)
                        wpd        = zp
                End If

                If (dlg.Check_ttf = 0) Then
                        alfa        = 0
                        cosa        = 1
                        sina        = 0
                Else
                        alfa        = wpd * Units.GetGeometryUnitToSI() * (2 * Pi * Freq) / (beta * CLight)
                        cosa        = Cos(alfa)
                        sina        = Sin(alfa)
                End If

                lok        = GetFieldVector ( xp, yp, zp, vxre, vyre, vzre, vxim, vyim, vzim )
                dw        = (wp2-wp1)* Units.GetGeometryUnitToSI()

                If ( Not lok ) Then
                        MsgBox "Problem in reading fieldvalue" + vbCrLf + vbCrLf + _
                                "maybe integration path outside of calculation box (symmetries...)" + vbCrLf + vbCrLf + _
                                "Abort"
                        Exit All
                Else
                        If idir=1 Then
                                vwsumre        = vwsumre + dw * (vxre*cosa - vxim*sina)
                                vwsumim        = vwsumim + dw * (vxim*cosa + vxre*sina)
                        ElseIf idir=2 Then
                                vwsumre        = vwsumre + dw * (vyre*cosa - vyim*sina)
                                vwsumim        = vwsumim + dw * (vyim*cosa + vyre*sina)
                        ElseIf idir=3 Then
                                vwsumre        = vwsumre + dw * (vzre*cosa - vzim*sina)
                                vwsumim        = vwsumim + dw * (vzim*cosa + vzre*sina)
                        End If
                End If

                Print #1,CStr(iw)+"        "+CStr(Format(dw,"0.000E+00"))+"        "+CStr(Format(vxre,"0.000E+00"))+"        "+ _
                  CStr(Format(vxim,"0.000E+00"))+"        "+CStr(Format(cosa,"0.000E+00"))+"        "+CStr(Format(sina,"0.000E+00"))+ _
                  "        "+CStr(Format(vwsumre,"0.000E+00"))+"        "+CStr(Format(vwsumim,"0.000E+00"))

        Next iw

        Close #1

        ' Shell("notepad.exe " + datafile, 1)


        MsgBox _
                                "Summary of electric voltage-integration:" + vbCrLf + _
                                "===============================" + vbCrLf + vbCrLf + _
                                "esymbol        = " + esymbol + vbCrLf + _
                                "frequency        = " + CStr(Freq) + vbCrLf + _
                                "beta         = " + dlg.beta + vbCrLf + _
                                "direction         = " + xyzstring(idir) + vbCrLf + vbCrLf + _
                                "V_real        = " + CStr(vwsumre) + vbCrLf + vbCrLf + _
                                "V_imag        = " + CStr(vwsumim)

        Exit All


                MsgBox "iw                = " + CStr(iw) + vbCrLf _
                     + "alfa        = " + CStr(alfa) + vbCrLf _
                     + "cosa        = " + CStr(cosa) + vbCrLf _
                     + "sina        = " + CStr(sina) + vbCrLf _
                     + "dw                = " + CStr(dw) + vbCrLf _
                     + "vxre        = " + CStr(vxre) + vbCrLf _
                     + "vxim        = " + CStr(vxim) + vbCrLf _
                     + "vyre        = " + CStr(vyre) + vbCrLf _
                     + "vyim        = " + CStr(vyim) + vbCrLf _
                     + "vzre        = " + CStr(vzre) + vbCrLf _
                     + "vzim        = " + CStr(vzim) + vbCrLf _
                     + "vwsumre        = " + CStr(vwsumre) + vbCrLf _
                     + "vwsumim        = " + CStr(vwsumim) + vbCrLf _
                     + "xp                = " + CStr(xp) + vbCrLf _
                     + "yp                = " + CStr(yp) + vbCrLf _
                     + "zp                = " + CStr(zp) + vbCrLf

End Sub

'-----------------------------------------------------------------------------------------------------------------------------

Function DialogFunc%(Item As String, Action As Integer, Value As Integer)

        Select Case Action
                Case 1 ' Dialog box initialization
                Case 2 ' Value changing or button pressed
                        Select Case Item
                                Case "Check_ttf"
                                        If Value=0 Then
                                           DlgText "beta",  "not used"
                                        ElseIf Value=1 Then
                                           DlgText "beta",  "1"
                                        End If
                                        DialogFunc%= True
                                Case "xyzDropListBox"
                                        If Value=0 Then
                                           DlgText "wlow",  CStr(Mesh.GetX(0))
                                           DlgText "whigh", CStr(Mesh.GetX(nx-1))
                                        ElseIf Value=1 Then
                                           DlgText "wlow",  CStr(Mesh.GetY(0))
                                           DlgText "whigh", CStr(Mesh.GetY(ny-1))
                                        ElseIf Value=2 Then
                                           DlgText "wlow",  CStr(Mesh.GetZ(0))
                                           DlgText "whigh", CStr(Mesh.GetZ(nz-1))
                                        End If
                                        DialogFunc%= True
                        End Select
                Case 3 ' ComboBox or TextBox Value changed
                Case 4 ' Focus changed
                Case 5 ' Idle
        End Select
End Function
