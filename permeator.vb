Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("PermeatorExtn.Permeator")> Public Class Permeator
    ' reference to the 2.1 type library has been set
    ' change the project name and the class name in the VB _
    'to reflect your entry in the Objects Manager _
    'ProgID/CLSID field of the EDF such that the names _
    'correspond to: _
    '<project name>.<class name>

    Private myContainer As HYSYS.ExtnUnitOperationContainer
    'Step 1 - Complete the variable declarations
    '***************************************************************************'
    '                                                                           '
    '                                 EDF Variables                             '
    '                                                                           '
    '***************************************************************************'
    Private edfInlet As HYSYS.ProcessStream
    Private edfPermeate As HYSYS.ProcessStream
    Private edfRetentate As HYSYS.ProcessStream
    Private edfPermIn As HYSYS.ProcessStream
    'Private edfThick As InternalRealVariable
    'Private edfDiam As InternalRealVariable
    'Private edfLen As InternalRealVariable
    'Private edfAperm As InternalRealVariable
    'Private edfNtubes As InternalRealVariable
    'Private edfn As InternalRealVariable
    'Private edfDT As InternalRealVariable
    'Private edfDH As InternalRealVariable
    'Private edfComposition As InternalRealFlexVariable
    'Private edfCompName As InternalTextFlexVariable
    'Private edfStreamName As InternalTextFlexVariable
    'Private edfVapFrac As InternalRealFlexVariable
    'Private edfTemperature As InternalRealFlexVariable
    'Private edfPressure As InternalRealFlexVariable
    'Private edfMolFlow As InternalRealFlexVariable
    'Private edfMassFlow As InternalRealFlexVariable
    'Private edfPressDrop As InternalRealVariable
    'Private edfPermPressDrop As InternalRealVariable
    'Private edfDiamExt As InternalRealVariable
    'Private edfLengthPos As InternalRealFlexVariable
    'Private edfCompT As InternalRealFlexVariable
    'Private edfCompH As InternalRealFlexVariable
    'Private edfNpoints As InternalRealVariable
    'Private edfk As InternalRealVariable

    Private pressureRFV As HYSYS.InternalRealFlexVariable
    Private flowRFV As HYSYS.InternalRealFlexVariable
    Private NumberOfPoints As HYSYS.InternalRealVariable
    Dim myPlotName As HYSYS.InternalTextVariable
    Dim myPlot As HYSYS.TwoDimensionalPlot



    Public Function Initialize(ByRef Container As HYSYS.ExtnUnitOperationContainer, ByRef IsRecalling As Boolean) As Integer

        On Error GoTo ErrorTrap
        ' Initialize container
        myContainer = Container
        ' Initialize EDF variables
        Call PointedEDFVariables()

        ' Recall check
        If Not IsRecalling Then
            'Step 3 - set the NumberOfPoints variable to 0.
            NumberOfPoints.Value = 0
        End If

        Call CreatePlot()

        Initialize = HYSYS.CurrentExtensionVersion_enum.extnCurrentVersion
        Exit Function

ErrorTrap:
        MsgBox("Initialize Error")
    End Function

    Public Sub Execute(ByRef Forgetting As Boolean)
        ' execute gets hit twice, once on a forgetting pass and then on _
        'a calculate pass
        On Error GoTo ErrorTrap

        If Forgetting Then Exit Sub


        'Step 5 - Check that we have enough information to Calculate
        If edfInlet Is Nothing Then Exit Sub
        If edfPermeate Is Nothing Then Exit Sub
        If edfRetentate Is Nothing Then Exit Sub
        If NumberOfPoints.Value <= 1 Then Exit Sub
        If Not edfInlet.Pressure.IsKnown Then Exit Sub
        If Not edfInlet.Temperature.IsKnown And Not edfPermeate.Temperature.IsKnown Then Exit Sub

        'Check to see that a Composition is specified
        Dim bv As Object
        Dim k As Short
        Dim compOK As Boolean
        compOK = True
        bv = edfInlet.ComponentMolarFraction.IsKnown
        For k = LBound(bv) To UBound(bv)
            'UPGRADE_WARNING: Couldn't resolve default property of object bv(k). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If bv(k) = False Then compOK = False
            Exit For
        Next k
        If compOK = False Then
            compOK = True
            bv = edfPermeate.ComponentMolarFraction.IsKnown
            For k = LBound(bv) To UBound(bv)
                'UPGRADE_WARNING: Couldn't resolve default property of object bv(k). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If bv(k) = False Then compOK = False
                Exit For
            Next k
        End If
        If compOK = False Then Exit Sub



        'Check that all pressure flow data is valid
        Dim DataIsOK As Boolean
        Dim I As Short
        DataIsOK = True
        For I = 0 To NumberOfPoints.Value - 1
            If pressureRFV.Values(I) = HYSYS.EmptyValue_enum.HEmpty Or flowRFV.Values(I) = HYSYS.EmptyValue_enum.HEmpty Then
                DataIsOK = False
                Exit For
            End If
        Next I
        If Not DataIsOK Then Exit Sub

        'Step 6 - Check the Flow and Pressure specs of the operation
        Dim specs As Integer
        specs = 0
        If edfInlet.StdGasFlow.IsKnown Then specs = specs + 1
        If edfPermeate.StdGasFlow.IsKnown Then specs = specs + 1
        If edfPermeate.Pressure.IsKnown Then specs = specs + 1

        'Step 7 - If only one specifaction is known, then execute this code
        Dim press As Double
        Dim flow As Double
        If specs = 1 Then
            If edfPermeate.Pressure.IsKnown Then
                'Calculate the Flow from the PQ Data
                flow = LinearInterpolation(pressureRFV, flowRFV, (edfPermeate.Pressure))
                edfPermeate.MolarFlow.Calculate(flow * 3600, "m3/h_(gas)")
            ElseIf edfPermeate.StdGasFlow.IsKnown Then
                'Calculate the Pressure from the PQ Data
                press = LinearInterpolation(flowRFV, pressureRFV, (edfPermeate.StdGasFlow))
                edfPermeate.Pressure.Calculate(press)
            Else
                'Calculate the Pressure from the PQ Data
                press = LinearInterpolation(flowRFV, pressureRFV, (edfInlet.StdGasFlow))
                edfPermeate.Pressure.Calculate(press)
            End If
        End If
        'Step 8 - Complete the Balance code.
        Dim StreamsList(1) As HYSYS.ProcessStream
        StreamsList(0) = edfInlet
        StreamsList(1) = edfPermeate
        myContainer.Balance(HYSYS.BalanceType_enum.btTotalBalance, 1, StreamsList)

        'Check if the Feed and Product streams are completely solved
        If edfInlet.DuplicateFluid.IsUpToDate And edfPermeate.DuplicateFluid.IsUpToDate Then
            myContainer.SolveComplete()
        End If
        Exit Sub
ErrorTrap:
        MsgBox("Execute Error")
    End Sub



    Sub StatusQuery(ByRef Status As HYSYS.ObjectStatus)
        On Error GoTo ErrorTrap
        'If the object is ignored then Exit the Subroutine
        If myContainer.ExtensionInterface.IsIgnored Then Exit Sub
        'Error messsages
        If edfInlet Is Nothing Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 1, "Feed stream is missing")
        End If
        If edfPermeate Is Nothing Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 2, "Permeate stream is missing")
        End If
        If edfRetentate Is Nothing Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 3, "Non permeating outlet stream is missing")
        End If
        'If edfLen < 0 Or edfDiam < 0 Or edfThick < 0 Or edfk < 0 Or edfNtubes < 0 Or edfAperm < 0 Then
        '    Call Status.AddStatusCondition(slMissingRequiredInformation, 4, "Physical parameters missing or incorrect")
        'End If
        'If flagRet Then
        '    Call Status.AddStatusCondition(slError, 5, "All feed flow is being permeated")
        'End If
        'If flagPerm Then
        '    Call Status.AddStatusCondition(slError, 6, "No species permeating")
        'End If

        ' DELETE
        If NumberOfPoints.Value <= 1 Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 7, "Not enough PQ Data Points")
        End If

        Dim DataIsOK As Boolean
        Dim I As Short
        DataIsOK = True
        For I = 0 To NumberOfPoints.Value - 1
            If pressureRFV.Values(I) = HYSYS.EmptyValue_enum.HEmpty Or flowRFV.Values(I) = HYSYS.EmptyValue_enum.HEmpty Then
                DataIsOK = False
                Exit For
            End If
        Next I
        If DataIsOK = False Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slWarning, 4, "PQ Data is incomplete")
        End If

        'Check Specs Again
        Dim specs As Integer
        specs = 0
        If Not edfInlet Is Nothing And Not edfPermeate Is Nothing Then
            If edfInlet.StdGasFlow.IsKnown Then specs = specs + 1
            If edfPermeate.StdGasFlow.IsKnown Then specs = specs + 1
            If edfPermeate.Pressure.IsKnown Then specs = specs + 1
        End If

        'Step 11 - If specs < 1, give a suitable status message.
        If specs < 1 Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 5, "Requires 1 flow or pressure spec")
        End If

        Exit Sub
ErrorTrap:
        MsgBox("Status Query Error")
    End Sub

    Sub Terminate()
        edfInlet = Nothing
        edfPermeate = Nothing
        edfRetentate = Nothing
        edfPermIn = Nothing
        pressureRFV = Nothing
        flowRFV = Nothing
        NumberOfPoints = Nothing
        myContainer.DeletePlot("PQPlot")
        myPlot = Nothing
        myPlotName = Nothing

        'edfThick = Nothing
        'edfLen = Nothing
        'edfDiam = Nothing
        'edfAperm = Nothing
        'edfNtubes = Nothing
        'edfn = Nothing
        'edfDT = Nothing
        'edfDH = Nothing
        'edfComposition = Nothing
        'edfCompName = Nothing
        'edfStreamName = Nothing
        'edfVapFrac = Nothing
        'edfTemperature = Nothing
        'edfPressure = Nothing
        'edfMolFlow = Nothing
        'edfMassFlow = Nothing
        'edfPressDrop = Nothing
        'edfPermPressDrop = Nothing
        'edfLengthPos = Nothing
        'edfCompT = Nothing
        'edfCompH = Nothing
        'edfNpoints = Nothing
        'myPlotH = Nothing
        'myPlotNameH = Nothing
        'myPlotT = Nothing
        'myPlotNameT = Nothing
        'edfDiamExt = Nothing
        'edfk = Nothing
    End Sub

    Sub VariableChanged(ByRef Variable As HYSYS.InternalVariableWrapper)
        ' Called when the user modifies any edf variable. It is required to update these variales as needed

        On Error GoTo ErrorTrap
        Dim pressureVT As Object
        Dim flowVT As Object
        Dim I As Short
        Select Case Variable.Tag
            '%% Attachment variables
            Case "Inlet"
                edfInlet = myContainer.FindVariable("Inlet").Variable.object
            Case "Permeate"
                edfPermeate = myContainer.FindVariable("Permeate").Variable.object
            Case "Retentate"
                edfRetentate = myContainer.FindVariable("Retentate").Variable.object
            Case "PermIn"
                edfPermIn = myContainer.FindVariable("PermIn").Variable.object
            Case "ActualizePlot"
                'edfCompT.SetBounds(edfNpoints.Value + 1)
                'edfCompH.SetBounds(edfNpoints.Value + 1)
                'edfLengthPos.SetBounds(edfNpoints.Value + 1)
                'edfLengthPos.Values = GrafL
                'edfCompT.Values = FpermCells
                'edfCompH.Values = FpermCells

            '%% Geometry variables
            Case "Length"
                'edfLen = myContainer.FindVariable("Length").Variable
                'L = edfLen.Value * edfNtubes                    ' total length of tubes
                '    ' Recalculation of cross-area and total volume
                '    Area = pi * (Din ^ 2) / 4
                '    Volume = L * Area
                '    GrafL = LengthVector(edfLen, edfNpoints)
                '    ' Recalculation of total permeation surface
                '    Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                '    edfAperm.SetValue L * (2 * pi * Ravg)
                'Aperm = edfAperm.Value
            Case "Diam"
                'edfDiam = myContainer.FindVariable("Diam").Variable
                'Din = edfDiam.GetValue
                '    ' Recalculation of cross-area and total volume
                '    Area = pi * (Din ^ 2) / 4
                '    Volume = edfLen.Value * edfNtubes * Area
                '    ' Recalculation of total permeation surface
                '    Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                '    edfAperm.SetValue L * (2 * pi * Ravg)
                'Aperm = edfAperm.Value
            Case "Thickness"
                'edfThick = myContainer.FindVariable("Thickness").Variable
                'thick = edfThick.Value
                '    ' Recalculation of total permeation surface
                '    Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                '    edfAperm.SetValue L * (2 * pi * Ravg)
                'Aperm = edfAperm.Value
            Case "Aperm"
                'edfAperm = myContainer.FindVariable("Aperm").Variable
                'Aperm = edfAperm.Value
                '    ' Recalculation of total permeation surface
                '    Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                '    L = Aperm / (2 * pi * Ravg)                     ' total length of tubes
                '    edfLen.SetValue L / edfNtubes.GetValue
                '' Recalculation of cross-area and total volume
                '    Area = pi * (Din ^ 2) / 4
                '    Volume = L * Area
            Case "Ntubes"
                'edfNtubes = myContainer.FindVariable("Ntubes").Variable
                '' Recalculation of total length (mantaining length/tube (edfLen))
                'L = edfLen.Value * edfNtubes
                '    ' Recalculation of total permeation surface (last change affects area)
                '    Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                '    edfAperm.SetValue L * (2 * pi * Ravg)
                'Aperm = edfAperm.Value
                '    ' Recalculation of cross-area and total volume
                '    Area = pi * (Din ^ 2) / 4
                '    Volume = L * Area

            '%% Pressure drop and others
            Case "n"
                'edfn = myContainer.FindVariable("n").Variable
            Case "PressDrop"
                'edfPressDrop = myContainer.FindVariable("PressDrop").Variable
            Case "PermPressDrop"
                'edfPermPressDrop = myContainer.FindVariable("PermPressDrop").Variable
            Case "k"
                'edfk = myContainer.FindVariable("k").Variable
            Case "NumberOfPoints"
                'edfNpoints = myContainer.FindVariable("NumberOfPoints").Variable
                'GrafL = LengthVector(edfLen, edfNpoints)

            ' TO DELETE
            Case "NumberOfPoints"
                pressureVT = pressureRFV.ValuesflowVT = flowRFV.Values
                ReDim Preserve pressureVT(NumberOfPoints.Value - 1)
                ReDim Preserve flowVT(NumberOfPoints.Value - 1)
                For I = LBound(flowVT) To UBound(flowVT)
                    If IsNothing(flowVT(I)) Then
                        flowVT(I) = HYSYS.EmptyValue_enum.HEmpty
                    End If
                    If IsNothing(pressureVT(I)) Then
                        pressureVT(I) = HYSYS.EmptyValue_enum.HEmpty
                    End If
                Next I
                pressureRFV.SetBounds((NumberOfPoints.Value))
                flowRFV.SetBounds((NumberOfPoints.Value))
                pressureRFV.Values = pressureVT
                flowRFV.Values = flowVT

        End Select
        Exit Sub
ErrorTrap:
        myContainer.Trace(myContainer.name & ": Error in Variable Changed for variable " & Variable.Tag & ".", False)
        MsgBox("Variable Changed Error")
    End Sub

    Function VariableChanging(ByRef Variable As HYSYS.InternalVariableWrapper) As Boolean

        Select Case Variable.Tag

            Case "NumberOfPoints"
                If Variable.NewRealValue < 2 Or Variable.NewRealValue > 100 Then
                    MsgBox("Entered Value is out of range, must be between 2 and 100.")
                    VariableChanging = False
                    Exit Function
                End If

        End Select

        VariableChanging = True

    End Function
    Private Sub BasisChanged()
        ' Not sure if this works, never saw it running
    End Sub


    Private Sub PointedEDFVariables()
        With myContainer
            edfInlet = .FindVariable("Inlet").Variable.object
            edfPermeate = .FindVariable("Permeate").Variable.object
            edfRetentate = .FindVariable("Retentate").Variable.object
            edfPermIn = .FindVariable("PermIn").Variable.object
            pressureRFV = .FindVariable("PressureData").Variable
            flowRFV = .FindVariable("FlowData").Variable
            NumberOfPoints = .FindVariable("NumberOfPoints").Variable
            myPlotName = .FindVariable("PlotName").Variable
            'edfThick = .FindVariable("Thickness").Variable
            'edfLen = .FindVariable("Length").Variable
            'edfDiam = .FindVariable("Diam").Variable
            'edfAperm = .FindVariable("Aperm").Variable
            'edfNtubes = .FindVariable("Ntubes").Variable
            'edfn = .FindVariable("n").Variable
            'edfDT = .FindVariable("DT").Variable
            'edfDH = .FindVariable("DH").Variable
            'edfComposition = .FindVariable("Composition").Variable
            'edfCompName = .FindVariable("CompoName").Variable
            'edfStreamName = .FindVariable("StreamNames").Variable
            'edfVapFrac = .FindVariable("VapFrac").Variable
            'edfTemperature = .FindVariable("StreamTemp").Variable
            'edfPressure = .FindVariable("StreamPress").Variable
            'edfMolFlow = .FindVariable("StreamMolFlow").Variable
            'edfMassFlow = .FindVariable("StreamMassFlow").Variable
            'edfPressDrop = .FindVariable("PressDrop").Variable
            'edfPermPressDrop = .FindVariable("PermPressDrop").Variable
            'edfLengthPos = .FindVariable("LengthPos").Variable
            'edfCompT = .FindVariable("CompT").Variable
            'edfCompH = .FindVariable("CompH").Variable
            'edfNpoints = .FindVariable("NumberOfPoints").Variable
            'myPlotNameH = .FindVariable("PlotNameH").Variable
            'myPlotNameT = .FindVariable("PlotNameT").Variable
            'edfDiamExt = .FindVariable("DiamExtern").Variable
            ''molDensity = .FindVariable("molDensity").Variable
            'edfk = .FindVariable("k").Variable
        End With
    End Sub

    Function LinearInterpolation(ByRef xDataRFV As HYSYS.InternalRealFlexVariable, ByRef yDataRFV As HYSYS.InternalRealFlexVariable, ByRef xPoint As HYSYS.RealVariable) As Double
        'This method linear interpolates to find the y point that coresponds to the known
        'x point for the given x and y data sets.

        Dim xData As Object
        Dim yData As Object
        Dim x As Double
        Dim y As Double

        On Error GoTo ErrorTrap

        Dim High As Integer
        Dim Low As Integer
        Dim number As Integer

        y = HYSYS.EmptyValue_enum.HEmpty
        LinearInterpolation = y

        'UPGRADE_WARNING: Couldn't resolve default property of object xDataRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xData = xDataRFV.Values
        'UPGRADE_WARNING: Couldn't resolve default property of object yDataRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yData = yDataRFV.Values
        x = xPoint.Value

        High = UBound(xData)
        Low = LBound(xData)
        number = High - Low + 1

        'There must be more than 1 data point to Linearly Interpolate
        If number <= 1 Then Exit Function
        'Check that the x and y Data have the same bounds
        If High <> UBound(yData) Or Low <> LBound(yData) Then Exit Function
        'Sort the x Data from low to high
        Call Sort(xData, yData)

        'Check to see that the x point is within the x Data Range
        'UPGRADE_WARNING: Couldn't resolve default property of object xData(High). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xData(Low). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If x < xData(Low) Or x > xData(High) Then
            MsgBox("The point is outside the data range")
            Exit Function
        End If

        Dim I As Short
        'Search the data until the x Point is between two x Data points
        For I = Low To High
            'UPGRADE_WARNING: Couldn't resolve default property of object xData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If x < xData(I) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object yData(I - 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object yData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xData(I - 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object yData(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                y = yData(I) - ((xData(I) - x) / (xData(I) - xData(I - 1)) * (yData(I) - yData(I - 1)))
                Exit For
            End If
        Next I
        'UPGRADE_WARNING: Couldn't resolve default property of object xData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xDataRFV.Values = xData
        'UPGRADE_WARNING: Couldn't resolve default property of object yData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yDataRFV.Values = yData
        LinearInterpolation = y
        Exit Function

ErrorTrap:
        MsgBox("Interpolation Error")
    End Function

    Private Sub Sort(ByRef KeyArray As Object, ByRef OtherArray As Object)

        'Description: Sorts the arrays passed so that smallest values occur first in KeyArray()
        '             does the same rearrangements on OtherArray() so values still correspond
        '             - Uses a Ripple type sort (Good for smallish data sets)
        '
        On Error GoTo ErrorTrap
        'Declare Variables------------------------------------------------------------------------------

        Dim I As Object
        Dim J As Short 'Counters
        Dim Temp As Object 'used to swap values

        'Procedure--------------------------------------------------------------------------------------

        For I = LBound(KeyArray) To UBound(KeyArray) - 1

            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For J = I + 1 To UBound(KeyArray)

                'UPGRADE_WARNING: Couldn't resolve default property of object KeyArray(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object KeyArray(J). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If KeyArray(J) < KeyArray(I) Then
                    'swap the xdata
                    'UPGRADE_WARNING: Couldn't resolve default property of object KeyArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Temp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Temp = KeyArray(J)
                    'UPGRADE_WARNING: Couldn't resolve default property of object KeyArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    KeyArray(J) = KeyArray(I)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Temp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object KeyArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    KeyArray(I) = Temp

                    'and the corresponding ydata
                    'UPGRADE_WARNING: Couldn't resolve default property of object OtherArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Temp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Temp = OtherArray(J)
                    'UPGRADE_WARNING: Couldn't resolve default property of object OtherArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    OtherArray(J) = OtherArray(I)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Temp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object OtherArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    OtherArray(I) = Temp

                End If

            Next  'J

        Next  'I
        Exit Sub

ErrorTrap:
        MsgBox("Sorting Error")
    End Sub

    Sub CreatePlot()

        If Not myPlot Is Nothing Then
            myContainer.DeletePlot("PQPlot")
            'UPGRADE_NOTE: Object myPlot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            myPlot = Nothing
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object myPlot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        myContainer.BuildPlot2("PQPlot", myPlot, HYSYS.HPlotType_enum.hptTwoDimensionalPlot)
        myPlotName.Value = "PQPlot"

        With myPlot
            .TitleData = "Wellhead PQ Relationship"

            .SetAxisLabelData(HYSYS.HAxisType_enum.hatXAxis, "Flow")
            .SetAxisLabelData(HYSYS.HAxisType_enum.hatYAxis, "Pressure")
            .SetAxisLabelVisible(HYSYS.HAxisType_enum.hatXAxis, True)
            .SetAxisLabelVisible(HYSYS.HAxisType_enum.hatYAxis, True)

            .LegendVisible = True

            .CreateXYDataSet(1, "PQData")
            .SetDataSetXData(1, flowRFV)
            .SetDataSetYData(1, pressureRFV)
            .SetDataSetColour(1, "Red")

        End With

    End Sub
End Class