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
    Private Feed As HYSYS.ProcessStream
    Private Product As HYSYS.ProcessStream
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
        If Feed Is Nothing Then Exit Sub
        If Product Is Nothing Then Exit Sub
        If NumberOfPoints.Value <= 1 Then Exit Sub
        If Not Feed.Pressure.IsKnown Then Exit Sub
        If Not Feed.Temperature.IsKnown And Not Product.Temperature.IsKnown Then Exit Sub

        'Check to see that a Composition is specified
        Dim bv As Object
        Dim k As Short
        Dim compOK As Boolean
        compOK = True
        'UPGRADE_WARNING: Couldn't resolve default property of object Feed.ComponentMolarFraction.IsKnown. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object bv. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bv = Feed.ComponentMolarFraction.IsKnown
        For k = LBound(bv) To UBound(bv)
            'UPGRADE_WARNING: Couldn't resolve default property of object bv(k). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If bv(k) = False Then compOK = False
            Exit For
        Next k
        If compOK = False Then
            compOK = True
            'UPGRADE_WARNING: Couldn't resolve default property of object Product.ComponentMolarFraction.IsKnown. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object bv. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            bv = Product.ComponentMolarFraction.IsKnown
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
        If Feed.StdGasFlow.IsKnown Then specs = specs + 1
        If Product.StdGasFlow.IsKnown Then specs = specs + 1
        If Product.Pressure.IsKnown Then specs = specs + 1

        'Step 7 - If only one specifaction is known, then execute this code
        Dim press As Double
        Dim flow As Double
        If specs = 1 Then
            If Product.Pressure.IsKnown Then
                'Calculate the Flow from the PQ Data
                flow = LinearInterpolation(pressureRFV, flowRFV, (Product.Pressure))
                Product.MolarFlow.Calculate(flow * 3600, "m3/h_(gas)")
            ElseIf Product.StdGasFlow.IsKnown Then
                'Calculate the Pressure from the PQ Data
                press = LinearInterpolation(flowRFV, pressureRFV, (Product.StdGasFlow))
                Product.Pressure.Calculate(press)
            Else
                'Calculate the Pressure from the PQ Data
                press = LinearInterpolation(flowRFV, pressureRFV, (Feed.StdGasFlow))
                Product.Pressure.Calculate(press)
            End If
        End If
        'Step 8 - Complete the Balance code.
        Dim StreamsList(1) As HYSYS.ProcessStream
        StreamsList(0) = Feed
        StreamsList(1) = Product
        myContainer.Balance(HYSYS.BalanceType_enum.btTotalBalance, 1, StreamsList)

        'Check if the Feed and Product streams are completely solved
        If Feed.DuplicateFluid.IsUpToDate And Product.DuplicateFluid.IsUpToDate Then
            myContainer.SolveComplete()
        End If
        Exit Sub
ErrorTrap:
        MsgBox("Execute Error")
    End Sub



    Sub StatusQuery(ByRef Status As HYSYS.ObjectStatus)
        On Error GoTo ErrorTrap
        'Step 9 - If the object is ignored then Exit the Subroutine
        'UPGRADE_WARNING: Couldn't resolve default property of object myContainer.ExtensionInterface.IsIgnored. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If myContainer.ExtensionInterface.IsIgnored Then Exit Sub

        'Step 10 - Complete the following If ... Then statements. Hint: look at the error messsages
        If Feed Is Nothing Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 1, "Requires an Inlet Stream")
        End If

        If Product Is Nothing Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 2, "Requires an Outlet Stream")
        End If

        If NumberOfPoints.Value <= 1 Then
            Call Status.AddStatusCondition(HYSYS.StatusLevel_enum.slMissingRequiredInformation, 3, "Not enough PQ Data Points")
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
        If Not Feed Is Nothing And Not Product Is Nothing Then
            If Feed.StdGasFlow.IsKnown Then specs = specs + 1
            If Product.StdGasFlow.IsKnown Then specs = specs + 1
            If Product.Pressure.IsKnown Then specs = specs + 1
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
        'UPGRADE_NOTE: Object Feed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Feed = Nothing
        'UPGRADE_NOTE: Object Product may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Product = Nothing
        'UPGRADE_NOTE: Object pressureRFV may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pressureRFV = Nothing
        'UPGRADE_NOTE: Object flowRFV may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        flowRFV = Nothing
        'UPGRADE_NOTE: Object NumberOfPoints may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        NumberOfPoints = Nothing
        myContainer.DeletePlot("PQPlot")
        'UPGRADE_NOTE: Object myPlot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        myPlot = Nothing
        'UPGRADE_NOTE: Object myPlotName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        myPlotName = Nothing
    End Sub

    Sub VariableChanged(ByRef Variable As HYSYS.InternalVariableWrapper)

        On Error GoTo ErrorTrap
        Dim pressureVT As Object
        Dim flowVT As Object
        Dim I As Short
        Select Case Variable.Tag

            Case "Inlet"
                'UPGRADE_WARNING: Couldn't resolve default property of object myContainer.FindVariable().Variable.object. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Feed = myContainer.FindVariable("Inlet").Variable.object
            Case "NoPermOut"
                'UPGRADE_WARNING: Couldn't resolve default property of object myContainer.FindVariable().Variable.object. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Product = myContainer.FindVariable("NoPermOut").Variable.object

            Case "NumberOfPoints"
                'UPGRADE_WARNING: Couldn't resolve default property of object pressureRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object pressureVT. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pressureVT = pressureRFV.Values
                'UPGRADE_WARNING: Couldn't resolve default property of object flowRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object flowVT. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                flowVT = flowRFV.Values
                ReDim Preserve pressureVT(NumberOfPoints.Value - 1)
                ReDim Preserve flowVT(NumberOfPoints.Value - 1)
                For I = LBound(flowVT) To UBound(flowVT)
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If IsNothing(flowVT(I)) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object flowVT(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        flowVT(I) = HYSYS.EmptyValue_enum.HEmpty
                    End If
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If IsNothing(pressureVT(I)) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object pressureVT(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        pressureVT(I) = HYSYS.EmptyValue_enum.HEmpty
                    End If
                Next I
                pressureRFV.SetBounds((NumberOfPoints.Value))
                flowRFV.SetBounds((NumberOfPoints.Value))
                'UPGRADE_WARNING: Couldn't resolve default property of object pressureVT. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pressureRFV.Values = pressureVT
                'UPGRADE_WARNING: Couldn't resolve default property of object flowVT. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                flowRFV.Values = flowVT

        End Select
        Exit Sub
ErrorTrap:
        MsgBox("Variable Changed Error")
    End Sub


    Private Sub PointedEDFVariables()
        With myContainer
            Feed = .FindVariable("Inlet").Variable.object
            Product = .FindVariable("NoPermOut").Variable.object
            pressureRFV = .FindVariable("PressureData").Variable
            flowRFV = .FindVariable("FlowData").Variable
            NumberOfPoints = .FindVariable("NumberOfPoints").Variable
            myPlotName = .FindVariable("PlotName").Variable
            'edfThick = .FindVariable("Thickness").Variable
            'edfLen = .FindVariable("Length").Variable
            'edfDiam = .FindVariable("Diam").Variable
            'edfAperm = .FindVariable("Aperm").Variable
            'edfNtubes = .FindVariable("Ntubes").Variable
            'edfInlet = .FindVariable("Inlet").Variable.object
            'edfPermeate = .FindVariable("PermOut").Variable.object
            'edfRetentate = .FindVariable("NoPermOut").Variable.object
            'edfPermIn = .FindVariable("PermIn").Variable.object
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
End Class