Option Strict Off
Option Explicit On
Imports HYSYS

<System.Runtime.InteropServices.ProgId("PermeatorExtn.Permeator")> Public Class Permeator

    Private myContainer As ExtnUnitOperationContainer
    '***************************************************************************'
    '                             VB Variables                                  '
    '***************************************************************************'
    ' Indices of permeated components: in this order: H2, HD, HT, D2, DT, T2
    Private ReadOnly nPerm As Short = 6          ' number of permeating species (6 if all hydrogens)
    Private ReadOnly nHeteroNuclear As Short = 3
    Private nPermAtom As Short          ' number of permeating atoms (3 if H, D and T)
    Private ReadOnly permCoeffs As Double(,) = { ' array of coefficients to compute partial pressures
        {1, 0.5, 0.5, 0, 0, 0},                  ' contribution per diatomic molecule
        {0, 0.5, 0, 1, 0.5, 0},
        {0, 0, 0.5, 0, 0.5, 1}
    }
    Private permIndicesHomo As Short()
    Private permIndicesHetero As Short(,) ' array of indices to locate heteronuclear molecules
    ' for each atom in HYSYS, with the shape: {{-1, HD, HT}, {HD, -1, DT}, {HT, DT, -1}}
    Private permIndices As Double()     ' array of indices to adress each diatomic species in
    ' hysys component list

    ' Permeation error flags
    Private flagRet As Boolean, flagPerm As Boolean
    ' Plot
    Dim myPlotNameH As InternalTextVariable
    Dim myPlotH As TwoDimensionalPlot
    Dim myPlotNameT As InternalTextVariable
    Dim myPlotT As TwoDimensionalPlot

    '***************************************************************************'
    '                            EDF Variables                                  '
    '***************************************************************************'
    Private edfInlet As ProcessStream
    Private edfPermeate As ProcessStream
    Private edfRetentate As ProcessStream
    Private edfPermIn As ProcessStream
    Private edfThick As InternalRealVariable
    Private edfDiam As InternalRealVariable
    Private edfLen As InternalRealVariable
    Private edfAperm As InternalRealVariable
    Private edfNtubes As InternalRealVariable
    Private edfk As InternalRealVariable
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
    Private edfPressDrop As InternalRealVariable
    Private edfPermPressDrop As InternalRealVariable
    'Private edfDiamExt As InternalRealVariable
    'Private edfLengthPos As InternalRealFlexVariable
    'Private edfCompT As InternalRealFlexVariable
    'Private edfCompH As InternalRealFlexVariable
    Private edfNpoints As InternalRealVariable

    '***************************************************************************'
    '                            Physical Variables                             '
    '***************************************************************************'
    ' Geometry
    Dim L As Double, thick As Double, Din As Double, Aperm As Double
    ' Volumes
    Private Volume As Double, Area As Double
    ' Streams
    Dim fluidList() As Fluid
    Dim StreamList() As ProcessStream
    ' Constants
    Private Const R As Double = 8.314472  ' [kJ/kmol-K]
    Private Const MOLFLOW_UNITS As String = "kgmole/s"

    '***************************************************************************'
    '                           OLD - to be deleted                             '
    '***************************************************************************'
    Private pressureRFV As InternalRealFlexVariable
    Private flowRFV As InternalRealFlexVariable
    Private NumberOfPoints As InternalRealVariable
    Dim myPlotName As InternalTextVariable
    Dim myPlot As TwoDimensionalPlot


    Public Function Initialize(ByRef Container As ExtnUnitOperationContainer, ByRef IsRecalling As Boolean) As Integer
        Dim IRV As InternalRealVariable
        On Error GoTo ErrorTrap
        ' Initialize container
        myContainer = Container
        ' Initialize EDF variables
        Call PointedEDFVariables()
        ' Recall check
        If Not IsRecalling Then
            'Step 3 - set the NumberOfPoints variable to 0.
            NumberOfPoints.Value = 0
        Else
            ' Visibility controllers
            IRV = myContainer.FindVariable("Design_enum").Variable
            IRV.Value = 0
            'Set the initial num value: Karlsruhe experiment for WDS and TEP (Welte, 2009)
            thick = 0.0001              'm
            edfThick.SetValue(thick)
            Din = 0.0023                'm
            edfDiam.SetValue(Din)
            edfNtubes.SetValue(183)      '-
            edfLen.SetValue(0.9)         'm
            L = edfLen.GetValue() * edfNtubes.GetValue()
            Dim Ravg As Double
            Ravg = ((Din + (Din + 2 * thick)) / 2) / 2
            Aperm = L * (2 * Math.PI * Ravg)
            edfAperm.SetValue(Aperm)     'm2
        End If
        ' Global variables: first calculation
        thick = edfThick.GetValue()
        Din = edfDiam.GetValue()
        Aperm = edfAperm.GetValue()
        Area = Math.PI * (Din ^ 2) / 4
        L = edfLen.GetValue() * edfNtubes.GetValue()
        Volume = L * Area
        Call CreatePlot()
        ' Loop for setting index for components of interest (based on Inlet's basis manager)
        IniCompIndex()
        ' Return Initialize
        Initialize = CurrentExtensionVersion_enum.extnCurrentVersion
        Exit Function

ErrorTrap:
        MsgBox("Initialize Error")
    End Function

    Public Sub Execute(ByRef Forgetting As Boolean)
        ' execute gets hit twice, once on a forgetting pass and then on _
        'a calculate pass
        Dim flashPerm As Double, flashNoPerm As Double
        Dim Tin As Double, retPressure As Double, permPressDrop As Double
        Dim i As Long, nComp As Long
        On Error GoTo ErrorTrap

        ' Step 1 - Forgetting check
        If Forgetting Then Exit Sub

        ' Step 2 - Check that we have enough information to Calculate
        If edfInlet Is Nothing Then Exit Sub
        If edfPermeate Is Nothing Then Exit Sub
        If edfRetentate Is Nothing Then Exit Sub
        If NumberOfPoints.Value <= 1 Then Exit Sub
        If Not edfInlet.Pressure.IsKnown Then Exit Sub
        If Not edfInlet.Temperature.IsKnown And Not edfPermeate.Temperature.IsKnown Then Exit Sub

        ' Step 3 - Build vector of Streams and Fluids (StreamList and fluidList)
        ReDim Preserve StreamList(0 To 2)
        StreamList(0) = edfInlet                ' In
        StreamList(1) = edfRetentate            ' Out non Permeated
        StreamList(2) = edfPermeate             ' Out Permeated
        If edfPermIn IsNot Nothing Then         ' (optional) 2nd stream in
            ReDim Preserve StreamList(0 To 3)
            StreamList(3) = edfPermIn
        End If
        ReDim Preserve fluidList(0 To UBound(StreamList))
        For i = 0 To UBound(StreamList)
            fluidList(i) = StreamList(i).DuplicateFluid
        Next i

        ' Step 4 - Calculate Permeation
        nComp = UBound(edfInlet.ComponentMolarFlowValue) + 1
        Dim permMolarFlows() As Double
        ReDim permMolarFlows(nComp - 1)
        permMolarFlows = Permeation() ' kmol/s

        ' Step 5 - Set the calculated values to the list of fluids and product streams
        SetPermeationFlows(StreamList, permMolarFlows)

        '' Step 6 - Pressure drop and flash specifications
        Tin = edfInlet.TemperatureValue
        retPressure = edfInlet.PressureValue - edfPressDrop.Value
        permPressDrop = edfInlet.PressureValue - edfPermeate.PressureValue
        edfRetentate.Temperature.Calculate(Tin)
        edfPermeate.Temperature.Calculate(Tin)
        edfRetentate.Pressure.Calculate(retPressure)
        edfPermPressDrop.Calculate(permPressDrop)

        ' Step 7 - Final Balance, checks and EDF visualization
        ' Now use the .Balance method of the container object to make Hysys perform a Balance
        ' The 1 parameter means that the first entry in the array is a feed stream and the rest 
        ' are products. Do a total balance so that if temps are specified then we'll do a heat
        myContainer.Balance(BalanceType_enum.btTotalBalance, 1, StreamList)
        ' Composition and Condition functions are for visualizing in the EDF
        ' TODO: Note that the next fluidList is not updated at this point without flashes
        ' TODO: Comp = Composition(fluidList) *********
        ' TODO: cond = Condition(fluidList) ************

        ' Step 8 - If we are here it means we solved the unit properly
        ' Check if the Product streams are completely solved
        If edfPermeate.DuplicateFluid.IsUpToDate And edfRetentate.DuplicateFluid.IsUpToDate Then
            myContainer.SolveComplete()
        End If
        Exit Sub
ErrorTrap:
        ' TODO: you should remove this msgbox since some Execute runs always throw an error
        MsgBox(Err.GetException().ToString)
    End Sub



    Sub StatusQuery(ByRef Status As ObjectStatus)
        On Error GoTo ErrorTrap
        'If the object is ignored then Exit the Subroutine
        If myContainer.ExtensionInterface.IsIgnored Then Exit Sub
        'Error messsages
        If edfInlet Is Nothing Then
            Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 1, "Feed stream is missing")
        End If
        If edfPermeate Is Nothing Then
            Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 2, "Permeate stream is missing")
        End If
        If edfRetentate Is Nothing Then
            Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 3, "Non permeating outlet stream is missing")
        End If
        If edfLen.Value < 0 Or edfDiam.Value < 0 Or edfThick.Value < 0 Or edfk.Value < 0 Or edfNtubes.Value < 0 Or edfAperm.Value < 0 Then
            Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 4, "Physical parameters missing or incorrect")
        End If
        'If flagRet Then
        '    Call Status.AddStatusCondition(slError, 5, "All feed flow is being permeated")
        'End If
        'If flagPerm Then
        '    Call Status.AddStatusCondition(slError, 6, "No species permeating")
        'End If

        '' DELETE
        'If NumberOfPoints.Value <= 1 Then
        '    Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 7, "Not enough PQ Data Points")
        'End If

        'Dim DataIsOK As Boolean
        'Dim I As Short
        'DataIsOK = True
        'For I = 0 To NumberOfPoints.Value - 1
        '    If pressureRFV.Values(I) = EmptyValue_enum.HEmpty Or flowRFV.Values(I) = EmptyValue_enum.HEmpty Then
        '        DataIsOK = False
        '        Exit For
        '    End If
        'Next I
        'If DataIsOK = False Then
        '    Call Status.AddStatusCondition(StatusLevel_enum.slWarning, 4, "PQ Data is incomplete")
        'End If

        ''Check Specs Again
        'Dim specs As Integer
        'specs = 0
        'If Not edfInlet Is Nothing And Not edfPermeate Is Nothing Then
        '    If edfInlet.StdGasFlow.IsKnown Then specs = specs + 1
        '    If edfPermeate.StdGasFlow.IsKnown Then specs = specs + 1
        '    If edfPermeate.Pressure.IsKnown Then specs = specs + 1
        'End If

        ''Step 11 - If specs < 1, give a suitable status message.
        'If specs < 1 Then
        '    Call Status.AddStatusCondition(StatusLevel_enum.slMissingRequiredInformation, 5, "Requires 1 flow or pressure spec")
        'End If

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

        edfThick = Nothing
        edfLen = Nothing
        edfDiam = Nothing
        edfAperm = Nothing
        edfNtubes = Nothing
        edfk = Nothing
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
        edfPressDrop = Nothing
        edfPermPressDrop = Nothing
        'edfLengthPos = Nothing
        'edfCompT = Nothing
        'edfCompH = Nothing
        edfNpoints = Nothing
        myPlotH = Nothing
        myPlotNameH = Nothing
        myPlotT = Nothing
        myPlotNameT = Nothing
        'edfDiamExt = Nothing
    End Sub

    Sub VariableChanged(ByRef Variable As InternalVariableWrapper)
        ' Called when the user modifies any edf variable. It is required to update these variales as needed

        On Error GoTo ErrorTrap
        Dim pressureVT As Object
        Dim flowVT As Object
        Dim I As Short
        Dim Ravg As Double
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
                edfLen = myContainer.FindVariable("Length").Variable
                L = edfLen.Value * edfNtubes.Value                    ' total length of tubes
                ' Recalculation of cross-area and total volume
                Area = Math.PI * (Din ^ 2) / 4
                Volume = L * Area
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' GrafL = LengthVector(edfLen, edfNpoints)
                ' Recalculation of total permeation surface
                Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                edfAperm.SetValue(L * (2 * Math.PI * Ravg))
                Aperm = edfAperm.Value
            Case "Diam"
                edfDiam = myContainer.FindVariable("Diam").Variable
                Din = edfDiam.GetValue()
                ' Recalculation of cross-area and total volume
                Area = Math.PI * (Din ^ 2) / 4
                Volume = edfLen.Value * edfNtubes.Value * Area
                ' Recalculation of total permeation surface
                Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                edfAperm.SetValue(L * (2 * Math.PI * Ravg))
                Aperm = edfAperm.Value
            Case "Thickness"
                edfThick = myContainer.FindVariable("Thickness").Variable
                thick = edfThick.GetValue()
                ' Recalculation of total permeation surface
                Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                edfAperm.SetValue(L * (2 * Math.PI * Ravg))
                Aperm = edfAperm.Value
            Case "Aperm"
                edfAperm = myContainer.FindVariable("Aperm").Variable
                Aperm = edfAperm.Value
                ' Recalculation of total permeation surface
                Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                L = Aperm / (2 * Math.PI * Ravg)                     ' total length of tubes
                edfLen.SetValue(L / edfNtubes.GetValue())
                ' Recalculation of cross-area and total volume
                Area = Math.PI * (Din ^ 2) / 4
                Volume = L * Area
            Case "Ntubes"
                edfNtubes = myContainer.FindVariable("Ntubes").Variable
                ' Recalculation of total length (mantaining length/tube (edfLen))
                Dim len As Double
                L = edfLen.Value * edfNtubes.Value
                ' Recalculation of total permeation surface (last change affects area)
                Ravg = ((Din + (Din + 2 * thick)) / 2) / 2      ' auxiliar average radius between Dint and Dext
                edfAperm.SetValue(L * (2 * Math.PI * Ravg))
                Aperm = edfAperm.Value
                ' Recalculation of cross-area and total volume
                Area = Math.PI * (Din ^ 2) / 4
                Volume = L * Area

            '%% Pressure drop and others
            Case "n"
                'edfn = myContainer.FindVariable("n").Variable
            Case "PressDrop"
                edfPressDrop = myContainer.FindVariable("PressDrop").Variable
            Case "PermPressDrop"
                edfPermPressDrop = myContainer.FindVariable("PermPressDrop").Variable
            Case "k"
                edfk = myContainer.FindVariable("k").Variable
            Case "NumberOfPoints"
                'edfNpoints = myContainer.FindVariable("NumberOfPoints").Variable
                'GrafL = LengthVector(edfLen, edfNpoints)

        End Select
        Exit Sub
ErrorTrap:
        myContainer.Trace(myContainer.name & ": Error in Variable Changed for variable " & Variable.Tag & ".", False)
        MsgBox("Variable Changed Error")
    End Sub

    Function VariableChanging(ByRef Variable As InternalVariableWrapper) As Boolean

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
            edfThick = .FindVariable("Thickness").Variable
            edfLen = .FindVariable("Length").Variable
            edfDiam = .FindVariable("Diam").Variable
            edfAperm = .FindVariable("Aperm").Variable
            edfNtubes = .FindVariable("Ntubes").Variable
            edfk = .FindVariable("k").Variable
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
            edfPressDrop = .FindVariable("PressDrop").Variable
            edfPermPressDrop = .FindVariable("PermPressDrop").Variable
            'edfLengthPos = .FindVariable("LengthPos").Variable
            'edfCompT = .FindVariable("CompT").Variable
            'edfCompH = .FindVariable("CompH").Variable
            edfNpoints = .FindVariable("NumberOfPoints").Variable
            myPlotNameH = .FindVariable("PlotNameH").Variable
            myPlotNameT = .FindVariable("PlotNameT").Variable
            'edfDiamExt = .FindVariable("DiamExtern").Variable
            ''molDensity = .FindVariable("molDensity").Variable
        End With
    End Sub



    '***************************************************************************'
    '                            Main Functions                                 '
    '***************************************************************************'
    Private Function Permeation()
        ' Calculate vector of permeated species in default HYSYS magnitude: "kmol/s"
        '   Ffeed(nComp):       vector      inlet molar flow (per component)
        '   Fperm(nComp):       vector      permeated molar flow in one cell (per component)
        '   FpermTotal(nComp):  vector      aggregates molar flow permeated for each cell (per component)
        '   Fcell(nComp):       vector      similar to "Ffeed" but for each cell calculated from previous cell
        '   FpermCells(Npoints):vector      collect total permeation in each cell for plotting purposes
        '
        '   FpermAtom:          Double(3)   Output of the diffusion part of code (H, D, T mole flows)
        '   FfeedPerm:
        '   PfeedAtom:
        '   PfeedPermAtom:
        '   X:
        '
        '   flagLoop:           Boolean     Used to assert if H, D or T permeation finished and break loop
        '   permIndicesSorted:  Short(3)    Sorts indices to address FpermAtom in ascending order
        '   FpermAtomFlag:      Double(3)
        '   permIndicesHeteroCopy: Short(3,3)

        ' Step 1 - Declarations
        Dim Fperm() As Double, FpermComp() As Double, Ffeed() As Double, Fcell() As Double
        Dim FpermAtom() As Double, X() As Double, PfeedPermAtom() As Double
        Dim Qfeed As Double, Tfeed As Double, Tfeed_K As Double, Pfeed As Double
        Dim PfeedPerm As Double, molFrac As Double
        'Dim dx As Double, Ravg As Double, ApermCell As Double
        Dim i As Long, j As Long, iPerm As Long, nComp As Long, nCell As Long
        nComp = UBound(edfInlet.ComponentMolarFlowValue) + 1
        ReDim Fperm(nComp - 1), FpermComp(nComp - 1), Fcell(nComp - 1)
        ReDim FpermAtom(nPermAtom - 1), X(nPermAtom - 1), PfeedPermAtom(nPerm - 1)

        ' Step 2 - Initializations
        ' First check to see if feed flow is unsuitable
        If edfInlet.MolarFlow.Value <= 0 Then
            ' Return the same product values as feed values
            '''''''RetExtProdMoleFracs = ExtFeedMoleFracs
            '''''''RetProdTotalMoleFlow = FeedTotalMoleFlow
            ''''''' TODO: Complete stuff here: return 0 value for permeation
            Exit Function
        End If

        ' Get EDF parameters into double VB variables
        ' nCell = edfNpoints.GetValue()
        ' ReDim FpermCells(nCell - 1)

        ' Get inlet stream parameters
        Pfeed = edfInlet.Pressure.GetValue("kPa")  ' TODO: check units of permeability to match this pressure's
        Tfeed = edfInlet.Temperature.GetValue("C")
        Tfeed_K = edfInlet.Temperature.GetValue("K")
        Qfeed = edfInlet.ActualVolumeFlowValue
        Ffeed = edfInlet.ComponentMolarFlowValue              ' molar flow per component

        '' Geometric calculations
        'ApermCell = Aperm / nCell                  ' permeation surface per differential cell
        'Fcell = Ffeed

        ' TODO: setup Permeabilities - calculate them here and write them in EDF (as read-only)
        Dim P() As Double
        P = {0.00000000002, 0.000000000012, 0.0000000000095} ' UNITS SHOULD BE: KMOL · m-1 · s -1 · Pa-0.5
        ' (instead of mol · m-1 · s -1 · Pa-0.5)
        ' TODO: MULTIPLY P TO AUTOMATICALLY GET FLOWS OF kmol/s (HYSYS default units)

        ' TODO: Set calculated PERMEABILITY values in edf for user's visualization
        ' edfDT.SetValue(DT)
        ' edfDH.SetValue(DH)

        ' LOOPS: use "i" for atoms (H, D, T) and "j" for molecules (H2, HD, HT, D2, DT, T2)
        ' Get total inlet pressure of permeating species ONLY. Apply Dalton's law
        PfeedPerm = 0  ' p1 in thesis
        For j = 0 To nPerm - 1
            iPerm = permIndices(j)
            PfeedPerm += Pfeed * edfInlet.ComponentMolarFractionValue(iPerm)
        Next

        ' Step 3 - Calculate atomic diffusion flows
        For i = 0 To nPermAtom - 1
            ' Loop through nPerm (should be 6) to get contribution to partial pressure of atom "i"
            PfeedPermAtom(i) = 0
            For j = 0 To nPerm - 1
                molFrac = edfInlet.ComponentMolarFractionValue(permIndices(j))
                ' permCoeffs is 1 or 0.5 depending on molecule being homonuclear or heteronuclear
                PfeedPermAtom(i) += permCoeffs(i, j) * Pfeed * molFrac
            Next
            X(i) = PfeedPermAtom(i) / PfeedPerm
            ' PERMEATION FORMULA: F = P(i) * A / t * (X(i) * sqrt(p_in) - Y(i) * sqrt(p_out))
            ' Assume negligible output pressure (TODO)
            FpermAtom(i) = P(i) * Aperm / thick * (X(i) * Math.Sqrt(PfeedPerm))
        Next

        ' Step 4 - Distribute molecular flow depending on the calculated atomic diffusion flow
        ' Check if we have all permeating species available in the inlet in `isFeedComplete`
        Dim isFeedComplete As Boolean = True
        For i = 0 To nPerm - 1
            iPerm = permIndices(i)
            If Ffeed(iPerm) = 0 Then isFeedComplete = False
        Next
        ' Select heuristic function to assign molecular distribution
        If isFeedComplete Then
            ' OPTION 1: We have all molecules in the feed and there is enough of all of them
            PermeateHeuristic1(FpermComp, FpermAtom)
        Else
            ' OPTION 2: Some molecules in the feed are missing
            PermeateHeuristic2(FpermComp, Ffeed, FpermAtom)
        End If

        ' Step 5 - Last checks
        ' Check that the total permeated flow is the same from both Atom and Molecular sides
        ' Use the rounded relative error to measure the closeness of both sums
        Dim SumComp As Double = Sum(FpermComp)
        Dim SumAtom As Double = Sum(FpermAtom)
        Dim RelError As Double
        RelError = Math.Abs((SumAtom - SumComp) / SumAtom)
        If Math.Round(RelError, 10) > 0 Then  ' 10 is an arbitrary number of decimals 
            MsgBox("Permeation function did not match required atomic permeation (FpermAtom) with" _
                   + "the actual output (FpermComp)")
        End If

        ' LOOP over cells
        'For i = 0 To nCell - 1
        '    ' Calculate concentrations [kgmole/m3]
        '    CHTsi = Fcell(idxHT) / Qfeed
        '    CHsi = Fcell(idxH2) / Qfeed
        '    CTsi = Fcell(idxT2) / Qfeed
        '    CTso = 0
        '    CHso = 0
        '    CHTso = 0
        '    If CTi = -32767 Then CTi = 0
        '    If CTo = -32767 Then CTo = 0
        '    If CHTi = -32767 Then CHTi = 0
        '    If CHTo = -32767 Then CHTo = 0
        '    If CHi = -32767 Then CHi = 0
        '    If CHo = -32767 Then CHo = 0
        '    ' Permeation calculation [at/s]. Richardson's law: [kmol/s�m2] -> multiply by "ApermCell" -> [kmol/s]
        '    Fperm(idxT2) = (DT / thick) * ((CHTsi + CTsi) - (CTso + CHTso)) * ApermCell
        '    Fperm(idxH2) = (DH / thick) * ((CHTsi + CHsi) - (CHso + CHTso)) * ApermCell
        '    '   Old calc. (pi * L * DH / Log(1 + (thick / (Din / 2))) * ((CHTsi + CHsi) - (CHso + CHTso)))
        '    ' Molar flow component vector in next cell
        '    Fcell = vectorSubtractIf(Fcell, Fperm)      ' Special function for dealing with undesired negarive permeate flow
        '    FpermTotal = vectorSum(Fperm, FpermTotal)   ' Aggregate permeation from previous cells
        '    FpermCells(i) = sumVectorElements(Fperm)
        'Next i

        Permeation = FpermComp
    End Function

    Private Sub PermeateHeuristic1(FpermComp() As Double, FpermAtom() As Double)
        ' Assign molecular molar flow (FpermComp) based on previosly calculated diffusion atomic
        ' flow (FpermAtom)
        ' OPTION 1: Case in which we have H, D and T flow, now calculate H2, HD, HT, D2, DT, T2 flows
        Dim i As Long, j As Long, iPerm As Long
        For j = 0 To nPerm - 1
            iPerm = permIndices(j)
            For i = 0 To nPermAtom - 1
                ' Calculate contribution of each species
                ' Divide by 2 means e.g. 1 (H2) + 0.5 (HD) + 0.5 (HT) = 2
                FpermComp(iPerm) += FpermAtom(i) * (permCoeffs(i, j) / 2)
            Next
        Next
        ' TODO: Check if the calculated permeation is greater than the inlet to warn the user as we to in PermeateHeuristic2
        '       and design and check extreme cases where this happens
    End Sub

    Private Sub PermeateHeuristic2(FpermComp() As Double, Ffeed() As Double, FpermAtom() As Double)
        ' Assign molecular molar flow (FpermComp) based on previosly calculated diffusion atomic
        ' flow (FpermAtom)
        ' OPTION 2: not all 6 species are available in the inlet

        Dim i As Long
        ' Loop in ascending order of permeation flow rate per atom
        Dim flagLoop As Boolean
        Dim permIndicesSorted As Short(), iAtom As Short, k As Short, iMolec As Short
        ' We will subtract the flows we go assigning from this array
        Dim FpermAtomFlag As Double(), permIndicesHeteroCopy As Short(,)
        ReDim FpermAtomFlag(nPermAtom - 1), permIndicesHeteroCopy(nPermAtom - 1, nPermAtom - 1)
        Array.Copy(FpermAtom, FpermAtomFlag, nPermAtom)
        Array.Copy(permIndicesHetero, permIndicesHeteroCopy, 9)

        ' Sort flows in ascending order and return the indices of the ordered list
        permIndicesSorted = SortIndices(FpermAtom)

        ' Heteronuclears loop first
        For i = 0 To nPermAtom - 1
            iAtom = permIndicesSorted(i)
            flagLoop = False
            For k = 0 To nHeteroNuclear - 1
                ' With heteronuclears we need double the flow rate bc they contribute by half
                iMolec = permIndicesHeteroCopy(iAtom, k)
                ' Continue to next iteration if element is -1
                If iMolec < 0 Then Continue For
                ' First check if there is feed flow rate at all for "iMolec"
                If Ffeed(iMolec) > 0 Then
                    If FpermAtomFlag(iAtom) >= 2 * Ffeed(iMolec) Then
                        ' We exhaust all inlet flow of this molecule in permeation
                        FpermComp(iMolec) += 2 * Ffeed(iMolec)
                    Else
                        ' Permeation of atom "i" is finished
                        FpermComp(iMolec) += 2 * FpermAtomFlag(iAtom)
                        flagLoop = True
                    End If
                    ' Subtract flow from our flag permeation array for both atoms affected
                    ' without the x2!
                    FpermAtomFlag(iAtom) -= FpermComp(iMolec) / 2   ' current atom
                    FpermAtomFlag(k) -= FpermComp(iMolec) / 2       ' the other atom affected ' TODO: this could become negative!
                    ' Remove molecule from matrix to avoid passing through it twice
                    ' (since permIndicesHetero is upper diagonal)
                    permIndicesHeteroCopy(iAtom, k) = -1
                    permIndicesHeteroCopy(k, iAtom) = -1
                    ' Exit For loop in case we finished 'iAtom' permeation
                    If flagLoop Then
                        Exit For
                    End If
                End If
            Next
        Next

        ' Homonuclears loop last (simpler loop)
        For i = 0 To nPermAtom - 1
            iAtom = permIndicesSorted(i)
            iMolec = permIndicesHomo(iAtom)
            ' First check if there is feed flow rate at all for "iMolec"
            If Ffeed(iMolec) > 0 Then
                If FpermAtomFlag(iAtom) >= Ffeed(iMolec) Then
                    ' We exhaust all inlet flow of this molecule in permeation
                    FpermComp(iMolec) += Ffeed(iMolec)
                Else
                    ' Permeation of atom "i" is finished
                    FpermComp(iMolec) += FpermAtomFlag(iAtom)
                End If
                ' Subtract flow from our flag permeation array
                FpermAtomFlag(iAtom) -= FpermComp(iMolec)
            End If
        Next

        ' At this point, FpermAtomFlag should be empty
        If Not IsZero(FpermAtomFlag) Then
            MsgBox("Permeation function did not exhaust FpermAtom completely")
        End If
    End Sub



    '***************************************************************************'
    '                         Auxiliary Functions                               '
    '***************************************************************************'
    Private Sub IniCompIndex()
        ' Loop for setting index for components that can permeate (based on Inlet's basis manager)
        Dim ComponentList As Components
        Dim isH As Short, isD As Short, isT As Short
        Dim i As Long, nComp As Long
        ' Init the array of the indices of permeating species to -1
        permIndices = {-1, -1, -1, -1, -1, -1}
        ' Init component list
        If edfInlet Is Nothing Then
            ComponentList = myContainer.Flowsheet.FluidPackage.Components
        Else
            ComponentList = edfInlet.DuplicateFluid.Components
        End If
        nComp = ComponentList.Count         ' Number of Components
        ' Loop over component list and search for hydrogen isotopes
        For i = 0 To nComp - 1
            Select Case ComponentList.Item(i).Name
                Case "Hydrogen"
                    permIndices(0) = i
                    isH = 1
                Case "Hydrogen*"
                    permIndices(0) = i
                    isH = 1
                Case "HD*"
                    permIndices(1) = i
                    isH = 1
                    isD = 1
                Case "HT*"
                    permIndices(2) = i
                    isH = 1
                    isT = 1
                Case "Deuterium*"
                    permIndices(3) = i
                    isD = 1
                Case "DT*"
                    permIndices(4) = i
                    isD = 1
                    isT = 1
                Case "Tritium*"
                    permIndices(5) = i
                    isT = 1
            End Select
        Next i
        permIndicesHetero = {
            {-1, permIndices(1), permIndices(2)},
            {permIndices(1), -1, permIndices(4)},
            {permIndices(2), permIndices(4), -1}
        }
        permIndicesHomo = {permIndices(0), permIndices(3), permIndices(5)}
        nPermAtom = isH + isD + isT
    End Sub
    Private Sub SetPermeationFlows(streams As ProcessStream(), permMolarFlows As Double())
        ' Updates vector of streams (vector of 3 or 4 Fluids) and,
        ' specifically, the products: Retentate (index 1) and Permeate
        ' (index 2), with the new permeated composition.
        '
        ' To do this, updates component molar flows AND total molar flow.
        '
        Dim inletMolarFlows() As Double, retMolarFlows() As Double
        Dim i As Long, nComp As Long
        nComp = UBound(permMolarFlows) + 1
        ' Calculate retentate component flows
        inletMolarFlows = streams(0).ComponentMolarFlowValue
        retMolarFlows = SubtractVectorsIf(inletMolarFlows, permMolarFlows)
        ' Set the fictitious molar flow vectors to the actual ones
        streams(1).ComponentMolarFlow.Calculate(retMolarFlows, MOLFLOW_UNITS)  ' Retentate
        streams(2).ComponentMolarFlow.Calculate(permMolarFlows, MOLFLOW_UNITS) ' Permeate
        ' Set total molar flow to each fluid
        Dim totalRetMolarFLow As Double = 0
        Dim totalPermMolarFlow As Double = 0
        For i = 0 To nComp - 1
            totalRetMolarFLow += retMolarFlows(i)
            totalPermMolarFlow += permMolarFlows(i)
        Next
        streams(1).MolarFlow.Calculate(totalRetMolarFLow, MOLFLOW_UNITS)       ' Retentate
        streams(2).MolarFlow.Calculate(totalPermMolarFlow, MOLFLOW_UNITS)      ' Permeate
    End Sub
    Private Function LengthVector(L, n) As Double()
        ' Builds a vector of equidistant elements representing the position [m] of each cell
        Dim i As Long
        Dim lenvec() As Double, dx As Double
        ReDim lenvec(n - 1)
        dx = L / n
        lenvec(0) = dx
        For i = 1 To n - 1
            lenvec(i) = lenvec(i - 1) + dx
        Next i
        LengthVector = lenvec
    End Function
    Private Function Sum(myArray() As Double) As Double
        Dim i As Long
        For i = 0 To UBound(myArray)
            Sum += myArray(i)
        Next
    End Function
    Private Function SubtractVectorsIf(ByRef A1, ByRef A2) As Double()
        ' Subtract 1D arrays avoiding negative outputs.
        '
        ' e.g.  we want to calculate the operation Fret = Finlet - Fperm
        '       in a per-component basis.
        '       Thus, A1 = Finlet, A2 = Fperm
        '       If Finlet = (4) and Fperm = (10); Fret = (-6)
        '       But avoiding negative numbers, the result we should get
        '       is Fperm = (4) and Fret = (0)
        '
        ' To do this, we use ByRef arguments and modify A2 (here Fperm)
        ' if necessary.
        Dim n As Long, i As Long
        Dim A() As Double
        ' Get sizes
        n = UBound(A1, 1) + 1
        ReDim A(n - 1)
        ' Check condition that allow sum/subtract
        If (UBound(A2, 1) + 1) <> n Then
            MsgBox("Sizes do not match in array sum function.")
            Exit Function
        End If
        ' Do calculation
        For i = 0 To n - 1
            If A1(i) >= A2(i) Then
                A(i) = A1(i) - A2(i)
            Else
                ' Molecule i permeated till its limits
                A(i) = 0  ' no flow in retentate
                A2(i) = A1(i)  ' adjust permeate flow to the exact input
            End If
        Next i
        SubtractVectorsIf = A
    End Function
    Private Function IsZero(myArray As Double()) As Boolean
        ' Check if the input 1D array is full of zeros
        Dim i As Long
        IsZero = True
        For i = 0 To UBound(myArray)
            If myArray(i) <> 0 Then
                IsZero = False
            End If
        Next
    End Function
    '    Function LinearInterpolation(ByRef xDataRFV As InternalRealFlexVariable, ByRef yDataRFV As InternalRealFlexVariable, ByRef xPoint As RealVariable) As Double
    '        'This method linear interpolates to find the y point that coresponds to the known
    '        'x point for the given x and y data sets.

    '        Dim xData As Object
    '        Dim yData As Object
    '        Dim x As Double
    '        Dim y As Double

    '        On Error GoTo ErrorTrap

    '        Dim High As Integer
    '        Dim Low As Integer
    '        Dim number As Integer

    '        y = EmptyValue_enum.HEmpty
    '        LinearInterpolation = y

    '        'UPGRADE_WARNING: Couldn't resolve default property of object xDataRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'UPGRADE_WARNING: Couldn't resolve default property of object xData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        xData = xDataRFV.Values
    '        'UPGRADE_WARNING: Couldn't resolve default property of object yDataRFV.Values. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'UPGRADE_WARNING: Couldn't resolve default property of object yData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        yData = yDataRFV.Values
    '        x = xPoint.Value

    '        High = UBound(xData)
    '        Low = LBound(xData)
    '        number = High - Low + 1

    '        'There must be more than 1 data point to Linearly Interpolate
    '        If number <= 1 Then Exit Function
    '        'Check that the x and y Data have the same bounds
    '        If High <> UBound(yData) Or Low <> LBound(yData) Then Exit Function
    '        'Sort the x Data from low to high
    '        Call Sort(xData, yData)

    '        'Check to see that the x point is within the x Data Range
    '        'UPGRADE_WARNING: Couldn't resolve default property of object xData(High). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'UPGRADE_WARNING: Couldn't resolve default property of object xData(Low). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        If x < xData(Low) Or x > xData(High) Then
    '            MsgBox("The point is outside the data range")
    '            Exit Function
    '        End If

    '        Dim I As Short
    '        'Search the data until the x Point is between two x Data points
    '        For I = Low To High
    '            'UPGRADE_WARNING: Couldn't resolve default property of object xData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            If x < xData(I) Then
    '                'UPGRADE_WARNING: Couldn't resolve default property of object yData(I - 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                'UPGRADE_WARNING: Couldn't resolve default property of object yData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                'UPGRADE_WARNING: Couldn't resolve default property of object xData(I - 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                'UPGRADE_WARNING: Couldn't resolve default property of object xData(I). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                'UPGRADE_WARNING: Couldn't resolve default property of object yData(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '                y = yData(I) - ((xData(I) - x) / (xData(I) - xData(I - 1)) * (yData(I) - yData(I - 1)))
    '                Exit For
    '            End If
    '        Next I
    '        'UPGRADE_WARNING: Couldn't resolve default property of object xData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        xDataRFV.Values = xData
    '        'UPGRADE_WARNING: Couldn't resolve default property of object yData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        yDataRFV.Values = yData
    '        LinearInterpolation = y
    '        Exit Function

    'ErrorTrap:
    '        MsgBox("Interpolation Error")
    '    End Function

    Private Function SortIndices(ByVal SrcArray As Double())

        'Description: Sorts the arrays passed so that smallest values occur first in KeyArray()
        '             does the same rearrangements on OtherArray() so values still correspond
        '             - Uses a Ripple type sort (Good for smallish data sets)
        '
        On Error GoTo ErrorTrap
        'Declare Variables------------------------------------------------------------------------------

        Dim I As Short
        Dim J As Short 'Counters
        Dim Temp As Short 'used to swap values
        Dim TempIndices As Short() = {0, 1, 2}

        'Procedure--------------------------------------------------------------------------------------
        For I = 0 To UBound(SrcArray) - 1
            For J = I + 1 To UBound(SrcArray)
                If SrcArray(J) < SrcArray(I) Then
                    ' Swap data
                    'Temp = SortedArray(J)
                    'SortedArray(J) = SortedArray(I)
                    'SortedArray(I) = Temp
                    ' Swap indices
                    Temp = TempIndices(J)
                    TempIndices(J) = TempIndices(I)
                    TempIndices(I) = Temp
                End If
            Next  'J
        Next  'I
        SortIndices = TempIndices
        Exit Function

ErrorTrap:
        MsgBox("Sorting Error")
    End Function

    Sub CreatePlot()
        If myPlotH IsNot Nothing Then
            myContainer.DeletePlot("LCPlotH")
            myPlotH = Nothing
        End If
        myContainer.BuildPlot2("LCPlotH", myPlotH, HPlotType_enum.hptTwoDimensionalPlot)
        myPlotNameH.Value = "LCPlotH"
        With myPlotH
            .TitleData = "Length vs Non Permeation"
            .SetAxisLabelData(HAxisType_enum.hatXAxis, "Length (m)")
            .SetAxisLabelData(HAxisType_enum.hatYAxis, "NoPerm Flow")
            .SetAxisLabelVisible(HAxisType_enum.hatXAxis, True)
            .SetAxisLabelVisible(HAxisType_enum.hatYAxis, True)
            .LegendVisible = True
            .CreateXYDataSet(1, "PQData")
            ''''''''''''''''''''''''''''''''''' .SetDataSetXData(1, edfLengthPos)
            ''''''''''''''''''''''''''''''''''' .SetDataSetYData(1, edfCompH)
            .SetDataSetColour(1, "Red")
        End With
        If myPlotT IsNot Nothing Then
            myContainer.DeletePlot("LCPlotT")
            myPlotT = Nothing
        End If
        myContainer.BuildPlot2("LCPlotT", myPlotT, HPlotType_enum.hptTwoDimensionalPlot)
        myPlotNameT.Value = "LCPlotT"
        With myPlotT
            .TitleData = "Length vs Permeation"

            .SetAxisLabelData(HAxisType_enum.hatXAxis, "Length (m)")
            .SetAxisLabelData(HAxisType_enum.hatYAxis, "Perm Flow")
            .SetAxisLabelVisible(HAxisType_enum.hatXAxis, True)
            .SetAxisLabelVisible(HAxisType_enum.hatYAxis, True)

            .LegendVisible = True

            .CreateXYDataSet(1, "PQData")
            '''''''''''''''''''''''''''''''''''.SetDataSetXData(1, edfLengthPos)
            '''''''''''''''''''''''''''''''''''.SetDataSetYData(1, edfCompT)
            .SetDataSetColour(1, "Red")
        End With
    End Sub
End Class