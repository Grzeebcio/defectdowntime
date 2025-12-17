Option Explicit

Private mHasRO As Boolean
Private mHasST As Boolean
Private mHasPrasa As Boolean
Private mHasProces As Boolean
Private mSelectedROST As String
Private mSelectedPrasaProces As String

' Initializes ComboBoxLinia with line headers from row 1 of the cfg_projekt sheet
' and loads ComboBoxProjekt with the projects under the currently selected line.
Private Sub UserForm_Initialize()
    InitializeLiniaIProjekty
    InitializeBrygadaIZmiana
    HideAllProblems
End Sub

' Shows the form modelessly so Excel stays interactive.
Public Sub ShowModeless()
    Me.Show vbModeless
End Sub

' Fills the line and project combo boxes when the form opens.
Private Sub InitializeLiniaIProjekty()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_projekt")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    With Me.ComboBoxLinia
        .Clear
        For col = 1 To lastCol
            If Trim$(ws.Cells(1, col).Value) <> "" Then
                .AddItem CStr(ws.Cells(1, col).Value)
            End If
        Next col
    End With

    ' Preload projects for the first available line
    If Me.ComboBoxLinia.ListCount > 0 Then
        Me.ComboBoxLinia.Value = Me.ComboBoxLinia.List(0)
        LoadProjectsForLine Me.ComboBoxLinia.Value
        LoadOperatorsForLine Me.ComboBoxLinia.Value
        UpdateButtonVisibility
    End If
End Sub

' Populates ComboBoxBrygada and ComboBoxZmiana with static choices.
Private Sub InitializeBrygadaIZmiana()
    With Me.ComboBoxBrygada
        .Clear
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
        .AddItem "D"
    End With

    With Me.ComboBoxZmiana
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
    End With
End Sub

' Populates ComboBoxProjekt with the projects found in the column for the given line.
Public Sub LoadProjectsForLine(ByVal linia As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_projekt")

    Dim colIndex As Variant
    colIndex = Application.Match(linia, ws.Rows(1), 0)
    If IsError(colIndex) Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(colIndex)).End(xlUp).Row

    Dim row As Long
    With Me.ComboBoxProjekt
        .Clear
        For row = 2 To lastRow
            If Trim$(ws.Cells(row, CLng(colIndex)).Value) <> "" Then
                .AddItem CStr(ws.Cells(row, CLng(colIndex)).Value)
            End If
        Next row
    End With
End Sub

' Event handler to keep the project list in sync when the line selection changes.
Private Sub ComboBoxLinia_Change()
    LoadProjectsForLine Me.ComboBoxLinia.Value
    LoadOperatorsForLine Me.ComboBoxLinia.Value
    UpdateButtonVisibility
    HideAllProblems
End Sub

' Refreshes button visibility based on the selected line and project.
Private Sub ComboBoxProjekt_Change()
    UpdateButtonVisibility
    HideAllProblems
End Sub

' Resets all process buttons to hidden and default styling.
Private Sub HideAllProcessButtons()
    mSelectedROST = ""
    mSelectedPrasaProces = ""
    ResetButtonStyle Me.CommandButtonRO
    ResetButtonStyle Me.CommandButtonST
    ResetButtonStyle Me.CommandButtonPrasa
    ResetButtonStyle Me.CommandButtonProces

    Me.CommandButtonRO.Visible = False
    Me.CommandButtonST.Visible = False
    Me.CommandButtonPrasa.Visible = False
    Me.CommandButtonProces.Visible = False
End Sub

' Applies availability rules from cfg_dostepnosc for the chosen line and project.
Private Sub UpdateButtonVisibility()
    HideAllProcessButtons

    Dim linia As String: linia = Trim$(Me.ComboBoxLinia.Value)
    If linia = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_dostepnosc")

    Dim colLinia As Variant, colProjekt As Variant
    Dim colRO As Variant, colST As Variant, colPrasa As Variant, colProces As Variant

    colLinia = Application.Match("Linia", ws.Rows(1), 0)
    colProjekt = Application.Match("Projekt", ws.Rows(1), 0) ' optional
    colRO = Application.Match("RO", ws.Rows(1), 0)
    colST = Application.Match("ST", ws.Rows(1), 0)
    colPrasa = Application.Match("Prasa", ws.Rows(1), 0)
    colProces = Application.Match("Proces", ws.Rows(1), 0)

    If IsError(colLinia) Or IsError(colRO) Or IsError(colST) Or _
       IsError(colPrasa) Or IsError(colProces) Then
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(colLinia)).End(xlUp).Row

    mHasRO = False
    mHasST = False
    mHasPrasa = False
    mHasProces = False

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If Trim$(ws.Cells(rowIndex, CLng(colLinia)).Value) = linia Then
            If Not IsError(colProjekt) Then
                Dim projValue As String
                projValue = Trim$(ws.Cells(rowIndex, CLng(colProjekt)).Value)
                If projValue <> "" And projValue <> Trim$(Me.ComboBoxProjekt.Value) Then
                    GoTo ContinueNext
                End If
            End If

            mHasRO = (Val(ws.Cells(rowIndex, CLng(colRO)).Value) = 1)
            mHasST = (Val(ws.Cells(rowIndex, CLng(colST)).Value) = 1)
            mHasPrasa = (Val(ws.Cells(rowIndex, CLng(colPrasa)).Value) = 1)
            mHasProces = (Val(ws.Cells(rowIndex, CLng(colProces)).Value) = 1)

            ShowIfAvailable Me.CommandButtonRO, mHasRO
            ShowIfAvailable Me.CommandButtonST, mHasST
            ShowIfAvailable Me.CommandButtonPrasa, mHasPrasa
            ShowIfAvailable Me.CommandButtonProces, mHasProces
            Exit For
        End If
ContinueNext:
    Next rowIndex
End Sub

' Sets button visible when the availability flag is 1 and resets its style.
Private Sub ShowIfAvailable(ByVal btn As MSForms.CommandButton, ByVal flagValue As Variant)
    btn.Visible = CBool(flagValue)
    If btn.Visible Then
        ResetButtonStyle btn
    End If
End Sub

' Loads operator choices for ComboBoxOp1-ComboBoxOp4 based on the selected line.
Private Sub LoadOperatorsForLine(ByVal linia As String)
    ClearOperatorCombos

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_operator")

    Dim colIndex As Variant
    colIndex = Application.Match(linia, ws.Rows(1), 0)
    If IsError(colIndex) Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(colIndex)).End(xlUp).Row

    Dim ops As Collection
    Set ops = New Collection

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim opValue As String
        opValue = Trim$(ws.Cells(rowIndex, CLng(colIndex)).Value)
        If opValue <> "" Then
            ops.Add opValue
        End If
    Next rowIndex

    If ops.Count = 0 Then Exit Sub

    SetComboOptions Me.ComboBoxOp1, ops
    SetComboOptions Me.ComboBoxOp2, ops
    SetComboOptions Me.ComboBoxOp3, ops
    SetComboOptions Me.ComboBoxOp4, ops
End Sub

' Clears operator combo boxes.
Private Sub ClearOperatorCombos()
    Me.ComboBoxOp1.Clear
    Me.ComboBoxOp2.Clear
    Me.ComboBoxOp3.Clear
    Me.ComboBoxOp4.Clear
End Sub

' Populates a combo box with operator names.
Private Sub SetComboOptions(ByVal comboBox As MSForms.ComboBox, ByVal operators As Collection)
    Dim idx As Long

    comboBox.Clear
    For idx = 1 To operators.Count
        comboBox.AddItem operators(idx)
    Next idx
End Sub

' Restores default appearance for a button.
Private Sub ResetButtonStyle(ByVal btn As MSForms.CommandButton)
    btn.BackColor = vbButtonFace
End Sub

' Highlights within the RO/ST pair without clearing Prasa/Proces selection.
Private Sub HighlightROST(ByVal selectedButton As MSForms.CommandButton)
    ResetButtonStyle Me.CommandButtonRO
    ResetButtonStyle Me.CommandButtonST
    selectedButton.BackColor = RGB(0, 176, 80)
End Sub

' Highlights within the Prasa/Proces pair without clearing RO/ST selection.
Private Sub HighlightPrasaProces(ByVal selectedButton As MSForms.CommandButton)
    ResetButtonStyle Me.CommandButtonPrasa
    ResetButtonStyle Me.CommandButtonProces
    selectedButton.BackColor = RGB(0, 176, 80)
End Sub

' Validates the date entered in TextBoxDay using the DD.MM.RRRR format.
Private Sub ValidateDayInput()
    Dim rawValue As String
    rawValue = Trim$(Me.TextBoxDay.Value)

    If rawValue = "" Then
        Me.TextBoxDay.BackColor = vbWhite
        Exit Sub
    End If

    If Not rawValue Like "##.##.####" Then
        Me.TextBoxDay.BackColor = RGB(255, 0, 0)
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(rawValue, ".")

    If UBound(parts) <> 2 Then
        Me.TextBoxDay.BackColor = RGB(255, 0, 0)
        Exit Sub
    End If

    On Error GoTo InvalidDate
    Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    Dim parsedDate As Date

    dayPart = CInt(parts(0))
    monthPart = CInt(parts(1))
    yearPart = CInt(parts(2))

    parsedDate = DateSerial(yearPart, monthPart, dayPart)

    If Format$(parsedDate, "dd.mm.yyyy") = rawValue Then
        Me.TextBoxDay.BackColor = vbWhite
        Exit Sub
    End If

InvalidDate:
    Me.TextBoxDay.BackColor = RGB(255, 0, 0)
End Sub

' Safely checks for the presence of a control by name on the form.
Private Function ControlExists(ByVal controlName As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = Me.Controls(controlName)
    ControlExists = (Err.Number = 0)
    Err.Clear
End Function

Private Sub CommandButtonRO_Click()
    HighlightROST Me.CommandButtonRO
    mSelectedROST = "RO"
    If mHasRO And mHasST Then
        Me.CommandButtonST.Visible = False
    End If
    UpdateProblems
End Sub

Private Sub CommandButtonST_Click()
    HighlightROST Me.CommandButtonST
    mSelectedROST = "ST"
    If mHasRO And mHasST Then
        Me.CommandButtonRO.Visible = False
    End If
    UpdateProblems
End Sub

Private Sub CommandButtonPrasa_Click()
    HighlightPrasaProces Me.CommandButtonPrasa
    mSelectedPrasaProces = "Prasa"
    If mHasPrasa And mHasProces Then
        Me.CommandButtonProces.Visible = False
    End If
    UpdateProblems
End Sub

Private Sub CommandButtonProces_Click()
    HighlightPrasaProces Me.CommandButtonProces
    mSelectedPrasaProces = "Proces"
    If mHasPrasa And mHasProces Then
        Me.CommandButtonPrasa.Visible = False
    End If
    UpdateProblems
End Sub

Private Sub CommandButtonSave_Click()
    If Not ValidateRequiredInputs() Then Exit Sub
    SaveFormData
End Sub

Private Sub CommandButtonPostoj_Click()
    If Not ValidateRequiredInputs() Then Exit Sub
    If Not TryShowForm("UserFormAwarie") Then
        MsgBox "Nie można otworzyć formularza UserFormAwarie.", vbExclamation
    End If
End Sub

' Saves all form entries into the "data" sheet. Columns B, C, and D are cleared on
' each save and used for shifts 1, 2, and 3 respectively; column I is cleared and
' reused for totals. Rows are matched by stable labels in column A, creating new
' rows when labels are missing so data always aligns to the same descriptors.
Private Sub SaveFormData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("data")

    ' Ensure a valid shift and compute its target column.
    Dim shiftVal As Long
    shiftVal = CLng(Val(Me.ComboBoxZmiana.Value))
    If shiftVal < 1 Or shiftVal > 3 Then
        MsgBox "Wybierz zmianę 1, 2 lub 3 przed zapisem.", vbExclamation
        Exit Sub
    End If

    Dim targetCol As Long
    targetCol = 1 + shiftVal ' 1->B, 2->C, 3->D

    If Not EnsureProcessSelections() Then Exit Sub

    ' Clear previous entries so each save starts from a clean slate.
    ws.Columns("B:D").ClearContents
    ws.Columns("I").ClearContents

    ' Core identifiers and selections.
    WriteField ws, "Data", Me.TextBoxDay.Value, targetCol
    WriteField ws, "Linia", Me.ComboBoxLinia.Value, targetCol
    WriteField ws, "Projekt", Me.ComboBoxProjekt.Value, targetCol
    WriteField ws, "Brygada", Me.ComboBoxBrygada.Value, targetCol
    WriteField ws, "Zmiana", Me.ComboBoxZmiana.Value, targetCol
    WriteField ws, "RO/ST", mSelectedROST, targetCol
    WriteField ws, "Prasa/Proces", mSelectedPrasaProces, targetCol

    ' Operators.
    WriteField ws, "Operator 1", Me.ComboBoxOp1.Value, targetCol
    WriteField ws, "Operator 2", Me.ComboBoxOp2.Value, targetCol
    WriteField ws, "Operator 3", Me.ComboBoxOp3.Value, targetCol
    WriteField ws, "Operator 4", Me.ComboBoxOp4.Value, targetCol

    ' Plan and hourly plan.
    WriteField ws, "Plan", Me.TextBoxPlan.Value, targetCol

    Dim idx As Long
    For idx = 1 To 8
        WriteField ws, "Plan H" & idx, Me.Controls("TextBoxP" & idx).Value, targetCol
    Next idx

    ' Hourly execution and total.
    For idx = 1 To 8
        WriteField ws, "Wykonanie H" & idx, Me.Controls("TextBoxH" & idx).Value, targetCol
    Next idx

    WriteField ws, "Suma wykonania", Me.TextBoxSum.Value, targetCol, 9 ' Column I

    SaveProblemEntries ws, targetCol

    Dim wsReport As Worksheet
    Set wsReport = GetReportSheet(Trim$(Me.ComboBoxLinia.Value), shiftVal, _
                                  mSelectedROST, mSelectedPrasaProces)

    If Not wsReport Is Nothing Then
        UpdateReportValues wsReport, targetCol
        wsReport.Activate
    End If
End Sub

' Verifies required selections and inputs before running save or downtime actions.
Private Function ValidateRequiredInputs() As Boolean
    Dim missing As Collection
    Set missing = New Collection

    If Trim$(Me.ComboBoxLinia.Value) = "" Then missing.Add "linia"
    If Trim$(Me.ComboBoxProjekt.Value) = "" Then missing.Add "projekt"
    If Trim$(Me.ComboBoxBrygada.Value) = "" Then missing.Add "brygada"
    If Trim$(Me.ComboBoxZmiana.Value) = "" Then missing.Add "zmiana"
    If Trim$(Me.TextBoxPlan.Value) = "" Then missing.Add "plan"
    If Trim$(Me.TextBoxSum.Value) = "" Then missing.Add "realizacja"

    If missing.Count > 0 Then
        Dim parts() As String
        ReDim parts(0 To missing.Count - 1)

        Dim idx As Long
        For idx = 1 To missing.Count
            parts(idx - 1) = missing(idx)
        Next idx

        MsgBox "Uzupełnij pola: " & Join(parts, ", ") & ".", vbExclamation
        ValidateRequiredInputs = False
        Exit Function
    End If

    If Not EnsureProcessSelections() Then
        ValidateRequiredInputs = False
        Exit Function
    End If

    ValidateRequiredInputs = True
End Function

' Ensures RO/ST and Prasa/Proces selections exist. If only one option is available
' in a pair, it is auto-selected; otherwise, the user is prompted to choose.
Private Function EnsureProcessSelections() As Boolean
    EnsureProcessSelections = True

    If mSelectedROST = "" Then
        If mHasRO And Not mHasST Then
            CommandButtonRO_Click
        ElseIf mHasST And Not mHasRO Then
            CommandButtonST_Click
        Else
            MsgBox "Wybierz RO lub ST przed zapisem.", vbExclamation
            EnsureProcessSelections = False
            Exit Function
        End If
    End If

    If mSelectedPrasaProces = "" Then
        If mHasPrasa And Not mHasProces Then
            CommandButtonPrasa_Click
        ElseIf mHasProces And Not mHasPrasa Then
            CommandButtonProces_Click
        Else
            MsgBox "Wybierz PRASA lub PROCES przed zapisem.", vbExclamation
            EnsureProcessSelections = False
            Exit Function
        End If
    End If
End Function

' Unhides and activates the sheet matching line, shift, RO/ST, and Prasa/Proces, e.g.,
' "PS3-1zm-RO-PRASA".
Private Sub UnhideFormSheet(ByVal linia As String, ByVal shiftVal As Long, _
                            ByVal rost As String, ByVal prasaProces As String)
    Dim wsTarget As Worksheet
    Set wsTarget = GetReportSheet(linia, shiftVal, rost, prasaProces)

    If wsTarget Is Nothing Then Exit Sub

    wsTarget.Activate
End Sub

Private Function GetReportSheet(ByVal linia As String, ByVal shiftVal As Long, _
                                ByVal rost As String, ByVal prasaProces As String) As Worksheet
    If linia = "" Or rost = "" Or prasaProces = "" Then Exit Function

    Dim safeLinia As String
    Dim safeRost As String
    Dim safePrasaProces As String

    safeLinia = Trim$(linia)
    safeRost = UCase$(Trim$(rost))
    safePrasaProces = UCase$(Trim$(prasaProces))

    Dim sheetName As String
    sheetName = safeLinia & "-" & shiftVal & "zm-" & safeRost & "-" & safePrasaProces

    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then Exit Function

    wsTarget.Visible = xlSheetVisible
    Set GetReportSheet = wsTarget
End Function

Private Sub UpdateReportValues(ByVal wsReport As Worksheet, ByVal sourceCol As Long)
    If sourceCol < 2 Or sourceCol > 4 Then Exit Sub

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("data")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        Dim label As String
        label = Trim$(wsData.Cells(rowIndex, 1).Value)

        If label <> "" Then
            Dim valueToCopy As Variant
            valueToCopy = wsData.Cells(rowIndex, sourceCol).Value

            Dim rangeName As String
            rangeName = ToRangeName(label)

            If rangeName <> "" Then
                Dim targetRange As Range
                On Error Resume Next
                Set targetRange = wsReport.Range(rangeName)
                On Error GoTo 0

                If Not targetRange Is Nothing Then
                    targetRange.Value = valueToCopy
                End If
            End If
        End If
    Next rowIndex
End Sub

' Builds the name of a dependent UserForm based on the current selections, e.g.,
' "PS3-1zm-RO-PRASA". Returns an empty string when required selections are missing.
Private Function BuildDependentFormName() As String
    Dim linia As String
    linia = Trim$(Me.ComboBoxLinia.Value)

    If linia = "" Then Exit Function

    Dim shiftVal As Long
    shiftVal = CLng(Val(Me.ComboBoxZmiana.Value))
    If shiftVal < 1 Or shiftVal > 3 Then Exit Function

    If mSelectedROST = "" Or mSelectedPrasaProces = "" Then Exit Function

    BuildDependentFormName = linia & "-" & shiftVal & "zm-" & _
                             UCase$(mSelectedROST) & "-" & UCase$(mSelectedPrasaProces)
End Function

' Attempts to show the dependent UserForm that matches the current selections.
' Displays a warning if any selection is missing or if the form cannot be loaded.
Public Sub ShowDependentUserForm()
    If Not EnsureProcessSelections() Then Exit Sub

    Dim formName As String
    formName = BuildDependentFormName

    If formName = "" Then
        MsgBox "Uzupełnij linię, zmianę oraz wybory RO/ST i PRASA/PROCES, aby otworzyć powiązany formularz.", _
               vbExclamation
        Exit Sub
    End If

    If Not TryShowForm(formName) Then
        MsgBox "Nie znaleziono formularza " & formName & ".", vbExclamation
    End If
End Sub

' Creates and shows a UserForm by name if it exists in the project. Returns True on success.
Private Function TryShowForm(ByVal formName As String) As Boolean
    Dim frm As Object

    ' First, try to create a fresh instance by name.
    On Error Resume Next
    Set frm = VBA.UserForms.Add(formName)
    If Err.Number = 0 And Not frm Is Nothing Then
        frm.Show vbModeless
        TryShowForm = True
        Exit Function
    End If

    Err.Clear
    Set frm = Nothing

    ' Next, see if an instance is already loaded and show it.
    Dim loaded As Object
    For Each loaded In VBA.UserForms
        If StrComp(loaded.Name, formName, vbTextCompare) = 0 Then
            loaded.Show vbModeless
            TryShowForm = True
            Exit Function
        End If
    Next loaded

    ' Final attempt: call the form's Show method directly by name.
    On Error Resume Next
    Application.Run formName & ".Show", vbModeless
    If Err.Number = 0 Then
        TryShowForm = True
    Else
        TryShowForm = False
    End If
    Err.Clear
End Function

' Persists all known problem labels (from cfg_dane) into column A and writes any
' entered values for the active selection into the chosen shift column.
Private Sub SaveProblemEntries(ByVal ws As Worksheet, ByVal targetCol As Long)
    Dim filledValues As Object
    Set filledValues = CreateObject("Scripting.Dictionary")

    Dim idx As Long
    For idx = 1 To 20
        Dim labelName As String
        Dim textName As String

        labelName = "LabelProb" & idx
        textName = "TextBoxProb" & idx

        If ControlExists(labelName) And ControlExists(textName) Then
            With Me.Controls(labelName)
                If .Visible And Trim$(.Caption) <> "" Then
                    filledValues(.Caption) = Me.Controls(textName).Value
                End If
            End With
        End If
    Next idx

    Dim allLabels As Variant
    allLabels = GetAllProblemLabels()

    If IsEmpty(allLabels) Then Exit Sub

    For idx = LBound(allLabels) To UBound(allLabels)
        Dim probLabel As String
        probLabel = allLabels(idx)

        Dim valueToWrite As String
        If filledValues.Exists(probLabel) Then
            valueToWrite = filledValues(probLabel)
        Else
            valueToWrite = ""
        End If

        WriteField ws, probLabel, valueToWrite, targetCol
    Next idx
End Sub

' Returns an alphabetized array of unique problem labels across cfg_dane.
Private Function GetAllProblemLabels() As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_dane")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim col As Long, rowIndex As Long
    For col = 1 To lastCol
        For rowIndex = 4 To lastRow
            Dim probValue As String
            probValue = Trim$(ws.Cells(rowIndex, col).Value)
            If probValue <> "" Then
                If Not dict.Exists(probValue) Then dict.Add probValue, True
            End If
        Next rowIndex
    Next col

    If dict.Count = 0 Then Exit Function

    Dim items() As String
    ReDim items(0 To dict.Count - 1)

    Dim idx As Long
    idx = 0
    Dim key As Variant
    For Each key In dict.Keys
        items(idx) = CStr(key)
        idx = idx + 1
    Next key

    SortStrings items
    GetAllProblemLabels = items
End Function

' Simple bubble sort for small arrays of strings.
Private Sub SortStrings(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim temp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub TextBoxDay_Change()
    ValidateDayInput
End Sub

Private Sub TextBoxPlan_Change()
    UpdateHourlyPlan
    UpdateHourlyActuals
    UpdateProblems
End Sub

' Distributes the plan quantity evenly across 8 hourly text boxes, rounding up.
Private Sub UpdateHourlyPlan()
    Dim totalPlan As Double
    Dim hourlyPlan As Long
    Dim idx As Long

    If Trim$(Me.TextBoxPlan.Value) = "" Or Not IsNumeric(Me.TextBoxPlan.Value) Then
        ClearHourlyPlan
        Exit Sub
    End If

    totalPlan = CDbl(Me.TextBoxPlan.Value)
    If totalPlan <= 0 Then
        ClearHourlyPlan
        Exit Sub
    End If

    hourlyPlan = CLng(Application.WorksheetFunction.RoundUp(totalPlan / 8, 0))

    For idx = 1 To 8
        Me.Controls("TextBoxP" & idx).Value = CStr(hourlyPlan)
    Next idx
End Sub

' Clears hourly plan fields when there is no valid plan input.
Private Sub ClearHourlyPlan()
    Dim idx As Long
    For idx = 1 To 8
        Me.Controls("TextBoxP" & idx).Value = ""
    Next idx
    UpdateHourlyActuals
End Sub

' Sums operator hourly quantities and colors them based on comparison to plan.
Private Sub UpdateHourlyActuals()
    Dim totalActual As Double
    Dim idx As Long

    totalActual = 0

    For idx = 1 To 8
        Dim actualBox As MSForms.TextBox
        Dim planBox As MSForms.TextBox
        Dim actualVal As Double
        Dim planVal As Double

        Set actualBox = Me.Controls("TextBoxH" & idx)
        Set planBox = Me.Controls("TextBoxP" & idx)

        If IsNumeric(actualBox.Value) Then
            actualVal = CDbl(actualBox.Value)
            totalActual = totalActual + actualVal
        Else
            actualVal = 0
        End If

        If IsNumeric(planBox.Value) And CDbl(planBox.Value) > 0 And IsNumeric(actualBox.Value) Then
            planVal = CDbl(planBox.Value)
            If actualVal >= planVal Then
                actualBox.BackColor = RGB(0, 176, 80)
            Else
                actualBox.BackColor = RGB(255, 0, 0)
            End If
        Else
            actualBox.BackColor = vbWhite
        End If
    Next idx

    Me.TextBoxSum.Value = CStr(totalActual)
End Sub

Private Sub TextBoxH1_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH2_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH3_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH4_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH5_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH6_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH7_Change()
    UpdateHourlyActuals
End Sub

Private Sub TextBoxH8_Change()
    UpdateHourlyActuals
End Sub

' Hides and clears all problem labels and text boxes.
Private Sub HideAllProblems()
    Dim idx As Long
    For idx = 1 To 20
        Dim labelName As String
        Dim textName As String

        labelName = "LabelProb" & idx
        textName = "TextBoxProb" & idx

        If ControlExists(labelName) Then
            With Me.Controls(labelName)
                .Caption = ""
                .Visible = False
            End With
        End If

        If ControlExists(textName) Then
            With Me.Controls(textName)
                .Value = ""
                .Visible = False
            End With
        End If
    Next idx
End Sub

' Populates problem labels and inputs based on cfg_dane for the chosen line and process selections.
Private Sub UpdateProblems()
    HideAllProblems

    Dim linia As String
    linia = Trim$(Me.ComboBoxLinia.Value)
    If linia = "" Then Exit Sub

    If mSelectedROST = "" Or mSelectedPrasaProces = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_dane")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim col As Long
    Dim filled As Long

    For col = 1 To lastCol
        If Trim$(ws.Cells(1, col).Value) = linia _
            And Trim$(ws.Cells(2, col).Value) = mSelectedPrasaProces _
            And Trim$(ws.Cells(3, col).Value) = mSelectedROST Then

            Dim rowIndex As Long
            For rowIndex = 4 To lastRow
                Dim probValue As String
                probValue = Trim$(ws.Cells(rowIndex, col).Value)
                If probValue <> "" Then
                    filled = filled + 1
                    If filled > 20 Then Exit For

                    Dim labelName As String
                    Dim textName As String
                    labelName = "LabelProb" & filled
                    textName = "TextBoxProb" & filled

                    If ControlExists(labelName) Then
                        With Me.Controls(labelName)
                            .Caption = probValue
                            .Visible = True
                        End With
                    End If

                    If ControlExists(textName) Then
                        With Me.Controls(textName)
                            .Value = ""
                            .Visible = True
                        End With
                    End If
                End If
            Next rowIndex

            Exit For
        End If
    Next col
End Sub

' Finds (or creates) the row whose column A matches the given label.
Private Function FindOrCreateRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        If CStr(ws.Cells(rowIndex, 1).Value) = label Then
            FindOrCreateRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    FindOrCreateRow = lastRow + 1
    ws.Cells(FindOrCreateRow, 1).Value = label
End Function

Private Function FindRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim normalized As String
    normalized = NormalizeLabel(label)

    If normalized = "" Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        If CStr(ws.Cells(rowIndex, 1).Value) = normalized Then
            FindRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Writes a value into the row identified by label, using targetCol or a fixed column.
Private Sub WriteField(ByVal ws As Worksheet, ByVal label As String, ByVal value As Variant, _
                       ByVal targetCol As Long, Optional ByVal fixedCol As Long = 0)
    Dim normalizedLabel As String
    normalizedLabel = NormalizeLabel(label)

    If normalizedLabel = "" Then Exit Sub

    Dim rowIndex As Long
    rowIndex = FindOrCreateRow(ws, normalizedLabel)

    Dim colToUse As Long
    If fixedCol > 0 Then
        colToUse = fixedCol
    Else
        colToUse = targetCol
    End If

    ws.Cells(rowIndex, colToUse).Value = value
End Sub

Private Function GetFieldValue(ByVal ws As Worksheet, ByVal label As String, _
                               ByVal targetCol As Long) As Variant
    If targetCol <= 0 Then Exit Function

    Dim rowIndex As Long
    rowIndex = FindRow(ws, label)

    If rowIndex = 0 Then Exit Function

    GetFieldValue = ws.Cells(rowIndex, targetCol).Value
End Function

' Normalizes labels for column A by replacing spaces with underscores.
Private Function NormalizeLabel(ByVal label As String) As String
    Dim trimmed As String
    trimmed = Trim$(label)

    If trimmed = "" Then
        NormalizeLabel = ""
        Exit Function
    End If

    Dim cleaned As String
    cleaned = Replace(trimmed, " ", "_")
    NormalizeLabel = cleaned
End Function

' Converts a stored label into a valid range name by normalizing spaces and
' replacing other characters Excel disallows in named ranges.
Private Function ToRangeName(ByVal label As String) As String
    Dim cleaned As String
    cleaned = NormalizeLabel(label)

    If cleaned = "" Then Exit Function

    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "-", "_")
    cleaned = Replace(cleaned, ".", "_")

    ToRangeName = cleaned
End Function
