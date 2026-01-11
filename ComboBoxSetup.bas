Option Explicit

Private mHasRO As Boolean
Private mHasST As Boolean
Private mHasPrasa As Boolean
Private mHasProces As Boolean
Private mSelectedROST As String
Private mSelectedPrasaProces As String
Private mAppState As Object

Private Sub MultiPage1_Change()

End Sub

' Initializes ComboBoxLinia with line headers from row 1 of the cfg_projekt sheet
' and loads ComboBoxProjekt with the projects under the currently selected line.
' Initializes the main form, populating combos and syncing shared context.
Private Sub UserForm_Initialize()
    ' Only run the main initialization when the required controls exist.
    If ControlExists("ComboBoxLinia") And ControlExists("ComboBoxProjekt") Then
        InitializeLiniaIProjekty
        InitializeBrygadaIZmiana
        HideAllProblems
    End If

    EnsureAppState
    LoadStateFromSettings
    ApplyStateToForm Me
    RefreshSharedContext Me
End Sub

' Shows the form modelessly so Excel stays interactive.
' Shows the form modelessly so Excel stays interactive.
Public Sub ShowModeless()
    Me.Show vbModeless
End Sub

' Fills the line and project combo boxes when the form opens.
' Fills the line and project combo boxes when the form opens.
Private Sub InitializeLiniaIProjekty()
    On Error GoTo ExitInit

    Dim cboLinia As MSForms.comboBox
    Dim cboProjekt As MSForms.comboBox

    Set cboLinia = GetComboIfExists("ComboBoxLinia")
    Set cboProjekt = GetComboIfExists("ComboBoxProjekt")

    If cboLinia Is Nothing Or cboProjekt Is Nothing Then GoTo ExitInit

    Dim ws As Worksheet
    Set ws = TryGetWorksheet("cfg_projekt")
    If ws Is Nothing Then GoTo ExitInit

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    With cboLinia
        .Clear
        For col = 1 To lastCol
            If Trim$(ws.Cells(1, col).value) <> "" Then
                .AddItem CStr(ws.Cells(1, col).value)
            End If
        Next col
    End With

    ' Preload projects for the first available line
    If cboLinia.ListCount > 0 Then
        cboLinia.value = cboLinia.List(0)
        LoadProjectsForLine cboLinia.value
        LoadOperatorsForLine cboLinia.value
        UpdateButtonVisibility
    End If

    RefreshSharedContext Me

ExitInit:
End Sub

' Populates ComboBoxBrygada and ComboBoxZmiana with static choices.
' Populates crew (brygada) and shift (zmiana) combo boxes with static choices.
Private Sub InitializeBrygadaIZmiana()
    If Not ControlExists("ComboBoxBrygada") Or Not ControlExists("ComboBoxZmiana") Then
        Exit Sub
    End If

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
' Populates ComboBoxProjekt with the projects found for the given line.
Public Sub LoadProjectsForLine(ByVal linia As String)
    Dim ws As Worksheet
    Set ws = TryGetWorksheet("cfg_projekt")
    If ws Is Nothing Then Exit Sub

    Dim colIndex As Variant
    colIndex = Application.Match(linia, ws.Rows(1), 0)
    If IsError(colIndex) Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(colIndex)).End(xlUp).row

    Dim row As Long
    With Me.ComboBoxProjekt
        .Clear
        For row = 2 To lastRow
            If Trim$(ws.Cells(row, CLng(colIndex)).value) <> "" Then
                .AddItem CStr(ws.Cells(row, CLng(colIndex)).value)
            End If
        Next row
    End With
End Sub

' Event handler to keep the project list in sync when the line selection changes.
' Keeps project and operator choices in sync when line changes.
Private Sub ComboBoxLinia_Change()
    LoadProjectsForLine Me.ComboBoxLinia.value
    LoadOperatorsForLine Me.ComboBoxLinia.value
    UpdateButtonVisibility
    HideAllProblems
    RefreshSharedContext Me
End Sub

' Refreshes button visibility based on the selected line and project.
' Refreshes availability and problems when project changes.
Private Sub ComboBoxProjekt_Change()
    UpdateButtonVisibility
    HideAllProblems
    RefreshSharedContext Me
End Sub

' Resets all process buttons to hidden and default styling.
' Resets process buttons and hides them before re-evaluating availability.
Private Sub HideAllProcessButtons()
    mSelectedROST = ""
    mSelectedPrasaProces = ""
    ResetButtonStyle Me.CommandButtonRO
    ResetButtonStyle Me.CommandButtonSt
    ResetButtonStyle Me.CommandButtonPrasa
    ResetButtonStyle Me.CommandButtonProces

    Me.CommandButtonRO.Visible = False
    Me.CommandButtonSt.Visible = False
    Me.CommandButtonPrasa.Visible = False
    Me.CommandButtonProces.Visible = False
End Sub

' Applies availability rules from cfg_dostepnosc for the chosen line and project.
' Applies availability rules from cfg_dostepnosc for the chosen line/project.
Private Sub UpdateButtonVisibility()
    HideAllProcessButtons

    Dim linia As String: linia = Trim$(Me.ComboBoxLinia.value)
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
    lastRow = ws.Cells(ws.Rows.Count, CLng(colLinia)).End(xlUp).row

    mHasRO = False
    mHasST = False
    mHasPrasa = False
    mHasProces = False

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        If Trim$(ws.Cells(rowIndex, CLng(colLinia)).value) = linia Then
            If Not IsError(colProjekt) Then
                Dim projValue As String
                projValue = Trim$(ws.Cells(rowIndex, CLng(colProjekt)).value)
                If projValue <> "" And projValue <> Trim$(Me.ComboBoxProjekt.value) Then
                    GoTo ContinueNext
                End If
            End If

            mHasRO = (val(ws.Cells(rowIndex, CLng(colRO)).value) = 1)
            mHasST = (val(ws.Cells(rowIndex, CLng(colST)).value) = 1)
            mHasPrasa = (val(ws.Cells(rowIndex, CLng(colPrasa)).value) = 1)
            mHasProces = (val(ws.Cells(rowIndex, CLng(colProces)).value) = 1)

            ShowIfAvailable Me.CommandButtonRO, mHasRO
            ShowIfAvailable Me.CommandButtonSt, mHasST
            ShowIfAvailable Me.CommandButtonPrasa, mHasPrasa
            ShowIfAvailable Me.CommandButtonProces, mHasProces
            Exit For
        End If
ContinueNext:
    Next rowIndex
End Sub

' Sets button visible when the availability flag is 1 and resets its style.
' Shows a process button when its availability flag is true.
Private Sub ShowIfAvailable(ByVal btn As MSForms.CommandButton, ByVal flagValue As Variant)
    btn.Visible = CBool(flagValue)
    If btn.Visible Then
        ResetButtonStyle btn
    End If
End Sub

' Loads operator choices for ComboBoxOp1-ComboBoxOp4 based on the selected line.
' Loads operator names for ComboBoxOp1-Op4 based on the selected line.
Private Sub LoadOperatorsForLine(ByVal linia As String)
    ClearOperatorCombos

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_operator")

    Dim colIndex As Variant
    colIndex = Application.Match(linia, ws.Rows(1), 0)
    If IsError(colIndex) Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, CLng(colIndex)).End(xlUp).row

    Dim ops As Collection
    Set ops = New Collection

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim opValue As String
        opValue = Trim$(ws.Cells(rowIndex, CLng(colIndex)).value)
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
' Clears operator combo boxes.
Private Sub ClearOperatorCombos()
    Me.ComboBoxOp1.Clear
    Me.ComboBoxOp2.Clear
    Me.ComboBoxOp3.Clear
    Me.ComboBoxOp4.Clear
End Sub

' Populates a combo box with operator names.
' Populates a combo box with operator names.
Private Sub SetComboOptions(ByVal comboBox As MSForms.comboBox, ByVal operators As Collection)
    Dim idx As Long

    comboBox.Clear
    For idx = 1 To operators.Count
        comboBox.AddItem operators(idx)
    Next idx
End Sub

' Refreshes the dependent ComboBoxBoxyOp list whenever operator fields change.
' Refreshes the dependent ComboBoxBoxyOp list whenever operator fields change.
Private Sub RefreshBoxyOp()
    PopulateBoxOperatorList Me, Me
    RefreshSharedContext Me
End Sub

Private Sub ComboBoxOp1_Change(): RefreshBoxyOp: End Sub
Private Sub ComboBoxOp2_Change(): RefreshBoxyOp: End Sub
Private Sub ComboBoxOp3_Change(): RefreshBoxyOp: End Sub
Private Sub ComboBoxOp4_Change(): RefreshBoxyOp: End Sub

Private Sub TextBoxOp1_Change(): RefreshBoxyOp: End Sub
Private Sub TextBoxOp2_Change(): RefreshBoxyOp: End Sub
Private Sub TextBoxOp3_Change(): RefreshBoxyOp: End Sub
Private Sub TextBoxOp4_Change(): RefreshBoxyOp: End Sub

' Rebuilds the shared context whenever operator fields are updated.
Private Sub ComboBoxBrygada_Change()
    RefreshSharedContext Me
End Sub

Private Sub ComboBoxZmiana_Change()
    RefreshSharedContext Me
End Sub

' Collects operators from ComboBoxOp1-Op4 and TextBoxOp1-Op4 without duplicates.
' Collects operators from ComboBoxOp1-Op4 and TextBoxOp1-Op4 without duplicates.
Private Function CollectOperatorNames(Optional ByVal sourceForm As Object = Nothing) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim formObj As Object
    If sourceForm Is Nothing Then
        Set formObj = Me
    Else
        Set formObj = sourceForm
    End If

    AddOperatorValue dict, GetTextIfExists(formObj, "ComboBoxOp1")
    AddOperatorValue dict, GetTextIfExists(formObj, "ComboBoxOp2")
    AddOperatorValue dict, GetTextIfExists(formObj, "ComboBoxOp3")
    AddOperatorValue dict, GetTextIfExists(formObj, "ComboBoxOp4")

    AddOperatorValue dict, GetTextIfExists(formObj, "TextBoxOp1")
    AddOperatorValue dict, GetTextIfExists(formObj, "TextBoxOp2")
    AddOperatorValue dict, GetTextIfExists(formObj, "TextBoxOp3")
    AddOperatorValue dict, GetTextIfExists(formObj, "TextBoxOp4")

    Set CollectOperatorNames = dict
End Function

' Adds a single operator value to a dictionary when it is non-empty.
' Adds a single operator value to a dictionary when it is non-empty.
Private Sub AddOperatorValue(ByVal dict As Object, ByVal rawValue As String)
    Dim name As String
    name = Trim$(rawValue)

    If name = "" Then Exit Sub
    If dict.exists(name) Then Exit Sub

    dict.Add name, True
End Sub

' Populates ComboBoxBoxyOp on another form with the gathered operators.
' Populates ComboBoxBoxyOp on another form with gathered operators.
Public Sub PopulateBoxOperatorList(ByVal targetForm As Object, Optional ByVal sourceForm As Object = Nothing)
    If targetForm Is Nothing Then Exit Sub

    Dim dict As Object
    Set dict = CollectOperatorNames(sourceForm)

    Dim cbo As MSForms.comboBox
    Set cbo = GetComboOnForm(targetForm, "ComboBoxBoxyOp")
    If cbo Is Nothing Then Exit Sub

    cbo.Clear

    Dim key As Variant
    For Each key In dict.Keys
        cbo.AddItem CStr(key)
    Next key

    If cbo.ListCount > 0 Then cbo.value = cbo.List(0)
End Sub

' Prepares a related form by pushing current operator choices into its
' ComboBoxBoxyOp list. Safe to call even when the combo is absent.
' Prepares a related form by pushing current operators and date.
Public Sub InitializeBoxForm(ByVal targetForm As Object)
    PopulateBoxOperatorList targetForm, Me

    Dim planDate As String
    planDate = Trim$(GetTextIfExists(Me, "TextBoxDay"))
    If planDate <> "" Then SetTextIfExists targetForm, "TextBoxDay", planDate
End Sub

' Opens the box logging form modelessly and seeds it with current operators and date.
' Opens the box logging form modelessly and seeds it with context.
Public Sub ShowBoxLogForm()
    Dim frm As Object
    If Not TryShowForm("BoxLogForm", frm) Then
        MsgBox "Nie można otworzyć formularza BoxLogForm.", vbExclamation
        Exit Sub
    End If

    InitializeBoxForm frm
End Sub

' Saves box data (date, operator, box number, current and added quantities) to the
' "data" sheet starting at row 2, column AK and moving right.
' Saves box data (date, operator, box number, quantities) to the AK+ log area.
Public Sub SaveBoxEntry(ByVal sourceForm As Object, Optional ByVal sourceMainForm As Object = Nothing)
    Dim ws As Worksheet
    Set ws = TryGetWorksheet("data")
    If ws Is Nothing Then
        MsgBox "Brak arkusza 'data'.", vbExclamation
        Exit Sub
    End If

    Dim contextForm As Object
    If sourceMainForm Is Nothing Then
        Set contextForm = Me
    Else
        Set contextForm = sourceMainForm
    End If

    Dim planDate As String
    planDate = GetStateOrControl(contextForm, "Data", "TextBoxDay")
    If planDate = "" Then planDate = Trim$(GetTextIfExists(sourceForm, "TextBoxDay"))

    Dim operatorName As String
    operatorName = Trim$(GetTextIfExists(sourceForm, "ComboBoxBoxyOp"))

    Dim brygada As String
    brygada = GetStateOrControl(contextForm, "Brygada", "ComboBoxBrygada")
    If brygada = "" Then brygada = Trim$(GetTextIfExists(sourceForm, "ComboBoxBrygada"))

    Dim boxNumber As String
    boxNumber = Trim$(GetTextIfExists(sourceForm, "TextBoxbox1"))

    Dim qtyCurrent As String
    qtyCurrent = Trim$(GetTextIfExists(sourceForm, "TextBoxboxilosc1"))

    Dim qtyAdded As String
    qtyAdded = Trim$(GetTextIfExists(sourceForm, "TextBoxboxilosc2"))

    If planDate = "" Or operatorName = "" Or boxNumber = "" Or _
       qtyCurrent = "" Or qtyAdded = "" Or brygada = "" Then
        MsgBox "Uzupełnij datę, brygadę, operatora, numer boxa oraz ilości przed zapisem.", _
               vbExclamation
        Exit Sub
    End If

    Dim startCol As Long
    startCol = ws.Columns("AK").Column

    Dim targetRow As Long
    targetRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).row
    If targetRow < 2 Then
        targetRow = 2
    Else
        targetRow = targetRow + 1
    End If

    ws.Cells(targetRow, startCol).value = planDate
    ws.Cells(targetRow, startCol + 1).value = brygada
    ws.Cells(targetRow, startCol + 2).value = operatorName
    ws.Cells(targetRow, startCol + 3).value = boxNumber
    ws.Cells(targetRow, startCol + 4).value = qtyCurrent
    ws.Cells(targetRow, startCol + 5).value = qtyAdded
End Sub

' Convenience wrapper for box forms to call from their save buttons.
' Convenience wrapper for box forms to call from their save buttons.
Public Sub SaveBoxEntryFromBoxForm(ByVal boxForm As Object)
    SaveBoxEntry boxForm, Me
End Sub

' Click handler for the dedicated "dodaj box" button to append a box entry.
' Click handler for the dedicated "dodaj box" button to append a box entry.
Public Sub CommandButtondodajbox_Click()
    SaveBoxEntryFromBoxForm Me
End Sub

' Restores default appearance for a button.
' Restores default appearance for a button.
Private Sub ResetButtonStyle(ByVal btn As MSForms.CommandButton)
    btn.BackColor = vbButtonFace
End Sub

' Highlights within the RO/ST pair without clearing Prasa/Proces selection.
' Highlights within the RO/ST pair without clearing Prasa/Proces selection.
Private Sub HighlightROST(ByVal selectedButton As MSForms.CommandButton)
    ResetButtonStyle Me.CommandButtonRO
    ResetButtonStyle Me.CommandButtonSt
    selectedButton.BackColor = RGB(0, 176, 80)
    RefreshSharedContext Me
End Sub

' Highlights within the Prasa/Proces pair without clearing RO/ST selection.
' Highlights within the Prasa/Proces pair without clearing RO/ST selection.
Private Sub HighlightPrasaProces(ByVal selectedButton As MSForms.CommandButton)
    ResetButtonStyle Me.CommandButtonPrasa
    ResetButtonStyle Me.CommandButtonProces
    selectedButton.BackColor = RGB(0, 176, 80)
    RefreshSharedContext Me
End Sub

' Validates the date entered in TextBoxDay using the DD.MM.RRRR format.
' Validates the date entered in TextBoxDay using the DD.MM.RRRR format.
Private Sub ValidateDayInput()
    Dim rawValue As String
    rawValue = Trim$(Me.TextBoxDay.value)

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

' Returns a normalized date value (dd.mm.yyyy) when TextBoxDay contains a valid
' date, or an empty string otherwise.
' Returns a validated DD.MM.RRRR date value or an empty string.
Private Function GetValidatedDayValue(Optional ByVal sourceForm As Object = Nothing) As String
    Dim formObj As Object
    If sourceForm Is Nothing Then
        Set formObj = Me
    Else
        Set formObj = sourceForm
    End If

    Dim rawValue As String
    rawValue = Trim$(GetTextIfExists(formObj, "TextBoxDay"))

    If rawValue = "" Then Exit Function
    If Not rawValue Like "##.##.####" Then Exit Function

    Dim parts() As String
    parts = Split(rawValue, ".")
    If UBound(parts) <> 2 Then Exit Function

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
        GetValidatedDayValue = rawValue
        Exit Function
    End If

InvalidDate:
    GetValidatedDayValue = ""
End Function

' Ensures the in-memory application state dictionary exists.
Private Sub EnsureAppState()
    If mAppState Is Nothing Then
        Set mAppState = CreateObject("Scripting.Dictionary")
    End If
End Sub

' Stores a simple value in the AppState dictionary.
Private Sub SetAppStateValue(ByVal key As String, ByVal value As String)
    EnsureAppState
    mAppState(key) = value
End Sub

' Stores an object value in the AppState dictionary.
Private Sub SetAppStateObject(ByVal key As String, ByVal value As Object)
    EnsureAppState
    Set mAppState(key) = value
End Sub

' Retrieves a string value from AppState.
Private Function GetAppStateValue(ByVal key As String) As String
    If mAppState Is Nothing Then Exit Function
    If Not mAppState.exists(key) Then Exit Function
    GetAppStateValue = CStr(mAppState(key))
End Function

' Applies AppState values to any form that exposes matching controls.
Private Sub ApplyStateToForm(ByVal targetForm As Object)
    If targetForm Is Nothing Then Exit Sub
    EnsureAppState

    SafeSetCombo targetForm, "ComboBoxLinia", GetAppStateValue("Linia")
    SafeSetCombo targetForm, "ComboBoxProjekt", GetAppStateValue("Projekt")
    SafeSetCombo targetForm, "ComboBoxBrygada", GetAppStateValue("Brygada")
    SafeSetCombo targetForm, "ComboBoxZmiana", GetAppStateValue("Zmiana")
    SetTextIfExists targetForm, "TextBoxDay", GetAppStateValue("Data")
    SetTextIfExists targetForm, "TextBoxPlan", GetAppStateValue("Plan")
    SetTextIfExists targetForm, "TextBoxSum", GetAppStateValue("Suma")

    PopulateBoxOperatorListFromContext targetForm
End Sub

' Loads persisted AppState values from the Settings sheet when present.
Private Sub LoadStateFromSettings()
    Dim ws As Worksheet
    Set ws = TryGetSettingsSheet()
    If ws Is Nothing Then Exit Sub

    EnsureAppState

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        Dim key As String
        key = Trim$(CStr(ws.Cells(rowIndex, 1).value))
        If key <> "" Then
            mAppState(key) = CStr(ws.Cells(rowIndex, 2).value)
        End If
    Next rowIndex
End Sub

' Persists AppState values to a Settings sheet so data survives reopen.
Private Sub SaveStateToSettings()
    Dim ws As Worksheet
    Set ws = TryGetSettingsSheet()
    If ws Is Nothing Then Exit Sub
    If mAppState Is Nothing Then Exit Sub

    Dim key As Variant
    Dim rowIndex As Long
    rowIndex = 1

    ws.Columns("A:B").ClearContents
    For Each key In mAppState.Keys
        If VarType(mAppState(key)) <> vbObject Then
            ws.Cells(rowIndex, 1).value = CStr(key)
            ws.Cells(rowIndex, 2).value = CStr(mAppState(key))
            rowIndex = rowIndex + 1
        End If
    Next key
End Sub

' Returns the settings sheet when present (Settings or data), or Nothing if absent.
Private Function TryGetSettingsSheet() As Worksheet
    Set TryGetSettingsSheet = TryGetWorksheet("Settings")
End Function

' Returns a value from AppState, falling back to a control value when missing.
Private Function GetStateOrControl(ByVal formObj As Object, ByVal key As String, ByVal controlName As String) As String
    Dim stateValue As String
    stateValue = GetAppStateValue(key)
    If Trim$(stateValue) <> "" Then
        GetStateOrControl = stateValue
    Else
        GetStateOrControl = Trim$(GetTextIfExists(formObj, controlName))
    End If
End Function

' Keeps a shared snapshot of the main form so other user forms can read the latest
' entries while UserForm1 stays open.
' Keeps a shared snapshot of the main form for dependent forms.
Public Sub RefreshSharedContext(ByVal sourceForm As Object)
    If sourceForm Is Nothing Then Exit Sub
    EnsureAppState

    SetAppStateValue "Linia", Trim$(GetTextIfExists(sourceForm, "ComboBoxLinia"))
    SetAppStateValue "Projekt", Trim$(GetTextIfExists(sourceForm, "ComboBoxProjekt"))
    SetAppStateValue "Brygada", Trim$(GetTextIfExists(sourceForm, "ComboBoxBrygada"))
    SetAppStateValue "Zmiana", Trim$(GetTextIfExists(sourceForm, "ComboBoxZmiana"))
    SetAppStateValue "Data", GetValidatedDayValue(sourceForm)

    Dim ops As Object
    Set ops = CollectOperatorNames(sourceForm)
    SetAppStateObject "Operatorzy", ops

    SetAppStateValue "ROST", mSelectedROST
    SetAppStateValue "PrasaProces", mSelectedPrasaProces
    SetAppStateValue "Plan", Trim$(GetTextIfExists(sourceForm, "TextBoxPlan"))
    SetAppStateValue "Suma", Trim$(GetTextIfExists(sourceForm, "TextBoxSum"))

    SaveStateToSettings
End Sub

' Applies the shared context to any form that exposes matching controls.
' Applies the shared context to any form that exposes matching controls.
Public Sub ApplySharedContext(ByVal targetForm As Object)
    ApplyStateToForm targetForm
End Sub

' Populates ComboBoxBoxyOp using the cached shared operator list.
Private Sub PopulateBoxOperatorListFromContext(ByVal targetForm As Object)
    If targetForm Is Nothing Then Exit Sub
    If mAppState Is Nothing Then Exit Sub
    If Not mAppState.exists("Operatorzy") Then Exit Sub

    Dim ops As Object
    Set ops = mAppState("Operatorzy")
    If ops Is Nothing Then Exit Sub

    Dim combo As MSForms.comboBox
    Set combo = GetComboOnForm(targetForm, "ComboBoxBoxyOp")
    If combo Is Nothing Then Exit Sub

    combo.Clear
    Dim key As Variant
    For Each key In ops.Keys
        combo.AddItem CStr(key)
    Next key
    If combo.ListCount > 0 Then combo.value = combo.List(0)
End Sub

' Safely sets a combo value, adding it if missing.
Private Sub SafeSetCombo(ByVal targetForm As Object, ByVal controlName As String, ByVal newValue As String)
    If targetForm Is Nothing Then Exit Sub
    If newValue = "" Then Exit Sub

    Dim combo As MSForms.comboBox
    Set combo = GetComboOnForm(targetForm, controlName)
    If combo Is Nothing Then Exit Sub

    Dim idx As Long, exists As Boolean
    For idx = 0 To combo.ListCount - 1
        If StrComp(CStr(combo.List(idx)), newValue, vbTextCompare) = 0 Then
            exists = True
            Exit For
        End If
    Next idx
    If Not exists Then combo.AddItem newValue
    combo.value = newValue
End Sub

' Retrieves a value from the shared context dictionary.
Private Function GetSharedValue(ByVal key As String) As String
    GetSharedValue = GetAppStateValue(key)
End Function

' Safely checks for the presence of a control by name on the form.
Private Function ControlExists(ByVal controlName As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = Me.Controls(controlName)
    ControlExists = (Err.Number = 0)
    Err.Clear
End Function

' Safely retrieves the text value from a control on a given form when present.
Private Function GetTextIfExists(ByVal formObj As Object, ByVal controlName As String) As String
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = formObj.Controls(controlName)

    If Err.Number = 0 Then
        GetTextIfExists = CStr(ctrl.value)
    Else
        Err.Clear
    End If

    On Error GoTo 0
End Function

' Safely sets a control's Value when it exists.
Private Sub SetTextIfExists(ByVal formObj As Object, ByVal controlName As String, ByVal newValue As String)
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = formObj.Controls(controlName)
    If Err.Number = 0 Then ctrl.value = newValue Else Err.Clear
    On Error GoTo 0
End Sub

' Safely returns a ComboBox control when it exists and is the right type.
Private Function GetComboIfExists(ByVal controlName As String) As MSForms.comboBox
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = Me.Controls(controlName)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    If TypeOf ctrl Is MSForms.comboBox Then
        Set GetComboIfExists = ctrl
    End If
End Function

' Returns a ComboBox from another form when it exists and is the right type.
Private Function GetComboOnForm(ByVal formObj As Object, ByVal controlName As String) As MSForms.comboBox
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = formObj.Controls(controlName)
    If Err.Number = 0 Then
        If TypeOf ctrl Is MSForms.comboBox Then Set GetComboOnForm = ctrl
    Else
        Err.Clear
    End If

    On Error GoTo 0
End Function

' Returns a worksheet when it exists; otherwise returns Nothing without raising.
Private Function TryGetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub CommandButtonRO_Click()
    HighlightROST Me.CommandButtonRO
    mSelectedROST = "RO"
    If mHasRO And mHasST Then
        Me.CommandButtonSt.Visible = False
    End If
    UpdateProblems
End Sub

Private Sub CommandButtonST_Click()
    HighlightROST Me.CommandButtonSt
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

    ' Persist the main shift data first.
    SaveFormData

    ' Append a summary row in the AK+ log for audit/history.
    AppendMainEntryToDataAK

    MsgBox "Dane zapisane.", vbInformation
End Sub


' Switches the multipage control to the awarie/postój page if it exists.
Private Function ShowAwariePageOnMultiPage() As Boolean
    Dim mp As MSForms.MultiPage
    Set mp = GetMultiPageIfExists("MultiPage1")
    If mp Is Nothing Then Exit Function

    Dim pageIndex As Long
    pageIndex = FindAwariePageIndex(mp)

    If pageIndex >= 0 Then
        mp.value = pageIndex
        ShowAwariePageOnMultiPage = True
    End If
End Function

' Attempts to find the awarie/postój page by caption or name; falls back to page 1.
Private Function FindAwariePageIndex(ByVal mp As MSForms.MultiPage) As Long
    Dim idx As Long
    Dim captionText As String

    FindAwariePageIndex = -1

    For idx = 0 To mp.Pages.Count - 1
        captionText = LCase$(mp.Pages(idx).Caption & " " & mp.Pages(idx).name)
        If InStr(captionText, "awari") > 0 Or InStr(captionText, "post") > 0 Then
            FindAwariePageIndex = idx
            Exit Function
        End If
    Next idx

    If mp.Pages.Count > 1 Then FindAwariePageIndex = 1
End Function

Private Function GetMultiPageIfExists(ByVal controlName As String) As MSForms.MultiPage
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = Me.Controls(controlName)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    If TypeOf ctrl Is MSForms.MultiPage Then
        Set GetMultiPageIfExists = ctrl
    End If
End Function

' Saves all form entries into the "data" sheet. Columns B, C, and D are cleared on
' each save and used for shifts 1, 2, and 3 respectively; column I is cleared and
' reused for totals. Rows are matched by stable labels in column A, creating new
' rows when labels are missing so data always aligns to the same descriptors.
Private Sub SaveFormData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("data")

    ' Ensure a valid shift and compute its target column.
    Dim shiftVal As Long
    shiftVal = CLng(val(GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana")))
    If shiftVal < 1 Or shiftVal > 3 Then
        MsgBox "Wybierz zmianę 1, 2 lub 3 przed zapisem.", vbExclamation
        Exit Sub
    End If

    Dim targetCol As Long
    targetCol = 1 + shiftVal ' 1->B, 2->C, 3->D

    Dim dayValue As String
    dayValue = GetAppStateValue("Data")
    If dayValue = "" Then dayValue = GetValidatedDayValue()
    If dayValue = "" Then
        MsgBox "Wprowadź poprawną datę (DD.MM.RRRR) przed zapisem.", vbExclamation
        Exit Sub
    End If

    If Not EnsureProcessSelections() Then Exit Sub

    Dim liniaValue As String
    Dim projektValue As String
    Dim brygadaValue As String
    Dim zmianaValue As String
    Dim planValue As String
    Dim sumaValue As String

    liniaValue = GetStateOrControl(Me, "Linia", "ComboBoxLinia")
    projektValue = GetStateOrControl(Me, "Projekt", "ComboBoxProjekt")
    brygadaValue = GetStateOrControl(Me, "Brygada", "ComboBoxBrygada")
    zmianaValue = GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana")
    planValue = GetStateOrControl(Me, "Plan", "TextBoxPlan")
    sumaValue = GetStateOrControl(Me, "Suma", "TextBoxSum")

    ' Clear previous entries so each save starts from a clean slate.
    ws.Columns("B:D").ClearContents
    ws.Columns("I").ClearContents

    ' Core identifiers and selections.
    WriteField ws, "Data", dayValue, targetCol
    WriteField ws, "Linia", liniaValue, targetCol
    WriteField ws, "Projekt", projektValue, targetCol
    WriteField ws, "Brygada", brygadaValue, targetCol
    WriteField ws, "Zmiana", zmianaValue, targetCol
    WriteField ws, "RO/ST", mSelectedROST, targetCol
    WriteField ws, "Prasa/Proces", mSelectedPrasaProces, targetCol

    ' Operators.
    WriteField ws, "Operator 1", Me.ComboBoxOp1.value, targetCol
    WriteField ws, "Operator 2", Me.ComboBoxOp2.value, targetCol
    WriteField ws, "Operator 3", Me.ComboBoxOp3.value, targetCol
    WriteField ws, "Operator 4", Me.ComboBoxOp4.value, targetCol

    ' Plan and hourly plan.
    WriteField ws, "Plan", planValue, targetCol

    Dim idx As Long
    For idx = 1 To 8
        WriteField ws, "Plan H" & idx, Me.Controls("TextBoxP" & idx).value, targetCol
    Next idx

    ' Hourly execution and total.
    For idx = 1 To 8
        WriteField ws, "Wykonanie H" & idx, Me.Controls("TextBoxH" & idx).value, targetCol
    Next idx

    WriteField ws, "Suma wykonania", sumaValue, targetCol, 9 ' Column I

    SaveProblemEntries ws, targetCol

    Dim wsReport As Worksheet
    Set wsReport = GetReportSheet(Trim$(liniaValue), shiftVal, _
                                  mSelectedROST, mSelectedPrasaProces)

    If Not wsReport Is Nothing Then
        UpdateReportValues wsReport, targetCol
        wsReport.Activate
    End If
End Sub

' Appends a summary row to the AK+ log with the current shift context.
Private Sub AppendMainEntryToDataAK()
    Dim ws As Worksheet
    Set ws = TryGetWorksheet("data")
    If ws Is Nothing Then
        MsgBox "Brak arkusza 'data'.", vbExclamation
        Exit Sub
    End If

    Dim dayValue As String
    dayValue = GetAppStateValue("Data")
    If dayValue = "" Then dayValue = GetValidatedDayValue()
    If dayValue = "" Then
        MsgBox "Wprowadź poprawną datę (DD.MM.RRRR) przed zapisem.", vbExclamation
        Exit Sub
    End If

    Dim startCol As Long
    startCol = ws.Columns("AK").Column

    Dim targetRow As Long
    targetRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).row
    If targetRow < 2 Then
        targetRow = 2
    Else
        targetRow = targetRow + 1
    End If

    ws.Cells(targetRow, startCol + 0).value = dayValue                            ' AK Data
    ws.Cells(targetRow, startCol + 1).value = GetStateOrControl(Me, "Linia", "ComboBoxLinia")   ' AL Linia
    ws.Cells(targetRow, startCol + 2).value = GetStateOrControl(Me, "Projekt", "ComboBoxProjekt") ' AM Projekt
    ws.Cells(targetRow, startCol + 3).value = GetStateOrControl(Me, "Brygada", "ComboBoxBrygada") ' AN Brygada
    ws.Cells(targetRow, startCol + 4).value = GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana")  ' AO Zmiana
    ws.Cells(targetRow, startCol + 5).value = mSelectedROST                       ' AP RO/ST
    ws.Cells(targetRow, startCol + 6).value = mSelectedPrasaProces                ' AQ Prasa/Proces

    ws.Cells(targetRow, startCol + 7).value = Trim$(Me.ComboBoxOp1.value)         ' AR Op1
    ws.Cells(targetRow, startCol + 8).value = Trim$(Me.ComboBoxOp2.value)         ' AS Op2
    ws.Cells(targetRow, startCol + 9).value = Trim$(Me.ComboBoxOp3.value)         ' AT Op3
    ws.Cells(targetRow, startCol + 10).value = Trim$(Me.ComboBoxOp4.value)        ' AU Op4

    ws.Cells(targetRow, startCol + 11).value = Trim$(GetTextIfExists(Me, "TextBoxboxilosc1")) ' AV Ilość box (stan)
    ws.Cells(targetRow, startCol + 12).value = GetStateOrControl(Me, "Suma", "TextBoxSum")    ' AW Realizacja
End Sub

' Verifies required selections and inputs before running save or downtime actions.
Private Function ValidateRequiredInputs() As Boolean
    Dim missing As Collection
    Set missing = New Collection

    Dim dayValue As String
    dayValue = GetAppStateValue("Data")
    If dayValue = "" Then dayValue = GetValidatedDayValue()
    If dayValue = "" Then missing.Add "data"

    If GetStateOrControl(Me, "Linia", "ComboBoxLinia") = "" Then missing.Add "linia"
    If GetStateOrControl(Me, "Projekt", "ComboBoxProjekt") = "" Then missing.Add "projekt"
    If GetStateOrControl(Me, "Brygada", "ComboBoxBrygada") = "" Then missing.Add "brygada"
    If GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana") = "" Then missing.Add "zmiana"
    If GetStateOrControl(Me, "Plan", "TextBoxPlan") = "" Then missing.Add "plan"
    If GetStateOrControl(Me, "Suma", "TextBoxSum") = "" Then missing.Add "realizacja"

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
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        Dim label As String
        label = Trim$(wsData.Cells(rowIndex, 1).value)

        If label <> "" Then
            Dim valueToCopy As Variant
            valueToCopy = wsData.Cells(rowIndex, sourceCol).value

            Dim rangeName As String
            rangeName = ToRangeName(label)

            If rangeName <> "" Then
                Dim targetRange As Range
                Set targetRange = GetNamedRange(wsReport, rangeName)

                If Not targetRange Is Nothing Then
                    targetRange.value = valueToCopy
                End If
            End If
        End If
    Next rowIndex

    CopyAlarmValues wsData, wsReport, sourceCol
End Sub

' Returns a named range scoped to the worksheet or workbook, or Nothing if the name
' is not defined. This avoids accidental column/row references when a name like
' "JC" or "A" is absent from the Names collection.
Private Function GetNamedRange(ByVal ws As Worksheet, ByVal rangeName As String) As Range
    Dim nm As name
    Dim qualified As String
    qualified = "'" & ws.name & "'!" & rangeName

    On Error Resume Next
    Set nm = ws.Names(rangeName)
    If nm Is Nothing Then Set nm = ws.Parent.Names(qualified)
    If nm Is Nothing Then Set nm = ws.Parent.Names(rangeName)
    On Error GoTo 0

    If Not nm Is Nothing Then
        On Error Resume Next
        Set GetNamedRange = nm.RefersToRange
        On Error GoTo 0
    End If
End Function

Private Function SafeGetRange(ByVal ranges As Collection, ByVal index As Long) As Range
    On Error Resume Next
    Set SafeGetRange = ranges.Item(index)
    On Error GoTo 0
End Function

Private Function TopLeftCell(ByVal rng As Range) As Range
    If rng Is Nothing Then Exit Function

    If rng.MergeCells Then
        Set TopLeftCell = rng.MergeArea.Cells(1, 1)
    Else
        Set TopLeftCell = rng.Cells(1, 1)
    End If
End Function

Private Sub ClearVerticalRange(ByVal rng As Range, ByVal depth As Long)
    Dim anchor As Range
    Dim i As Long
    Dim cellToClear As Range

    Set anchor = TopLeftCell(rng)
    If anchor Is Nothing Then Exit Sub

    If depth < 1 Then depth = 1

    For i = 0 To depth - 1
        Set cellToClear = TopLeftCell(anchor.Worksheet.Cells(anchor.row + i, anchor.Column))
        If Not cellToClear Is Nothing Then
            cellToClear.MergeArea.ClearContents
        End If
    Next i
End Sub

Private Function SafeGetLong(ByVal values As Collection, ByVal index As Long) As Long
    On Error Resume Next

    If values Is Nothing Then
        SafeGetLong = 0
        On Error GoTo 0
        Exit Function
    End If

    SafeGetLong = CLng(values.Item(index))
    If Err.Number <> 0 Then
        Err.Clear
        SafeGetLong = 0
    End If

    On Error GoTo 0
End Function

' Copies all alarm entries for the given shift into the report sheet, stacking them
' under the existing named ranges without requiring additional named cells.
Private Sub CopyAlarmValues(ByVal wsData As Worksheet, ByVal wsReport As Worksheet, _
                            ByVal shiftCol As Long)
    Dim startCol As Long, endCol As Long

    Select Case shiftCol
        Case 2 ' Shift 1 -> G:N
            startCol = 7: endCol = 14
        Case 3 ' Shift 2 -> P:W
            startCol = 16: endCol = 23
        Case 4 ' Shift 3 -> Z:AG
            startCol = 26: endCol = 33
        Case Else
            Exit Sub
    End Select

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, startCol).End(xlUp).row

    Dim colIndex As Long
    Dim header As String
    Dim rangeName As String
    Dim targetRange As Range
    Dim anchorCell As Range
    Dim sourceCol As Long

    Dim headers As Collection
    Dim sourceCols As Collection
    Dim targetRanges As Collection

    Set headers = New Collection
    Set sourceCols = New Collection
    Set targetRanges = New Collection

    For colIndex = startCol To endCol
        header = Trim$(wsData.Cells(1, colIndex).value)

        If header <> "" Then
            rangeName = ToRangeName(header)

            If rangeName <> "" Then
                Set targetRange = GetNamedRange(wsReport, rangeName)

                If Not targetRange Is Nothing Then
                    headers.Add header
                    sourceCols.Add colIndex
                    targetRanges.Add targetRange
                End If
            End If
        End If
    Next colIndex

    Dim headerCount As Long
    headerCount = targetRanges.Count

    If headerCount = 0 Then Exit Sub

    Dim entryCount As Long
    Dim dataRows() As Long
    Dim rowIndex As Long

    For rowIndex = 2 To lastRow
        Dim hasData As Boolean
        For colIndex = startCol To endCol
            If Trim$(wsData.Cells(rowIndex, colIndex).value) <> "" Then
                hasData = True
                Exit For
            End If
        Next colIndex

        If hasData Then
            entryCount = entryCount + 1
            ReDim Preserve dataRows(1 To entryCount)
            dataRows(entryCount) = rowIndex
        End If
    Next rowIndex

    If entryCount = 0 Then
        ' Clear existing cells to avoid stale values when no alarms are present.
        Dim clearDepth As Long
        clearDepth = lastRow - 1

        If clearDepth < 1 Then clearDepth = 1

        For colIndex = 1 To headerCount
            Set targetRange = SafeGetRange(targetRanges, colIndex)
            If Not targetRange Is Nothing Then
                ClearVerticalRange targetRange, clearDepth
            End If
        Next colIndex
        Exit Sub
    End If

    ' Clear enough rows to cover the current dataset so old entries do not linger.
    Dim maxDepth As Long
    maxDepth = Application.Max(entryCount, lastRow - 1)

    For colIndex = 1 To headerCount
        Set targetRange = SafeGetRange(targetRanges, colIndex)
        If Not targetRange Is Nothing Then
            ClearVerticalRange targetRange, maxDepth
        End If
    Next colIndex

    Dim entryIndex As Long
    For entryIndex = 1 To entryCount
        rowIndex = dataRows(entryIndex)

        For colIndex = 1 To headerCount
            Set targetRange = SafeGetRange(targetRanges, colIndex)
            Set anchorCell = TopLeftCell(targetRange)
            sourceCol = SafeGetLong(sourceCols, colIndex)

            If Not anchorCell Is Nothing And sourceCol > 0 Then
                Dim targetCell As Range
                Set targetCell = TopLeftCell(anchorCell.Worksheet.Cells(anchorCell.row + (entryIndex - 1), anchorCell.Column))

                If Not targetCell Is Nothing Then
                    targetCell.value = wsData.Cells(rowIndex, sourceCol).value
                End If
            End If
        Next colIndex
    Next entryIndex
End Sub

' Builds the name of a dependent UserForm based on the current selections, e.g.,
' "PS3-1zm-RO-PRASA". Returns an empty string when required selections are missing.
Private Function BuildDependentFormName() As String
    Dim linia As String
    linia = GetStateOrControl(Me, "Linia", "ComboBoxLinia")

    If linia = "" Then Exit Function

    Dim shiftVal As Long
    shiftVal = CLng(val(GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana")))
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

' Applies the current form's line and project selections to another form if matching
' combo boxes exist there. Missing controls are silently ignored.
Private Sub ApplyContextToForm(ByVal targetForm As Object)
    If targetForm Is Nothing Then Exit Sub

    Dim linia As String
    Dim projekt As String

    If ControlExists("ComboBoxLinia") Then linia = Me.ComboBoxLinia.value
    If ControlExists("ComboBoxProjekt") Then projekt = Me.ComboBoxProjekt.value

    If linia <> "" Then SetComboValueIfExists targetForm, "ComboBoxLinia", linia
    If projekt <> "" Then SetComboValueIfExists targetForm, "ComboBoxProjekt", projekt

    ' Also share operator selections with forms that expose ComboBoxBoxyOp.
    InitializeBoxForm targetForm
End Sub

' Safe context application that does not depend on the caller having the same controls.
Private Sub ApplyContextToFormSafe(ByVal targetForm As Object, ByVal sourceForm As Object)
    If targetForm Is Nothing Then Exit Sub
    If sourceForm Is Nothing Then Exit Sub

    Dim linia As String
    Dim projekt As String

    linia = Trim$(GetTextIfExists(sourceForm, "ComboBoxLinia"))
    projekt = Trim$(GetTextIfExists(sourceForm, "ComboBoxProjekt"))

    If linia <> "" Then SetComboValueIfExists targetForm, "ComboBoxLinia", linia
    If projekt <> "" Then SetComboValueIfExists targetForm, "ComboBoxProjekt", projekt

    PopulateBoxOperatorList targetForm, sourceForm
    Dim planDate As String
    planDate = GetValidatedDayValue(sourceForm)
    If planDate <> "" Then SetTextIfExists targetForm, "TextBoxDay", planDate
    Dim brygada As String
    brygada = Trim$(GetTextIfExists(sourceForm, "ComboBoxBrygada"))
    If brygada <> "" Then SetComboValueIfExists targetForm, "ComboBoxBrygada", brygada
End Sub

' Sets a combo box value on another form, adding the item if it is not already present.
Private Sub SetComboValueIfExists(ByVal targetForm As Object, ByVal controlName As String, ByVal value As String)
    If targetForm Is Nothing Then Exit Sub
    If value = "" Then Exit Sub

    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = targetForm.Controls(controlName)
    On Error GoTo 0

    If ctrl Is Nothing Then Exit Sub
    If Not TypeOf ctrl Is MSForms.comboBox Then Exit Sub

    Dim idx As Long, exists As Boolean
    For idx = 0 To ctrl.ListCount - 1
        If StrComp(CStr(ctrl.List(idx)), value, vbTextCompare) = 0 Then
            exists = True
            Exit For
        End If
    Next idx

    If Not exists Then ctrl.AddItem value
    ctrl.value = value
End Sub

' Creates and shows a UserForm by name if it exists in the project. Returns True on success.
Private Function TryShowForm(ByVal formName As String, Optional ByRef openedForm As Object, _
                             Optional ByVal contextForm As Object = Nothing) As Boolean
    Dim frm As Object
    Set openedForm = Nothing

    ' First, try to create a fresh instance by name.
    On Error Resume Next
    Set frm = VBA.UserForms.Add(formName)
    If Err.Number = 0 And Not frm Is Nothing Then
        frm.Show vbModeless
        If Not contextForm Is Nothing Then ApplyContextToFormSafe frm, contextForm
        Set openedForm = frm
        TryShowForm = True
        Exit Function
    End If

    Err.Clear
    Set frm = Nothing

    ' Next, see if an instance is already loaded and show it.
    Dim loaded As Object
    For Each loaded In VBA.UserForms
        If StrComp(loaded.name, formName, vbTextCompare) = 0 Then
            loaded.Show vbModeless
            If Not contextForm Is Nothing Then ApplyContextToFormSafe loaded, contextForm
            Set openedForm = loaded
            TryShowForm = True
            Exit Function
        End If
    Next loaded

    ' Final attempt: call the form's Show method directly by name.
    On Error Resume Next
    Application.Run formName & ".Show", vbModeless
    If Err.Number = 0 Then
        ' Attempt to capture the instance that was shown via Application.Run.
        For Each loaded In VBA.UserForms
            If StrComp(loaded.name, formName, vbTextCompare) = 0 Then
                If Not contextForm Is Nothing Then ApplyContextToFormSafe loaded, contextForm
                Set openedForm = loaded
                Exit For
            End If
        Next loaded

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
                    filledValues(.Caption) = Me.Controls(textName).value
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
        If filledValues.exists(probLabel) Then
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
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim col As Long, rowIndex As Long
    For col = 1 To lastCol
        For rowIndex = 4 To lastRow
            Dim probValue As String
            probValue = Trim$(ws.Cells(rowIndex, col).value)
            If probValue <> "" Then
                If Not dict.exists(probValue) Then dict.Add probValue, True
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

' Re-validates the date and updates shared context when TextBoxDay changes.
Private Sub TextBoxDay_Change()
    ValidateDayInput
    RefreshSharedContext Me
End Sub

' Recalculate hourly plan, actuals, and problems when plan changes.
Private Sub TextBoxPlan_Change()
    UpdateHourlyPlan
    UpdateHourlyActuals
    UpdateProblems
    SetAppStateValue "Plan", Trim$(Me.TextBoxPlan.value)
    RefreshSharedContext Me
End Sub

' Distributes the plan quantity evenly across 8 hourly text boxes, rounding up.
' Distributes the plan quantity evenly across 8 hourly text boxes, rounding up.
Private Sub UpdateHourlyPlan()
    Dim totalPlan As Double
    Dim hourlyPlan As Long
    Dim idx As Long

    If Trim$(Me.TextBoxPlan.value) = "" Or Not IsNumeric(Me.TextBoxPlan.value) Then
        ClearHourlyPlan
        Exit Sub
    End If

    totalPlan = CDbl(Me.TextBoxPlan.value)
    If totalPlan <= 0 Then
        ClearHourlyPlan
        Exit Sub
    End If

    hourlyPlan = CLng(Application.WorksheetFunction.RoundUp(totalPlan / 8, 0))

    For idx = 1 To 8
        Me.Controls("TextBoxP" & idx).value = CStr(hourlyPlan)
    Next idx
End Sub

' Clears hourly plan fields when there is no valid plan input.
' Clears hourly plan fields when there is no valid plan input.
Private Sub ClearHourlyPlan()
    Dim idx As Long
    For idx = 1 To 8
        Me.Controls("TextBoxP" & idx).value = ""
    Next idx
    UpdateHourlyActuals
End Sub

' Sums operator hourly quantities and colors them based on comparison to plan.
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

        If IsNumeric(actualBox.value) Then
            actualVal = CDbl(actualBox.value)
            totalActual = totalActual + actualVal
        Else
            actualVal = 0
        End If

        If IsNumeric(planBox.value) And CDbl(planBox.value) > 0 And IsNumeric(actualBox.value) Then
            planVal = CDbl(planBox.value)
            If actualVal >= planVal Then
                actualBox.BackColor = RGB(0, 176, 80)
            Else
                actualBox.BackColor = RGB(255, 0, 0)
            End If
        Else
            actualBox.BackColor = vbWhite
        End If
    Next idx

    Me.TextBoxSum.value = CStr(totalActual)
    SetAppStateValue "Suma", Me.TextBoxSum.value
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
                .value = ""
                .Visible = False
            End With
        End If
    Next idx
End Sub

' Populates problem labels and inputs based on cfg_dane for the chosen line and process selections.
' Populates problem labels and inputs based on cfg_dane selections.
Private Sub UpdateProblems()
    HideAllProblems

    Dim linia As String
    linia = Trim$(Me.ComboBoxLinia.value)
    If linia = "" Then Exit Sub

    If mSelectedROST = "" Or mSelectedPrasaProces = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("cfg_dane")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim col As Long
    Dim filled As Long

    For col = 1 To lastCol
        If Trim$(ws.Cells(1, col).value) = linia _
            And Trim$(ws.Cells(2, col).value) = mSelectedPrasaProces _
            And Trim$(ws.Cells(3, col).value) = mSelectedROST Then

            Dim rowIndex As Long
            For rowIndex = 4 To lastRow
                Dim probValue As String
                probValue = Trim$(ws.Cells(rowIndex, col).value)
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
                            .value = ""
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
' Finds (or creates) the row whose column A matches the given label.
Private Function FindOrCreateRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        If CStr(ws.Cells(rowIndex, 1).value) = label Then
            FindOrCreateRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    FindOrCreateRow = lastRow + 1
    ws.Cells(FindOrCreateRow, 1).value = label
End Function

' Finds an existing row whose column A matches the given label.
Private Function FindRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim normalized As String
    normalized = NormalizeLabel(label)

    If normalized = "" Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        If CStr(ws.Cells(rowIndex, 1).value) = normalized Then
            FindRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Writes a value into the row identified by label, using targetCol or a fixed column.
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

    ws.Cells(rowIndex, colToUse).value = value
End Sub

' Reads a stored value for the label from the specified column.
Private Function GetFieldValue(ByVal ws As Worksheet, ByVal label As String, _
                               ByVal targetCol As Long) As Variant
    If targetCol <= 0 Then Exit Function

    Dim rowIndex As Long
    rowIndex = FindRow(ws, label)

    If rowIndex = 0 Then Exit Function

    GetFieldValue = ws.Cells(rowIndex, targetCol).value
End Function

' Normalizes labels for column A by replacing spaces with underscores.
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
' Converts a stored label into a valid range name.
Private Function ToRangeName(ByVal label As String) As String
    Dim cleaned As String
    cleaned = NormalizeLabel(label)

    If cleaned = "" Then Exit Function

    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "-", "_")
    cleaned = Replace(cleaned, ".", "_")

    ToRangeName = cleaned
End Function
' ============================================================
'  AWARIE / POSTOJ - wszystko w UserForm1 (bez modułów)
'  Źródło: cfg_awarie_long / tabela tblAwarie: (1)Linia (2)Maszyna (3)Opis
'  Zapis: data -> zmiana1: G:N, zmiana2: P:W, zmiana3: Z:AG
' ============================================================

Private Const AW_SHEET As String = "cfg_awarie_long"
Private Const AW_TABLE As String = "tblAwarie"

' ---------- Zdarzenia kontrolek awarii (w UserForm1) ----------

Private Sub CommandButtonPostoj_Click()
    On Error GoTo EH

    If Not ValidateRequiredInputs() Then Exit Sub

    ' przełącz stronę multipage na awarie (jak masz)
    ShowAwariePageOnMultiPage

    ' zainicjuj listy awarii w tym samym UserForm1
    Awarie_Init

    Exit Sub
EH:
    MsgBox "Błąd otwarcia awarii: " & Err.Description, vbExclamation
End Sub

Private Sub cboA_Maszyna_Change()
    Awarie_LoadAwarieForMaszyna Trim$(Me.cboA_Maszyna.value)
End Sub

Private Sub cmdA_Zapisz_Click()
    Awarie_SavePostoj
End Sub

Private Sub cmdA_Wyczysc_Click()
    Awarie_ClearFields
End Sub

' ---------- Główne procedury awarii ----------

Private Sub Awarie_Init()
    ' Ustal linię i zmianę z kontekstu głównego (u Ciebie jest mAppState)
    Dim linia As String
    linia = GetStateOrControl(Me, "Linia", "ComboBoxLinia")
    If Trim$(linia) = "" Then
        MsgBox "Brak wybranej linii w głównym formularzu.", vbExclamation
        Exit Sub
    End If

    Awarie_ClearUI

    ' Załaduj maszyny dla linii
    Awarie_LoadMaszynyForLinia linia

    ' Auto: jeśli jest pierwsza maszyna, załaduj awarie
    If Me.cboA_Maszyna.ListCount > 0 Then
        Me.cboA_Maszyna.ListIndex = 0
        Awarie_LoadAwarieForMaszyna Trim$(Me.cboA_Maszyna.value)
    End If
End Sub

Private Sub Awarie_ClearUI()
    On Error Resume Next
    Me.cboA_Maszyna.Clear
    Me.lstA_Awarie.Clear
    Me.lblA_Maszyna.Caption = vbNullString
    Awarie_ClearFields
    On Error GoTo 0
End Sub

Private Sub Awarie_ClearFields()
    On Error Resume Next
    Me.txtA_CzasMin.value = vbNullString
    Me.txtA_Komentarz.value = vbNullString
    Me.txtA_Uwagi.value = vbNullString
    On Error GoTo 0
End Sub

Private Sub Awarie_LoadMaszynyForLinia(ByVal linia As String)
    Dim lo As ListObject
    Set lo = Awarie_GetTable()
    If lo Is Nothing Then
        MsgBox "Brak tabeli '" & AW_TABLE & "' w arkuszu '" & AW_SHEET & "'.", vbExclamation
        Exit Sub
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range.Cells(1, 1).value)) = Trim$(linia) Then
            Dim m As String
            m = Trim$(CStr(r.Range.Cells(1, 2).value))
            If m <> "" Then dict(m) = True
        End If
    Next r

    Me.cboA_Maszyna.Clear

    Dim k As Variant
    For Each k In dict.Keys
        Me.cboA_Maszyna.AddItem CStr(k)
    Next k

    If Me.cboA_Maszyna.ListCount = 0 Then
        MsgBox "Brak maszyn dla linii '" & linia & "' w tabeli '" & AW_TABLE & "'.", vbExclamation
    End If
End Sub

Private Sub Awarie_LoadAwarieForMaszyna(ByVal maszyna As String)
    Me.lstA_Awarie.Clear
    Me.lblA_Maszyna.Caption = maszyna

    Dim linia As String
    linia = GetStateOrControl(Me, "Linia", "ComboBoxLinia")

    If Trim$(linia) = "" Or Trim$(maszyna) = "" Then Exit Sub

    Dim lo As ListObject
    Set lo = Awarie_GetTable()
    If lo Is Nothing Then Exit Sub

    Dim r As ListRow
    For Each r In lo.ListRows
        If Trim$(CStr(r.Range.Cells(1, 1).value)) = Trim$(linia) _
           And Trim$(CStr(r.Range.Cells(1, 2).value)) = Trim$(maszyna) Then

            Dim opis As String
            opis = Trim$(CStr(r.Range.Cells(1, 3).value))
            If opis <> "" Then Me.lstA_Awarie.AddItem opis
        End If
    Next r

    If Me.lstA_Awarie.ListCount > 0 Then Me.lstA_Awarie.ListIndex = 0
End Sub

Private Sub Awarie_SavePostoj()
    On Error GoTo EH

    Dim linia As String, maszyna As String, opis As String
    Dim czasMin As Double

    linia = Trim$(GetStateOrControl(Me, "Linia", "ComboBoxLinia"))
    maszyna = Trim$(Me.cboA_Maszyna.value)

    If Me.lstA_Awarie.ListIndex >= 0 Then
        opis = CStr(Me.lstA_Awarie.List(Me.lstA_Awarie.ListIndex))
    Else
        opis = ""
    End If

    If linia = "" Then
        MsgBox "Brak linii.", vbExclamation
        Exit Sub
    End If
    If maszyna = "" Then
        MsgBox "Wybierz maszynę.", vbExclamation
        Exit Sub
    End If
    If opis = "" Then
        MsgBox "Wybierz awarię/postój.", vbExclamation
        Exit Sub
    End If

    Dim rawCzas As String
    rawCzas = Replace(Trim$(Me.txtA_CzasMin.value), ",", ".")
    If rawCzas = "" Or Not IsNumeric(rawCzas) Then
        MsgBox "Podaj czas postoju w minutach (liczba).", vbExclamation
        Exit Sub
    End If
    czasMin = CDbl(rawCzas)
    If czasMin <= 0 Then
        MsgBox "Czas postoju musi być > 0.", vbExclamation
        Exit Sub
    End If

    Dim wsD As Worksheet
    Set wsD = TryGetWorksheet("data")
    If wsD Is Nothing Then
        MsgBox "Brak arkusza 'data'.", vbExclamation
        Exit Sub
    End If

    ' start kolumny wg zmiany (1->G, 2->P, 3->Z)
    Dim shiftVal As Long
    shiftVal = CLng(val(GetStateOrControl(Me, "Zmiana", "ComboBoxZmiana")))
    If shiftVal < 1 Or shiftVal > 3 Then
        MsgBox "Wybierz zmianę 1/2/3 w głównym formularzu.", vbExclamation
        Exit Sub
    End If

    Dim startCol As Long
    Select Case shiftVal
        Case 1: startCol = wsD.Columns("G").Column ' 7
        Case 2: startCol = wsD.Columns("P").Column ' 16
        Case 3: startCol = wsD.Columns("Z").Column ' 26
    End Select

    ' gdzie wstawiamy (pierwszy pusty w startCol od wiersza 2)
    Dim targetRow As Long
    targetRow = wsD.Cells(wsD.Rows.Count, startCol).End(xlUp).row
    If targetRow < 2 Then
        targetRow = 2
    ElseIf Trim$(CStr(wsD.Cells(2, startCol).value)) = "" Then
        targetRow = 2
    Else
        targetRow = targetRow + 1
    End If

    ' czasy
    Const MINUTES_PER_DAY As Double = 24# * 60#
    Dim dtEnd As Date, dtStart As Date
    dtEnd = Now
    dtStart = dtEnd - (czasMin / MINUTES_PER_DAY)

    ' zapis układem zgodnym z Twoim CopyAlarmValues (8 kolumn: startCol..startCol+7)
    With wsD
        .Cells(targetRow, startCol + 0).value = dtStart                 ' G/P/Z
        .Cells(targetRow, startCol + 1).value = dtEnd                   ' H/Q/AA
        .Cells(targetRow, startCol + 2).value = linia                   ' I/R/AB
        .Cells(targetRow, startCol + 3).value = maszyna                 ' J/S/AC
        .Cells(targetRow, startCol + 4).value = czasMin                 ' K/T/AD
        .Cells(targetRow, startCol + 5).value = opis                    ' L/U/AE
        .Cells(targetRow, startCol + 6).value = Trim$(Me.txtA_Komentarz.value) ' M/V/AF
        .Cells(targetRow, startCol + 7).value = Trim$(Me.txtA_Uwagi.value)     ' N/W/AG
    End With

    MsgBox "Zapisano postój (wiersz " & targetRow & ").", vbInformation
    Awarie_ClearFields
    Exit Sub

EH:
    MsgBox "Błąd zapisu postoju: " & Err.Description, vbExclamation
End Sub

' ---------- Tabela cfg_awarie_long / tblAwarie ----------

Private Function Awarie_GetTable() As ListObject
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = TryGetWorksheet(AW_SHEET)
    If ws Is Nothing Then Exit Function

    Set Awarie_GetTable = ws.ListObjects(AW_TABLE)
    Exit Function
EH:
    Set Awarie_GetTable = Nothing
End Function


