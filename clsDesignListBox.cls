VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDesignListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pHeadersCol As Collection
Private pDesignListBoxObjectsCol As New Collection
Private pFreezeColumnsCol As Collection
Private pFreezeRowsCol As Collection
Private pInpArr() As Variant
Private pParent As UserForm
Private WithEvents pFrame As MSForms.Frame
Attribute pFrame.VB_VarHelpID = -1
Private pParamDict As New Scripting.Dictionary
Private pIsEmpty As Boolean
'Public Events
Public Event BeforeClick()
Public Event Change()
Public Event Click()

'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'Properties
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
Public Property Get IsEmpty() As Boolean

IsEmpty = pIsEmpty

End Property
Public Property Get Frame() As MSForms.Frame

Set Frame = pFrame

End Property
Public Property Get ColumnsCount() As Long

'has to return 0 if empty. No Error Handler required
ColumnsCount = pParamDict("ColumnsCount")

End Property
Public Property Get RowsCount() As Long

'has to return 0 if empty. No Error Handler required
RowsCount = pParamDict("RowsCount")

End Property
Public Property Get RowHeight() As Long

'has to return 0 if empty. No Error Handler required
RowHeight = pParamDict("RowHeight")

End Property
Public Property Let RowHeight(InpHeight As Long)

If InpHeight < 1 Then InpHeight = 1

pParamDict("RowHeight") = InpHeight

Dim AllLabelsCol As Collection
Set AllLabelsCol = Me.AllLabels

Dim Labl
For Each Labl In AllLabelsCol
    Labl.Height = InpHeight
Next Labl

End Property
Public Property Get ColumnWidths() As String

'has to return vbnullstring if empty. No Error Handler required
ColumnWidths = pParamDict("ColumnWidths")

End Property
Public Property Let ColumnWidths(strWidths As String)

Dim k As Long

If Me.ColumnsCount = 0 Then
    pParamDict("ColumnsCount") = UBound(Split(strWidths, ";")) + 1
End If

'Adapt Widths to ColumnsCount
If Me.ColumnsCount > UBound(Split(strWidths, ";")) Then
    Do While Me.ColumnsCount > UBound(Split(strWidths, ";"))
        If Right(strWidths, 1) = ";" Then
            strWidths = strWidths + "0;"
        Else
            strWidths = strWidths + ";0;"
        End If
    Loop
ElseIf Me.ColumnsCount < UBound(Split(strWidths, ";")) Then
    Dim pivotString As String, i As Long
    i = 0
    Do Until Me.ColumnsCount < UBound(Split(pivotString, ";"))
        pivotString = pivotString & Split(strWidths, ";")(i) & ";"
        i = i + 1
    Loop
    strWidths = pivotString
End If

'Eliminate last ;
If Right(strWidths, 1) = ";" Then
    strWidths = Left(strWidths, Len(strWidths) - 1)
End If

For k = 0 To Me.ColumnsCount - 1
    If Not IsNumeric(Split(strWidths, ";")(k)) Then
        Err.Raise 123, "clsDesignListBox", "Input for ColumnWidths is not correct.", "Correct notation example: '10; 10; 10;'"
        Exit Property
    End If
Next k

pParamDict("ColumnWidths") = strWidths

On Error GoTo ErrHandler
If Not UBound(pInpArr, 1) = 0 Then
    Call Me.Fill(pInpArr)
End If

Exit Property
ErrHandler:

End Property
Public Property Get Headers() As Boolean

Headers = pParamDict("Headers")

End Property
Public Property Let Headers(isOn As Boolean)

pParamDict("Headers") = isOn

If Not Me.IsEmpty Then
    Call ActivateDeactivateHeaders(isOn)
    If Me.FreezeRows = 0 Then
        Me.FreezeRows = 1
    End If
End If

End Property
Public Property Get AllLabels() As Collection

Dim ResultCol As New Collection

Dim Labl
For Each Labl In pFrame.Controls
    ResultCol.Add Labl
Next Labl

Set AllLabels = ResultCol

End Property
Public Property Get RowLabels(RowNumber As Long, Optional InludingHeaders As Boolean) As Collection

Dim ResultCol As New Collection

If Not CheckInputRowNumber(RowNumber) Then Exit Property

Dim Labl
For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Or InludingHeaders Then
        If Split(Labl.Name, ";")(0) = RowNumber Then
            ResultCol.Add Labl
        End If
    End If
Next Labl

Set RowLabels = ResultCol

End Property
Public Property Get ColumnLabels(ColumnNumber As Long, Optional InludingHeaders As Boolean) As Collection

If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property

Dim ResultCol As New Collection

Dim Labl
For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Or InludingHeaders Then
        If Split(Labl.Name, ";")(1) = ColumnNumber Then
            ResultCol.Add Labl
        End If
    End If
Next Labl

Set ColumnLabels = ResultCol

End Property
Public Property Get ExactLabel(RowNumber As Long, ColumnNumber As Long) As Object

If Not CheckInputRowNumber(RowNumber) Then Exit Property
If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property

Dim Labl
For Each Labl In pFrame.Controls
    If Split(Labl.Name, ";")(0) = RowNumber And Split(Labl.Name, ";")(1) = ColumnNumber Then
        Set ExactLabel = Labl
        Exit Property
    End If
Next Labl

End Property
Public Property Get HeadersLabels() As Collection

Set HeadersLabels = pHeadersCol

End Property
Public Property Get ColumnSource() As Long

Dim TestNumber As Long
TestNumber = Round(pParamDict("ColumnSource"))

If TestNumber < 0 Then TestNumber = 0
If TestNumber > Me.ColumnsCount Then TestNumber = Me.ColumnsCount

ColumnSource = TestNumber

End Property
Public Property Let ColumnSource(ColumnNumber As Long)

pParamDict("ColumnSource") = Round(ColumnNumber)

End Property
Public Property Get TrueSelectedValue() As Variant

Dim ColumnSourceNumber As Long
ColumnSourceNumber = Me.ColumnSource

Dim Labl
For Each Labl In pFrame.Controls
    If Labl.BackColor = Me.SelectionColor And Split(Labl.Name, ";")(1) = ColumnSourceNumber Then
        SelectedValue = Labl.Caption
    End If
Next Labl

End Property
Public Property Get SelectedValue() As Variant

Dim ColumnSourceNumber As Long
ColumnSourceNumber = Me.ColumnSource + GetDeviation("Columns")

Dim RowNumber As Long
RowNumber = Me.ListIndex + GetDeviation("Rows")

SelectedValue = pInpArr(RowNumber, ColumnSourceNumber)

End Property
Public Property Get SelectionColor() As Long

SelectionColor = pParamDict("SelectionColor")

End Property
Public Property Let SelectionColor(ColorNumber As Long)

If Not Me.IsEmpty Then Call DeselectAllLabels

pParamDict("SelectionColor") = ColorNumber

End Property
Public Property Get ListIndex() As Long

Dim Result As Long
Result = -1

Dim Labl
For Each Labl In pFrame.Controls
    If Labl.BackColor = Me.SelectionColor Then
        Result = Split(Labl.Name, ";")(0)
        Exit For
    End If
Next Labl

ListIndex = Result

End Property
Public Property Let ListIndex(RowNumber As Long)

If RowNumber = -1 Then
    Call DeselectAllLabels
Else
    Call Me.SelectRow(RowNumber)
End If

End Property
Public Property Get FreezeRows() As Long

FreezeRows = pParamDict("FreezeRows")

End Property
Public Property Let FreezeRows(RowsFromTop As Long)

RowsFromTop = Round(RowsFromTop)
If RowsFromTop < 0 Then RowsFromTop = 0
If RowsFromTop > Me.RowsCount Then RowsFromTop = Me.RowsCount

pParamDict("FreezeRows") = RowsFromTop
Set pFreezeRowsCol = New Collection

pFrame.Scroll fmScrollActionBegin, fmScrollActionBegin

If RowsFromTop = 0 Then Exit Property

Dim i As Long
For i = 0 To RowsFromTop - 1
    pFreezeRowsCol.Add Me.RowLabels(i, True)
Next i

Dim Col, Labl

For Each Col In pFreezeColumnsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

For Each Col In pFreezeRowsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

End Property
Public Property Get FreezeColumns() As Long

FreezeColumns = pParamDict("FreezeColumns")

End Property
Public Property Let FreezeColumns(ColumnsFromLeft As Long)

ColumnsFromLeft = Round(ColumnsFromLeft)
If ColumnsFromLeft < 0 Then ColumnsFromLeft = 0
If ColumnsFromLeft > Me.ColumnsCount Then ColumnsFromLeft = Me.ColumnsCount

pParamDict("FreezeColumns") = ColumnsFromLeft
Set pFreezeColumnsCol = New Collection

pFrame.Scroll fmScrollActionBegin, fmScrollActionBegin

If ColumnsFromLeft = 0 Then Exit Property

Dim i As Long
For i = 0 To ColumnsFromLeft - 1
    pFreezeColumnsCol.Add Me.ColumnLabels(i, True)
Next i

Dim Col, Labl

For Each Col In pFreezeColumnsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

For Each Col In pFreezeRowsCol
    For Each Labl In Col
        Labl.ZOrder msoBringToFront
    Next Labl
Next Col

End Property

'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'Events
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////

Private Sub Class_Initialize()

pIsEmpty = True
Me.SelectionColor = RGB(204, 255, 255)
Me.ColumnSource = 0
Set pFreezeColumnsCol = New Collection
Set pFreezeRowsCol = New Collection

End Sub
Private Sub pFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)

Dim Col, Labl

For Each Col In pFreezeColumnsCol
    For Each Labl In Col
        Labl.Left = Labl.Left + ActualDx
        If Labl.Left < 0 Then Labl.Left = 0
    Next Labl
Next Col

For Each Col In pFreezeRowsCol
    For Each Labl In Col
        Labl.Top = Labl.Top + ActualDy
        If Labl.Top < 0 Then Labl.Top = 0
    Next Labl
Next Col

End Sub

'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'Public Procedures
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////

Public Sub Create(Parent As UserForm, inpTop As Long, inpLeft As Long, InpHeight As Long, inpWidth As Long, InpArr As Variant)

Call AddControlsToUserform(Parent, Parent, TypeName(Parent), "Frame", "DesignListBox", InpHeight, inpLeft, inpTop, inpWidth, , , , , , , , , , , , True)

Set pFrame = Parent.DesignListBox
pFrame.ScrollBars = fmScrollBarsBoth

Set pParent = Parent

Call Fill(InpArr)

End Sub
Public Sub Fill(InpArr As Variant)

If Not Right(TypeName(InpArr), 2) = "()" Then
    Err.Raise 123, "clsDesignListBox", "Input is not an Array", "Use Arrays only."
    Exit Sub
End If
    
Call Me.Clear

'Store inpArr
pInpArr = InpArr

'Store Deviation
pParamDict("DeviationRows") = LBound(InpArr, 1)
pParamDict("DeviationColumns") = LBound(InpArr, 2)

Dim ArrayDimensions As Integer
ArrayDimensions = NumberOfArrayDimensions(InpArr)

'ColumnsCount
If ArrayDimensions = 1 Then
    pParamDict("ColumnsCount") = 1
ElseIf ArrayDimensions = 2 Then
    pParamDict("ColumnsCount") = UBound(InpArr, 2) - LBound(InpArr, 2) + 1
Else
    Err.Raise 123, "clsDesignListBox", "Input Array has more than 2 dimensions", "Use 1 or 2 dimension arrays."
End If

'RowsCount
pParamDict("RowsCount") = UBound(InpArr, 1) - LBound(InpArr, 1) + 1

'ColumnWidths
If Me.ColumnWidths = vbNullString Then
    Dim strWidths As String
    Dim strWidth As String
    strWidth = CStr(Round((pFrame.Width - 20) / pParamDict("ColumnsCount"), 2)) 'subtract 20 because of Scroll Bar
    Dim i As Long
    For i = 1 To pParamDict("ColumnsCount")
        strWidths = strWidths + strWidth & ";"
    Next i
    'Eliminate last ;
    If Right(strWidths, 1) = ";" Then
        strWidths = Left(strWidths, Len(strWidths) - 1)
    End If
    pParamDict("ColumnWidths") = strWidths
End If

'RowHeight
pParamDict("RowHeight") = 12

Call ExecuteFill(InpArr)
 
End Sub
Public Sub Sort(ColumnNumber As Long, Optional Descending As Boolean)

If Not CheckInputColumnNumber(ColumnNumber) Then Exit Sub

Call DeselectAllLabels

'If Headers then remove Arrow Char
If Me.Headers Then
    Call SortArray(pInpArr, LBound(pInpArr, 1) + 1, , ColumnNumber + GetDeviation("Columns"), Descending)
Else
    Call SortArray(pInpArr, , , ColumnNumber + GetDeviation("Columns"), Descending)
End If

Dim FoundElementsDict As New Scripting.Dictionary
Dim PivotCol As Collection

'Reorganize Labels
Dim Labl, i As Long
For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Then
        Labl.Tag = GetPositionInArr(Labl.Caption, Split(Labl.Name, ";")(1), FoundElementsDict(Labl.Caption))
        If TypeName(FoundElementsDict(Labl.Caption)) = "Collection" Then
            FoundElementsDict(Labl.Caption).Add Labl.Tag
        Else
            Set PivotCol = New Collection
            PivotCol.Add Labl.Tag
            Set FoundElementsDict(Labl.Caption) = PivotCol
        End If
        Labl.Name = i
    Else
        'Remove or add Arrow Char
        If Right(Labl.Caption, 1) = ChrW(&H2191) Or Right(Labl.Caption, 1) = ChrW(&H2193) Then
            Labl.Caption = Left(Labl.Caption, Len(Labl.Caption) - 2)
        End If
        If CLng(Split(Labl.Name, ";")(1)) = ColumnNumber Then
            If Descending Then
                Labl.Caption = Labl.Caption & Space(1) & ChrW(&H2191)
            Else
                Labl.Caption = Labl.Caption & Space(1) & ChrW(&H2193)
            End If
        End If
    End If
    i = i + 1
Next Labl

For Each Labl In pFrame.Controls
    If Not Labl.Tag = "Header" Then
        Labl.Name = Labl.Tag
        Labl.Top = (Split(Labl.Name, ";")(0)) * Me.RowHeight
        'Debug.Assert Labl.Top <> 0
    End If
Next Labl

pFrame.Repaint

End Sub
Public Sub SelectRow(RowNumber As Long)

If Not CheckInputRowNumber(RowNumber) Then Exit Sub

Call DeselectAllLabels
Call SelectCorrectLabels(RowNumber)

RaiseEvent Change

End Sub
Public Sub Clear()

Dim Ctrl As Control
For Each Ctrl In pFrame.Controls
    pFrame.Controls.Remove Ctrl.Name
Next Ctrl

pIsEmpty = True

End Sub
Public Sub RaiseEventBeforeClick()

RaiseEvent BeforeClick

End Sub
Public Sub RaiseEventClick()

RaiseEvent Click

End Sub

'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'Private Procedures
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
Private Sub ExecuteFill(ByVal InpArr As Variant)

Dim i As Long, j As Long, k As Long
Dim setTop, setLeft, setHeight, setWidth
Dim CreatedLabel As MSForms.Label, InpBorderStyle As Long, InpIsBold As Boolean
Dim DLBObject As clsDesignListBoxObject

If Me.Headers = True Then
    Set pHeadersCol = New Collection
End If

For i = LBound(InpArr, 1) To UBound(InpArr, 1)
    For j = LBound(InpArr, 2) To UBound(InpArr, 2)
        setTop = (i - GetDeviation("Rows")) * Me.RowHeight
        setHeight = Me.RowHeight
        setLeft = 0
        For k = LBound(InpArr, 2) To j
            If k = LBound(InpArr, 2) Then
                setLeft = 0
            ElseIf k - GetDeviation("Columns") <= UBound(Split(Me.ColumnWidths, ";")) Then
                setLeft = setLeft + CLng(Split(Me.ColumnWidths, ";")(k - GetDeviation("Columns") - 1))
            End If
        Next k
        If j - GetDeviation("Columns") <= UBound(Split(Me.ColumnWidths, ";")) Then
            setWidth = CLng(Split(Me.ColumnWidths, ";")(j - GetDeviation("Columns")))
        Else
            setWidth = 0
        End If
        If setWidth = 0 Then GoTo NextIteration
        If Me.Headers = True And i = LBound(InpArr, 1) Then
            InpBorderStyle = 1
            InpIsBold = True
        Else
            InpBorderStyle = 0
            InpIsBold = False
        End If
        Set CreatedLabel = AddControlsToUserform(pParent, pParent.DesignListBox, "DesignListBox", "Label", i - GetDeviation("Rows") & ";" & j - GetDeviation("Columns"), setHeight, setLeft, setTop, setWidth, InpArr(i, j), InpArr(i, j), , , , , , , , InpBorderStyle, , , InpIsBold)
        If Me.Headers = True And i = LBound(InpArr, 1) Then
            pHeadersCol.Add CreatedLabel
            CreatedLabel.Tag = "Header"
        End If
        Set DLBObject = New clsDesignListBoxObject
        DLBObject.Create CreatedLabel, Me
        pDesignListBoxObjectsCol.Add DLBObject
NextIteration:
    Next j
Next i

pIsEmpty = False

If Me.Headers Then
    Me.FreezeRows = 1
End If

pFrame.ScrollWidth = setLeft + 10
pFrame.ScrollHeight = setTop + Me.RowHeight

End Sub
Private Function GetDeviation(Mode As String) As Long
'Mode can be "Rows" or "Columns"

If Mode = "Rows" Then
    GetDeviation = pParamDict("DeviationRows")
ElseIf Mode = "Columns" Then
    GetDeviation = pParamDict("DeviationColumns")
End If



End Function
Private Sub ActivateDeactivateHeaders(isOn As Boolean)

Dim Labl

If isOn Then
    Set pHeadersCol = New Collection
    Set pHeadersCol = Me.RowLabels(0)
    For Each Labl In pHeadersCol
        Labl.BorderStyle = 1
        Labl.Font.Bold = True
        Labl.Tag = "Header"
    Next Labl
Else
    For Each Labl In pHeadersCol
        Labl.BorderStyle = 0
        Labl.Font.Bold = False
        Labl.Tag = ""
    Next Labl
    Set pHeadersCol = New Collection
End If

End Sub
Private Sub SelectCorrectLabels(RowNumber As Long)

Dim Labl
For Each Labl In pParent.Controls
    If Split(Labl.Name, ";")(0) = RowNumber Then
        Labl.BackColor = Me.SelectionColor
    End If
Next Labl

End Sub
Private Sub DeselectAllLabels()

Dim Labl
For Each Labl In pFrame.Controls
    If Labl.BackColor = Me.SelectionColor Then
        Labl.BackColor = vbWhite
    End If
Next Labl

End Sub
Private Function CheckInputRowNumber(ByRef RowNumber As Long) As Boolean

Err.Clear

RowNumber = Round(RowNumber)

If Me.IsEmpty Then
    Call IsEmptyErr
ElseIf RowNumber < 0 Then
    Err.Raise 123, "clsDesignListBox", "RowNumber must be bigger than 0"
ElseIf RowNumber > Me.RowsCount - 1 Then
    Err.Raise 123, "clsDesignListBox", "RowNumber must be smaller than " & Me.RowsCount - 1, "Rows are zero based"
End If

If Err.Number = 0 Then
    CheckInputRowNumber = True
End If

End Function
Private Function CheckInputColumnNumber(ByRef ColumnNumber As Long) As Boolean

Err.Clear

ColumnNumber = Round(ColumnNumber)

If Me.IsEmpty Then
    Call IsEmptyErr
ElseIf ColumnNumber < 0 Then
    Err.Raise 123, "clsDesignListBox", "ColumnNumber must be bigger than 0"
ElseIf ColumnNumber > Me.ColumnsCount - 1 Then
    Err.Raise 123, "clsDesignListBox", "ColumnNumber must be smaller than " & Me.ColumnsCount - 1, "Columns are zero based"
End If

If Err.Number = 0 Then
    CheckInputColumnNumber = True
End If


End Function
Private Sub IsEmptyErr()

Err.Raise 123, "clsDesignListBox", "DesignListBox is empty", "Use .create or .fill to instantiate DesignListBox"

End Sub

'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'Special Procedures
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////

Private Function NumberOfArrayDimensions(arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(arr, Ndx)
Loop Until Err.Number <> 0
Err.Clear

NumberOfArrayDimensions = Ndx - 1

End Function
Private Function AddControlsToUserform(UForm As UserForm, OutpLocation As Object, strOutputLocation As String, InpType, InpName, InpHeight, inpLeft, inpTop, _
inpWidth, Optional InpCaption, Optional InpValue, Optional InpTextAlign, Optional InpTextSize, _
Optional InpForeColor, Optional InpItems As Collection, Optional InpListIndex, Optional InpTipText, _
Optional InpSpecialEffect, Optional InpBorderStyle, Optional InpVisibility As Boolean = True, Optional InpParent As Boolean, Optional InpIsBold) As MSForms.Control
'InpType can be: CheckBox, ComboBox, CommandButton,Frame,Image,Label,ListBox,MultiPage,OptionButton,ScrollBar
'SpinButton,TabStrip,TextBox,ToggleButton
'''''''''''''''''''''''
'If a problem with adding objects to a frame arises (automation error), then its probably because the frame is not the last
'created frame on the userform. The frame to fill has to be the last created (on runtime or on designtime)
'''''''''''''''''''''''
'ControlClass defines the WithEvents Class assigned to the control. 0 means, that no class is assigned

Dim ProgId As String
ProgId = "Forms." & InpType & ".1"

Dim NewForm, ControlToEdit

If Not InpParent Then
    Set NewForm = OutpLocation.Controls.Add(ProgId, InpName)
Else
    Set NewForm = OutpLocation.Add(ProgId, InpName)
End If
Set ControlToEdit = NewForm

With ControlToEdit
    On Error Resume Next
    .Height = InpHeight
    .Left = inpLeft
    .Top = inpTop
    .Width = inpWidth
    If Not IsMissing(InpCaption) Then
        .Caption = InpCaption
    End If
    If Not IsMissing(InpValue) Then
        .Value = InpValue
    End If
    If Not IsMissing(InpTextAlign) Then
        .TextAlign = InpTextAlign
    End If
    If Not IsMissing(InpTextSize) Then
        .Font.Size = InpTextSize
    End If
    If Not IsMissing(InpForeColor) Then
        .ForeColor = InpForeColor
    End If
    If Not IsMissing(InpItems) Then
        Dim Item
        For Each Item In InpItems
            .AddItem Item
        Next Item
    End If
    If Not IsMissing(InpListIndex) Then
        .ListIndex = InpListIndex
    End If
    If Not IsMissing(InpTipText) Then
        .ControlTipText = InpTipText
        .MousePointer = fmMousePointerHelp
    End If
    If Not IsMissing(InpSpecialEffect) Then
        .SpecialEffect = InpSpecialEffect
    End If
    If Not IsMissing(InpBorderStyle) Then
        .BorderStyle = InpBorderStyle
    End If
    .Font.Bold = InpIsBold
    .Visible = InpVisibility
End With

OutpLocation.Repaint

Set AddControlsToUserform = ControlToEdit

End Function
Private Sub SortArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0, Optional Descending As Boolean)

If NumberOfArrayDimensions(InpArr) = 1 Then
    Call QuickSortSingleDimArray(InpArr, lngMin, lngMax)
Else
    Call QuickSortMultiDimArray(InpArr, lngMin, lngMax, lngColumn)
End If

If Descending Then
    Call ReverseArrayInPlace(InpArr, lngMin, lngMax)
End If

End Sub
Private Sub QuickSortSingleDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next

    'Sort a 1-Dimensional array

    ' SampleUsage: sort arrData
    '
    '   QuickSortSingleDimArray arrData

    '
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming


    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j

        While InpArr(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < InpArr(j) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the item
            varX = InpArr(i)
            InpArr(i) = InpArr(j)
            InpArr(j) = varX

            i = i + 1
            j = j - 1
        End If

    Wend

    If (lngMin < j) Then Call QuickSortSingleDimArray(InpArr, lngMin, j)
    If (i < lngMax) Then Call QuickSortSingleDimArray(InpArr, i, lngMax)

End Sub
Private Sub QuickSortMultiDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While InpArr(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < InpArr(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(InpArr, 2) To UBound(InpArr, 2))
            For lngColTemp = LBound(InpArr, 2) To UBound(InpArr, 2)
                arrRowTemp(lngColTemp) = InpArr(i, lngColTemp)
                InpArr(i, lngColTemp) = InpArr(j, lngColTemp)
                InpArr(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortMultiDimArray(InpArr, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortMultiDimArray(InpArr, i, lngMax, lngColumn)
    
End Sub
Private Function ReverseArrayInPlace(InputArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, _
    Optional Reversed As Boolean = False, Optional DoubleReversed As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' If Reversed is true then it does reorder it in the other dimension
' in the calling procedure is reversed. This works only on arrays
' of simple data types (String, Single, Double, Integer, Long). It will not work
' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long
Dim i

If lngMin = -1 Then lngMin = LBound(InputArray, 1)
If lngMax = -1 Then lngMax = UBound(InputArray, 1)

'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

Ndx2 = lngMax
''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
If Reversed = False Then
    For Ndx = lngMin To ((lngMax - lngMin + 1) \ 2)
        For i = LBound(InputArray, 2) To UBound(InputArray, 2)
            'swap the elements
            Temp = InputArray(Ndx, i)
            InputArray(Ndx, i) = InputArray(Ndx2, i)
            InputArray(Ndx2, i) = Temp
        Next i
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
    If DoubleReversed = True Then GoTo Rev
ElseIf Reversed = True Or DoubleReversed = True Then
Rev:
    Ndx2 = UBound(InputArray, 2)
    For Ndx = LBound(InputArray, 2) To Int(((UBound(InputArray, 2) - LBound(InputArray, 2) + 1) \ 2))
        For i = LBound(InputArray, 1) To UBound(InputArray, 1)
            'swap the elements
            Temp = InputArray(i, Ndx)
            InputArray(i, Ndx) = InputArray(i, Ndx2)
            InputArray(i, Ndx2) = Temp
        Next i
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
End If

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayInPlace = True

End Function

Private Function GetPositionInArr(Content As Variant, ColumnNumber, LastRowInput As Variant) As String

Dim PivotCol As Collection

Dim i, Element
For i = LBound(pInpArr, 1) To UBound(pInpArr, 1)
    If CStr(Content) = CStr(pInpArr(i, ColumnNumber + GetDeviation("Columns"))) Then
        If TypeName(LastRowInput) = "Collection" Then
            Set PivotCol = LastRowInput
            If Not IsInCollection(PivotCol, , i - GetDeviation("Rows") & ";" & ColumnNumber) Then
                GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
                'Debug.Assert i - GetDeviation("Rows") <> 0
                Exit Function
            End If
        Else
            GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
            'Debug.Assert i - GetDeviation("Rows") <> 0
            Exit Function
        End If
    End If
Next i

Debug.Print "Result: None"

End Function
Private Function IsInCollection(Kollection As Collection, Optional Key As Variant, Optional Item As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Kollection, the collection we would like to examine
'            : (Optional) Key, the Key we want to find in the collection
'            : (Optional) Item, the Item we want to find in the collection
'OUTPUT      : True if Key or Item is found, False if not
'SPECIAL CASE: If both Key and Item are missing, return False
    
    Dim strKey As String
    Dim var As Variant

    'First, investigate assuming a Key was provided
    If Not IsMissing(Key) Then
    
        strKey = CStr(Key)
        
        'Handling errors is the strategy here
        On Error Resume Next
            IsInCollection = True
            var = Kollection(strKey) '<~ this is where our (potential) error will occur
            If Err.Number = 91 Then GoTo CheckForObject
            If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function

CheckForObject:
        If IsObject(Kollection(strKey)) Then
            IsInCollection = True
            On Error GoTo 0
            Exit Function
        End If

NotFound:
        IsInCollection = False
        On Error GoTo 0
        Exit Function
        
    'If the Item was provided but the Key was not, then...
    ElseIf Not IsMissing(Item) Then
    
        IsInCollection = False '<~ assume that we will not find the item
    
        'We have to loop through the collection and check each item against the passed-in Item
        For Each var In Kollection
            If TypeName(var) = "clsSymbol" Then var = var.ID
            If TypeName(var) = "clsDividedSymbol" Then var = var.CombinedId
            If var = Item Then
                IsInCollection = True
                Exit Function
            End If
        Next var
    
    'Otherwise, no Key OR Item was provided, so we default to False
    Else
        IsInCollection = False
    End If
    
End Function

