OrgtreeMaker.vba

Private rowOrgTree As Integer 'Keeps track of which row in OrgTree we should write to
Private colManager As Integer 'Keeps track of which column the Managers are listed
Private colName As Integer 'Keeps track of which column the Name are listed
Private colEmail As Integer 'Keeps track of which column the Email are listed
Private maxCol As Integer 'Keeps track of the total amount of columns being used
Private wsEI As Worksheet ' refers to employee info
Private wsOT As Worksheet ' refers to org tree
Private initTitle As String ' stores A1 title name

Sub Setup() 'sets most of the variables in order to get started

 rowOrgTree = 1
 maxCol = 0
 colManager = 0
 colName = 0
 colEmail = 0

 Set wsEI = Worksheets("Employee Info")
 Set wsOT = Worksheets("Org Tree")
 wsOT.UsedRange.ClearContents ' deletes any tree that was previously created in OrgTree

 initTitle = wsEI.Cells(1, 1).Value 'storing the initial column title

 For i = 1 To wsEI.Columns.Count 'traverses through each column and stops when it reaches a blank
    If (wsEI.Cells(1, i).Value = "Full_Name") Then
        colName = i 'storing column of name
        maxCol = maxCol + 1
    ElseIf (wsEI.Cells(1, i).Value = "Email_Address") Then
        colEmail = i 'storing column of email
        maxCol = maxCol + 1
    ElseIf (wsEI.Cells(1, i).Value = "Manager") Then
        colManager = i 'storing column of manager
        maxCol = maxCol + 1
    ElseIf (IsEmpty(wsEI.Cells(1, i)) <> True) Then
        maxCol = maxCol + 1
    Else
        i = wsEI.Columns.Count 'stops the for loop
    End If
 Next i

 If (maxCol = 0 Or colManager = 0 Or colName = 0 Or colEmail = 0) Then
     MsgBox ("Something is wrong, check column titles")
 Else
     Call FirstCall("Boss", 1)
 End If
End Sub

Sub FirstCall(CEOname As String, colOrgTree As Integer) 'creates the first spot in the OrgTree
Dim firstCol As String 'stores address of A1
Dim lastCol As String 'stores address of whatever the last used column in A is
Dim rangeCol As String 'will contain a string of "A1:Alastcolumn"

firstCol = wsEI.Cells(1, 1).Address()
lastCol = wsEI.Cells(1, maxCol).Address()
rangeCol = firstCol + ":" + lastCol
rangeCol = Replace(rangeCol, "$", "") 'removing $ from the string to be used by AutoFilter

'removes filter
If wsEI.AutoFilterMode Then wsEI.AutoFilter.ShowAllData

'filters by whatever name you stored in CEOname
wsEI.Range(rangeCol).AutoFilter Field:=colName, Criteria1:="=" + CEOname

 Dim r As Range
 'goes through each of the now filtered ranges and writes name and email to OrgTree
    For Each r In wsEI.UsedRange.SpecialCells(xlCellTypeVisible).Rows 'Enables traversal of only used cells
        'Stops the first row and any blank rows from being written to the OrgTree
        If (wsEI.Cells(r.Row, 1) <> initTitle And IsEmpty(wsEI.Cells(r.Row, 1)) <> True) Then
            wsOT.Cells(rowOrgTree, colOrgTree).Value = wsEI.Cells(r.Row, colName) 'Write name to Orgtree
            wsOT.Cells(rowOrgTree, colOrgTree + 1).Value = wsEI.Cells(r.Row, colEmail) 'Write Email To OrgTree
            Call FillIn(CEOname, colOrgTree + 3) 'Go to the next level and print all subordinates, leaving space for Batch#
        End If
    Next r
    
 'removes filter
If wsEI.AutoFilterMode Then wsEI.AutoFilter.ShowAllData
End Sub

Sub FillIn(managerName As String, colOrgTree As Integer) 'Fills in the rest of the OrgTree
Dim firstCol As String 'stores address of A1
Dim lastCol As String 'stores address of whatever the last used column in A is
Dim rangeCol As String 'will contain a string of "A1:Alastcolumn"
Dim temp As Integer 'will store temp rowOrgTree

temp = rowOrgTree

firstCol = wsEI.Cells(1, 1).Address()
lastCol = wsEI.Cells(1, maxCol).Address()
rangeCol = firstCol + ":" + lastCol
rangeCol = Replace(rangeCol, "$", "") 'removing $ from the string to be used by AutoFilter

'removes filter
If wsEI.AutoFilterMode Then wsEI.AutoFilter.ShowAllData

'filters by whatever name you stored in managerName
wsEI.Range(rangeCol).AutoFilter Field:=colManager, Criteria1:="=" + managerName

 Dim r As Range
 'goes through each of the now filtered ranges and writes name and email to OrgTree
    For Each r In wsEI.UsedRange.SpecialCells(xlCellTypeVisible).Rows
        'Stops the first row and any blank rows from being written to the OrgTree
        If (wsEI.Cells(r.Row, 1) <> initTitle And IsEmpty(wsEI.Cells(r.Row, colName)) <> True) Then
            wsOT.Cells(rowOrgTree, colOrgTree).Value = wsEI.Cells(r.Row, colName)
            wsOT.Cells(rowOrgTree, colOrgTree + 1).Value = wsEI.Cells(r.Row, colEmail)
            Call FillIn(wsEI.Cells(r.Row, colName).Value, colOrgTree + 3) 'Go to the next level and print all subordinates, leaving space for Batch#
        End If
    Next r
    
    If (temp = rowOrgTree) Then    'If the children of this row didn't have any children, then you must go down a column
            rowOrgTree = rowOrgTree + 1
        End If
 'removes filter
If wsEI.AutoFilterMode Then wsEI.AutoFilter.ShowAllData
End Sub


