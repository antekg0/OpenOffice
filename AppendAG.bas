Rem  *****  OpenOffice Calc BASIC  *****
Option Explicit

Private document, dispatcher As Object
Private args(1) As new com.sun.star.beans.PropertyValue
Private cEMPTY%

Sub AppendAG
Rem The range of data must be separated by empty columns (left and right). Left edge of the sheet can be a border of the range. 
Rem The final column is the same As selected before start the macro

Rem ----------------------------------------------------------------------
Rem define variables
Dim cDATE%, NullDateCorrection%, StartCol%

Rem Correction to the date related to the NullDate issue As recommended by Villeroy https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=33400
With ThisComponent.NullDate
  NullDateCorrection = DateSerial(1899, 12, 30) - DateSerial(.Year, .Month, .Day)
End With

With com.sun.star
  cEMPTY = .table.CellContentType.EMPTY
  cDATE = .util.NumberFormat.DATE
End With
 
Rem ----------------------------------------------------------------------
Rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

Rem ----------------------------------------------------------------------

  fDispatch "Deselect" 'the statement makes the current selection to be reduced to an active cell.
rem ThisComponent.CurrentController.Select(ThisComponent.createInstance("com.sun.star.sheet.SheetCellRanges"))
With ThisComponent.CurrentSelection
  StartCol = .CellAddress.Column  

Rem testing the first column and move there
  Dim NotFirstCol As Boolean
  args(0).Name = "By"
  args(0).Value = 1
  args(1).Name = "Sel"
  args(1).Value = false
  fDispatch "GoLeft"
  NotFirstCol = IsCurrentCellNotEmpty
  fDispatch "GoRight"
  If NotFirstCol Then fDispatch "GoLeftToStartOfData"

Rem ----------------------------------------------------------------------
  args(1).Value = true
  fDispatch "GoRightToEndOfData"
  
Rem ----------------------------------------------------------------------
  fDispatch "Copy"

Rem ----------------------------------------------------------------------
  args(1).Value = false
  fDispatch "GoLeftToStartOfData"

Rem Move to the free line
  fDispatch "GoDown"
  If IsCurrentCellNotEmpty then
	Rem Go one row up to have always not empty row below
	fDispatch "GoUp"
	fDispatch "GoDownToEndOfData"
	fDispatch "GoDown"
  End If

Rem ----------------------------------------------------------------------
  fDispatch "Paste"

Rem Remove selection -> Select one cell 
  args(0).Value = 0
  fDispatch "GoRight"

Rem Finding the date in a row by moving right and assigning the current date
  args(0).Value = 1
  Do While IsCurrentCellNotEmpty
	If (ThisComponent.NumberFormats.getByKey(.NumberFormat).Type and cDATE) <> 0 Then
	  .Value = Date + NullDateCorrection
	  Exit Do
	End If
	fDispatch "GoRight"
  Loop

Rem Move to final position = initial column
  args(0).Value = StartCol - .CellAddress.Column
  fDispatch "GoRight"
  
End With

End Sub

Sub fDispatch(Cmnd As String)
  dispatcher.executeDispatch(document, ".uno:"& Cmnd, "", 0, args())
End Sub

Function IsCurrentCellNotEmpty As Boolean
  IsCurrentCellNotEmpty = ThisComponent.CurrentSelection.Type <> cEmpty
End Function