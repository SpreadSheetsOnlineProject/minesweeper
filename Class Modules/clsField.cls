VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem Define where to start the game board
Private Enum enumBoardStartAt
    start_at_row = 4
    start_at_col = 17
End Enum

Private m_row_num As Integer
Private m_col_num As Integer
Private m_number_of_mines As Integer
Private m_number_of_remaining_mines As Integer
Private m_field As Variant
Private m_untouched_fields As Integer

Private m_board_start_row As Integer
Private m_board_start_col As Integer

' ----------------------------------------------------------------
' Procedure Name: setTouchedCells
' Purpose: Set up class variables and create the array of the field
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter row_size (Integer): Number of rows in the field.
' Parameter col_size (Integer): Number of columns in the field.
' Parameter mine_number (Integer): Number of mines in the field.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub setField(ByVal row_size As Integer, _
                    ByVal col_size As Integer, _
                    ByVal mine_number As Integer)

    m_row_num = row_size
    m_col_num = col_size
    m_number_of_mines = mine_number
    m_number_of_remaining_mines = mine_number
    m_untouched_fields = row_size * col_size - mine_number
    
    m_board_start_row = enumBoardStartAt.start_at_row
    m_board_start_col = enumBoardStartAt.start_at_col
    
    Rem Set random mines on the field
    Call setMines
    Rem Calculate numbers of the field based on mines
    Call setNumbers

End Sub

' ----------------------------------------------------------------
' Procedure Name: isValidField (Boolean)
' Purpose: Return true if the selected range inside of the game board.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter range_ (Range): The selected cell
' Parameter col_size (Integer): Number of columns in the field.
' Parameter mine_number (Integer): Number of mines in the field.
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function isValidField(ByRef range_ As Range) As Boolean

    Rem If range contains multiple cells then return false
    If range_.Count > 1 Then
        isValidField = False
        Exit Function
    End If
    
    Dim c_row As Integer: c_row = range_.row
    Dim c_col As Integer: c_col = range_.Column
    
    If c_row < LBound(m_field, 1) Then Exit Function
    If c_col < LBound(m_field, 2) Then Exit Function
    If c_row > UBound(m_field, 1) Then Exit Function
    If c_col > UBound(m_field, 2) Then Exit Function
    
    isValidField = True

End Function

' ----------------------------------------------------------------
' Procedure Name: getCellValue (Integer)
' Purpose: Return the value of the selected cell of field.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter range_ (Range): The selected cell
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function getCellValue(ByRef range_ As Range) As Integer

    Rem If the range contains multiple cells, then throw an error.
    If range_.Count > 1 Then Err.Raise vbObjectError + 533, "SSOP-minesweeper", "Single range is allowed!", Erl

    getCellValue = m_field(range_.row, range_.Column)

End Function

' ----------------------------------------------------------------
' Procedure Name: isMine (Boolean)
' Purpose: Return true if the selected cell is mine.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter range_ (Range): The selected cell
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function isMine(ByRef range_ As Range) As Boolean

    Rem If the range contains multiple cells, then throw an error.
    If range_.Count > 1 Then Err.Raise vbObjectError + 533, "SSOP-minesweeper", "Single range is allowed!", Erl

    isMine = IIf(m_field(range_.row, range_.Column) = -1, True, False)

End Function

' ----------------------------------------------------------------
' Procedure Name: getMines (Collection)
' Purpose: Return a collection which contains range of the field where mines inside.
' Procedure Kind: Function
' Procedure Access: Public
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function getMines() As Collection

    Set getMines = New Collection
    
    Dim lv_row As Integer
    Dim lv_col As Integer
    
    Rem Loop through the whole field, and add range where mine found.
    For lv_row = LBound(m_field, 1) To UBound(m_field, 1)
    
        For lv_col = LBound(m_field, 2) To UBound(m_field, 2)
        
            If m_field(lv_row, lv_col) = -1 Then
                getMines.Add shField.Range(Cells(lv_row, lv_col), Cells(lv_row, lv_col))
            End If
        
        Next lv_col
    
    Next lv_row

End Function

' ----------------------------------------------------------------
' Procedure Name: collItemExists (Boolean)
' Purpose: Return true if a collection has the key.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter coll (Collection): Collection, where the search the key
' Parameter key (String): The key
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Private Function collItemExists(ByRef coll As Collection, _
                                ByVal key As String) As Boolean

On Error GoTo eh
    
    coll.Item key
    
done:
    collItemExists = True
    Exit Function
eh:
    collItemExists = False

End Function

' ----------------------------------------------------------------
' Procedure Name: getEmptyCells
' Purpose: Recursive function to fetch empty cells and cells with number around the empty cells.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter lv_row (Integer): The actual loop variable for row
' Parameter lv_col (Integer): The actual loop variable for column
' Parameter coll (Collection): The collection which store empty and sorrounding cells without mines
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Function getEmptyCells(ByVal lv_row As Integer, _
                              ByVal lv_col As Integer, _
                              ByRef coll As Collection)

    Rem This will be the key of the collection item.
    Dim coll_key As String: coll_key = lv_row & "," & lv_col

    Rem If the value of the field is number, then store the cell.
    If m_field(lv_row, lv_col) > 0 Then
        Rem If the item does not exist in the collection, has been neven found, then add to the collection
        If Not collItemExists(coll, coll_key) Then
            coll.Add shField.Range(Cells(lv_row, lv_col), Cells(lv_row, lv_col)), coll_key
        End If
        Rem I do not want to go further, I looking for the edge of the empty cells
        Exit Function
    Rem If the current cell is empty, then add to collection and continue recursive function.
    ElseIf m_field(lv_row, lv_col) = 0 Then
        Rem If the item does not exist in the collection, has been neven found, then add to the collection
        If Not collItemExists(coll, coll_key) Then
            coll.Add shField.Range(Cells(lv_row, lv_col), Cells(lv_row, lv_col)), coll_key
        End If
    End If

    Rem go down
    If lv_row + 1 <= UBound(m_field, 1) Then
        If Not collItemExists(coll, lv_row + 1 & "," & lv_col) Then
            Call getEmptyCells(lv_row + 1, lv_col, coll)
        End If
    End If
    Rem go right
    If lv_col + 1 <= UBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row & "," & lv_col + 1) Then
            Call getEmptyCells(lv_row, lv_col + 1, coll)
        End If
    End If
    Rem go up
    If lv_row - 1 >= LBound(m_field, 1) Then
        If Not collItemExists(coll, lv_row - 1 & "," & lv_col) Then
            Call getEmptyCells(lv_row - 1, lv_col, coll)
        End If
    End If
    Rem go left
    If lv_col - 1 >= LBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row & "," & lv_col - 1) Then
            Call getEmptyCells(lv_row, lv_col - 1, coll)
        End If
    End If
    Rem go right down
    If lv_row + 1 <= UBound(m_field, 1) And lv_col + 1 <= UBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row + 1 & "," & lv_col + 1) Then
            Call getEmptyCells(lv_row + 1, lv_col + 1, coll)
        End If
    End If
    Rem go left down
    If lv_row - 1 >= LBound(m_field, 1) And lv_col - 1 >= LBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row - 1 & "," & lv_col - 1) Then
            Call getEmptyCells(lv_row - 1, lv_col - 1, coll)
        End If
    End If
    Rem go left up
    If lv_row + 1 <= UBound(m_field, 1) And lv_col - 1 >= LBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row + 1 & "," & lv_col - 1) Then
            Call getEmptyCells(lv_row + 1, lv_col - 1, coll)
        End If
    End If
    Rem go right up
    If lv_row - 1 >= LBound(m_field, 1) And lv_col + 1 <= UBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row - 1 & "," & lv_col + 1) Then
            Call getEmptyCells(lv_row - 1, lv_col + 1, coll)
        End If
    End If
    Rem go right up
    If lv_row - 1 >= UBound(m_field, 1) And lv_col + 1 <= UBound(m_field, 2) Then
        If Not collItemExists(coll, lv_row + 1 & "," & lv_col + 1) Then
            Call getEmptyCells(lv_row + 1, lv_col + 1, coll)
        End If
    End If
    
End Function

' ----------------------------------------------------------------
' Procedure Name: setMines
' Purpose: Randomly find the palces of mines.
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Private Sub setMines()

    Rem Size of the array
    Dim upper_row As Integer: upper_row = m_row_num + enumBoardStartAt.start_at_row - 1
    Dim upper_col As Integer: upper_col = m_col_num + enumBoardStartAt.start_at_col - 1
    Dim lower_row As Integer: lower_row = enumBoardStartAt.start_at_row
    Dim lower_col As Integer: lower_col = enumBoardStartAt.start_at_col

    Rem Set up the size and the start index of the field array.
    ReDim m_field(lower_row To upper_row, lower_col To upper_col)
    
    Dim c_mine_number As Integer
    
    Dim lv_row As Integer
    Dim lv_col As Integer
    
    Rem Loop until the number of placed mines reach the defined number of mines fot the current table size
    Do Until c_mine_number = m_number_of_mines
        
        Rem Pick a row and column where to place the mine
        lv_row = Int((upper_row - lower_row + 1) * Rnd + lower_row)
        lv_col = Int((upper_col - lower_col + 1) * Rnd + lower_col)
        
        Rem Check if the cell has already contain a mine
        If m_field(lv_row, lv_col) <> -1 Then
            c_mine_number = c_mine_number + 1
            m_field(lv_row, lv_col) = -1
        End If
        
    Loop

End Sub

' ----------------------------------------------------------------
' Procedure Name: setNumbers
' Purpose: Set value of each cell based on the number of mines around them
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Private Sub setNumbers()

    Dim lv_row As Integer
    Dim lv_col As Integer
    
    Dim a_row As Integer
    Dim b_row As Integer
    Dim l_col As Integer
    Dim r_col As Integer
    
    Rem Loop through the whole field
    For lv_row = LBound(m_field, 1) To UBound(m_field, 1)
    
        For lv_col = LBound(m_field, 2) To UBound(m_field, 2)
        
            Rem If the current cell contains mine then add 1 to sorrounding cells
            If m_field(lv_row, lv_col) = -1 Then
                
                Rem row above
                a_row = IIf(lv_row - 1 >= LBound(m_field, 1), lv_row - 1, lv_row)
                Rem row below
                b_row = IIf(lv_row + 1 <= UBound(m_field, 1), lv_row + 1, lv_row)
                Rem col left
                l_col = IIf(lv_col - 1 >= LBound(m_field, 2), lv_col - 1, lv_col)
                Rem col right
                r_col = IIf(lv_col + 1 <= UBound(m_field, 2), lv_col + 1, lv_col)
                
                Rem above left
                If a_row <> lv_row And l_col <> lv_col Then
                    m_field(a_row, l_col) = IIf(m_field(a_row, l_col) <> -1, m_field(a_row, l_col) + 1, m_field(a_row, l_col))
                End If
                Rem above center
                If a_row <> lv_row Then
                    m_field(a_row, lv_col) = IIf(m_field(a_row, lv_col) <> -1, m_field(a_row, lv_col) + 1, m_field(a_row, lv_col))
                End If
                Rem above right
                If a_row <> lv_row And r_col <> lv_col Then
                    m_field(a_row, r_col) = IIf(m_field(a_row, r_col) <> -1, m_field(a_row, r_col) + 1, m_field(a_row, r_col))
                End If
                Rem center left
                If l_col <> lv_col Then
                    m_field(lv_row, l_col) = IIf(m_field(lv_row, l_col) <> -1, m_field(lv_row, l_col) + 1, m_field(lv_row, l_col))
                End If
                Rem center right
                If r_col <> lv_col Then
                    m_field(lv_row, r_col) = IIf(m_field(lv_row, r_col) <> -1, m_field(lv_row, r_col) + 1, m_field(lv_row, r_col))
                End If
                Rem below left
                If b_row <> lv_row And l_col <> lv_col Then
                    m_field(b_row, l_col) = IIf(m_field(b_row, l_col) <> -1, m_field(b_row, l_col) + 1, m_field(b_row, l_col))
                End If
                Rem below center
                If b_row <> lv_row Then
                    m_field(b_row, lv_col) = IIf(m_field(b_row, lv_col) <> -1, m_field(b_row, lv_col) + 1, m_field(b_row, lv_col))
                End If
                Rem below right
                If b_row <> lv_row And r_col <> lv_col Then
                    m_field(b_row, r_col) = IIf(m_field(b_row, r_col) <> -1, m_field(b_row, r_col) + 1, m_field(b_row, r_col))
                End If
                
            End If
        
        Next lv_col
    
    Next lv_row

End Sub

' ----------------------------------------------------------------
' Procedure Name: fieldToSheet
' Purpose: Cheet function to show the field on a sheet.
' Procedure Kind: Function
' Procedure Access: Public
' Parameter target_sheet (Worksheet): Where to show the field
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Public Sub fieldToSheet(ByRef target_sheet As Worksheet)

    Rem If the target is the gameboard itself then throw error
    If target_sheet.Name = shField.Name Then Err.Raise vbObjectError + 533, "SSOP-minesweeper", "This sheet is not allowed!", Erl

    Dim lv_row As Integer
    Dim lv_col As Integer
    
    For lv_row = LBound(m_field, 1) To UBound(m_field, 1)
    
        For lv_col = LBound(m_field, 2) To UBound(m_field, 2)
        
            target_sheet.Cells(lv_row, lv_col).value = m_field(lv_row, lv_col)
        
        Next lv_col
        
    Next lv_row

End Sub


' ----------------------------------------------------------------
' Properties
' Author: Zoltan Sepa
' Date: 04-Mar-2022
' ----------------------------------------------------------------
Property Get field() As Variant: field = m_field: End Property
Property Let rowNum(ByVal new_value As Integer): m_row_num = new_value: End Property
Property Get rowNum() As Integer: rowNum = m_row_num: End Property
Property Let colNum(ByVal new_value As Integer): m_col_num = new_value: End Property
Property Get colNum() As Integer: colNum = m_col_num: End Property
Property Let remainingMines(ByVal new_value As Integer): m_number_of_remaining_mines = new_value: End Property
Property Get remainingMines() As Integer: remainingMines = m_number_of_remaining_mines: End Property
Property Get boardStartRow() As Integer: boardStartRow = m_board_start_row: End Property
Property Get boardStartCol() As Integer: boardStartCol = m_board_start_col: End Property
Property Let untouchedFields(ByVal new_value As Integer): m_untouched_fields = new_value: End Property
Property Get untouchedFields() As Integer: untouchedFields = m_untouched_fields: End Property

