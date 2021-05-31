Attribute VB_Name = "Spread"



'function prototypes
Declare Function SpreadGetDataFillData Lib "Spread20.VBX" (SS As control, Var As Variant, ByVal VType As Integer) As Integer
Declare Function SpreadSetDataFillData Lib "Spread20.VBX" (SS As control, Var As Variant) As Integer
Declare Function SpreadSaveTabFile Lib "Spread20.VBX" (SS As control, ByVal FileName As String) As Integer
Declare Function SpreadSetCellDirtyFlag Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long, ByVal Dirty As Integer) As Integer
Declare Function SpreadGetCellDirtyFlag Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long) As Integer
Declare Function SpreadGetMultiSelItem Lib "Spread20.VBX" (SS As control, ByVal SelPrev As Long) As Long

Declare Function SpreadAddCustomFunction Lib "Spread20.VBX" (hCtl As control, ByVal lpszFunctionName As String, ByVal nParameterCnt As Integer) As Integer
Declare Function SpreadCFGetDoubleParam Lib "Spread20.VBX" (hCtl As control, ByVal dParam As Integer) As Double
Declare Function SpreadCFGetLongParam Lib "Spread20.VBX" (hCtl As control, ByVal dParam As Integer) As Long
Declare Function SpreadCFGetParamInfo Lib "Spread20.VBX" (hCtl As control, ByVal dParam As Integer, wType As Integer, wStatus As Integer) As Integer
Declare Function SpreadCFGetStringParam Lib "Spread20.VBX" (hCtl As control, ByVal dParam As Integer) As String
Declare Sub SpreadCFSetResult Lib "Spread20.VBX" (hCtl As control, Var As Variant)
Declare Function SpreadColNumberToLetter Lib "Spread20.VBX" (ByVal lHeaderNumber As Long) As String
Declare Sub SpreadColWidthToTwips Lib "Spread20.VBX" (SS As control, ByVal fColWidth As Single, lpTwips As Long)
Declare Sub SpreadGetBottomRightCell Lib "Spread20.VBX" (SS As control, lpCol As Long, lpRow As Long)
Declare Sub SpreadGetCellFromScreenCoord Lib "Spread20.VBX" (SS As control, lpCol As Long, lpRow As Long, ByVal X As Long, ByVal Y As Long)
Declare Function SpreadGetCellPos Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long, lpx As Long, lpy As Long, lpWidth As Long, lpHeight As Long) As Integer
Declare Sub SpreadGetClientArea Lib "Spread20.VBX" (SS As control, lplWidth As Long, lplHeight As Long)
Declare Function SpreadGetColItemData Lib "Spread20.VBX" (SS As control, ByVal col As Long) As Long
Declare Sub SpreadGetFirstValidCell Lib "Spread20.VBX" (SS As control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetItemData Lib "Spread20.VBX" (SS As control) As Long
Declare Sub SpreadGetLastValidCell Lib "Spread20.VBX" (SS As control, lpCol As Long, lpRow As Long)
Declare Function SpreadGetRowItemData Lib "Spread20.VBX" (SS As control, ByVal Row As Long) As Long
Declare Function SpreadGetText Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long, Var As Variant) As Integer
Declare Function SpreadIsCellSelected Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long) As Integer
Declare Function SpreadIsFormulaValid Lib "Spread20.VBX" (SS As control, hszFormula As String) As Integer
Declare Function SpreadIsVisible Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long, ByVal Partial As Integer) As Integer
Declare Sub SpreadRowHeightToTwips Lib "Spread20.VBX" (SS As control, ByVal Row As Long, ByVal fRowHeight As Single, lpTwips As Long)
Declare Sub SpreadSaveDesignInfo Lib "Spread20.VBX" (ByVal Their_hWnd As Integer, ByVal My_hWnd As Integer, ByVal finit As Integer)
Declare Sub SpreadSetColItemData Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal lpVar As Long)
Declare Sub SpreadSetItemData Lib "Spread20.VBX" (SS As control, ByVal lpVar As Long)
Declare Sub SpreadSetRowItemData Lib "Spread20.VBX" (SS As control, ByVal Row As Long, ByVal lpVar As Long)
Declare Sub SpreadSetText Lib "Spread20.VBX" (SS As control, ByVal col As Long, ByVal Row As Long, lpVar As Variant)
Declare Sub SpreadTwipsToColWidth Lib "Spread20.VBX" (SS As control, ByVal Twips As Long, fColWidth As Single)
Declare Sub SpreadTwipsToRowHeight Lib "Spread20.VBX" (SS As control, ByVal Row As Long, ByVal Twips As Long, fRowHeight As Single)

'spreadsheet actions
Global Const SS_ACTION_ACTIVE_CELL = 0
Global Const SS_ACTION_GOTO_CELL = 1
Global Const SS_ACTION_SELECT_BLOCK = 2
Global Const SS_ACTION_CLEAR = 3
Global Const SS_ACTION_DELETE_COL = 4
Global Const SS_ACTION_DELETE_ROW = 5
Global Const SS_ACTION_INSERT_COL = 6
Global Const SS_ACTION_INSERT_ROW = 7
Global Const SS_ACTION_LOAD_SPREAD_SHEET = 8
Global Const SS_ACTION_SAVE_ALL = 9
Global Const SS_ACTION_SAVE_VALUES = 10
Global Const SS_ACTION_RECALC = 11
Global Const SS_ACTION_CLEAR_TEXT = 12
Global Const SS_ACTION_PRINT = 13
Global Const SS_ACTION_DESELECT_BLOCK = 14
Global Const SS_ACTION_DSAVE = 15
Global Const SS_ACTION_SET_CELL_BORDER = 16
Global Const SS_ACTION_ADD_MULTISELBLOCK = 17
Global Const SS_ACTION_GET_MULTI_SELECTION = 18
Global Const SS_ACTION_COPY_RANGE = 19
Global Const SS_ACTION_MOVE_RANGE = 20
Global Const SS_ACTION_SWAP_RANGE = 21
Global Const SS_ACTION_CLIPBOARD_COPY = 22
Global Const SS_ACTION_CLIPBOARD_CUT = 23
Global Const SS_ACTION_CLIPBOARD_PASTE = 24
Global Const SS_ACTION_SORT = 25
Global Const SS_ACTION_COMBO_CLEAR = 26
Global Const SS_ACTION_COMBO_REMOVE = 27
Global Const SS_ACTION_RESET = 28

'SelectBlockOptions
Global Const SS_SELBLOCKOPT_COLS = 1
Global Const SS_SELBLOCKOPT_ROWS = 2
Global Const SS_SELBLOCKOPT_BLOCKS = 4
Global Const SS_SELBLOCKOPT_ALL = 8

'DAutoSize settings
'Global Const SS_AUTOSIZE_NONE = 0
Global Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Global Const SS_AUTOSIZE_BEST_GUESS = 2

'cell type
Global Const SS_CELL_TYPE_DATE = 0
Global Const SS_CELL_TYPE_EDIT = 1
Global Const SS_CELL_TYPE_FLOAT = 2
Global Const SS_CELL_TYPE_INTEGER = 3
Global Const SS_CELL_TYPE_PIC = 4
Global Const SS_CELL_TYPE_STATIC_TEXT = 5
Global Const SS_CELL_TYPE_TIME = 6
Global Const SS_CELL_TYPE_BUTTON = 7
Global Const SS_CELL_TYPE_COMBOBOX = 8
Global Const SS_CELL_TYPE_PICTURE = 9
Global Const SS_CELL_TYPE_CHECKBOX = 10
Global Const SS_CELL_TYPE_OWNER_DRAWN = 11

'cell border types
Global Const SS_BORDER_TYPE_NONE = 0
Global Const SS_BORDER_TYPE_OUTLINE = 16
Global Const SS_BORDER_TYPE_LEFT = 1
Global Const SS_BORDER_TYPE_RIGHT = 2
Global Const SS_BORDER_TYPE_TOP = 4
Global Const SS_BORDER_TYPE_BOTTOM = 8

'cell border style
Global Const SS_BORDER_STYLE_DEFAULT = 0
Global Const SS_BORDER_STYLE_SOLID = 1
Global Const SS_BORDER_STYLE_DASH = 2
Global Const SS_BORDER_STYLE_DOT = 3
Global Const SS_BORDER_STYLE_DASH_DOT = 4
Global Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Global Const SS_BORDER_STYLE_BLANK = 6

'Bound. auto col sizing
Global Const SS_BOUND_COL_NO_SIZE = 0
Global Const SS_BOUND_COL_MAX_SIZE = 1
Global Const SS_BOUND_COL_SMART_SIZE = 2

'row and column header settings
Global Const SS_HEADER_BLANK = 0
Global Const SS_HEADER_NUMBERS = 1
Global Const SS_HEADER_LETTERS = 2

'check box text relative to the check box
Global Const SS_CHECKBOX_TEXT_LEFT = 0
Global Const SS_CHECKBOX_TEXT_RIGHT = 1

'cursorstyle
Global Const SS_CURSOR_STYLE_USER_DEFINED = 0
Global Const SS_CURSOR_STYLE_DEFAULT = 1
Global Const SS_CURSOR_STYLE_ARROW = 2
Global Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Global Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

'cursortype
Global Const SS_CURSOR_TYPE_DEFAULT = 0
Global Const SS_CURSOR_TYPE_COLRESIZE = 1
Global Const SS_CURSOR_TYPE_ROWRESIZE = 2
Global Const SS_CURSOR_TYPE_BUTTON = 3
Global Const SS_CURSOR_TYPE_GRAYAREA = 4
Global Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Global Const SS_CURSOR_TYPE_COLHEADER = 6
Global Const SS_CURSOR_TYPE_ROWHEADER = 7

'operation mode
Global Const SS_OP_MODE_NORMAL = 0
Global Const SS_OP_MODE_READONLY = 1
Global Const SS_OP_MODE_ROWMODE = 2
Global Const SS_OP_MODE_SINGLE_SELECT = 3
Global Const SS_OP_MODE_MULTI_SELECT = 4

'sort order
Global Const SS_SORT_ORDER_NONE = 0
Global Const SS_SORT_ORDER_ASCENDING = 1
Global Const SS_SORT_ORDER_DESCENDING = 2

'Sort By
Global Const SS_SORT_BY_ROW = 0
Global Const SS_SORT_BY_COL = 1

'user resize row and columns
Global Const SS_USER_RESIZE_COL = 1
Global Const SS_USER_RESIZE_ROW = 2

'user resize row and columns
Global Const SS_USER_RESIZE_DEFAULT = 0
Global Const SS_USER_RESIZE_ON = 1
Global Const SS_USER_RESIZE_OFF = 2

'style of the virtual mode special scroll bar vscrollspecialtypes
Global Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Global Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Global Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

'position settings
Global Const SS_POSITION_UPPER_LEFT = 0
Global Const SS_POSITION_UPPER_CENTER = 1
Global Const SS_POSITION_UPPER_RIGHT = 2
Global Const SS_POSITION_CENTER_LEFT = 3
Global Const SS_POSITION_CENTER_CENTER = 4
Global Const SS_POSITION_CENTER_RIGHT = 5
Global Const SS_POSITION_BOTTOM_LEFT = 6
Global Const SS_POSITION_BOTTOM_CENTER = 7
Global Const SS_POSITION_BOTTOM_RIGHT = 8

'scroll bar
Global Const SS_SCROLLBAR_NONE = 0
Global Const SS_SCROLLBAR_H_ONLY = 1
Global Const SS_SCROLLBAR_V_ONLY = 2
Global Const SS_SCROLLBAR_BOTH = 3

'print type
Global Const SS_PRINT_ALL = 0
Global Const SS_PRINT_CELL_RANGE = 1
Global Const SS_PRINT_CURRENT_PAGE = 2
Global Const SS_PRINT_PAGE_RANGE = 3

'button types
Global Const SS_CELL_BUTTON_NORMAL = 0
Global Const SS_CELL_BUTTON_TWO_STATE = 1

'button picture align
Global Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Global Const SS_CELL_BUTTON_ALIGN_TOP = 1
Global Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Global Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

'button draw mode
Global Const SS_BDM_ALWAYS = 0
Global Const SS_BDM_CURRENT_CELL = 1
Global Const SS_BDM_CURRENT_COLUMN = 2
Global Const SS_BDM_CURRENT_ROW = 4

'date formats
Global Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Global Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Global Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Global Const SS_CELL_DATE_FORMAT_YYMMDD = 3

'Edit case
Global Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Global Const SS_CELL_EDIT_CASE_NO_CASE = 1
Global Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

'Edit char set
Global Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Global Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Global Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Global Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

'Static text vertical alignment
Global Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Global Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Global Const SS_CELL_STATIC_V_ALIGN_TOP = 2

'Time
Global Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Global Const SS_CELL_TIME_24_HOUR_CLOCK = 1

'Unit type
Global Const SS_CELL_UNIT_NORMAL = 0
Global Const SS_CELL_UNIT_VGA = 1
Global Const SS_CELL_UNIT_TWIPS = 2

'horizontal align
Global Const SS_CELL_H_ALIGN_LEFT = 0
Global Const SS_CELL_H_ALIGN_RIGHT = 1
Global Const SS_CELL_H_ALIGN_CENTER = 2

'EditmodeAction
Global Const SS_CELL_EDITMODE_EXIT_NONE = 0
Global Const SS_CELL_EDITMODE_EXIT_UP = 1
Global Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Global Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Global Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Global Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Global Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6

'Custom function parameter type
Global Const SS_VALUE_TYPE_LONG = 0
Global Const SS_VALUE_TYPE_DOUBLE = 1
Global Const SS_VALUE_TYPE_STR = 2

'The return status of a custom function
Global Const SS_VALUE_STATUS_OK = 0
Global Const SS_VALUE_STATUS_ERROR = 1
Global Const SS_VALUE_STATUS_EMPTY = 2
Global Const SS_VALUE_STATUS_CLEAR = 3
Global Const SS_VALUE_STATUS_NONE = 4



Public Sub Func_PromediosCol(SprSor As control, SprDendo As control, sprRes As control, col As Integer, FD As Integer, FH As Integer)
Dim i
Dim Sor As Variant, Dendo As Variant

For i = FD To FH
    SprSor.GetText col, i, Sor
    SprDendo.GetText col, i, Dendo
    
    If Val(Dendo) > 0 Then
        sprRes.SetText col, i, str(Sor / Dendo)
    End If
Next

End Sub

Public Sub Spead_VaciarGrilla(grid As control)
    grid.MaxRows = 0
End Sub
Public Sub spread_LockGrilla(spr As control, valor As Boolean, CD As Integer, CH As Integer)

    'Select column(s)
    spr.col = CD
    spr.Row = -1
    spr.Col2 = CH
    spr.Row2 = -1
    spr.BlockMode = True

    'Lock cells
    
    spr.Lock = valor

    'Turn block mode off
    spr.BlockMode = False

    
End Sub

Public Sub spread_ResaltarCelda(spr As control, col As Long, fila As Variant)
Dim dato As Variant

    spr.GetText col, fila, dato
    If Val(dato) <= 0 Then
        spr.Row = fila
        spr.col = col
        spr.ForeColor = RGB(255, 0, 0)
    Else
        spr.Row = fila
        spr.col = col
        spr.ForeColor = RGB(0, 0, 255)
    
    End If

End Sub

Public Sub Spread_TotalesGrillas(ByRef spr As control, CantCol As Integer, ColInit As Integer)
Dim i
    spr.MaxRows = spr.MaxRows + 1
    
    Spread_TotalesLinea spr
    
    spr.Row = spr.MaxRows
    For i = ColInit - 1 To CantCol
        spr.col = i + 1
        spr.Formula = "sum(" & Chr(65 + i) & "1:" & Chr(65 + i) & Trim(str(spr.MaxRows - 1)) & ")"
    Next
End Sub
Public Sub Spread_TotalesGrillasCol(ByRef spr As control, colD As Integer, colH As Integer, fila As Integer, col As Integer)
Dim i
    
    spr.Row = fila

    spr.col = col
    
    spr.Formula = "sum(" & Chr(64 + colD) & Trim(str(fila)) & ":" & Chr(64 + colH) & Trim(str(fila)) & ")"
    
End Sub

Public Sub Spread_TotalesGrillaAcum(ByRef spr As control, col As Integer, ColPoner As Integer, fila As Integer)
Dim i
    
    spr.Row = fila
    
    spr.col = ColPoner
    spr.Formula = "sum(" & Chr(64 + col) & "1:" & Chr(64 + col) & Trim(str(fila)) & ")"
    
End Sub


Public Sub Spread_TotalesLinea(spr As control)

    'Select a block of cells
    spr.col = 1
    spr.Row = spr.MaxRows
    spr.Col2 = spr.MaxCols
    spr.Row2 = spr.MaxRows
    spr.BlockMode = True

    'Determine the color of background, foreground and border color
    spr.ForeColor = RGB(0, 0, 255)
    spr.BackColor = RGB(242, 242, 242)
    spr.CellBorderColor = RGB(255, 255, 255)
    
    'Turn block mode off
    spr.BlockMode = False

End Sub



Public Sub Spread_PintaLinea(spr As control, f As Integer, c As Integer, FF As Integer, CC As Integer)

    'Select a block of cells
    spr.col = c
    spr.Row = f
    spr.Col2 = CC
    spr.Row2 = FF
    spr.BlockMode = True


'    spr.ForeColor = RGB(0, 0, 255)
    spr.BackColor = RGB(254, 240, 240)
'    spr.CellBorderColor = RGB(255, 255, 255)
    
    'Turn block mode off
    spr.BlockMode = False

End Sub



Public Sub Spread_PintarfinSemana(ByRef spr As control)
Dim i As Integer
Dim valor As Variant

For i = 1 To spr.MaxRows
    spr.GetText 1, i, valor
    If IsDate(valor) Then
        If WeekDay(CDate(valor)) = DOMINGO Or WeekDay(CDate(valor)) = SABADO Then
            Spread_PintaLinea spr, i, 1, i, spr.MaxCols
        End If
    End If
Next

End Sub

Public Sub Spread_CargarGrilla2(rs As Recordset, grid As control)
Dim Row As Integer, i As Integer
If rs.RecordCount > 0 Then
 Row = 0
 grid.MaxRows = Row
 rs.MoveFirst
 While Not rs.EOF
   Row = Row + 1
   grid.MaxRows = Row
   grid.Row = grid.MaxRows
   For i = 0 To rs.Fields.Count - 1
       grid.col = i + 1
         If grid.CellType <> SS_CELL_TYPE_BUTTON Then
                If Not IsNull(rs.Fields(i)) Then
                            Select Case grid.col
                                Case 1
                                    If rs!TIPO_NOV = "B" Then
                                          grid.Text = "---"
                                     Else
                                          grid.Text = rs.Fields(i)
                                     End If
                                Case 2
                                     grid.Text = rs.Fields(i)
                                Case 3
                                     grid.Text = Format$(rs.Fields(i), "DD/MM/YY")
                                Case 6
                                     grid.Text = rs.Fields(i)
                                     If rs.Fields(i + 1) = "G" Then
                                        grid.Text = grid.Text & " - G"
                                     End If
                                Case 7
                                    grid.Text = rs.Fields(i)
                             End Select
                  End If
       End If
   Next i
   rs.MoveNext
 Wend
End If
End Sub


Public Sub Spread_AddRow(grid As control)
  grid.MaxRows = grid.MaxRows + 1
  grid.Row = grid.MaxRows
End Sub

Public Sub Spread_DelRow(grid As control)
Dim i As Integer
grid.ReDraw = False
grid.Row = grid.SelBlockRow
For i = grid.SelBlockRow To grid.SelBlockRow2
  grid.Action = SS_ACTION_DELETE_ROW
  grid.MaxRows = grid.MaxRows - 1
Next i
grid.ReDraw = True
End Sub

Public Sub Spread_CargarGrilla(rs As Recordset, grid As control)
'Diseñado por Gerado Rossel
' Proposito Cargar una grilla con el contenido de
' un record set
' parametros: rs el recordset con los datso
'             grid la grilla a cargar
Dim Row As Integer, i As Integer
If rs.RecordCount > 0 Then
 Row = 0
 grid.MaxRows = Row
 rs.MoveFirst
 While Not rs.EOF
   Row = Row + 1
   grid.MaxRows = Row
   grid.Row = grid.MaxRows
   For i = 0 To rs.Fields.Count - 1
      grid.col = i + 1
      If grid.CellType <> SS_CELL_TYPE_BUTTON Then
          If Not IsNull(rs.Fields(i)) Then
                 If grid.CellType = SS_CELL_TYPE_DATE Then
                     grid.Text = Format$(rs.Fields(i), "DD/MM/YY")
                 Else
                     grid.Text = rs.Fields(i)
                 End If
          End If
      End If
   Next i
   rs.MoveNext
 Wend
End If
End Sub

Public Function Spread_FilaOcupada(g As control, fila As Long) As Boolean
Dim i
Dim valor As Variant
Dim Result As Boolean

Result = True

'If g.MaxRows > 0 Then

    For i = 1 To g.MaxCols
        g.col = i
        If Not g.ColHidden Then
            If g.CellType <> SS_CELL_TYPE_BUTTON Then
                g.GetText i, fila, valor
                If valor = "" Then
                    Result = False
                    Exit For
                End If
            End If
        End If
    Next
'End If

    Spread_FilaOcupada = Result

End Function

Public Function Spread_ColOcupada(g As control, col As Long) As Boolean
Dim i
Dim valor As Variant
Dim Result As Boolean

Result = True

g.col = col

If g.MaxRows > 0 Then

    For i = 1 To g.MaxRows
        g.Row = i
        If Not g.ColHidden Then
            If g.CellType <> SS_CELL_TYPE_BUTTON Then
                g.GetText col, i, valor
                If valor = "" Then
                    Result = False
                    Exit For
                End If
            End If
        End If
    Next
End If

    Spread_ColOcupada = Result

End Function


Public Function Spread_MaximoValor(fila As Long, col As Integer, spr As control) As Long
Dim i, j
Dim Max As Long
Dim valor As Variant

Max = 0

For i = 1 To fila
    For j = 2 To col + 1
        spr.GetText j, i, valor
        If Val(valor) > Max Then
            Max = valor
        End If
    Next
Next

Spread_MaximoValor = Max
End Function

Public Sub Spread_SetPosCombo(grid As control, col As Long, Row As Long, str As String)
Dim i As Integer, rowold As Integer, colold As Integer
'setea la posicion del combo en la grilla
'de la celda col, row y  cuya descripcion esta en str
colold = grid.col
rowold = grid.Row
grid.col = col
grid.Row = Row
For i = 0 To grid.TypeComboBoxCount - 1
  grid.TypeComboBoxIndex = i
  If grid.TypeComboBoxString = str Then
       grid.TypeComboBoxCurSel = i
       grid.col = colold
       grid.Row = rowold
       Exit Sub
  End If
Next i
grid.col = colold
grid.Row = rowold
End Sub


Public Sub Spread_CargarGrillaGauge(rs As Recordset, grid As control, gauge As control)
'Diseñado por Gerado Rossel
' Proposito Cargar una grilla con el contenido de
' un record set
' parametros: rs el recordset con los datso
'             grid la grilla a cargar
Dim Row As Integer, i As Integer, j As Integer
'gauge.caption = ""
gauge.Value = 0
If rs.RecordCount > 0 Then
 ' setear gauge valores iniciales
' gauge.caption = ""
 gauge.Value = 0
 gauge.Min = 0
 rs.MoveLast
 If (rs.RecordCount - 1) <= 0 Then
   gauge.Max = 1
 Else
   gauge.Max = rs.RecordCount - 1
 End If
 '**
 Row = 0
 grid.MaxRows = Row
 rs.MoveFirst
 j = 1 ' variable de control para el gauge
 While Not rs.EOF
   Row = Row + 1
   grid.MaxRows = Row
   grid.Row = grid.MaxRows
   For i = 0 To rs.Fields.Count - 1
      grid.col = i + 1
      If Not IsNull(rs.Fields(i)) Then
         If grid.CellType = SS_CELL_TYPE_DATE Then
             grid.Text = Format$(rs.Fields(i), FMTFECHA)
         Else
             grid.Text = rs.Fields(i)
         End If
      End If
   Next i
   rs.MoveNext
   ' cargar estado actual en gauge
   If gauge.Max >= j Then
     gauge.Value = j
     j = j + 1
   End If
'   gauge.caption = str$(((gauge.Value - gauge.Min) / (gauge.Max - gauge.Min)) * 100) + " %"
   '**
   
 Wend
End If
End Sub

Public Sub Spread_DelOneRow(grid As control, Row As Long)
Dim i As Integer
'grid.ReDraw = False
If grid.MaxRows > 0 Then
    grid.Row = Row
    grid.Action = SS_ACTION_DELETE_ROW
    grid.MaxRows = grid.MaxRows - 1
    'grid.ReDraw = True
End If
End Sub

Public Sub Spead_LimpiarGrilla(grid As control, FD As Integer, FH As Integer, CD As Integer, CH As Integer)
' Select a block of cells

grid.Row = FD

grid.col = CD

grid.Row2 = FH

grid.Col2 = CH

grid.BlockMode = True

' Clear the data and format of the cells

grid.Action = SS_ACTION_CLEAR_TEXT

' Turn block mode off

grid.BlockMode = False

End Sub


Public Sub spread_OrdenarGrilla(spr As control, col As Integer, orden As Integer)
        spr.Row = 1
        spr.col = 1
        spr.Row2 = spr.MaxRows - 1
        spr.Col2 = spr.MaxCols
        
        ' Set sort definition for key 1
        spr.SortBy = SS_SORT_BY_ROW

        spr.SortKey(1) = col
        spr.SortKeyOrder(1) = orden 'SS_SORT_ORDER_ASCENDING
        spr.Action = SS_ACTION_SORT
End Sub

