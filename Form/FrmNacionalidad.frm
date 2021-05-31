VERSION 5.00
Begin VB.Form FrmNacionalidad 
   Caption         =   "Rubros"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4065
      TabIndex        =   14
      Text            =   "1"
      Top             =   3930
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame FrmBotones 
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   840
      TabIndex        =   11
      Top             =   3510
      Width           =   2805
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   465
         Left            =   1590
         TabIndex        =   13
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   465
         Left            =   255
         TabIndex        =   12
         Top             =   210
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el rubro a Imprimir"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3435
      Left            =   240
      TabIndex        =   0
      Top             =   30
      Width           =   4125
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todos las nacionalidades"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   3090
         Width           =   2385
      End
      Begin VB.Frame FrmRubros 
         Height          =   2760
         Left            =   195
         TabIndex        =   1
         Top             =   195
         Width           =   3780
         Begin VB.CheckBox ChkRubro 
            Caption         =   "BOLIVIANO"
            Height          =   240
            Index           =   19
            Left            =   2130
            TabIndex        =   26
            Top             =   2265
            Width           =   1485
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "VENEZOLANO"
            Height          =   195
            Index           =   15
            Left            =   2130
            TabIndex        =   25
            Top             =   1380
            Width           =   1425
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "PERUANO"
            Height          =   225
            Index           =   16
            Left            =   2130
            TabIndex        =   24
            Top             =   1605
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "PARAGUAYO"
            Height          =   210
            Index           =   18
            Left            =   2130
            TabIndex        =   23
            Top             =   2070
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ECUATORIANO"
            Height          =   240
            Index           =   17
            Left            =   2130
            TabIndex        =   22
            Top             =   1830
            Width           =   1485
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ESPAÑOL"
            Height          =   195
            Index           =   10
            Left            =   2130
            TabIndex        =   21
            Top             =   270
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ITALIANO"
            Height          =   195
            Index           =   9
            Left            =   105
            TabIndex        =   20
            Top             =   2280
            Width           =   1500
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "O.AMERICANOS"
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   19
            Top             =   1380
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "N. AMERICANO"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   18
            Top             =   1140
            Width           =   1500
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ALEMAN"
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   17
            Top             =   2055
            Width           =   1740
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "RESTO DEL MUNDO"
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   16
            Top             =   1815
            Width           =   1950
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "O.EUROPEOS"
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   15
            Top             =   1605
            Width           =   1425
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "COLOMBIANO"
            Height          =   240
            Index           =   14
            Left            =   2130
            TabIndex        =   10
            Top             =   1155
            Width           =   1470
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "MEXICANO"
            Height          =   210
            Index           =   13
            Left            =   2130
            TabIndex        =   9
            Top             =   945
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "INGLES"
            Height          =   225
            Index           =   12
            Left            =   2130
            TabIndex        =   8
            Top             =   705
            Width           =   1545
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "FRANCES"
            Height          =   195
            Index           =   11
            Left            =   2130
            TabIndex        =   7
            Top             =   495
            Width           =   1350
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "CHILENO"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   5
            Top             =   930
            Width           =   1425
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "URUGUAYO"
            Height          =   240
            Index           =   1
            Left            =   105
            TabIndex        =   4
            Top             =   480
            Width           =   1350
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "BRASILEÑO"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   3
            Top             =   705
            Width           =   1350
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ARGENTINO"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Top             =   270
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "FrmNacionalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public Rubros As String

Private Sub L_TratarExcel(Titulo As String, subTit As String, Esp As String, CantCol As Integer)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i, n As Integer
Dim Tit As Variant
Dim nombre As String
Dim color As Integer
Dim PeriodoActual As String
Dim PeriodoAnterior As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

FrmRubro.caption = Aplicacion.SeteoProceso(FrmRubro.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    ''AppExcel.application.Visible = True

    ReDim titCol(frmVtaRubroLocal.SprRubroImporte(0).MaxCols)
    Col = 1
    fila = 1
    color = Exl_Gris
    For i = 1 To frmVtaRubroLocal.SprRubroImporte(0).MaxCols
        frmVtaRubroLocal.SprRubroImporte(0).GetText i, 0, Tit
        titCol(i) = Tit
    Next
       
       PeriodoActual = "Período Actual : " & frmVtaRubroLocal.mskFDesde.FormattedText & " - " & frmVtaRubroLocal.mskFHasta.FormattedText
       PeriodoAnterior = "Período Anterior : " & frmVtaRubroLocal.mskFDesdeAnt.FormattedText & " - " & frmVtaRubroLocal.mskFHastaAnt.FormattedText
    
       Exl_PonerValor AppExcel, fila, Col, PeriodoActual
       fila = fila + 1
       Exl_PonerValor AppExcel, fila, Col, PeriodoAnterior
       fila = fila + 1
    
     
If ChkTodos.Value = 1 Then
  For n = 0 To 7
        Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
            
        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
                        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '------------------------
        fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
        fila = fila + 1
        '--------------------------------------
  
  Next n
Else
    For n = 0 To 7
       If ChkRubro(n).Value = 1 Then
         
         Select Case Left(ChkRubro(n).caption, 3)
            
            Case "ACC"
                Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          Case "BEB"
                Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "CIG"
               Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COM"
               Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COS"
               Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
            
            Case "ELE"
            
                Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          
          
          Case "PER"
               Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "TAB"
               Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, Col, frmVtaRubroLocal.TabSubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, Col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
         
         
         End Select
                
     End If

   Next n
 
End If
   ' Exl.Exl_AnchoCol AppExcel, frmVtaRubroLocal.SprRubroImporte(0).MaxCols, frmVtaRubroLocal.SprRubroImporte(0).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 2, 2, 12
    Exl.Exl_AnchoCol AppExcel, 3, 3, 12
    Exl.Exl_AnchoCol AppExcel, 4, 4, 10
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    AppExcel.Application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    'If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    '    AppExcel.PrintOut
    'End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    FrmRubro.caption = Aplicacion.SeteoFin
    Exit Sub


End Sub

Private Sub L_TratarExcel_2(Titulo As String, subTit As String, Esp As String, CantCol As Integer)

Dim AppExcel As Object
Dim titCol() As String
Dim rango As String
Dim Col As Integer
Dim fila As Integer, filaant As Integer
Dim i, n As Integer
Dim Tit As Variant
Dim nombre As String
Dim color As Integer
Dim PeriodoActual As String
Dim PeriodoAnterior As String

On Error GoTo ErrorExl:


nombre = frmDir.NombreArchivo()
DoEvents

FrmNacionalidad.caption = Aplicacion.SeteoProceso(FrmNacionalidad.caption)

If nombre <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    ''AppExcel.application.Visible = True

    ReDim titCol(frmVtaNacionRubro.SprRubroImporte(0).MaxCols)
    Col = 1
    fila = 1
    color = Exl_Gris
    For i = 1 To frmVtaNacionRubro.SprRubroImporte(0).MaxCols
        frmVtaNacionRubro.SprRubroImporte(0).GetText i, 0, Tit
        titCol(i) = Tit
    Next
       
       PeriodoActual = "Período Actual   : " & frmVtaNacionRubro.mskFDesde.FormattedText & " - " & frmVtaNacionRubro.mskFHasta.FormattedText
       PeriodoAnterior = "Período Anterior : " & frmVtaNacionRubro.mskFDesdeAnt.FormattedText & " - " & frmVtaNacionRubro.mskFHastaAnt.FormattedText
       
       Exl_PonerValor AppExcel, fila, Col, "Venta por Nacionalidad Rubro"
       fila = fila + 2
        
       Exl_PonerValor AppExcel, fila, Col, PeriodoActual
       fila = fila + 1
       Exl_PonerValor AppExcel, fila, Col, PeriodoAnterior
       fila = fila + 2
    
       Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.L_Locales
       fila = fila + 2
     
       rango = Exl_rangos(1, 6, 1, 2)
       Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
     
If ChkTodos.Value = 1 Then
  For n = 0 To 7
        Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(0)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
            
        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroImporte(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila - 1), (fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows), 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(1)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
                        
        Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroUnidades(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '------------------------
        fila = fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(3)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroPartic(n), AppExcel, fila, Col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        fila = fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows
        fila = fila + 1
        '--------------------------------------
  
  Next n
Else
  For n = 0 To 19
         If ChkRubro(n).Value = 1 Then
         
         'Select Case Left(ChkRubro(n).caption, 3)
            
         '   Case "ARG"
         Exl_PonerValor AppExcel, fila, Col, ChkRubro(n).caption
         rango = Exl_rangos(fila, fila, 1, 4)
         Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
         Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
         AppExcel.Application.Range(rango).Merge
         Exl_Lineas AppExcel, rango, Exl_Linsimple
         Exl_ColorInt AppExcel, rango, color
                   
         AppExcel.Application.Range(rango).Merge
         Exl_Lineas AppExcel, rango, Exl_Linsimple
               
         fila = fila + 2
                
         Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(0)
         rango = Exl_rangos(fila, fila, 1, 4)
         Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
         Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
         Exl_ColorInt AppExcel, rango, color
         AppExcel.Application.Range(rango).Merge
                    
         fila = fila + 1
            
         Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroImporte(n), AppExcel, fila, Col, titCol
         rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows, 2, 4)
         Exl_Format AppExcel, rango
                                
         rango = Exl_rangos((fila - 1), (fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows), 4, 4)
         Exl_ColorInt AppExcel, rango, color
                
        '-------------------------
         fila = fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows
         fila = fila + 1
                
         Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(1)
         rango = Exl_rangos(fila, fila, 1, 4)
         Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
         Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
         Exl_ColorInt AppExcel, rango, color
         AppExcel.Application.Range(rango).Merge
    
         fila = fila + 1
                                
         Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroUnidades(n), AppExcel, fila, Col, titCol
         rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows, 2, 4)
         Exl_Format AppExcel, rango
                
         rango = Exl_rangos(fila - 1, fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows, 4, 4)
         Exl_ColorInt AppExcel, rango, color
                
         '------------------------
         fila = fila + frmVtaNacionRubro.SprRubroUnidades(n).MaxRows
         fila = fila + 1
              
         Exl_PonerValor AppExcel, fila, Col, frmVtaNacionRubro.TabSubrubro(n).TabCaption(3)
         rango = Exl_rangos(fila, fila, 1, 4)
         Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
         Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
         Exl_ColorInt AppExcel, rango, color
         AppExcel.Application.Range(rango).Merge
    
         fila = fila + 1
                
         Exl_BajarGrillaExel frmVtaNacionRubro.SprRubroPartic(n), AppExcel, fila, Col, titCol
         rango = Exl_rangos(fila + 1, fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows, 2, 4)
         Exl_Format AppExcel, rango
                
         rango = Exl_rangos(fila - 1, fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows, 4, 4)
         Exl_ColorInt AppExcel, rango, color
                
         fila = fila + frmVtaNacionRubro.SprRubroPartic(n).MaxRows
         fila = fila + 1
         '--------------------------------------
          
                
         End If

   Next n
    'Elimina las columnas que sobran
    'rango = Exl_rangos(fila, fila + frmVtaNacionRubro.SprRubroImporte(n).MaxRows, 5, 7)
    rango = Exl_rangos(1, 10000, 5, 7)
    AppExcel.Application.Range(rango).Delete
 
End If
    'Elimina las columnas que sobran
    rango = Exl_rangos(1, 1000, 5, 7)
    AppExcel.Application.Range(rango).Delete
   
   ' Exl.Exl_AnchoCol AppExcel, frmVtaNacionrubro.SprRubroImporte(0).MaxCols, frmVtaNacionrubro.SprRubroImporte(0).MaxCols, 1
    Exl.Exl_AnchoCol AppExcel, 1, 1, 8
    Exl.Exl_AnchoCol AppExcel, 2, 2, 12
    Exl.Exl_AnchoCol AppExcel, 3, 3, 12
    Exl.Exl_AnchoCol AppExcel, 4, 4, 10
    
    AppExcel.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    AppExcel.Application.ActiveSheet.PageSetup.TopMargin = Exl_TopMargen
    AppExcel.Application.ActiveSheet.PageSetup.BottomMargin = Exl_BotMargen
    'If MsgBox("Quiere Imprimir la Planilla Generada", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    '    AppExcel.PrintOut
    'End If
    
    AppExcel.SaveAs nombre & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    FrmNacionalidad.caption = Aplicacion.SeteoFin
    Exit Sub


End Sub


Private Sub ChkRubro_Click(Index As Integer)
ChkTodos.Value = 0
End Sub

Private Sub ChkTodos_Click()
Dim n As Integer

If ChkTodos.Value = 1 Then
   For n = 0 To 7
      ChkRubro(n).Value = 0
   Next n
 Else

End If

End Sub


Private Sub CmdAceptar_Click()
If Text1.Text = 2 Then
  If ChkTodos.Value = 1 Then
        L_TratarExcel_2 "Ventas por Nacionalidad/Rubro", "Todas las Nacionalidades", "TODOS", 8
  Else
        L_TratarExcel_2 "Ventas por Nacionalidad/Rubro", "Todas las Nacionalidades", "", 8
  End If
End If

End Sub


Private Sub CmdCancelar_Click()
  Unload Me
End Sub


