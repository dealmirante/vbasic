VERSION 5.00
Begin VB.Form FrmRubro 
   Caption         =   "Rubros"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4020
      TabIndex        =   14
      Text            =   "1"
      Top             =   2670
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame FrmBotones 
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   840
      TabIndex        =   11
      Top             =   2415
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
      Height          =   2385
      Left            =   225
      TabIndex        =   0
      Top             =   30
      Width           =   4125
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todos los Rubros"
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
         Top             =   1995
         Width           =   1635
      End
      Begin VB.Frame FrmRubros 
         Height          =   1590
         Left            =   195
         TabIndex        =   1
         Top             =   195
         Width           =   3720
         Begin VB.CheckBox ChkRubro 
            Caption         =   "TABACOS"
            Height          =   360
            Index           =   7
            Left            =   2055
            TabIndex        =   10
            Top             =   1170
            Width           =   1050
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "PERFUMES"
            Height          =   360
            Index           =   6
            Left            =   2055
            TabIndex        =   9
            Top             =   870
            Width           =   1200
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ELECTRONICA"
            Height          =   360
            Index           =   5
            Left            =   2055
            TabIndex        =   8
            Top             =   555
            Width           =   1425
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "COSMETICOS"
            Height          =   360
            Index           =   4
            Left            =   2055
            TabIndex        =   7
            Top             =   270
            Width           =   1350
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "COMESTIBLES"
            Height          =   360
            Index           =   3
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1425
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "CIGARILLOS"
            Height          =   360
            Index           =   2
            Left            =   210
            TabIndex        =   4
            Top             =   855
            Width           =   1245
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "BEBIDAS"
            Height          =   360
            Index           =   1
            Left            =   210
            TabIndex        =   3
            Top             =   555
            Width           =   990
         End
         Begin VB.CheckBox ChkRubro 
            Caption         =   "ACCESORIOS"
            Height          =   360
            Index           =   0
            Left            =   210
            TabIndex        =   2
            Top             =   270
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmRubro"
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
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i, n As Integer
Dim tit As Variant
Dim NOMBRE As String
Dim color As Integer
Dim PeriodoActual As String
Dim PeriodoAnterior As String

On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmRubro.caption = Aplicacion.SeteoProceso(FrmRubro.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    ''AppExcel.application.Visible = True

    ReDim titCol(frmVtaRubroLocal.SprRubroImporte(0).MaxCols)
    col = 1
    fila = 1
    color = Exl_Gris
    For i = 1 To frmVtaRubroLocal.SprRubroImporte(0).MaxCols
        frmVtaRubroLocal.SprRubroImporte(0).GetText i, 0, tit
        titCol(i) = tit
    Next
       
       PeriodoActual = "Período Actual : " & frmVtaRubroLocal.mskFDesde.FormattedText & " - " & frmVtaRubroLocal.mskFHasta.FormattedText
       PeriodoAnterior = "Período Anterior : " & frmVtaRubroLocal.mskFDesdeAnt.FormattedText & " - " & frmVtaRubroLocal.mskFHastaAnt.FormattedText
    
       Exl_PonerValor AppExcel, fila, col, PeriodoActual
       fila = fila + 1
       Exl_PonerValor AppExcel, fila, col, PeriodoAnterior
       fila = fila + 1
    
     
If ChkTodos.Value = 1 Then
  For n = 0 To 7
        Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
            
        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
                        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '------------------------
        fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
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
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          Case "BEB"
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "CIG"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COM"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COS"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
            
            Case "ELE"
            
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          
          
          Case "PER"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroLocal.SprRubroPromedio(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "TAB"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroLocal.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroLocal.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroLocal.tabsubrubro(n).TabCaption(2)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroLocal.SprRubroPromedio(n), AppExcel, fila, col, titCol
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
    
    AppExcel.SaveAs NOMBRE & ".xls"
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
Dim col As Integer
Dim fila As Integer, filaant As Integer
Dim i, n As Integer
Dim tit As Variant
Dim NOMBRE As String
Dim color As Integer
Dim PeriodoActual As String
Dim PeriodoAnterior As String

'On Error GoTo ErrorExl:


NOMBRE = frmDir.NombreArchivo()
DoEvents

FrmRubro.caption = Aplicacion.SeteoProceso(FrmRubro.caption)

If NOMBRE <> "" Then
    Set AppExcel = CreateObject("excel.sheet")
    ''AppExcel.application.Visible = True

    ReDim titCol(frmVtaRubroNacion.SprRubroImporte(0).MaxCols)
    col = 1
    fila = 1
    color = Exl_Gris
    For i = 1 To frmVtaRubroNacion.SprRubroImporte(0).MaxCols
        frmVtaRubroNacion.SprRubroImporte(0).GetText i, 0, tit
        titCol(i) = tit
    Next
       
       PeriodoActual = "Período Actual   : " & frmVtaRubroNacion.mskFDesde.FormattedText & " - " & frmVtaRubroNacion.mskFHasta.FormattedText
       PeriodoAnterior = "Período Anterior : " & frmVtaRubroNacion.mskFDesdeAnt.FormattedText & " - " & frmVtaRubroNacion.mskFHastaAnt.FormattedText
       
       Exl_PonerValor AppExcel, fila, col, "Venta por Rubro Nacionalidad"
       fila = fila + 2
        
       Exl_PonerValor AppExcel, fila, col, PeriodoActual
       fila = fila + 1
       Exl_PonerValor AppExcel, fila, col, PeriodoAnterior
       fila = fila + 2
    
       Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.L_Locales
       fila = fila + 2
     
       rango = Exl_rangos(1, 6, 1, 2)
       Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
     
If ChkTodos.Value = 1 Then
  For n = 0 To 7
        Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        Exl_ColorInt AppExcel, rango, color
           
        AppExcel.Application.Range(rango).Merge
        Exl_Lineas AppExcel, rango, Exl_Linsimple
        
        fila = fila + 2
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge
            
        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
          
        rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '-------------------------
        fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
                        
        Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        '------------------------
        fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
        fila = fila + 1
        
        Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
        rango = Exl_rangos(fila, fila, 1, 4)
        Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
        Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
        Exl_ColorInt AppExcel, rango, color
        AppExcel.Application.Range(rango).Merge

        fila = fila + 1
        
        Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
        rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
        Exl_Format AppExcel, rango
        
        rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
        Exl_ColorInt AppExcel, rango, color
        
        fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
        fila = fila + 1
        '--------------------------------------
  
  Next n
Else
    For n = 0 To 7
       If ChkRubro(n).Value = 1 Then
         
         Select Case Left(ChkRubro(n).caption, 3)
            
            Case "ACC"
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                'Elimina las columnas que sobran
                rango = Exl_rangos(fila, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 5, 7)
                AppExcel.Application.Range(rango).Delete
                
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          Case "BEB"
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "CIG"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COM"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "COS"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
            
            Case "ELE"
            
                Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          
          
          
          Case "PER"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
          Case "TAB"
               Exl_PonerValor AppExcel, fila, col, ChkRubro(n).caption
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 14, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Centro, Exl_CentroVert, False
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                Exl_ColorInt AppExcel, rango, color
                   
                AppExcel.Application.Range(rango).Merge
                Exl_Lineas AppExcel, rango, Exl_Linsimple
                
                fila = fila + 2
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(0)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
                    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroImporte(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                  
                rango = Exl_rangos((fila - 1), (fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows), 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '-------------------------
                fila = fila + frmVtaRubroNacion.SprRubroImporte(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(1)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroUnidades(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                '------------------------
                fila = fila + frmVtaRubroNacion.SprRubroUnidades(n).MaxRows
                fila = fila + 1
                
                Exl_PonerValor AppExcel, fila, col, frmVtaRubroNacion.tabsubrubro(n).TabCaption(3)
                rango = Exl_rangos(fila, fila, 1, 4)
                Exl_Letra AppExcel, rango, NEGRITA, 12, "Ms Serif"
                Exl_Justificacion AppExcel, rango, Exl_Izq, Exl_CentroVert, False
                Exl_ColorInt AppExcel, rango, color
                AppExcel.Application.Range(rango).Merge
    
                fila = fila + 1
                
                Exl_BajarGrillaExel frmVtaRubroNacion.SprRubroPartic(n), AppExcel, fila, col, titCol
                rango = Exl_rangos(fila + 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 2, 4)
                Exl_Format AppExcel, rango
                
                rango = Exl_rangos(fila - 1, fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows, 4, 4)
                Exl_ColorInt AppExcel, rango, color
                
                fila = fila + frmVtaRubroNacion.SprRubroPartic(n).MaxRows
                fila = fila + 1
                '--------------------------------------
         
         
         End Select
                
     End If

   Next n
 
End If
    'Elimina las columnas que sobran
    rango = Exl_rangos(1, 1000, 5, 7)
    AppExcel.Application.Range(rango).Delete
   
   ' Exl.Exl_AnchoCol AppExcel, frmVtaRubronacion.SprRubroImporte(0).MaxCols, frmVtaRubronacion.SprRubroImporte(0).MaxCols, 1
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
    
    AppExcel.SaveAs NOMBRE & ".xls"
    Set AppExcel = Nothing
End If

ErrorExl:

    FrmRubro.caption = Aplicacion.SeteoFin
    Exit Sub


End Sub


Public Function Rubros()
Dim cadena As String

cadena = ""
If ChkTodos.Value = 1 Then
    cadena = "ACCBEBCIGCOMCOSELEPERTAB"
Else
    If ChkRubro(0).Value = 1 Then
        cadena = cadena & "ACC"
    End If
    If ChkRubro(1).Value = 1 Then
        cadena = cadena & "BEB"
    End If
    If ChkRubro(2).Value = 1 Then
        cadena = cadena & "CIG"
    End If
    If ChkRubro(3).Value = 1 Then
        cadena = cadena & "COM"
    End If
    If ChkRubro(4).Value = 1 Then
        cadena = cadena & "COS"
    End If
    If ChkRubro(5).Value = 1 Then
        cadena = cadena & "ELE"
    End If
    If ChkRubro(6).Value = 1 Then
        cadena = cadena & "PER"
    End If
    If ChkRubro(7).Value = 1 Then
        cadena = cadena & "TAB"
    End If
    
End If
        
    Rubros = cadena
    
End Function


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

If Text1.Text = 1 Then
  If ChkTodos.Value = 1 Then
        L_TratarExcel "Ventas por Rubro/Local", "Todos los Rubros", "TODOS", 8
  Else
        L_TratarExcel "Ventas por Rubro/Local", "Todos los Rubros", "", 8
  End If
ElseIf Text1.Text = 2 Then
  If ChkTodos.Value = 1 Then
        L_TratarExcel_2 "Ventas por Rubro/Nacionalidad", "Todos los Rubros", "TODOS", 8
  Else
        L_TratarExcel_2 "Ventas por Rubro/Nacionalidad", "Todos los Rubros", "", 8
  End If
ElseIf Text1.Text = 3 Then
        frmVtaRubroProv.Set_cadenaRubro (Rubros)
        Unload Me
End If

End Sub


Private Sub CmdCancelar_Click()
  Unload Me
End Sub


