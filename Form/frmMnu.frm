VERSION 5.00
Begin VB.Form frmMnu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7035
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuGral 
      Caption         =   "&Estadisticas"
      Index           =   0
      Begin VB.Menu mnuEstadis 
         Caption         =   "Ventas por &Local y Grupo"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGral 
      Caption         =   "&Salir"
      Index           =   10
   End
End
Attribute VB_Name = "frmMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuEstadis_Click(Index As Integer)
Select Case Index
    Case 0
        frmLGC.Show 1
        
End Select
End Sub


