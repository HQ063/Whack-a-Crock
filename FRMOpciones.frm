VERSION 5.00
Begin VB.Form FRMOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FRMOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDGuardopc 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nivel"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton OPT5 
         Caption         =   "5"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton OPT4 
         Caption         =   "4"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton OPT3 
         Caption         =   "3"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton OPT2 
         Caption         =   "2"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.OptionButton OPT1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "FRMOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDGuardopc_Click()
FRMOpciones.Visible = False

'propiedades segun opciones
'nivel
Select Case nivel
Case 1
FRMJuego.TIMVert.Interval = 120
FRMJuego.TIMcambiodir.Interval = 900
FRMJuego.TIMLateral.Interval = 1800
Case 2
FRMJuego.TIMVert.Interval = 100
FRMJuego.TIMcambiodir.Interval = 750
FRMJuego.TIMLateral.Interval = 1500
Case 3
FRMJuego.TIMVert.Interval = 80
FRMJuego.TIMcambiodir.Interval = 600
FRMJuego.TIMLateral.Interval = 1200
Case 4
FRMJuego.TIMVert.Interval = 60
FRMJuego.TIMcambiodir.Interval = 450
FRMJuego.TIMLateral.Interval = 900
Case 5
FRMJuego.TIMVert.Interval = 40
FRMJuego.TIMcambiodir.Interval = 300
FRMJuego.TIMLateral.Interval = 600
End Select
' fin nivel

' Fin opciones

End Sub

Private Sub OPT1_Click()
nivel = 1
End Sub

Private Sub OPT2_Click()
nivel = 2
End Sub

Private Sub OPT3_Click()
nivel = 3
End Sub

Private Sub OPT4_Click()
nivel = 4
End Sub

Private Sub OPT5_Click()
nivel = 5
End Sub
