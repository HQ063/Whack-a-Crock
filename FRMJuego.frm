VERSION 5.00
Begin VB.Form FRMJuego 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whack-A-Crock"
   ClientHeight    =   3405
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6255
   Icon            =   "FRMJuego.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FRMJuego.frx":164A
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FRMJuego.frx":1EA8
   ScaleHeight     =   3405
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TIMSegs 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Timer TIMGameOver 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1560
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      Picture         =   "FRMJuego.frx":68165
      ScaleHeight     =   1215
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton CMDParar 
         Caption         =   "Terminar juego"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CMDOptions 
         Caption         =   "Opciones"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CMDPlay 
         Caption         =   "Jugar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Restante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label LBLTiempo 
         BackStyle       =   0  'Transparent
         Caption         =   "0:30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Label LBLEfec 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.Label LBLefectividad 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectividad: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LBLTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Cocodrilos Totales:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label LBLTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LBLPtos 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.Label LBLGolpes 
         BackStyle       =   0  'Transparent
         Caption         =   "Cocodrilos Golpeados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Timer TIMVert 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   120
      Top             =   2760
   End
   Begin VB.Timer TIMcambiodir 
      Enabled         =   0   'False
      Interval        =   450
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer TIMLateral 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Label LBLGameover 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   5
      Left            =   6000
      Picture         =   "FRMJuego.frx":6DAFC
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   4
      Left            =   4800
      Picture         =   "FRMJuego.frx":6DB9A
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   3
      Left            =   3600
      Picture         =   "FRMJuego.frx":6DC38
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   2
      Left            =   2400
      Picture         =   "FRMJuego.frx":6DCD6
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   1
      Left            =   1200
      Picture         =   "FRMJuego.frx":6DD74
      Top             =   1200
      Width           =   270
   End
   Begin VB.Image PICCoco 
      Height          =   825
      Left            =   5160
      Picture         =   "FRMJuego.frx":6DE12
      Top             =   120
      Width           =   750
   End
   Begin VB.Image Barreras 
      Height          =   825
      Index           =   0
      Left            =   0
      Picture         =   "FRMJuego.frx":6E5F2
      Top             =   1200
      Width           =   270
   End
   Begin VB.Menu MNUArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MNUNuevo 
         Caption         =   "Nuevo Juego"
      End
      Begin VB.Menu MNUDetener 
         Caption         =   "Terminar Juego"
      End
      Begin VB.Menu MNUBarra 
         Caption         =   "-"
      End
      Begin VB.Menu MNUSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MNUHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MNUOpciones 
         Caption         =   "Opciones"
      End
   End
End
Attribute VB_Name = "FRMJuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim casa
Dim direccion
Dim Golpeado As Boolean
Dim efectividad As String
Dim tiemposegs As String

Private Sub CMDOptions_Click()
TIMVert.Enabled = False
TIMcambiodir.Enabled = False
TIMLateral.Enabled = False
TIMSegs.Enabled = False
TIMGameOver.Enabled = False
FRMOpciones.Visible = True
End Sub

Private Sub CMDParar_Click()
TIMVert.Enabled = False
TIMcambiodir.Enabled = False
TIMLateral.Enabled = False
TIMSegs.Enabled = False
TIMGameOver.Enabled = False
End Sub

Private Sub CMDPlay_Click()

PICCoco.Top = 360
direccion = "abajo"
tiemposegs = 30
TIMVert.Enabled = True
TIMcambiodir.Enabled = True
TIMLateral.Enabled = True
TIMSegs.Enabled = True
TIMGameOver.Enabled = True
LBLGameover.Visible = False
LBLPtos.Caption = "0"
LBLTotal.Caption = "0"
End Sub

Private Sub Form_Load()
direccion = "abajo"
tiemposegs = 30
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim salir As String
salir = MsgBox("¿Estas seguro que deseas salir?", vbYesNo, "Salir?")
If salir = vbYes Then
End
Else
Cancel = 1
End If

End Sub


Private Sub LBLTotal_Change()
If LBLTotal.Caption <> 0 Then
efectividad = LBLPtos.Caption / LBLTotal.Caption * 100
efectividad = Int(efectividad)
LBLEfec.Caption = efectividad + "%"
End If
End Sub

Private Sub MNUDetener_Click()
TIMVert.Enabled = False
TIMcambiodir.Enabled = False
TIMLateral.Enabled = False
TIMSegs.Enabled = False
TIMGameOver.Enabled = False
End Sub

Private Sub MNUNuevo_Click()

PICCoco.Top = 360
direccion = "abajo"
tiemposegs = 30
TIMVert.Enabled = True
TIMcambiodir.Enabled = True
TIMLateral.Enabled = True
TIMSegs.Enabled = True
TIMGameOver.Enabled = True
LBLPtos.Caption = "0"
LBLTotal.Caption = "0"

End Sub

Private Sub MNUOpciones_Click()
TIMVert.Enabled = False
TIMcambiodir.Enabled = False
TIMLateral.Enabled = False
TIMSegs.Enabled = False
TIMGameOver.Enabled = False
FRMOpciones.Visible = True

End Sub

Private Sub MNUSalir_Click()
Dim salir As String
salir = MsgBox("¿Estas seguro que deseas salir?", vbYesNo, "Salir?")
If salir = vbYes Then
End
End If
End Sub

Private Sub PICCOCO_Click()
If Golpeado = False Then
LBLPtos.Caption = LBLPtos.Caption + 1
Golpeado = True
End If
End Sub



Private Sub TIMcambiodir_Timer()
If (direccion = "abajo") Then
direccion = "arriba"
ElseIf (direccion = "arriba") Then
direccion = "abajo"
PICCoco.Top = 360
End If
End Sub

Private Sub TIMGameOver_Timer()
TIMVert.Enabled = False
TIMcambiodir.Enabled = False
TIMLateral.Enabled = False
TIMSegs.Enabled = False
TIMGameOver.Enabled = False
LBLTiempo.Caption = "0:00"
LBLGameover.Visible = True
Dim resultado
resultado = MsgBox("Felicidades, has golpeado " + LBLPtos.Caption + " de " + LBLTotal.Caption + " Cocodrilos, obteniendo un " + LBLEfec.Caption + " de efectividad.", vbOKOnly, "Felicitaciones")
End Sub

Private Sub TIMLateral_Timer()
casa = creanumerocasa()
Select Case casa
Case 1
PICCoco.Left = 360
Case 2
PICCoco.Left = 1560
Case 3
PICCoco.Left = 2760
Case 4
PICCoco.Left = 3960
Case 5
PICCoco.Left = 5160
End Select
LBLTotal.Caption = LBLTotal.Caption + 1
Golpeado = False
End Sub

Private Sub TIMSegs_Timer()
tiemposegs = tiemposegs - 1
If tiemposegs >= 10 Then
LBLTiempo.Caption = "0:" + tiemposegs
ElseIf tiemposegs <= 9 Then
LBLTiempo.Caption = "0:0" + tiemposegs
End If
End Sub

Private Sub TIMVert_Timer()
If (direccion = "abajo") Then
PICCoco.Top = PICCoco.Top + 100
ElseIf (direccion = "arriba") Then
PICCoco.Top = PICCoco.Top - 100
End If
End Sub
