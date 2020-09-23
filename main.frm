VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9345
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9030
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   5820
      TabIndex        =   15
      Top             =   4290
      Width           =   2670
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "obstacles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   1590
         TabIndex        =   19
         Top             =   2865
         Width           =   705
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "strafe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   375
         TabIndex        =   18
         Top             =   2865
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "rising"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   1590
         TabIndex        =   17
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "clasico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   375
         TabIndex        =   16
         Top             =   1440
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   3
         Left            =   1500
         Picture         =   "main.frx":08CA
         Top             =   1710
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   2
         Left            =   285
         Picture         =   "main.frx":4098
         Top             =   1710
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   1
         Left            =   1500
         Picture         =   "main.frx":7866
         Top             =   285
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   1185
         Index           =   0
         Left            =   285
         Picture         =   "main.frx":B034
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7965
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   9555
      Width           =   810
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   5820
      TabIndex        =   1
      Top             =   840
      Width           =   2685
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1305
         Left            =   675
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   5
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Modo: "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   2835
         Width           =   2310
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Puntaje:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   2475
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   210
         TabIndex        =   3
         Top             =   2115
         Width           =   1320
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   210
         TabIndex        =   2
         Top             =   1740
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7950
      Left            =   510
      ScaleHeight     =   528
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   915
      Width           =   4785
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "x 3"
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   1740
         TabIndex        =   14
         Top             =   3540
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Rising..."
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   1755
         TabIndex        =   13
         Top             =   3510
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "PAUSA"
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   2280
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "PREPARATE !"
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   1455
         TabIndex        =   8
         Top             =   3345
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "5"
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   2235
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   375
      Left            =   8325
      Top             =   8790
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   60
      Picture         =   "main.frx":E802
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   9335
      Left            =   10
      Top             =   10
      Width           =   9020
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " x"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8580
      TabIndex        =   10
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-=  TETRIS  =-"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   705
      TabIndex        =   9
      Top             =   90
      Width           =   7725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desarrollado Por Julio C. Gotuzzo
'Buenos Aires - Argentina
'gotuzzo@excite.com

Private Type Bloque
 X As Long
 Y As Long
 visible As Boolean
End Type

Private Type Pieza
 visible As Boolean
 color As Long
 tipo As Integer
 rot As Integer
 bloque1 As Bloque
 bloque2 As Bloque
 bloque3 As Bloque
 bloque4 As Bloque
End Type

Private piezas As Pieza
Private matris(20, 12) As Integer
Public pieza_sig As Integer
Public puntaje As Long
Public lineas As Long
Public nivel As Integer
Private modo As Integer
Private contador_piezas As Integer
Private CURX, CURY As Single

Public Sub iniciar_matris()
 Dim n As Integer
 Dim f As Integer
 n = 0
 Do While n <= 20
  f = 0
  Do While f <= 12
   If f = 0 And n <= 19 Then
    matris(n, f) = 1
   Else
    If f = 12 And n <= 19 Then
     matris(n, f) = 1
    Else
     If n = 20 Then
      matris(n, f) = 1
     Else
      matris(n, f) = 0
     End If
    End If
   End If
   f = f + 1
  Loop
  n = n + 1
 Loop
End Sub

Public Sub dibujar_matris()
 Dim n As Integer
 Dim f As Integer
 n = 0
 Do While n <= 19
  f = 1
  Do While f <= 11
   If matris(n, f) = 0 Then
    cuadrado_relleno f * 24, n * 24, 26, Picture1.BackColor, Picture1
   End If
   f = f + 1
  Loop
  n = n + 1
 Loop
 n = 0
 Do While n <= 20
  f = 0
  Do While f <= 12
   If matris(n, f) > 0 Then
    Select Case matris(n, f)
    Case Is = 1
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HC0C0C0, Picture1
    Case Is = 2
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HFF0000, Picture1
    Case Is = 3
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HFF&, Picture1
    Case Is = 4
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HFFFF&, Picture1
    Case Is = 5
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HFFFF00, Picture1
    Case Is = 6
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HFFFF00, Picture1
    Case Is = 7
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HC000&, Picture1
    Case Is = 8
     cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &HC000&, Picture1
    End Select
    cuadrado_vacio f * 24, n * 24, 24, 0, Picture1
    cuadrado_vacio (f * 24) + 1, (n * 24) + 1, 24, 0, Picture1
   End If
   f = f + 1
  Loop
  n = n + 1
 Loop
End Sub

Public Sub dibujar_matris_bloque(X As Long, Y As Long)
 cuadrado_vacio Y, X, 24, 0, Picture1
 cuadrado_vacio Y + 1, X + 1, 24, 0, Picture1
End Sub

Public Sub pieza_siguiente()
 pieza_sig = (Rnd(5) * 6) + 1
 Picture2.Cls
 Select Case pieza_sig
 Case Is = 1
  cuadrado_relleno 2, 2, 22, &HFF0000, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HFF0000, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
  cuadrado_relleno 2, 26, 22, &HFF0000, Picture2
  cuadrado_vacio 0, 24, 24, 0, Picture2
  cuadrado_vacio 1, 25, 24, 0, Picture2
  cuadrado_relleno 26, 26, 22, &HFF0000, Picture2
  cuadrado_vacio 24, 24, 24, 0, Picture2
  cuadrado_vacio 25, 25, 24, 0, Picture2
 Case Is = 2
  cuadrado_relleno 2, 2, 22, &HFF&, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 2, 26, 22, &HFF&, Picture2
  cuadrado_vacio 0, 24, 24, 0, Picture2
  cuadrado_vacio 1, 25, 24, 0, Picture2
  cuadrado_relleno 2, 50, 22, &HFF&, Picture2
  cuadrado_vacio 0, 48, 24, 0, Picture2
  cuadrado_vacio 1, 49, 24, 0, Picture2
 Case Is = 3
  cuadrado_relleno 2, 2, 22, &HFFFF&, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HFFFF&, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
  cuadrado_relleno 50, 2, 22, &HFFFF&, Picture2
  cuadrado_vacio 48, 0, 24, 0, Picture2
  cuadrado_vacio 49, 1, 24, 0, Picture2
  cuadrado_relleno 26, 26, 22, &HFFFF&, Picture2
  cuadrado_vacio 24, 24, 24, 0, Picture2
  cuadrado_vacio 25, 25, 24, 0, Picture2
 Case Is = 4
  cuadrado_relleno 2, 2, 22, &HFFFF00, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 2, 26, 22, &HFFFF00, Picture2
  cuadrado_vacio 0, 24, 24, 0, Picture2
  cuadrado_vacio 1, 25, 24, 0, Picture2
  cuadrado_relleno 2, 50, 22, &HFFFF00, Picture2
  cuadrado_vacio 0, 48, 24, 0, Picture2
  cuadrado_vacio 1, 49, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HFFFF00, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
 Case Is = 5
  cuadrado_relleno 2, 2, 22, &HFFFF00, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HFFFF00, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
  cuadrado_relleno 26, 26, 22, &HFFFF00, Picture2
  cuadrado_vacio 24, 24, 24, 0, Picture2
  cuadrado_vacio 25, 25, 24, 0, Picture2
  cuadrado_relleno 26, 50, 22, &HFFFF00, Picture2
  cuadrado_vacio 24, 48, 24, 0, Picture2
  cuadrado_vacio 25, 49, 24, 0, Picture2
 Case Is = 6
  cuadrado_relleno 2, 26, 22, &HC000&, Picture2
  cuadrado_vacio 0, 24, 24, 0, Picture2
  cuadrado_vacio 1, 25, 24, 0, Picture2
  cuadrado_relleno 26, 26, 22, &HC000&, Picture2
  cuadrado_vacio 24, 24, 24, 0, Picture2
  cuadrado_vacio 25, 25, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HC000&, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
  cuadrado_relleno 50, 2, 22, &HC000&, Picture2
  cuadrado_vacio 48, 0, 24, 0, Picture2
  cuadrado_vacio 49, 1, 24, 0, Picture2
 Case Is = 7
  cuadrado_relleno 2, 2, 22, &HC000&, Picture2
  cuadrado_vacio 0, 0, 24, 0, Picture2
  cuadrado_vacio 1, 1, 24, 0, Picture2
  cuadrado_relleno 26, 2, 22, &HC000&, Picture2
  cuadrado_vacio 24, 0, 24, 0, Picture2
  cuadrado_vacio 25, 1, 24, 0, Picture2
  cuadrado_relleno 26, 26, 22, &HC000&, Picture2
  cuadrado_vacio 24, 24, 24, 0, Picture2
  cuadrado_vacio 25, 25, 24, 0, Picture2
  cuadrado_relleno 50, 26, 22, &HC000&, Picture2
  cuadrado_vacio 48, 24, 24, 0, Picture2
  cuadrado_vacio 49, 25, 24, 0, Picture2
 End Select
End Sub

Public Sub iniciar_pieza(tipo As Integer)
 Dim mitad As Long
 mitad = 4 * 24
 piezas.visible = True
 piezas.tipo = tipo
 piezas.rot = 1
 Select Case tipo
 Case Is = 1
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = 24 + mitad
  piezas.bloque2.Y = 0
  piezas.bloque3.visible = True
  piezas.bloque3.X = mitad
  piezas.bloque3.Y = 24
  piezas.bloque4.visible = True
  piezas.bloque4.X = 24 + mitad
  piezas.bloque4.Y = 24
  piezas.color = &HFF0000
 Case Is = 2
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = mitad
  piezas.bloque2.Y = 24
  piezas.bloque3.visible = True
  piezas.bloque3.X = mitad
  piezas.bloque3.Y = 48
  piezas.bloque4.visible = False
 piezas.color = &HFF&
 Case Is = 3
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = 24 + mitad
  piezas.bloque2.Y = 0
  piezas.bloque3.visible = True
  piezas.bloque3.X = 48 + mitad
  piezas.bloque3.Y = 0
  piezas.bloque4.visible = True
  piezas.bloque4.X = 24 + mitad
  piezas.bloque4.Y = 24
 piezas.color = &HFFFF&
 Case Is = 4
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = mitad
  piezas.bloque2.Y = 24
  piezas.bloque3.visible = True
  piezas.bloque3.X = mitad
  piezas.bloque3.Y = 48
  piezas.bloque4.visible = True
  piezas.bloque4.X = 24 + mitad
  piezas.bloque4.Y = 0
 piezas.color = &HFFFF00
 Case Is = 5
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = 24 + mitad
  piezas.bloque2.Y = 0
  piezas.bloque3.visible = True
  piezas.bloque3.X = 24 + mitad
  piezas.bloque3.Y = 24
  piezas.bloque4.visible = True
  piezas.bloque4.X = 24 + mitad
  piezas.bloque4.Y = 48
 piezas.color = &HFFFF00
 Case Is = 6
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 24
  piezas.bloque2.visible = True
  piezas.bloque2.X = 24 + mitad
  piezas.bloque2.Y = 24
  piezas.bloque3.visible = True
  piezas.bloque3.X = 24 + mitad
  piezas.bloque3.Y = 0
  piezas.bloque4.visible = True
  piezas.bloque4.X = 48 + mitad
  piezas.bloque4.Y = 0
 piezas.color = &HC000&
 Case Is = 7
  piezas.bloque1.visible = True
  piezas.bloque1.X = mitad
  piezas.bloque1.Y = 0
  piezas.bloque2.visible = True
  piezas.bloque2.X = 24 + mitad
  piezas.bloque2.Y = 0
  piezas.bloque3.visible = True
  piezas.bloque3.X = 24 + mitad
  piezas.bloque3.Y = 24
  piezas.bloque4.visible = True
  piezas.bloque4.X = 48 + mitad
  piezas.bloque4.Y = 24
 piezas.color = &HC000&
 End Select
 dibujar_pieza
End Sub

Public Sub rotar_pieza()
 If detenido = True Then Exit Sub
 borrar_pieza
 Select Case piezas.tipo
 Case Is = 2
  If piezas.rot = 1 Then
   piezas.rot = 2
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X + 24
  Else
   piezas.rot = 1
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X - 24
  End If
 Case Is = 3
  Select Case piezas.rot
  Case Is = 1
   piezas.rot = 2
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y - 24
   piezas.bloque4.X = piezas.bloque4.X - 24
  Case Is = 2
   piezas.rot = 3
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y - 24
   piezas.bloque4.X = piezas.bloque4.X + 24
  Case Is = 3
   piezas.rot = 4
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y + 24
   piezas.bloque4.X = piezas.bloque4.X + 24
  Case Is = 4
   piezas.rot = 1
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y + 24
   piezas.bloque4.X = piezas.bloque4.X - 24
  End Select
 Case Is = 4
  Select Case piezas.rot
  Case Is = 1
   piezas.rot = 2
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y + 48
  Case Is = 2
   piezas.rot = 3
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X + 24
   piezas.bloque4.X = piezas.bloque4.X - 48
  Case Is = 3
   piezas.rot = 4
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y - 48
  Case Is = 4
   piezas.rot = 1
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X - 24
   piezas.bloque4.X = piezas.bloque4.X + 48
  End Select
 Case Is = 5
  Select Case piezas.rot
  Case Is = 1
   piezas.rot = 2
   piezas.bloque1.X = piezas.bloque1.X + 48
   piezas.bloque2.Y = piezas.bloque2.Y + 24
   piezas.bloque2.X = piezas.bloque2.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y - 24
   piezas.bloque4.X = piezas.bloque4.X - 24
  Case Is = 2
   piezas.rot = 3
   piezas.bloque1.Y = piezas.bloque1.Y + 48
   piezas.bloque2.Y = piezas.bloque2.Y + 24
   piezas.bloque2.X = piezas.bloque2.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y - 24
   piezas.bloque4.X = piezas.bloque4.X + 24
  Case Is = 3
   piezas.rot = 4
   piezas.bloque1.X = piezas.bloque1.X - 48
   piezas.bloque2.Y = piezas.bloque2.Y - 24
   piezas.bloque2.X = piezas.bloque2.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y + 24
   piezas.bloque4.X = piezas.bloque4.X + 24
  Case Is = 4
   piezas.rot = 1
   piezas.bloque1.Y = piezas.bloque1.Y - 48
   piezas.bloque2.Y = piezas.bloque2.Y - 24
   piezas.bloque2.X = piezas.bloque2.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y + 24
   piezas.bloque4.X = piezas.bloque4.X - 24
  End Select
 Case Is = 6
  If piezas.rot = 1 Then
   piezas.rot = 2
   piezas.bloque1.Y = piezas.bloque1.Y - 24
   piezas.bloque1.X = piezas.bloque1.X + 24
   piezas.bloque3.Y = piezas.bloque3.Y + 24
   piezas.bloque3.X = piezas.bloque3.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y + 48
  Else
   piezas.rot = 1
   piezas.bloque1.Y = piezas.bloque1.Y + 24
   piezas.bloque1.X = piezas.bloque1.X - 24
   piezas.bloque3.Y = piezas.bloque3.Y - 24
   piezas.bloque3.X = piezas.bloque3.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y - 48
  End If
 Case Is = 7
  If piezas.rot = 1 Then
   piezas.rot = 2
   piezas.bloque1.X = piezas.bloque1.X + 48
   piezas.bloque2.Y = piezas.bloque2.Y + 24
   piezas.bloque2.X = piezas.bloque2.X + 24
   piezas.bloque4.Y = piezas.bloque4.Y + 24
   piezas.bloque4.X = piezas.bloque4.X - 24
  Else
   piezas.rot = 1
   piezas.bloque1.X = piezas.bloque1.X - 48
   piezas.bloque2.Y = piezas.bloque2.Y - 24
   piezas.bloque2.X = piezas.bloque2.X - 24
   piezas.bloque4.Y = piezas.bloque4.Y - 24
   piezas.bloque4.X = piezas.bloque4.X + 24
  End If
 End Select
 dibujar_pieza
End Sub

Public Sub asentar_pieza()
 If detenido = True Then Exit Sub
 puntaje = puntaje + 10 + piezas.tipo
 Label4.Caption = "Puntaje: " + Trim(Str(puntaje))
 If piezas.bloque1.visible = True Then matris((piezas.bloque1.Y / 24), (piezas.bloque1.X / 24)) = piezas.tipo + 1
 If piezas.bloque2.visible = True Then matris((piezas.bloque2.Y / 24), (piezas.bloque2.X / 24)) = piezas.tipo + 1
 If piezas.bloque3.visible = True Then matris((piezas.bloque3.Y / 24), (piezas.bloque3.X / 24)) = piezas.tipo + 1
 If piezas.bloque4.visible = True Then matris((piezas.bloque4.Y / 24), (piezas.bloque4.X / 24)) = piezas.tipo + 1
End Sub

Public Sub dibujar_pieza()
 If piezas.bloque1.visible = True Then dibujar_bloque 1
 If piezas.bloque2.visible = True Then dibujar_bloque 2
 If piezas.bloque3.visible = True Then dibujar_bloque 3
 If piezas.bloque4.visible = True Then dibujar_bloque 4
End Sub

Public Sub borrar_pieza()
 If piezas.bloque1.visible = True Then borrar_bloque 1
 If piezas.bloque2.visible = True Then borrar_bloque 2
 If piezas.bloque3.visible = True Then borrar_bloque 3
 If piezas.bloque4.visible = True Then borrar_bloque 4
End Sub

Public Sub mover_pieza(dire As Integer)
 If detenido = True Then Exit Sub
 borrar_pieza
 Select Case dire
 Case Is = 1
  If piezas.bloque1.visible = True Then piezas.bloque1.Y = piezas.bloque1.Y + 24
  If piezas.bloque2.visible = True Then piezas.bloque2.Y = piezas.bloque2.Y + 24
  If piezas.bloque3.visible = True Then piezas.bloque3.Y = piezas.bloque3.Y + 24
  If piezas.bloque4.visible = True Then piezas.bloque4.Y = piezas.bloque4.Y + 24
 Case Is = 2
  If piezas.bloque1.visible = True Then piezas.bloque1.X = piezas.bloque1.X - 24
  If piezas.bloque2.visible = True Then piezas.bloque2.X = piezas.bloque2.X - 24
  If piezas.bloque3.visible = True Then piezas.bloque3.X = piezas.bloque3.X - 24
  If piezas.bloque4.visible = True Then piezas.bloque4.X = piezas.bloque4.X - 24
 Case Is = 3
  If piezas.bloque1.visible = True Then piezas.bloque1.X = piezas.bloque1.X + 24
  If piezas.bloque2.visible = True Then piezas.bloque2.X = piezas.bloque2.X + 24
  If piezas.bloque3.visible = True Then piezas.bloque3.X = piezas.bloque3.X + 24
  If piezas.bloque4.visible = True Then piezas.bloque4.X = piezas.bloque4.X + 24
 End Select
 dibujar_pieza
End Sub

Public Sub dibujar_bloque(numero As Integer)
 Select Case numero
 Case Is = 1
  cuadrado_relleno piezas.bloque1.X + 2, piezas.bloque1.Y + 2, 22, piezas.color, Picture1
  cuadrado_vacio piezas.bloque1.X, piezas.bloque1.Y, 24, 0, Picture1
  cuadrado_vacio piezas.bloque1.X + 1, piezas.bloque1.Y + 1, 24, 0, Picture1
 Case Is = 2
  cuadrado_relleno piezas.bloque2.X + 2, piezas.bloque2.Y + 2, 22, piezas.color, Picture1
  cuadrado_vacio piezas.bloque2.X, piezas.bloque2.Y, 24, 0, Picture1
  cuadrado_vacio piezas.bloque2.X + 1, piezas.bloque2.Y + 1, 24, 0, Picture1
 Case Is = 3
  cuadrado_relleno piezas.bloque3.X + 2, piezas.bloque3.Y + 2, 22, piezas.color, Picture1
  cuadrado_vacio piezas.bloque3.X, piezas.bloque3.Y, 24, 0, Picture1
  cuadrado_vacio piezas.bloque3.X + 1, piezas.bloque3.Y + 1, 24, 0, Picture1
 Case Is = 4
  cuadrado_relleno piezas.bloque4.X + 2, piezas.bloque4.Y + 2, 22, piezas.color, Picture1
  cuadrado_vacio piezas.bloque4.X, piezas.bloque4.Y, 24, 0, Picture1
  cuadrado_vacio piezas.bloque4.X + 1, piezas.bloque4.Y + 1, 24, 0, Picture1
 End Select
End Sub

Public Sub borrar_bloque(numero As Integer)
 Select Case numero
 Case Is = 1
  cuadrado_relleno piezas.bloque1.X, piezas.bloque1.Y, 26, Picture1.BackColor, Picture1
  If piezas.bloque1.Y > 0 Then
   If matris((piezas.bloque1.Y / 24) - 1, (piezas.bloque1.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y - 24, piezas.bloque1.X - 24
   If matris((piezas.bloque1.Y / 24) - 1, piezas.bloque1.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque1.Y - 24, piezas.bloque1.X
   If matris((piezas.bloque1.Y / 24) - 1, (piezas.bloque1.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y - 24, piezas.bloque1.X + 24
  End If
  If matris(piezas.bloque1.Y / 24, (piezas.bloque1.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y, piezas.bloque1.X + 24
  If matris((piezas.bloque1.Y / 24) + 1, (piezas.bloque1.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y + 24, piezas.bloque1.X + 24
  If matris((piezas.bloque1.Y / 24) + 1, piezas.bloque1.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque1.Y + 24, piezas.bloque1.X
  If matris((piezas.bloque1.Y / 24) + 1, (piezas.bloque1.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y + 24, piezas.bloque1.X - 24
  If matris(piezas.bloque1.Y / 24, (piezas.bloque1.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque1.Y, piezas.bloque1.X - 24
 Case Is = 2
  cuadrado_relleno piezas.bloque2.X, piezas.bloque2.Y, 26, Picture1.BackColor, Picture1
  If piezas.bloque2.Y > 0 Then
   If matris((piezas.bloque2.Y / 24) - 1, (piezas.bloque2.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y - 24, piezas.bloque2.X - 24
   If matris((piezas.bloque2.Y / 24) - 1, piezas.bloque2.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque2.Y - 24, piezas.bloque2.X
   If matris((piezas.bloque2.Y / 24) - 1, (piezas.bloque2.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y - 24, piezas.bloque2.X + 24
  End If
  If matris(piezas.bloque2.Y / 24, (piezas.bloque2.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y, piezas.bloque2.X + 24
  If matris((piezas.bloque2.Y / 24) + 1, (piezas.bloque2.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y + 24, piezas.bloque2.X + 24
  If matris((piezas.bloque2.Y / 24) + 1, piezas.bloque2.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque2.Y + 24, piezas.bloque2.X
  If matris((piezas.bloque2.Y / 24) + 1, (piezas.bloque2.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y + 24, piezas.bloque2.X - 24
  If matris(piezas.bloque2.Y / 24, (piezas.bloque2.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque2.Y, piezas.bloque2.X - 24
 Case Is = 3
  cuadrado_relleno piezas.bloque3.X, piezas.bloque3.Y, 26, Picture1.BackColor, Picture1
  If piezas.bloque3.Y > 0 Then
   If matris((piezas.bloque3.Y / 24) - 1, (piezas.bloque3.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y - 24, piezas.bloque3.X - 24
   If matris((piezas.bloque3.Y / 24) - 1, piezas.bloque3.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque3.Y - 24, piezas.bloque3.X
   If matris((piezas.bloque3.Y / 24) - 1, (piezas.bloque3.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y - 24, piezas.bloque3.X + 24
  End If
  If matris(piezas.bloque3.Y / 24, (piezas.bloque3.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y, piezas.bloque3.X + 24
  If matris((piezas.bloque3.Y / 24) + 1, (piezas.bloque3.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y + 24, piezas.bloque3.X + 24
  If matris((piezas.bloque3.Y / 24) + 1, piezas.bloque3.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque3.Y + 24, piezas.bloque3.X
  If matris((piezas.bloque3.Y / 24) + 1, (piezas.bloque3.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y + 24, piezas.bloque3.X - 24
  If matris(piezas.bloque3.Y / 24, (piezas.bloque3.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque3.Y, piezas.bloque3.X - 24
 Case Is = 4
  cuadrado_relleno piezas.bloque4.X, piezas.bloque4.Y, 26, Picture1.BackColor, Picture1
  If piezas.bloque4.Y > 0 Then
   If matris((piezas.bloque4.Y / 24) - 1, (piezas.bloque4.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y - 24, piezas.bloque4.X - 24
   If matris((piezas.bloque4.Y / 24) - 1, piezas.bloque4.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque4.Y - 24, piezas.bloque4.X
   If matris((piezas.bloque4.Y / 24) - 1, (piezas.bloque4.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y - 24, piezas.bloque4.X + 24
  End If
  If matris(piezas.bloque4.Y / 24, (piezas.bloque4.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y, piezas.bloque4.X + 24
  If matris((piezas.bloque4.Y / 24) + 1, (piezas.bloque4.X / 24) + 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y + 24, piezas.bloque4.X + 24
  If matris((piezas.bloque4.Y / 24) + 1, piezas.bloque4.X / 24) > 0 Then dibujar_matris_bloque piezas.bloque4.Y + 24, piezas.bloque4.X
  If matris((piezas.bloque4.Y / 24) + 1, (piezas.bloque4.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y + 24, piezas.bloque4.X - 24
  If matris(piezas.bloque4.Y / 24, (piezas.bloque4.X / 24) - 1) > 0 Then dibujar_matris_bloque piezas.bloque4.Y, piezas.bloque4.X - 24
 End Select
End Sub

Private Sub Form_Load()
 Randomize
 Cargar_Scores
 detenido = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label7.ForeColor = &H80FF& Then Label7.ForeColor = &HFF0000
End Sub

Private Sub preparate()
 Dim sec As Integer
 Dim count As Integer
 Label1.Top = 223
 Label1.Left = 97
 Label1.visible = True
 Label5.Top = 248
 Label5.Left = 149
 Label5.Caption = "3"
 Label5.visible = True
 sec = Second(Time)
 count = 3
 Do While count > 0
  DoEvents
  Do While sec = Second(Time)
  Loop
  sec = Second(Time)
  count = count - 1
  Label5.Caption = Trim(Str(count))
 Loop
 Label1.visible = False
 Label1.Top = 7320
 Label1.Left = 7725
 Label5.visible = False
 Label5.Top = 7695
 Label5.Left = 8505
 detenido = False
 Timer1.Enabled = True
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If detenido = False Then
  modo = Index + 1
  Select Case modo
  Case Is = 1
   Label9.Caption = "Modo: Clasico"
  Case Is = 2
   Label9.Caption = "Modo: Rising"
  Case Is = 3
   Label9.Caption = "Modo: Strafe"
  Case Is = 4
   Label9.Caption = "Modo: Obstacles"
  End Select
  iniciar_juego
 End If
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Me.Move Me.Left + (X - CURX), Me.Top + (Y - CURY)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label1.visible = False Then
  Guardar_Scores
  creditos.Enabled = True
  creditos.visible = True
  Unload Form1
 End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label7.ForeColor <> &H80FF& Then Label7.ForeColor = &H80FF&
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
 If Label1.visible = False Then
  Select Case KeyCode
  Case Is = vbKeyEscape
   Guardar_Scores
   creditos.Enabled = True
   creditos.visible = True
   Unload Form1
  Case Is = vbKeyDown
   If piezas.visible = True And detenido = False Then
    If colision_pieza(1) = False Then mover_pieza 1
   End If
  Case Is = vbKeyLeft
   If piezas.visible = True And detenido = False Then
    If colision_pieza(2) = False Then mover_pieza 2
   End If
  Case Is = vbKeyRight
   If piezas.visible = True And detenido = False Then
    If colision_pieza(3) = False Then mover_pieza 3
   End If
  Case Is = vbKeySpace
   If piezas.visible = True And detenido = False And piezas.tipo > 1 Then
    If colision_pieza(4) = False Then rotar_pieza
   End If
  Case Is = vbKeyP
   If Label1.visible = False Then
    If Label8.visible = False Then
     Label8.Top = 225
     Label8.Left = 126
     Label8.visible = True
     Timer1.Enabled = False
     detenido = True
    Else
     Label8.visible = False
     Label8.Top = 1000
     Label8.Left = 1000
     Timer1.Enabled = True
     detenido = False
    End If
   End If
  End Select
 End If
End Sub

Private Sub Timer1_Timer()
 If Label1.visible = False And detenido = False Then
 Text2.SetFocus
 If piezas.visible = True Then
  If colision_pieza(1) = True Then
   asentar_pieza
   check_lines
   If llego_arriba = True Then
    game_over
   Else
    If nivel = 1 And lineas > 10 Then
     nivel = 2
     Label2.Caption = "Nivel: 2"
     Timer1.Interval = 350
    Else
     If nivel = 2 And lineas > 20 Then
      nivel = 3
      Label2.Caption = "Nivel: 3"
      Timer1.Interval = 325
     Else
      If nivel = 3 And lineas > 30 Then
       nivel = 4
       Label2.Caption = "Nivel: 4"
       Timer1.Interval = 300
      Else
       If nivel = 4 And lineas > 40 Then
        nivel = 5
        Label2.Caption = "Nivel: 5"
        Timer1.Interval = 275
       Else
        If nivel = 5 And lineas > 50 Then
         nivel = 6
         Label2.Caption = "Nivel: 6"
         Timer1.Interval = 250
        Else
         If nivel = 6 And lineas > 60 Then
          nivel = 7
          Label2.Caption = "Nivel: 7"
          Timer1.Interval = 225
         Else
          If nivel = 7 And lineas > 70 Then
           nivel = 8
           Label2.Caption = "Nivel: 8"
           Timer1.Interval = 200
          Else
           If nivel = 8 And lineas > 80 Then
            nivel = 9
            Label2.Caption = "Nivel: 9"
            Timer1.Interval = 175
           Else
            If nivel = 9 And lineas > 90 Then
             nivel = 10
             Label2.Caption = "Nivel: 10"
             Timer1.Interval = 150
            Else
             If nivel = 10 And lineas > 100 Then
              ganaste
             End If
            End If
           End If
          End If
         End If
        End If
       End If
      End If
     End If
    End If
    iniciar_pieza pieza_sig
    pieza_siguiente
   End If
   If contador_piezas = 5 Then
    If modo > 1 And modo < 4 Then
     Label10.Top = 234
     Label10.Left = 117
     If modo = 2 Then
      Label9.Caption = "Modo: Rising"
      Label10.Caption = "Rising..."
      Label10.visible = True
      rise
     Else
      If modo = 3 Then
       Label9.Caption = "Modo: Strafe"
       Label10.Caption = "Strafing..."
       Label10.visible = True
       strafe
      End If
     End If
     Label10.visible = False
     Label10.Top = 6000
     Label10.Left = 6000
    End If
    contador_piezas = 0
   Else
    contador_piezas = contador_piezas + 1
    If modo = 2 Then
     Label9.Caption = "Modo: Rising " + Trim(Str(5 - contador_piezas))
    Else
     If modo = 3 Then
      Label9.Caption = "Modo: Strafe " + Trim(Str(5 - contador_piezas))
     End If
    End If
   End If
  Else
   mover_pieza 1
  End If
 End If
 End If
End Sub

Private Function colision_pieza(mov As Integer) As Boolean
 If detenido = True Then Exit Function
 Select Case piezas.tipo
 Case Is = 1
  If col1(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 2
  If col2(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 3
  If col3(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 4
  If col4(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 5
  If col5(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 6
  If col6(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 Case Is = 7
  If col7(mov) = True Then
   colision_pieza = True
  Else
   colision_pieza = False
  End If
 End Select
End Function

Private Function col1(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
 Select Case mov
 Case Is = 1 'abajo
  If matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
   col1 = False
  Else
   col1 = True
  End If
 Case Is = 2 'iz
  If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
   col1 = False
  Else
   col1 = True
  End If
 Case Is = 3 'der
  If matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
   col1 = False
  Else
   col1 = True
  End If
 End Select
End Function

Private Function col2(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
  If piezas.rot = 1 Then
   Select Case mov
   Case Is = 1
    If matris(ubi3_y + 1, ubi3_x) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 4
    If matris(ubi1_y + 1, ubi1_x - 1) = 0 And matris(ubi3_y - 1, ubi3_x + 1) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   End Select
  Else
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 3
    If matris(ubi3_y, ubi3_x + 1) = 0 Then
     col2 = False
    Else
     col2 = True
    End If
   Case Is = 4
    If ubi1_y > 0 Then
     If matris(ubi1_y - 1, ubi1_x + 1) = 0 And matris(ubi3_y + 1, ubi3_x - 1) = 0 Then
      col2 = False
     Else
      col2 = True
     End If
    Else
     col2 = True
    End If
   End Select
  End If
End Function

Private Function col3(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
  Select Case piezas.rot
  Case Is = 1
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 3
    If matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 4 'rotacion
    If ubi1_y > 0 Then
     If matris(ubi1_y - 1, ubi1_x + 1) = 0 And matris(ubi3_y + 1, ubi3_x - 1) = 0 And matris(ubi4_y - 1, ubi4_x - 1) = 0 Then
      col3 = False
     Else
      col3 = True
     End If
    Else
     col3 = True
    End If
   End Select
  Case Is = 2
   Select Case mov
   Case Is = 1
    If matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 4 'rotacion
     If matris(ubi1_y + 1, ubi1_x + 1) = 0 And matris(ubi3_y - 1, ubi3_x - 1) = 0 And matris(ubi4_y - 1, ubi4_x + 1) = 0 Then
      col3 = False
     Else
      col3 = True
     End If
   End Select
  Case Is = 3
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 2
    If matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 4 'rotacion
    If ubi4_y > 0 Then
     If matris(ubi1_y + 1, ubi1_x - 1) = 0 And matris(ubi3_y - 1, ubi3_x + 1) = 0 And matris(ubi4_y + 1, ubi4_x + 1) = 0 Then
      col3 = False
     Else
      col3 = True
     End If
    Else
     col3 = True
    End If
   End Select
  Case Is = 4
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y - 1, ubi1_x - 1) = 0 And matris(ubi3_y + 1, ubi3_x + 1) = 0 And matris(ubi4_y + 1, ubi4_x - 1) = 0 Then
     col3 = False
    Else
     col3 = True
    End If
   End Select
  End Select
End Function

Private Function col4(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
  Select Case piezas.rot
  Case Is = 1
   Select Case mov
   Case Is = 1
    If matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 3
    If matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y + 1, ubi1_x + 1) = 0 And matris(ubi3_y - 1, ubi3_x - 1) = 0 And matris(ubi4_y + 2, ubi4_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   End Select
  Case Is = 2
   Select Case mov
   Case Is = 1
    If matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 2
    If matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 4 'rotacion
    If ubi3_y > 0 Then
     If matris(ubi1_y + 1, ubi1_x - 1) = 0 And matris(ubi3_y - 1, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x - 2) = 0 Then
      col4 = False
     Else
      col4 = True
     End If
    Else
     col4 = True
    End If
   End Select
  Case Is = 3
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 2
    If matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y - 1, ubi1_x - 1) = 0 And matris(ubi3_y + 1, ubi3_x + 1) = 0 And matris(ubi4_y - 2, ubi4_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   End Select
  Case Is = 4
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 3
    If matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col4 = False
    Else
     col4 = True
    End If
   Case Is = 4 'rotacion
    If ubi4_y > 0 Then
     If matris(ubi1_y - 1, ubi1_x + 1) = 0 And matris(ubi3_y + 1, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x + 2) = 0 Then
      col4 = False
     Else
      col4 = True
     End If
    Else
     col4 = True
    End If
   End Select
  End Select
End Function

Private Function col5(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
 If detenido = True Then Exit Function
  Select Case piezas.rot
  Case Is = 1
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 3
    If matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y, ubi1_x + 2) = 0 And matris(ubi2_y + 1, ubi2_x + 1) = 0 And matris(ubi4_y - 1, ubi4_x - 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   End Select
  Case Is = 2
   Select Case mov
   Case Is = 1
    If matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi2_y, ubi2_x + 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 4 'rotacion
    If ubi1_y > 0 Then
     If matris(ubi1_y + 2, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x - 1) = 0 And matris(ubi4_y - 1, ubi4_x + 1) = 0 Then
      col5 = False
     Else
      col5 = True
     End If
    Else
     col5 = True
    End If
   End Select
  Case Is = 3
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 2
    If matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y, ubi1_x - 2) = 0 And matris(ubi2_y - 1, ubi2_x - 1) = 0 And matris(ubi4_y + 1, ubi4_x + 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   End Select
  Case Is = 4
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi2_y, ubi2_x - 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col5 = False
    Else
     col5 = True
    End If
   Case Is = 4 'rotacion
    If ubi4_y > 0 Then
     If matris(ubi1_y - 2, ubi1_x) = 0 And matris(ubi2_y - 1, ubi2_x + 1) = 0 And matris(ubi4_y + 1, ubi4_x - 1) = 0 Then
      col5 = False
     Else
      col5 = True
     End If
    Else
     col5 = True
    End If
   End Select
  End Select
End Function

Private Function col6(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
  If piezas.rot = 1 Then
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 3
    If matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 4 'rotacion
     If matris(ubi1_y - 1, ubi1_x + 1) = 0 And matris(ubi3_y + 1, ubi3_x + 1) = 0 And matris(ubi4_y + 2, ubi4_x) = 0 Then
      col6 = False
     Else
      col6 = True
     End If
   End Select
  Else
   Select Case mov
   Case Is = 1
    If matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi2_y, ubi2_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi3_y, ubi3_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y + 1, ubi1_x - 1) = 0 And matris(ubi3_y - 1, ubi3_x - 1) = 0 And matris(ubi4_y - 2, ubi4_x) = 0 Then
     col6 = False
    Else
     col6 = True
    End If
   End Select
  End If
End Function

Private Function col7(mov As Integer) As Boolean
 Dim ubi1_x As Integer
 Dim ubi1_y As Integer
 Dim ubi2_x As Integer
 Dim ubi2_y As Integer
 Dim ubi3_x As Integer
 Dim ubi3_y As Integer
 Dim ubi4_x As Integer
 Dim ubi4_y As Integer
 If detenido = True Then Exit Function
 ubi1_x = (piezas.bloque1.X / 24)
 ubi1_y = (piezas.bloque1.Y / 24)
 ubi2_x = (piezas.bloque2.X / 24)
 ubi2_y = (piezas.bloque2.Y / 24)
 ubi3_x = (piezas.bloque3.X / 24)
 ubi3_y = (piezas.bloque3.Y / 24)
 ubi4_x = (piezas.bloque4.X / 24)
 ubi4_y = (piezas.bloque4.Y / 24)
  If piezas.rot = 1 Then
   Select Case mov
   Case Is = 1
    If matris(ubi1_y + 1, ubi1_x) = 0 And matris(ubi3_y + 1, ubi3_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 3
    If matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 4 'rotacion
    If ubi1_y > 0 Then
     If matris(ubi1_y, ubi1_x + 2) = 0 And matris(ubi2_y + 1, ubi2_x + 1) = 0 And matris(ubi4_y + 1, ubi4_x - 1) = 0 Then
      col7 = False
     Else
      col7 = True
     End If
    Else
     col7 = True
    End If
   End Select
  Else
   Select Case mov
   Case Is = 1
    If matris(ubi2_y + 1, ubi2_x) = 0 And matris(ubi4_y + 1, ubi4_x) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 2
    If matris(ubi1_y, ubi1_x - 1) = 0 And matris(ubi3_y, ubi3_x - 1) = 0 And matris(ubi4_y, ubi4_x - 1) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 3
    If matris(ubi1_y, ubi1_x + 1) = 0 And matris(ubi2_y, ubi2_x + 1) = 0 And matris(ubi4_y, ubi4_x + 1) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   Case Is = 4 'rotacion
    If matris(ubi1_y, ubi1_x - 2) = 0 And matris(ubi2_y - 1, ubi2_x - 1) = 0 And matris(ubi4_y - 1, ubi4_x + 1) = 0 Then
     col7 = False
    Else
     col7 = True
    End If
   End Select
  End If
End Function

Private Sub check_lines()
 Dim n As Integer
 Dim cuantas As Integer
 Dim numerobloques As Integer
 Dim numero_de_lineas(19) As Integer
 Dim f As Integer
 If detenido = True Then Exit Sub
 n = 0
 Do While n <= 19
  numero_de_lineas(n) = 0
  n = n + 1
 Loop
 n = 0
 cuantas = 0
 Do While n <= 19
  f = 1
  numerobloques = 0
  Do While f <= 11
   If matris(n, f) > 0 Then numerobloques = numerobloques + 1
   f = f + 1
  Loop
  If numerobloques = 11 Then
   cuantas = cuantas + 1
   lineas = lineas + 1
   Label3.Caption = "Lineas: " + Trim(Str(lineas))
   numero_de_lineas(n) = 1
  End If
  n = n + 1
 Loop
 If cuantas > 0 Then
  If cuantas > 1 Then
   Label11.Caption = "x " + Trim(Str(cuantas))
   Label11.Top = 236
   Label11.Left = 116
   Label11.visible = True
  End If
  puntaje = puntaje + ((100 * cuantas) + (50 * (cuantas - 1)))
  Label4.Caption = "Puntaje: " + Trim(Str(puntaje))
  blinker_lineas numero_de_lineas()
  If cuantas > 1 Then
   Label11.visible = False
   Label11.Top = 6000
   Label11.Left = 6000
  End If
 End If
End Sub

Private Sub blinker_lineas(cuales() As Integer)
 Dim n As Integer
 If detenido = True Then Exit Sub
 n = 0
 Do While n <= 19
  If cuales(n) = 1 Then
   rectangulo_relleno 26, (n * 24) + 2, 24 * 12, ((n * 24) + 24) - 1, &HFFFF&, Picture1
   suprimir_linea n
  End If
  n = n + 1
 Loop
 n = Second(Time)
 Do While n = Second(Time)
  DoEvents
 Loop
 dibujar_matris
End Sub

Private Sub suprimir_linea(cual As Integer)
 Dim n As Integer
 Dim f As Integer
 n = cual
 If detenido = True Then Exit Sub
 Do While n > 0
  f = 1
  Do While f <= 11
   matris(n, f) = matris(n - 1, f)
   matris(n - 1, f) = 0
   f = f + 1
  Loop
  n = n - 1
 Loop
End Sub

Private Function llego_arriba() As Boolean
 Dim n As Integer
 If detenido = True Then Exit Function
 llego_arriba = False
 n = 1
 Do While n <= 11
  If matris(0, n) > 0 Then
   llego_arriba = True
   Exit Do
  End If
  n = n + 1
 Loop
End Function

Private Sub game_over()
 Dim n As Integer
 Dim f As Integer
 n = 0
 If detenido = True Then Exit Sub
 Do While n <= 19
  f = 1
  Do While f <= 11
   If matris(n, f) > 0 Then
    cuadrado_relleno (f * 24) + 2, (n * 24) + 2, 22, &H404040, Picture1
    cuadrado_vacio f * 24, n * 24, 24, 0, Picture1
    cuadrado_vacio (f * 24) + 1, (n * 24) + 1, 24, 0, Picture1
   End If
   f = f + 1
  Loop
  n = n + 1
 Loop
 MsgBox ("perdiste")
 iniciar_juego
End Sub

Private Sub ganaste()
 Dim n As Integer
 Dim hay As Boolean
 Dim cual As Integer
 If detenido = True Then Exit Sub
 Timer1.Enabled = False
 n = 0
 hay = False
 Do While n <= 9
  If score(n).score < puntaje Then
   hay = True
   cual = n
   Exit Do
  End If
  n = n + 1
 Loop
 If hay = True Then
  Form1.Enabled = False
  n = 0
  Do While n <= 9
   If n <> cual Then
    Form2.punt(n).Caption = Trim(Str(n + 1)) + ". " + score(n).Nombre + " - " + Trim(Str(score(n).score))
   Else
    Form2.punt(n).Caption = Trim(Str(n + 1)) + ".                                                        - " + Trim(Str(puntaje))
   End If
   n = n + 1
  Loop
  detenido = True
  Form2.Label2.Caption = Trim(Str(cual))
  Form2.Label3.Caption = Trim(Str(puntaje))
  Form2.Text1.Top = 570 + (cual * 300)
  Form2.Text1.Text = ""
  Form2.Text1.visible = True
  Form2.Enabled = True
  Form2.visible = True
  Form2.Text1.SetFocus
 End If
End Sub

Private Sub iniciar_juego()
 Timer1.Enabled = False
 detenido = True
 Picture1.Cls
 Picture2.Cls
 iniciar_matris
 puntaje = 0
 nivel = 1
 lineas = 0
 Timer1.Interval = 375
 contador_piezas = 0
 Label2.Caption = "Nivel: 1"
 Label3.Caption = "Lineas: 0"
 Label4.Caption = "Puntaje: 0"
 If modo = 4 Then
  poner_obstaculos
 End If
 dibujar_matris
 preparate
 iniciar_pieza (Rnd(5) * 6) + 1
 pieza_siguiente
End Sub

Private Sub poner_obstaculos()
 Dim cant_obstacle As Integer
 cant_obstacle = 0
 Do While cant_obstacle <= 2
  DoEvents
  cant_obstacle = cant_obstacle + 1
  matris((Rnd(9) * 10) + 7, (Rnd(9) * 10) + 1) = 1
 Loop
End Sub

Private Sub rise()
 Dim n As Integer
 Dim f As Integer
 If detenido = True Then Exit Sub
 n = 0
 Do While n < 19
  f = 1
  Do While f <= 11
   matris(n, f) = matris(n + 1, f)
   f = f + 1
  Loop
  n = n + 1
 Loop
 f = 1
 Do While f <= 11
  matris(19, f) = matris(18, f)
  f = f + 1
 Loop
 dibujar_matris
End Sub

Private Sub strafe()
 Dim n As Integer
 Dim f As Integer
 Dim valor As Integer
 If detenido = True Then Exit Sub
 n = 0
 Do While n <= 19
  f = 11
  valor = matris(n, 11)
  Do While f > 1
   matris(n, f) = matris(n, f - 1)
   f = f - 1
  Loop
  matris(n, 1) = valor
  n = n + 1
 Loop
 dibujar_matris
End Sub
