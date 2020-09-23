VERSION 5.00
Begin VB.Form creditos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2670
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   2580
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   4320
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   855
         Top             =   2160
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "creditos.frx":0000
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "gotuzzo@excite.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   60
         TabIndex        =   5
         Top             =   1545
         Width           =   4200
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Todos Los Derechos Reservados."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   2160
         Width           =   4185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "El 9 de Mayo de 2006"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   315
         TabIndex        =   3
         Top             =   1095
         Width           =   3840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Julio C. Gotuzzo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   675
         Width           =   4050
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tetris Desarrollado Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   285
         Width           =   4005
      End
   End
End
Attribute VB_Name = "creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desarrollado Por Julio C. Gotuzzo
'Buenos Aires - Argentina
'gotuzzo@excite.com

Dim reloj As Integer
Dim cont As Integer

Private Sub Form_Click()
 End
End Sub

Private Sub Form_Load()
 reloj = Second(Time)
 cont = 0
End Sub

Private Sub Timer1_Timer()
 If cont = 5 Then End
 If reloj <> Second(Time) Then
  reloj = Second(Time)
  cont = cont + 1
 End If
End Sub
