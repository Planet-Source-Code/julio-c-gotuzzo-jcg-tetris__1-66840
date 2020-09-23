VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   645
      MaxLength       =   30
      TabIndex        =   12
      Top             =   570
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5820
      TabIndex        =   14
      Top             =   105
      Width           =   300
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   360
      Left            =   3450
      TabIndex        =   13
      Top             =   3675
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   360
      Left            =   4890
      TabIndex        =   11
      Top             =   3675
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   4095
      Left            =   10
      Top             =   10
      Width           =   6180
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   9
      Left            =   420
      TabIndex        =   10
      Top             =   3300
      Width           =   5310
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   8
      Left            =   420
      TabIndex        =   9
      Top             =   3000
      Width           =   5280
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   7
      Left            =   420
      TabIndex        =   8
      Top             =   2700
      Width           =   5295
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   6
      Left            =   420
      TabIndex        =   7
      Top             =   2400
      Width           =   5250
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   5
      Left            =   420
      TabIndex        =   6
      Top             =   2100
      Width           =   5175
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   4
      Left            =   420
      TabIndex        =   5
      Top             =   1800
      Width           =   5280
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   3
      Left            =   420
      TabIndex        =   4
      Top             =   1500
      Width           =   5235
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   2
      Left            =   420
      TabIndex        =   3
      Top             =   1200
      Width           =   5205
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   900
      Width           =   5145
   End
   Begin VB.Label punt 
      BackStyle       =   0  'Transparent
      Caption         =   "1.                                                        - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Width           =   5145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PUNTUACIONES MAXIMAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   1275
      TabIndex        =   0
      Top             =   90
      Width           =   3525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desarrollado Por Julio C. Gotuzzo
'Buenos Aires - Argentina
'gotuzzo@excite.com

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 detenido = False
 Form1.Enabled = True
 Unload Form2
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
  If Text1.Text = "" Then
   MsgBox ("Escriba un nombre")
  Else
   Text1.visible = False
   punt(Val(Label2.Caption)).Caption = Trim(Str(Val(Label2.Caption) + 1)) + ". " + Text1.Text + " - " + Trim(Label3.Caption)
   score(Val(Label2.Caption)).Nombre = Text1.Text
   score(Val(Label2.Caption)).score = Val(Label3.Caption)
  End If
 End If
End Sub

Private Sub Text1_LostFocus()
 If Text1.Text = "" Then
  MsgBox ("Escriba un nombre")
  Text1.SetFocus
 End If
End Sub
