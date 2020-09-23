Attribute VB_Name = "Module1"
'Desarrollado Por Julio C. Gotuzzo
'Buenos Aires - Argentina
'gotuzzo@excite.com

Public Type Scores
 Nombre As String
 score As Long
End Type

Public score(9) As Scores
Public detenido As Boolean

Public Sub cuadrado_relleno(X As Long, Y As Long, tamano As Long, color As Long, donde As PictureBox)
 Dim n As Long
 donde.ForeColor = color
 n = 0
 Do While n <= tamano
  donde.Line (X, Y + n)-(X + tamano, Y + n)
  n = n + 1
 Loop
End Sub

Public Sub cuadrado_vacio(X As Long, Y As Long, tamano As Long, color As Long, donde As PictureBox)
 donde.ForeColor = color
 donde.Line (X, Y)-(X + tamano, Y)
 donde.Line (X, Y)-(X, Y + tamano)
 donde.Line (X, Y + tamano)-(X + tamano, Y + tamano)
 donde.Line (X + tamano, Y)-(X + tamano, Y + tamano)
End Sub

Public Sub rectangulo_relleno(x1 As Long, y1 As Long, x2 As Long, y2 As Long, color As Long, donde As PictureBox)
 Dim n As Long
 donde.ForeColor = color
 n = y1
 Do While n <= y2
  donde.Line (x1, n)-(x2, n)
  n = n + 1
 Loop
End Sub

Private Function encode(palabra As String) As String
 Dim n As Integer
 Dim letra As String
 n = 1
 encode = ""
 Do While n <= Len(palabra)
  letra = Mid(palabra, n, 1)
  letra = Chr(Asc(letra) + 20)
  encode = encode + letra
  n = n + 1
 Loop
End Function

Private Function decode(palabra As String) As String
 Dim n As Integer
 Dim letra As String
 n = 1
 decode = ""
 Do While n <= Len(palabra)
  letra = Mid(palabra, n, 1)
  letra = Chr(Asc(letra) - 20)
  decode = decode + letra
  n = n + 1
 Loop
End Function

Public Function Cargar_Scores() As Boolean
    Dim sLineIn As String
    Dim n As Integer
    n = 0
    On Error GoTo ErrLoadListFromFile
    
    Open App.Path + "\scores.dat" For Input As #1
    Do While n <= 9
        Line Input #1, sLineIn
        score(n).Nombre = decode(sLineIn)
        Line Input #1, sLineIn
        score(n).score = Val(Trim(decode(sLineIn)))
        n = n + 1
    Loop
    Close #1
 Cargar_Scores = True
AfterLoadListFromFile:
Exit Function

ErrLoadListFromFile:
    Cargar_Scores = False
    Resume AfterLoadListFromFile
    
End Function

Public Function Guardar_Scores() As Boolean
  
Dim iFile As Integer
Dim n As Integer

n = 0
iFile = FreeFile
Open App.Path + "\scores.dat" For Output As #iFile

Do While n <= 9
 Print #iFile, encode(score(n).Nombre)
 Print #iFile, encode(Trim(Str(score(n).score)))
 n = n + 1
Loop
Guardar_Scores = True

ErrorHandler:

Close #iFile
End Function
