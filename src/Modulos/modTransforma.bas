Attribute VB_Name = "Module1"
'****************************************************************************************
'* PROYECTO      : TRANSFORMACIONES
'* CONTENIDO     : ESTUDIA LAS TRANSFORMACIONES A UNA SOLUCION DE SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 18 DE JULIO DE 2014
'* ACTUALIZACION : 18 DE JULIO DE 2014
'****************************************************************************************
Option Explicit

Public miVectorSolucion(1 To 16) As Integer


Public Function Giro90(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    Giro90 = miAuxiliar
End Function

Public Function Giro180(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    Giro180 = miAuxiliar
End Function

Public Function Giro270(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    Giro270 = miAuxiliar
End Function

Public Function Filas12(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    Filas12 = miAuxiliar
End Function

Public Function Filas34(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    Filas34 = miAuxiliar
End Function

Public Function Filas1234(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    Filas1234 = miAuxiliar
End Function

Public Function Columnas12(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    Columnas12 = miAuxiliar
End Function

Public Function Columnas34(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    Columnas34 = miAuxiliar
End Function

Public Function Columnas1234(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    Columnas1234 = miAuxiliar
End Function

Public Function Niveles(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    Niveles = miAuxiliar
End Function

Public Function Torres(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    Torres = miAuxiliar
End Function

Public Function NivelesTorres(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    NivelesTorres = miAuxiliar
End Function

Public Function Horizontal(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    Horizontal = miAuxiliar
End Function

Public Function Vertical(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    Vertical = miAuxiliar
End Function

Public Function TransponerIzquierda(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    TransponerIzquierda = miAuxiliar
End Function

Public Function TransponerDerecha(miEnviado As String) As String
    Dim miAuxiliar As String
    miAuxiliar = ""
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 1, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 5, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 9, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 13, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 2, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 6, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 10, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 14, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 3, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 7, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 11, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 15, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 4, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 8, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 12, 1))
    miAuxiliar = miAuxiliar + Trim(Mid(miEnviado, 16, 1))
    TransponerDerecha = miAuxiliar
End Function

' ENCUENTRA CADA UNA DE LAS 24 TRANSPOSICIONES DE UN PROBLEMA O SOLUCION
Public Sub Transposicion(miEnviado As String)
    ' DECLARACIÓN DE VARIABLES PRIVADAS
    Dim miProblemaTranspuesto As String
    Dim miPrimero As Integer
    Dim miSegundo As Integer
    Dim miTercero As Integer
    Dim miCuarto As Integer
    Dim x As Integer
    
    ' INICIALIZACIÓN DE VARIABLES PRIVADAS
    miProblemaTranspuesto = ""
    
    ' BUSQUEDA DE LAS TRANSPOSICIONES
    For miPrimero = 1 To 4
        For miSegundo = 1 To 4
            If miPrimero <> miSegundo Then
            
                For miTercero = 1 To 4
                    If miPrimero <> miSegundo And _
                           miPrimero <> miTercero And _
                           miSegundo <> miTercero Then
                    
                        For miCuarto = 1 To 4
                            If miPrimero <> miCuarto And _
                               miSegundo <> miCuarto And _
                               miTercero <> miCuarto Then
                                
                                ' CAMBIO DE DATOS POR SU TRANSPOSICION CORRESPONDIENTE
                                ' CALCULA TODOS LOS PROBLEMAS TRANSPUESTOS
                                For x = 1 To 16
                                    If Val(Mid(miEnviado, x, 1)) = 0 Then
                                        miProblemaTranspuesto = miProblemaTranspuesto + "0"
                                    End If
                                    If Val(Mid(miEnviado, x, 1)) = 1 Then
                                        miProblemaTranspuesto = miProblemaTranspuesto + Trim(Str(miPrimero))
                                    End If
                                    If Val(Mid(miEnviado, x, 1)) = 2 Then
                                        miProblemaTranspuesto = miProblemaTranspuesto + Trim(Str(miSegundo))
                                    End If
                                    If Val(Mid(miEnviado, x, 1)) = 3 Then
                                        miProblemaTranspuesto = miProblemaTranspuesto + Trim(Str(miTercero))
                                    End If
                                    If Val(Mid(miEnviado, x, 1)) = 4 Then
                                        miProblemaTranspuesto = miProblemaTranspuesto + Trim(Str(miCuarto))
                                    End If
                                Next x
                                
                                'OJO OJO OJO ENLACE A OTRO PROCEDIMIENTO (LE QUITA GENERALIDAD)
                                Print #78, miProblemaTranspuesto
                                
                                miProblemaTranspuesto = ""
                                
                            End If
                        Next miCuarto
                    End If
                Next miTercero
            End If
        Next miSegundo
    Next miPrimero
End Sub


' FUNCION PARA DETERMINAR LA SOLUCION DEL SUDOKU 4X4
Public Function DeterminaNumeroSolucion(dato As String) As Integer
    Dim i As Integer
    Dim miArregloAnalisis(1 To 16) As Integer
    
    ' CARGA LOS DATOS PARA ANALIZARLOS
    For i = 1 To 16
        miArregloAnalisis(i) = Val(Mid(Trim(dato), i, 1))
    Next i
    
    Dim miSolucion(288) As String
    miSolucion(1) = "1234341221434321"
    miSolucion(2) = "1243431221343421"
    miSolucion(3) = "1324241331424231"
    miSolucion(4) = "1342421331242431"
    miSolucion(5) = "1423231441323241"
    miSolucion(6) = "1432321441232341"
    miSolucion(7) = "2134342112434312"
    miSolucion(8) = "2143432112343412"
    miSolucion(9) = "2314142332414132"
    miSolucion(10) = "2341412332141432"
    miSolucion(11) = "2413132442313142"
    miSolucion(12) = "2431312442131342"
    miSolucion(13) = "3124243113424213"
    miSolucion(14) = "3142423113242413"
    miSolucion(15) = "3214143223414123"
    miSolucion(16) = "3241413223141423"
    miSolucion(17) = "3412123443212143"
    miSolucion(18) = "3421213443121243"
    miSolucion(19) = "4123234114323214"
    miSolucion(20) = "4132324114232314"
    miSolucion(21) = "4213134224313124"
    miSolucion(22) = "4231314224131324"
    miSolucion(23) = "4312124334212134"
    miSolucion(24) = "4321214334121234"
    miSolucion(25) = "1234431221433421"
    miSolucion(26) = "1243341221344321"
    miSolucion(27) = "1324421331422431"
    miSolucion(28) = "1342241331244231"
    miSolucion(29) = "1423321441322341"
    miSolucion(30) = "1432231441233241"
    miSolucion(31) = "2134432112433412"
    miSolucion(32) = "2143342112344312"
    miSolucion(33) = "2314412332411432"
    miSolucion(34) = "2341142332144132"
    miSolucion(35) = "2413312442311342"
    miSolucion(36) = "2431132442133142"
    miSolucion(37) = "3124423113422413"
    miSolucion(38) = "3142243113244213"
    miSolucion(39) = "3214413223411423"
    miSolucion(40) = "3241143223144123"
    miSolucion(41) = "3412213443211243"
    miSolucion(42) = "3421123443122143"
    miSolucion(43) = "4123324114322314"
    miSolucion(44) = "4132234114233214"
    miSolucion(45) = "4213314224311324"
    miSolucion(46) = "4231134224133124"
    miSolucion(47) = "4312214334211234"
    miSolucion(48) = "4321124334122134"
    miSolucion(49) = "1234341241232341"
    miSolucion(50) = "1243431231242431"
    miSolucion(51) = "1324241341323241"
    miSolucion(52) = "1342421321343421"
    miSolucion(53) = "1423231431424231"
    miSolucion(54) = "1432321421434321"
    miSolucion(55) = "2134342142131342"
    miSolucion(56) = "2143432132141432"
    miSolucion(57) = "2314142342313142"
    miSolucion(58) = "2341412312343412"
    miSolucion(59) = "2413132432414132"
    miSolucion(60) = "2431312412434312"
    miSolucion(61) = "3124243143121243"
    miSolucion(62) = "3142423123141423"
    miSolucion(63) = "3214143243212143"
    miSolucion(64) = "3241413213242413"
    miSolucion(65) = "3412123423414123"
    miSolucion(66) = "3421213413424213"
    miSolucion(67) = "4123234134121234"
    miSolucion(68) = "4132324124131324"
    miSolucion(69) = "4213134234212134"
    miSolucion(70) = "4231314214232314"
    miSolucion(71) = "4312124324313124"
    miSolucion(72) = "4321214314323214"
    miSolucion(73) = "1234341223414123"
    miSolucion(74) = "1243431224313124"
    miSolucion(75) = "1324241332414132"
    miSolucion(76) = "1342421334212134"
    miSolucion(77) = "1423231442313142"
    miSolucion(78) = "1432321443212143"
    miSolucion(79) = "2134342113424213"
    miSolucion(80) = "2143432114323214"
    miSolucion(81) = "2314142331424231"
    miSolucion(82) = "2341412334121234"
    miSolucion(83) = "2413132441323241"
    miSolucion(84) = "2431312443121243"
    miSolucion(85) = "3124243112434312"
    miSolucion(86) = "3142423114232314"
    miSolucion(87) = "3214143221434321"
    miSolucion(88) = "3241413224131324"
    miSolucion(89) = "3412123441232341"
    miSolucion(90) = "3421213442131342"
    miSolucion(91) = "4123234112343412"
    miSolucion(92) = "4132324113242413"
    miSolucion(93) = "4213134221343421"
    miSolucion(94) = "4231314223141423"
    miSolucion(95) = "4312124331242431"
    miSolucion(96) = "4321214332141432"
    miSolucion(97) = "1234341243212143"
    miSolucion(98) = "1243431234212134"
    miSolucion(99) = "1324241342313142"
    miSolucion(100) = "1342421324313124"
    miSolucion(101) = "1423231432414132"
    miSolucion(102) = "1432321423414123"
    miSolucion(103) = "2134342143121243"
    miSolucion(104) = "2143432134121234"
    miSolucion(105) = "2314142341323241"
    miSolucion(106) = "2341412314323214"
    miSolucion(107) = "2413132431424231"
    miSolucion(108) = "2431312413424213"
    miSolucion(109) = "3124243142131342"
    miSolucion(110) = "3142423124131324"
    miSolucion(111) = "3214143241232341"
    miSolucion(112) = "3241413214232314"
    miSolucion(113) = "3412123421434321"
    miSolucion(114) = "3421213412434312"
    miSolucion(115) = "4123234132141432"
    miSolucion(116) = "4132324123141423"
    miSolucion(117) = "4213134231242431"
    miSolucion(118) = "4231314213242413"
    miSolucion(119) = "4312124321343421"
    miSolucion(120) = "4321214312343412"
    miSolucion(121) = "1234431234212143"
    miSolucion(122) = "1243341243212134"
    miSolucion(123) = "1324421324313142"
    miSolucion(124) = "1342241342313124"
    miSolucion(125) = "1423321423414132"
    miSolucion(126) = "1432231432414123"
    miSolucion(127) = "2134432134121243"
    miSolucion(128) = "2143342143121234"
    miSolucion(129) = "2314412314323241"
    miSolucion(130) = "2341142341323214"
    miSolucion(131) = "2413312413424231"
    miSolucion(132) = "2431132431424213"
    miSolucion(133) = "3124423124131342"
    miSolucion(134) = "3142243142131324"
    miSolucion(135) = "3214413214232341"
    miSolucion(136) = "3241143241232314"
    miSolucion(137) = "3412213412434321"
    miSolucion(138) = "3421123421434312"
    miSolucion(139) = "4123324123141432"
    miSolucion(140) = "4132234132141423"
    miSolucion(141) = "4213314213242431"
    miSolucion(142) = "4231134231242413"
    miSolucion(143) = "4312214312343421"
    miSolucion(144) = "4321124321343412"
    miSolucion(145) = "1234342121434312"
    miSolucion(146) = "1243432121343412"
    miSolucion(147) = "1324243131424213"
    miSolucion(148) = "1342423131242413"
    miSolucion(149) = "1423234141323214"
    miSolucion(150) = "1432324141232314"
    miSolucion(151) = "2134341212434321"
    miSolucion(152) = "2143431212343421"
    miSolucion(153) = "2314143232414123"
    miSolucion(154) = "2341413232141423"
    miSolucion(155) = "2413134242313124"
    miSolucion(156) = "2431314242131324"
    miSolucion(157) = "3124241313424231"
    miSolucion(158) = "3142421313242431"
    miSolucion(159) = "3214142323414132"
    miSolucion(160) = "3241412323141432"
    miSolucion(161) = "3412124343212134"
    miSolucion(162) = "3421214343121234"
    miSolucion(163) = "4123231414323241"
    miSolucion(164) = "4132321414232341"
    miSolucion(165) = "4213132424313142"
    miSolucion(166) = "4231312424131342"
    miSolucion(167) = "4312123434212143"
    miSolucion(168) = "4321213434121243"
    miSolucion(169) = "1234432121433412"
    miSolucion(170) = "1243342121344312"
    miSolucion(171) = "1324423131422413"
    miSolucion(172) = "1342243131244213"
    miSolucion(173) = "1423324141322314"
    miSolucion(174) = "1432234141233214"
    miSolucion(175) = "2134431212433421"
    miSolucion(176) = "2143341212344321"
    miSolucion(177) = "2314413232411423"
    miSolucion(178) = "2341143232144123"
    miSolucion(179) = "2413314242311324"
    miSolucion(180) = "2431134242133124"
    miSolucion(181) = "3124421313422431"
    miSolucion(182) = "3142241313244231"
    miSolucion(183) = "3214412323411432"
    miSolucion(184) = "3241142323144132"
    miSolucion(185) = "3412214343211234"
    miSolucion(186) = "3421124343122134"
    miSolucion(187) = "4123321414322341"
    miSolucion(188) = "4132231414233241"
    miSolucion(189) = "4213312424311342"
    miSolucion(190) = "4231132424133142"
    miSolucion(191) = "4312213434211243"
    miSolucion(192) = "4321123434122143"
    miSolucion(193) = "1234432131422413"
    miSolucion(194) = "1243342141322314"
    miSolucion(195) = "1324423121433412"
    miSolucion(196) = "1342243141233214"
    miSolucion(197) = "1423324121344312"
    miSolucion(198) = "1432234131244213"
    miSolucion(199) = "2134431232411423"
    miSolucion(200) = "2143341242311324"
    miSolucion(201) = "2314413212433421"
    miSolucion(202) = "2341143242133124"
    miSolucion(203) = "2413314212344321"
    miSolucion(204) = "2431134232144123"
    miSolucion(205) = "3124421323411432"
    miSolucion(206) = "3142241343211234"
    miSolucion(207) = "3214412313422431"
    miSolucion(208) = "3241142343122134"
    miSolucion(209) = "3412214313244231"
    miSolucion(210) = "3421124323144132"
    miSolucion(211) = "4123321424311342"
    miSolucion(212) = "4132231434211243"
    miSolucion(213) = "4213312414322341"
    miSolucion(214) = "4231132434122143"
    miSolucion(215) = "4312213414233241"
    miSolucion(216) = "4321123424133142"
    miSolucion(217) = "1234432124133142"
    miSolucion(218) = "1243342123144132"
    miSolucion(219) = "1324423134122143"
    miSolucion(220) = "1342243132144123"
    miSolucion(221) = "1423324143122134"
    miSolucion(222) = "1432234142133124"
    miSolucion(223) = "2134431214233241"
    miSolucion(224) = "2143341213244231"
    miSolucion(225) = "2314413234211243"
    miSolucion(226) = "2341143231244213"
    miSolucion(227) = "2413314243211234"
    miSolucion(228) = "2431134241233214"
    miSolucion(229) = "3124421314322341"
    miSolucion(230) = "3142241312344321"
    miSolucion(231) = "3214412324311342"
    miSolucion(232) = "3241142321344312"
    miSolucion(233) = "3412214342311324"
    miSolucion(234) = "3421124341322314"
    miSolucion(235) = "4123321413422431"
    miSolucion(236) = "4132231412433421"
    miSolucion(237) = "4213312423411432"
    miSolucion(238) = "4231132421433412"
    miSolucion(239) = "4312213432411423"
    miSolucion(240) = "4321123431422413"
    miSolucion(241) = "1234342143122143"
    miSolucion(242) = "1243432134122134"
    miSolucion(243) = "1324243142133142"
    miSolucion(244) = "1342423124133124"
    miSolucion(245) = "1423234132144132"
    miSolucion(246) = "1432324123144123"
    miSolucion(247) = "2134341243211243"
    miSolucion(248) = "2143431234211234"
    miSolucion(249) = "2314143241233241"
    miSolucion(250) = "2341413214233214"
    miSolucion(251) = "2413134231244231"
    miSolucion(252) = "2431314213244213"
    miSolucion(253) = "3124241342311342"
    miSolucion(254) = "3142421324311324"
    miSolucion(255) = "3214142341322341"
    miSolucion(256) = "3241412314322314"
    miSolucion(257) = "3412124321344321"
    miSolucion(258) = "3421214312344312"
    miSolucion(259) = "4123231432411432"
    miSolucion(260) = "4132321423411423"
    miSolucion(261) = "4213132431422431"
    miSolucion(262) = "4231312413422413"
    miSolucion(263) = "4312123421433421"
    miSolucion(264) = "4321213412433412"
    miSolucion(265) = "1234432134122143"
    miSolucion(266) = "1243342143122134"
    miSolucion(267) = "1324423124133142"
    miSolucion(268) = "1342243142133124"
    miSolucion(269) = "1423324123144132"
    miSolucion(270) = "1432234132144123"
    miSolucion(271) = "2134431234211243"
    miSolucion(272) = "2143341243211234"
    miSolucion(273) = "2314413214233241"
    miSolucion(274) = "2341143241233214"
    miSolucion(275) = "2413314213244231"
    miSolucion(276) = "2431134231244213"
    miSolucion(277) = "3124421324311342"
    miSolucion(278) = "3142241342311324"
    miSolucion(279) = "3214412314322341"
    miSolucion(280) = "3241142341322314"
    miSolucion(281) = "3412214312344321"
    miSolucion(282) = "3421124321344312"
    miSolucion(283) = "4123321423411432"
    miSolucion(284) = "4132231432411423"
    miSolucion(285) = "4213312413422431"
    miSolucion(286) = "4231132431422413"
    miSolucion(287) = "4312213412433421"
    miSolucion(288) = "4321123421433412"

    Dim j As Integer
    Dim miSirve As Boolean
    Dim miContadorIguales As Integer
    Dim miSolucionTemporal As String
    Dim miNumeroSolucionTemporal As Integer
    
    ' REVISA CONTRA TODAS LAS SOLUCIONES
    miContadorIguales = 0
    For j = 1 To 288
        miSirve = True
        For i = 1 To 16
            If miArregloAnalisis(i) <> 0 Then
                If miArregloAnalisis(i) <> Mid(Trim(miSolucion(j)), i, 1) Then
                    miSirve = False
                End If
            End If
        Next i
        If miSirve = True Then
            ' CUENTA LA SOLUCION
            miContadorIguales = miContadorIguales + 1
            miSolucionTemporal = miSolucion(j)
            miNumeroSolucionTemporal = j
        End If
    Next j
    
    ' NO ES UN PROBLEMA DE SUDOKU
    If miContadorIguales = 0 Then
        DeterminaNumeroSolucion = 0
    End If
    
    ' ES AMBIGUO
    If miContadorIguales > 1 Then
        DeterminaNumeroSolucion = 0
    End If
    
    ' EL PROBLEMA TIENE UNA SOLA SOLUCION
    If miContadorIguales = 1 Then
        DeterminaNumeroSolucion = miNumeroSolucionTemporal
        
        
        For i = 1 To 16
            miVectorSolucion(i) = Mid(Trim(miSolucion(miNumeroSolucionTemporal)), i, 1)
        Next i
    End If
End Function

