Attribute VB_Name = "Módulo1"
Sub calculo_rigidez_por_barra()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("datos")
    Set sh2 = ThisWorkbook.Sheets("rigidez_global_barra")
'cuenta el numero de barras
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    Dim b As Long
    b = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    b = b - 1

'crea las matrices de rigidez en local por barra
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    Dim RL() As Double
    Dim TempRG() As Variant
    Dim CC() As Double
    Dim Dat() As Double
    Dim Swap(1 To 6, 1 To 6) As Double
    Dim CCSwap(1 To 6, 1 To 6) As Double
    ReDim Dat(5, b)
    ReDim CC(6, 6, b)
    ReDim RL(6, 6, b)
'recoge los datos de las barras
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    For i = 1 To 5
        For j = 1 To b
            Dat(i, j) = sh.Cells(j + 1, i + 3)
        Next j
    Next i
    
'rellena las matrices KC
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    For k = 1 To b
        For j = 1 To 6
            For i = 1 To 6
                Select Case True
                    Case (i = 1 And j = 1)
                        RL(i, j, k) = (Dat(1, k) * Dat(3, k)) / Dat(2, k)
                    Case (i = 2 And j = 1)
                        RL(i, j, k) = 0
                    Case (i = 3 And j = 1)
                        RL(i, j, k) = 0
                    Case (i = 4 And j = 1)
                        RL(i, j, k) = -(Dat(1, k) * Dat(3, k)) / Dat(2, k)
                    Case (i = 5 And j = 1)
                        RL(i, j, k) = 0
                    Case (i = 6 And j = 1)
                        RL(i, j, k) = 0
                    Case (i = 1 And j = 2)
                        RL(i, j, k) = 0
                    Case (i = 2 And j = 2)
                        RL(i, j, k) = (12 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 3)
                    Case (i = 3 And j = 2)
                        RL(i, j, k) = (6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 4 And j = 2)
                        RL(i, j, k) = 0
                    Case (i = 5 And j = 2)
                        RL(i, j, k) = (-12 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 3)
                    Case (i = 6 And j = 2)
                        RL(i, j, k) = (6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 1 And j = 3)
                        RL(i, j, k) = 0
                    Case (i = 2 And j = 3)
                        RL(i, j, k) = (6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 3 And j = 3)
                        RL(i, j, k) = (4 * Dat(3, k) * Dat(4, k)) / (Dat(2, k))
                    Case (i = 4 And j = 3)
                        RL(i, j, k) = 0
                    Case (i = 5 And j = 3)
                        RL(i, j, k) = (-6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 6 And j = 3)
                        RL(i, j, k) = (2 * Dat(3, k) * Dat(4, k)) / (Dat(2, k))
                    Case (i = 1 And j = 4)
                        RL(i, j, k) = -(Dat(1, k) * Dat(3, k)) / Dat(2, k)
                    Case (i = 2 And j = 4)
                        RL(i, j, k) = 0
                    Case (i = 3 And j = 4)
                        RL(i, j, k) = 0
                    Case (i = 4 And j = 4)
                        RL(i, j, k) = (Dat(1, k) * Dat(3, k)) / Dat(2, k)
                    Case (i = 5 And j = 4)
                        RL(i, j, k) = 0
                    Case (i = 6 And j = 4)
                        RL(i, j, k) = 0
                    Case (i = 1 And j = 5)
                        RL(i, j, k) = 0
                    Case (i = 2 And j = 5)
                        RL(i, j, k) = (-12 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 3)
                    Case (i = 3 And j = 5)
                        RL(i, j, k) = (-6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 4 And j = 5)
                        RL(i, j, k) = 0
                    Case (i = 5 And j = 5)
                        RL(i, j, k) = (12 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 3)
                    Case (i = 6 And j = 5)
                        RL(i, j, k) = (-6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 1 And j = 6)
                        RL(i, j, k) = 0
                    Case (i = 2 And j = 6)
                        RL(i, j, k) = (6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 3 And j = 6)
                        RL(i, j, k) = (2 * Dat(3, k) * Dat(4, k)) / (Dat(2, k))
                    Case (i = 4 And j = 6)
                        RL(i, j, k) = 0
                    Case (i = 5 And j = 6)
                        RL(i, j, k) = (-6 * Dat(3, k) * Dat(4, k)) / (Dat(2, k) ^ 2)
                    Case (i = 6 And j = 6)
                        RL(i, j, k) = (4 * Dat(3, k) * Dat(4, k)) / (Dat(2, k))
                End Select
            Next i
        Next j
    Next k
    
'montar matrices cambio coord
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    For k = 1 To b
        For j = 1 To 6
            For i = 1 To 6
                Select Case True
                    Case (i = 1 And j = 1)
                        CC(i, j, k) = Round(Cos(Dat(5, k)), 10)
                    Case (i = 2 And j = 1)
                        CC(i, j, k) = Round(Sin(Dat(5, k)), 10)
                    Case (i = 1 And j = 2)
                        CC(i, j, k) = Round(-Sin(Dat(5, k)), 10)
                    Case (i = 2 And j = 2)
                        CC(i, j, k) = Round(Cos(Dat(5, k)), 10)
                    Case (i = 3 And j = 3)
                        CC(i, j, k) = 1
                    Case (i = 4 And j = 4)
                        CC(i, j, k) = Round(Cos(Dat(5, k)), 10)
                    Case (i = 5 And j = 4)
                        CC(i, j, k) = Round(Sin(Dat(5, k)), 10)
                    Case (i = 4 And j = 5)
                        CC(i, j, k) = Round(-Sin(Dat(5, k)), 10)
                    Case (i = 5 And j = 5)
                        CC(i, j, k) = Round(Cos(Dat(5, k)), 10)
                    Case (i = 6 And j = 6)
                        CC(i, j, k) = 1
                End Select
            Next i
        Next j
    Next k
    
'montar matrices swap y multiplicar
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    For k = 1 To b
        For j = 1 To 6
            For i = 1 To 6
                Swap(i, j) = RL(i, j, k)
                CCSwap(i, j) = CC(i, j, k)
            Next i
        Next j
        
        TempRG = WorksheetFunction.MMult(WorksheetFunction.MMult(CCSwap(), Swap()), WorksheetFunction.Transpose(CCSwap()))
        sh2.Cells(1 + (k - 1) * 8, 1) = "MATRIZ RIGIDEZ LOCAL BARRA " & k
        sh2.Cells(1 + (k - 1) * 8, 8) = "MATRIZ CAMBIO COORDENADAS BARRA " & k
        sh2.Cells(1 + (k - 1) * 8, 15) = "MATRIZ RIGIDEZ GLOBAL BARRA " & k
        For i = 1 To 6
            For j = 1 To 6
                sh2.Cells(1 + i + (8 * (k - 1)), j) = Swap(i, j)
                sh2.Cells(1 + i + (8 * (k - 1)), j).Interior.ColorIndex = 8
                sh2.Cells(1 + i + (8 * (k - 1)), j + 7) = CCSwap(i, j)
                sh2.Cells(1 + i + (8 * (k - 1)), j + 7).Interior.ColorIndex = 15
                sh2.Cells(1 + i + (8 * (k - 1)), j + 14) = TempRG(i, j)
                sh2.Cells(1 + i + (8 * (k - 1)), j + 14).Interior.ColorIndex = 17
            Next j
        Next i
    Next k
End Sub
Sub ensamblar_KE()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("datos")
    Set sh2 = ThisWorkbook.Sheets("rigidez_global_barra")
    Set sh3 = ThisWorkbook.Sheets("ensamblado_matrices_completo")
'cuenta el numero de barras y nudos
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

    Dim b As Long
    Dim n As Long
    b = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    b = b - 1
    n = sh.Range("L1", sh.Range("L1").End(xlDown)).Rows.Count
    n = n - 1
'   Dim Reactions() As Integer
'    Dim n1 As Integer
'    Dim n2 As Integer
'    Dim gdl As Integer
'    ReDim Reactions(n)
'    n1 = 0
'    n2 = 0
'calcula los gdl del sistema a resolver (esto sera para el paso siguiente)!!!!!
'    For i = 1 To n
'    Reactions(i) = 0
'    Reactions(i) = sh.Cells(i + 1, 13)
'        If Reactions(i) = "1" Then
'            n1 = n1 + 1
'        End If
'
'        If Reactions(i) = "2" Then
'            n2 = n2 + 1
'        End If
'    Next i
'    gdl = 3 * n - 3 * n1 - 2 * n2
    Dim KG As Variant
    Dim KT As Variant
    ReDim KG(6, 6, b)
    ReDim KT(3 * n, 3 * n)
    
    For k = 1 To b
        For i = 1 To 6
            For j = 1 To 6
                KG(i, j, k) = sh2.Cells(1 + i + (8 * (k - 1)), j + 14)
            Next j
        Next i
    Next k
    
    Dim nudoi As Integer
    Dim nudoj As Integer
    For i = 1 To 3 * n
        For j = 1 To 3 * n
            KT(i, j) = 0
        Next j
    Next i
    For k = 1 To b
        nudoi = sh.Cells(k + 1, 2)
        nudof = sh.Cells(k + 1, 3)
        For i = 1 To 3
            For j = 1 To 3
                KT((nudoi * 3 - 3 + i), (nudoi * 3 - 3 + j)) = KT((nudoi * 3 - 3 + i), (nudoi * 3 - 3 + j)) + KG(i, j, k)
                KT((nudof * 3 - 3 + i), (nudof * 3 - 3 + j)) = KT((nudof * 3 - 3 + i), (nudof * 3 - 3 + j)) + KG(i + 3, j + 3, k)
                KT((nudoi * 3 - 3 + i), (nudof * 3 - 3 + j)) = KT((nudoi * 3 - 3 + i), (nudof * 3 - 3 + j)) + KG(i, j + 3, k)
                KT((nudof * 3 - 3 + i), (nudoi * 3 - 3 + j)) = KT((nudof * 3 - 3 + i), (nudoi * 3 - 3 + j)) + KG(i + 3, j, k)
            Next j
        Next i
    Next k
'mostramos la matriz en la hoja
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    For i = 1 To 3 * n
            For j = 1 To 3 * n
                sh3.Cells(i, j) = KT(i, j)
                sh3.Cells(i, j).Interior.ColorIndex = 15
            Next j
        Next i
End Sub
Sub reducir_y_resolver()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("datos")
    Set sh2 = ThisWorkbook.Sheets("rigidez_global_barra")
    Set sh3 = ThisWorkbook.Sheets("ensamblado_matrices_completo")
    Set sh4 = ThisWorkbook.Sheets("ke_reducida_solucion")
    Set sh5 = ThisWorkbook.Sheets("esfuerzos_barras_diagramas")
'cuenta el numero de barras y nudos
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    Dim b As Long
    Dim n As Long
    b = sh.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
    b = b - 1
    n = sh.Range("L1", sh.Range("L1").End(xlDown)).Rows.Count
    n = n - 1
    
    Dim Reactions() As Integer
    Dim Emp() As Variant
    Dim n1 As Integer
    Dim n2 As Integer
    Dim gdl As Integer
    ReDim Reactions(3 * n)
    ReDim Emp(3 * n)
    Dim Calc() As Variant
    Dim FNudos() As Variant
    ReDim FNudos(3 * n)
    ReDim Calc(3 * n)
    n1 = 0
    n2 = 0
'calcula los gdl del sistema a resolver, monta vector Fnudos y F empotramiento perfecto
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
For i = 1 To n
    Calc(3 * i - 2) = "despX " & i
    Calc(3 * i - 1) = "despY " & i
    Calc(3 * i) = "giro " & i
    FNudos(3 * i - 2) = sh.Cells(i + 1, 16)
    FNudos(3 * i - 1) = sh.Cells(i + 1, 17)
    FNudos(3 * i) = sh.Cells(i + 1, 18)
    Emp(3 * i - 2) = sh.Cells(i + 1, 19)
    Emp(3 * i - 1) = sh.Cells(i + 1, 20)
    Emp(3 * i) = sh.Cells(i + 1, 21)
Next i
gdl = 3 * n
For i = 1 To n
    If sh.Cells(i + 1, 13) = 1 Then
        Calc(3 * i - 2) = 0
        FNudos(3 * i - 2) = "Rx " & i
        gdl = gdl - 1
    End If
    If sh.Cells(i + 1, 14) = 1 Then
        Calc(3 * i - 1) = 0
        FNudos(3 * i - 1) = "Ry " & i
        gdl = gdl - 1
    End If
    If sh.Cells(i + 1, 15) = 1 Then
        Calc(3 * i) = 0
        FNudos(3 * i) = "M " & i
        gdl = gdl - 1
    End If
Next i


'    For i = 1 To n
'    Reactions(i) = 0
'    Reactions(i) = sh.Cells(i + 1, 13)
'        If Reactions(i) = "0" Then
'            Calc(3 * i - 2) = "despX " & i
'            Calc(3 * i - 1) = "despY " & i
'            Calc(3 * i) = "giro " & i
'            FNudos(3 * i - 2) = sh.Cells(i + 1, 14)
'            FNudos(3 * i - 1) = sh.Cells(i + 1, 15)
'            FNudos(3 * i) = sh.Cells(i + 1, 16)
'            Emp(3 * i - 2) = sh.Cells(i + 1, 17)
'            Emp(3 * i - 1) = sh.Cells(i + 1, 18)
'            Emp(3 * i) = sh.Cells(i + 1, 19)
'        End If
'
'        If Reactions(i) = "1" Then
'            n1 = n1 + 1
'            Calc(3 * i - 2) = 0
'            Calc(3 * i - 1) = 0
'            Calc(3 * i) = 0
'            FNudos(3 * i - 2) = "Rx " & i
'            FNudos(3 * i - 1) = "Ry " & i
'            FNudos(3 * i) = "M " & i
'            Emp(3 * i - 2) = sh.Cells(i + 1, 17)
'            Emp(3 * i - 1) = sh.Cells(i + 1, 18)
'            Emp(3 * i) = sh.Cells(i + 1, 19)
'        End If
'
'        If Reactions(i) = "2" Then
'            n2 = n2 + 1
'            Calc(3 * i - 2) = 0
'            Calc(3 * i - 1) = 0
'            Calc(3 * i) = "giro " & i
'            FNudos(3 * i - 2) = "Rx " & i
'            FNudos(3 * i - 1) = "Ry " & i
'            FNudos(3 * i) = 0
'            Emp(3 * i - 2) = sh.Cells(i + 1, 17)
'            Emp(3 * i - 1) = sh.Cells(i + 1, 18)
'            'Emp(3 * i) = 0
'            Emp(3 * i) = sh.Cells(i + 1, 19)
'        End If
'    Next i
    
    
    
    For i = 1 To 3 * n
            sh3.Cells(i, 3 * n + 2) = FNudos(i)
            sh3.Cells(i, 3 * n + 2).Interior.ColorIndex = 20
            sh3.Cells(i, 3 * n + 4) = Calc(i)
            sh3.Cells(i, 3 * n + 4).Interior.ColorIndex = 17
            If IsNumeric(sh3.Cells(i, 3 * n + 2)) Then
                sh3.Cells(i, 3 * n + 2) = sh3.Cells(i, 3 * n + 2) - Emp(i)
            End If
'copia a una nueva hoja, reduce el sistema y lo resuelve
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    Next i
    sh3.Cells.Copy Destination:=sh4.Cells(1, 1)
    
    For i = 1 To 3 * n
        If Calc(3 * n + 1 - i) = 0 Then
            sh4.Rows(3 * n + 1 - i).Delete
            sh4.Columns(3 * n + 1 - i).Delete
        End If
    Next i
    
    Dim Solution() As Variant
    Dim MInversa() As Variant
    Dim KE_Red() As Double
    Dim V_Fuerzas() As Double
    
    ReDim KE_Red(1 To gdl, 1 To gdl)
    ReDim V_Fuerzas(1 To gdl, 1)
    ReDim Solution(1 To gdl, 1)
    For i = 1 To gdl
        For j = 1 To gdl
            KE_Red(i, j) = sh4.Cells(i, j)
            V_Fuerzas(i, 1) = sh4.Cells(i, gdl + 2)
            Solution(i, 1) = 0
        Next j
    Next i
    
    MInversa = WorksheetFunction.MInverse(KE_Red())
    For i = 1 To gdl
        For j = 1 To gdl
            Solution(i, 1) = Solution(i, 1) + MInversa(i, j) * V_Fuerzas(j, 1)
            sh4.Cells(i + 1 + gdl, j) = MInversa(i, j)
            sh4.Cells(i + 1 + gdl, j).Interior.ColorIndex = 28
            Next j
    Next i
        
    For i = 1 To gdl
        Solution(i, 1) = Round(Solution(i, 1), 10)
        sh4.Cells(i, gdl + 5) = Solution(i, 1)
        sh4.Cells(i, gdl + 5).Interior.ColorIndex = 6
    Next i
    sh4.Cells(i, gdl + 5) = "RESULTADOS"
    Dim VectorDesplazamientos() As Double
    ReDim VectorDesplazamientos(1 To 3 * n, 1 To 1)
    j = 1
    For i = 1 To n * 3
        
        If Not IsNumeric(sh3.Cells(i, 3 * n + 4)) Then
            VectorDesplazamientos(i, 1) = Solution(j, 1)
            j = j + 1
        Else
            VectorDesplazamientos(i, 1) = sh3.Cells(i, 3 * n + 4)
        End If
    Next i
    
    For i = 1 To n * 3
        sh3.Cells(i, 3 * n + 5) = VectorDesplazamientos(i, 1)
        sh3.Cells(i, 3 * n + 5).Interior.ColorIndex = 6
    Next i
    
    
        
'reacciones apoyos
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    Dim MatrizCalculoApoyos() As Double
    ReDim MatrizCalculoApoyos(1 To 1, 1 To 3 * n)
'    Dim MatrizCalculoApoyos() As Double
'    ReDim MatrizCalculoApoyos(1 To 3, 1 To 3 * n)
    Dim Reaccion(1 To 3, 1 To 1) As Variant
    Dim SwapDesplazamientos(1 To 9, 1 To 1) As Double
    
'Si encuentra reaccion en el nudo, monta su matriz  resolver, cogiendola de la matriz completa
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
For Z = 1 To 3 * n
    If Not IsNumeric(sh3.Cells(Z, 3 * n + 2)) Then
        For j = 1 To 3 * n
                MatrizCalculoApoyos(1, j) = sh3.Cells(Z, j)
        Next j
        
        Suma = 0
        For j = 1 To 3 * n
                Suma = Suma + MatrizCalculoApoyos(1, j) * VectorDesplazamientos(j, 1)
        Next j
        sh3.Cells(Z, 3 * n + 3) = Suma + Emp(Z)
        sh3.Cells(Z, 3 * n + 3).Interior.ColorIndex = 6
    End If
Next Z

'    For Z = 1 To n
'        If Reactions(Z) = "1" Then
'             For i = 1 To 3
'                For j = 1 To 3 * n
'                    MatrizCalculoApoyos(i, j) = sh3.Cells(3 * Z - (3 - i), j)
'                Next j
'                SwapDesplazamientos(i, 1) = VectorDesplazamientos(3 * Z - (3 - i), 1)
'            Next i
'
'        'Reaccion = WorksheetFunction.MMult(MatrizCalculoApoyos(), VectorDesplazamientos())
'
'        For i = 1 To 3
'        Suma = 0
'            For j = 1 To 3 * n
'                Suma = Suma + MatrizCalculoApoyos(i, j) * VectorDesplazamientos(j, 1)
'            Next j
'            sh3.Cells(3 * Z - (3 - i), 3 * n + 3) = Suma + Emp(3 * Z - 3 + i)
'            sh3.Cells(3 * Z - (3 - i), 3 * n + 3).Interior.ColorIndex = 6
'        Next i
'      End If
'    Next Z
            
            
        
    
'    Resultados = WorksheetFunction.MMult(WorksheetFunction.MInverse(KE_Red()), V_Fuerzas())
'    For i = 1 To gdl
'            sh4.Cells(i, 9) = Resultados(i, 1)
'            sh4.Cells(i, 9).Interior.ColorIndex = 28
'    Next i

'Calcula las reacciones para cada nudo para resolver los diagramas
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Dim Coord(1 To 6, 1 To 6) As Double
Dim KLoc(1 To 6, 1 To 6) As Double
Dim FGlob(1 To 6, 1 To 1) As Double
Dim Desp(1 To 6, 1 To 1) As Double
Dim Swap() As Variant
Dim FLoc() As Variant
Dim FBarra(1 To 6, 1 To 1) As Variant

For Z = 1 To b
    
    nudoi = sh.Cells(Z + 1, 2)
    nudof = sh.Cells(Z + 1, 3)
    
    sh5.Cells(8 * Z - 7, 2) = "ESFUERZOS LOCALES BARRA " & Z
    sh5.Cells(8 * Z - 7 + 1, 1) = "Fx" & nudoi
    sh5.Cells(8 * Z - 7 + 2, 1) = "Fy" & nudoi
    sh5.Cells(8 * Z - 7 + 3, 1) = "M" & nudoi
    sh5.Cells(8 * Z - 7 + 4, 1) = "Fx" & nudof
    sh5.Cells(8 * Z - 7 + 5, 1) = "Fy" & nudof
    sh5.Cells(8 * Z - 7 + 6, 1) = "M" & nudof
    
    
    For i = 1 To 3
        Desp(i, 1) = VectorDesplazamientos(3 * nudoi - 3 + i, 1)
        Desp(3 + i, 1) = VectorDesplazamientos(3 * nudof - 3 + i, 1)
        FGlob(i, 1) = sh.Cells(nudoi + 1, 18 + i)
        FGlob(i + 3, 1) = sh.Cells(nudof + 1, 18 + i)
    Next i
    For i = 1 To 6
        For j = 1 To 6
            KLoc(i, j) = sh2.Cells(8 * Z - 7 + i, j)
            Coord(j, i) = sh2.Cells(8 * Z - 7 + i, j + 7)
        Next j
    Next i
    FLoc = WorksheetFunction.MMult(Coord(), FGlob())
    Swap = WorksheetFunction.MMult(WorksheetFunction.MMult(KLoc(), Coord()), Desp())
    For i = 1 To 6
        FBarra(i, 1) = FLoc(i, 1) + Swap(i, 1)
        sh5.Cells(8 * Z - 7 + i, 4) = FLoc(i, 1)
        sh5.Cells(8 * Z - 7 + i, 6) = Swap(i, 1)
        sh5.Cells(8 * Z - 7 + i, 2) = FBarra(i, 1)
        sh5.Cells(8 * Z - 7 + i, 2).Interior.ColorIndex = 6
    Next i
    

Next Z

End Sub
Sub calcular_estructura()
    calculo_rigidez_por_barra
    ensamblar_KE
    reducir_y_resolver
End Sub
Sub limpiar_hoja()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("datos")
    Set sh2 = ThisWorkbook.Sheets("rigidez_global_barra")
    Set sh3 = ThisWorkbook.Sheets("ensamblado_matrices_completo")
    Set sh4 = ThisWorkbook.Sheets("ke_reducida_solucion")
    Set sh5 = ThisWorkbook.Sheets("esfuerzos_barras_diagramas")
    sh2.Cells.Clear
    sh3.Cells.Clear
    sh4.Cells.Clear
    sh5.Cells.Clear
End Sub

