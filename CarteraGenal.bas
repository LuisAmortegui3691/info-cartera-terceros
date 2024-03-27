Attribute VB_Name = "Módulo1"
Sub macroCartera()

    ' Medicion flujos
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim duraSegundos As Double

    Dim docuEntrada As String, docuSalida As String
    Dim arcAlfasis As String
    Dim i As Long, j As Long
    Dim ultimaFilaAlfa As Long, ultiFilaVi As Long, ultiFilaNy As Long
    Dim letViviana As String, remis As String, poli As String, fechRemi As String, ramo As String, abono As String
    Dim placa As String, responsa As String
    Dim plantilla As String
    Dim plantillaNydia As String
    
    ' Registrar tiempo inicio
    tiempoInicio = Timer
    
    docuEntrada = ThisWorkbook.Sheets("main").Range("C2").Value
    docuSalida = ThisWorkbook.Sheets("main").Range("C3").Value
    
    arcAlfasis = docuEntrada & "Informe Alfasis\"
    arcAlfasis = Dir(arcAlfasis)
    
    Application.DisplayAlerts = True
    Workbooks.Open Filename:=docuEntrada & "Informe Alfasis\" & arcAlfasis
    Application.DisplayAlerts = False
    
    Workbooks(arcAlfasis).Activate
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Vivina"
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = "Nydia"
    
    ' Titulos cabecera Viviana
    Workbooks(arcAlfasis).Sheets("Vivina").Range("A1").Value = "Remisión"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("B1").Value = "Póliza"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("C1").Value = "Fec.Remi"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("D1").Value = "Ramo"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("E1").Value = "Abono"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("F1").Value = "Responsable de pago"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("G1").Value = "Placas"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("H1").Value = "OBSERVACIONES"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("I1").Value = "RESULTADO"
    Workbooks(arcAlfasis).Sheets("Vivina").Range("J1").Value = "ENCARGADA DE AREA"
    
    ' Titulos cabecera Nydia
    Workbooks(arcAlfasis).Sheets("Nydia").Range("A1").Value = "Remisión"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("B1").Value = "Póliza"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("C1").Value = "Fec.Remi"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("D1").Value = "Ramo"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("E1").Value = "Abono"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("F1").Value = "Responsable de pago"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("G1").Value = "Placas"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("H1").Value = "OBSERVACIONES"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("I1").Value = "RESULTADO"
    Workbooks(arcAlfasis).Sheets("Nydia").Range("J1").Value = "ENCARGADA DE AREA"
    
    
    ultimaFilaAlfa = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("A" & Rows.Count).End(xlUp).Row
    ultiFilaVi = Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & Rows.Count).End(xlUp).Row
    ultiFilaNy = Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 4 To ultimaFilaAlfa
        letViviana = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("H" & i).Value
        remis = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("A" & i).Value
        poli = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("D" & i).Value
        fechRemi = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("F" & i).Value
        ramo = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("H" & i).Value
        abono = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("K" & i).Value
        responsa = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("AE" & i).Value
        placa = Workbooks(arcAlfasis).Sheets("GeneralCartera").Range("AF" & i).Value
        
        ' Ciclo Viviana
        
        ultiFilaVi = Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & Rows.Count).End(xlUp).Row
         
        If letViviana = "VD" Then
            For j = 1 To (ultiFilaVi + 1)
                Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & ultiFilaVi + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Vivina").Range("B" & ultiFilaVi + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Vivina").Range("C" & ultiFilaVi + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Vivina").Range("D" & ultiFilaVi + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Vivina").Range("E" & ultiFilaVi + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Vivina").Range("G" & ultiFilaVi + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Vivina").Range("F" & ultiFilaVi + 1).Value = responsa
            Next j
        ElseIf letViviana = "SOAT" Then
            For j = 1 To (ultiFilaVi + 1)
                Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & ultiFilaVi + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Vivina").Range("B" & ultiFilaVi + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Vivina").Range("C" & ultiFilaVi + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Vivina").Range("D" & ultiFilaVi + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Vivina").Range("E" & ultiFilaVi + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Vivina").Range("G" & ultiFilaVi + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Vivina").Range("F" & ultiFilaVi + 1).Value = responsa
            Next j
        ElseIf letViviana = "SALUD COMP" Then
            For j = 1 To (ultiFilaVi + 1)
                Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & ultiFilaVi + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Vivina").Range("B" & ultiFilaVi + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Vivina").Range("C" & ultiFilaVi + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Vivina").Range("D" & ultiFilaVi + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Vivina").Range("E" & ultiFilaVi + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Vivina").Range("G" & ultiFilaVi + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Vivina").Range("F" & ultiFilaVi + 1).Value = responsa
            Next j
        ElseIf letViviana = "SALUD" Then
            For j = 1 To (ultiFilaVi + 1)
                Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & ultiFilaVi + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Vivina").Range("B" & ultiFilaVi + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Vivina").Range("C" & ultiFilaVi + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Vivina").Range("D" & ultiFilaVi + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Vivina").Range("E" & ultiFilaVi + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Vivina").Range("G" & ultiFilaVi + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Vivina").Range("F" & ultiFilaVi + 1).Value = responsa
            Next j
        Else
        
        End If
        
        ' Ciclo Nydia
        
        ultiFilaNy = Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & Rows.Count).End(xlUp).Row
        
        If letViviana = "AUTOMO" Then
            For j = 1 To (ultiFilaNy + 1)
                Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & ultiFilaNy + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Nydia").Range("B" & ultiFilaNy + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Nydia").Range("C" & ultiFilaNy + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Nydia").Range("D" & ultiFilaNy + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Nydia").Range("E" & ultiFilaNy + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Nydia").Range("G" & ultiFilaNy + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Nydia").Range("F" & ultiFilaNy + 1).Value = responsa
            Next j
        ElseIf letViviana = "EXEQUIAL" Then
            For j = 1 To (ultiFilaNy + 1)
                Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & ultiFilaNy + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Nydia").Range("B" & ultiFilaNy + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Nydia").Range("C" & ultiFilaNy + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Nydia").Range("D" & ultiFilaNy + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Nydia").Range("E" & ultiFilaNy + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Nydia").Range("G" & ultiFilaNy + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Nydia").Range("F" & ultiFilaNy + 1).Value = responsa
            Next j
        ElseIf letViviana = "MOTO" Then
            For j = 1 To (ultiFilaNy + 1)
                Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & ultiFilaNy + 1).Value = remis
                Workbooks(arcAlfasis).Sheets("Nydia").Range("B" & ultiFilaNy + 1).Value = poli
                Workbooks(arcAlfasis).Sheets("Nydia").Range("C" & ultiFilaNy + 1).Value = fechRemi
                Workbooks(arcAlfasis).Sheets("Nydia").Range("D" & ultiFilaNy + 1).Value = ramo
                Workbooks(arcAlfasis).Sheets("Nydia").Range("E" & ultiFilaNy + 1).Value = abono
                Workbooks(arcAlfasis).Sheets("Nydia").Range("G" & ultiFilaNy + 1).Value = placa
                Workbooks(arcAlfasis).Sheets("Nydia").Range("F" & ultiFilaNy + 1).Value = responsa
            Next j
        Else
        
        End If
        
    Next i
    
    ' Apertura plantilla Viviana
    plantilla = docuEntrada & "plantilla\"
    plantilla = Dir(plantilla)
    
    
    ' Apertura plantilla Nydia
    plantillaNydia = docuEntrada & "plantilla Nydia\"
    plantillaNydia = Dir(plantillaNydia)
    
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=docuEntrada & "plantilla\" & plantilla
    Application.DisplayAlerts = True
    
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=docuEntrada & "plantilla Nydia\" & plantillaNydia
    Application.DisplayAlerts = True
    
    ultiFilaVi = Workbooks(arcAlfasis).Sheets("Vivina").Range("A" & Rows.Count).End(xlUp).Row
    ultiFilaNy = Workbooks(arcAlfasis).Sheets("Nydia").Range("A" & Rows.Count).End(xlUp).Row
    
    Workbooks(arcAlfasis).Sheets("Vivina").Range("A2:" & "K" & ultiFilaVi).Copy
    Workbooks(plantilla).Sheets("Vivina").Range("A2").PasteSpecial
    
    Workbooks(arcAlfasis).Sheets("Nydia").Range("A2:" & "K" & ultiFilaVi).Copy
    Workbooks(plantillaNydia).Sheets("Nydia").Range("A2").PasteSpecial
    
    ' Tiempo final
    tiempoFin = Timer
    ' calcular tiempo
    duraSegundos = tiempoFin - tiempoInicio
    
    ' Mostrar tiempo
    Debug.Print "Duracion en segundos es " & duraSegundos
    
    MsgBox "Procesos terminado"

End Sub
