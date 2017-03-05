Function PrimeiraLinhaLivreX(NomeSheet As String, _
                                    PrimLinha As Long, _
                                    Coluna As Integer, _
                                    Margem As Integer) As Long
    
    ' Encontra a 1ª linha livre num sheet, tendo em conta uma margem de segurança
    '
    ' Parametros:
    '   NomeSheet - Nome do sheet a analisar
    '   PrimLinha - 1ª linha com informação util no sheet
    '   Coluna    - Coluna do sheet a pesquisar
    '   Margem    - Margem de segurança
    
    Dim LinhaActiva As Long
    Dim Flag As Boolean, a As Integer
    
    Flag = False
    LinhaActiva = PrimLinha - 1
    Sheets(NomeSheet).Activate
    
    Do Until Flag = True
        LinhaActiva = LinhaActiva + 1
        ' Se descobrir uma célula vazia, vai analisar a margem de segurança
        If Trim(Cells(LinhaActiva, Coluna)) = "" Then
            ' Vai analisar a margem
            Flag = True ' Assume que está vazia e vai tentar provar o contrario
            For a = LinhaActiva + 1 To LinhaActiva - 1 + Margem
                If Trim(Cells(a, Coluna)) <> "" Then    ' A margem não está vazia
                    Flag = False
                    Exit For
                End If
            Next a
        End If
    Loop
    
    PrimeiraLinhaLivreX = LinhaActiva
End Function

Function checkRightFile() As Boolean
    checkRightFile = Sheets("Hoja1").Cells(6, 3) = "N° CC" And Sheets("Hoja1").Cells(6, 4) = "Descripcion"
End Function

Function workInProgress(line As Integer) As Boolean
    workInProgress = Sheets("Hoja1").Cells(line, 6) = "EN CURSO"
End Function

Function getBalDirectoryName() As String
    getBalDirectoryName = period_form.year & "-" & period_form.month
End Function

Function getBalFilename(code As String) As String
    getBalFilename = code & "_BAL_" & "_" & period_form.year & "_" & period_form.month & ".xlsx"
End Function

Function copyTemplate(line As Integer, balTemplate As String, balRoot As String) 

    Dim balFilename As String
    Dim balDirectory As String
    

    balDirectory = balRoot & "\" & period_form.year
        
    balFilename = getBalFilename(Sheets("Hoja1").Cells(line, 3).Value)
    
    If Len(Dir(balDirectory, vbDirectory)) = 0 Then
       MkDir balDirectory
    End If
    
    balDirectory = balDirectory & "\" & getBalDirectoryName()
    
    If Len(Dir(balDirectory, vbDirectory)) = 0 Then
       MkDir balDirectory
    End If
    
    balFilename = balDirectory & "\" & balFilename
    
    FileCopy balTemplate, balFilename

End Function

Sub inicio_BAL()

	' 0- PERIODO DE ANALISIS

	Load period_form
	
	period_form.Show 0

    
End Sub

Sub BALANCES_NP()

    Dim balRoot As String
    Dim balTemplate As String
    Dim workList As String
    Dim lRow
    Dim x As Integer
    Dim wbName As String
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim rngSource As Range
    Dim rngDest As Range
    Dim lastLine As Integer
    Dim i As Integer
    
    
    
    balRoot = "D:\Nicolas Perea\Desktop\Planeacion\01_Balancetes"
    balTemplate = "D:\Nicolas Perea\Desktop\Planeacion\BAL_Tipo.xlsx"
    workList = "D:\Nicolas Perea\Desktop\Planeacion\Listado-Obras-Rev03.xlsx"
    
    Workbooks.Open Filename:=workList
    
    Windows("Listado-Obras-Rev03.xlsx").Activate

    If checkRightFile() Then
    
        MsgBox "ARCHIVO CORRECTO", vbInformation
        
        lastLine = PrimeiraLinhaLivreX("Hoja1", 8, 3, 1)
        
        For i = 8 To lastLine

            If workInProgress(i) Then
            
               Call copyTemplate(i, balTemplate, balRoot)
            
              
            End If
                            
        Next i
    
    
    Else
    
        MsgBox "ARCHIVO EQUIVOCADO!! ", vbExclamation
    
    End If


End Sub



