Function ArrayIsEmpty(fArray()) As Boolean
    On Error Resume Next
    
    intUpper = UBound(fArray)
    ArrayIsEmpty = Err
    
End Function

Function rootDataing_data() As Variant
    
    Dim tblrootData As ListObject
    Dim rngrootData As Range
    Dim tableList() As Variant
    Dim sheetList() As Variant
    Dim sourceList() As Variant
    Dim targetList() As Variant
    Dim rootDataList() As Variant
    Dim insertRowIndex As Integer
    Dim rootDataSourceIndex As Integer
    Dim index As Integer
    Dim dataIndex As Integer
    
    Set tblrootData = Worksheets("routing_sheet").ListObjects("routing_table")
    Set rngrootData = tblrootData.ListColumns("source").DataBodyRange
    
    rootDataSourceIndex = 0
    For Each rootDataSource In rngrootData
        If rootDataSource = "idArticle" Then
            ReDim Preserve sheetList(rootDataSourceIndex)
            ReDim Preserve tableList(rootDataSourceIndex)
            sheetList(rootDataSourceIndex) = tblrootData.DataBodyRange(rootDataSource.Row - 1, tblrootData.ListColumns("targetSheet").index).Value
            tableList(rootDataSourceIndex) = tblrootData.DataBodyRange(rootDataSource.Row - 1, tblrootData.ListColumns("targetTable").index).Value
            rootDataSourceIndex = rootDataSourceIndex + 1
        End If
    Next
    
    insertRowIndex = 0
    dataIndex = 0
        For Each Table In tableList
            index = 0
            For Each rootDataSource In rngrootData
                If tblrootData.DataBodyRange(rootDataSource.Row - 1, tblrootData.ListColumns("targetTable").index).Value = Table Then
                ReDim Preserve sourceList(index)
                ReDim Preserve targetList(index)
                sourceList(index) = rootDataSource
                targetList(index) = tblrootData.DataBodyRange(rootDataSource.Row - 1, tblrootData.ListColumns("targetColumn").index).Value
                index = index + 1
                End If
            Next
            ReDim Preserve rootDataList(dataIndex)
            rootDataList(dataIndex) = Array(Table, sheetList(tableIndex), sourceList, targetList)
            tableIndex = tableIndex + 1
            dataIndex = dataIndex + 1
        Next
   
   rootDataing_data = rootDataList

End Function
Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function



Sub AddArticle()

    Dim tblSource As ListObject
    Dim rngSource As Range
    Dim targetWks As Worksheet
    Dim targetTable As ListObject
    Dim targetRng As Range
    Dim newInsertRow As ListRow
    
    Dim newRowList() As Variant
    Dim size As Integer
    size = 0
    
    Dim waitingInsertList() As Variant
    Dim insertSize As Integer
    insertSize = 0
    
    
    Dim duplicate As Boolean
    Dim rootDataResult() As Variant
    Dim printing As String
    Dim printingSource As String
    Dim printingTarget As String
    
    Set tblSource = ActiveSheet.ListObjects("Insertion")
    Set rngSource = tblSource.ListColumns("inserted").DataBodyRange
    
    ' Parcourir chaque ligne dans rngSource
    For Each Row In rngSource
        ' Vérifier si la valeur de la ligne est 0 et que la colonne "idArticle" n'est pas vide
        If Row.Value = 0 And tblSource.DataBodyRange(Row.Row - 1, tblSource.ListColumns("idArticle").index).Value <> "" Then
            ' Redimensionner le tableau newRowList pour ajouter la nouvelle ligne
            ReDim Preserve newRowList(size)
            newRowList(size) = Row.Row - 1
            size = size + 1
        End If
    Next

    ' Vérifier si newRowList n'est pas vide
    If Not ArrayIsEmpty(newRowList) Then
        ' Parcourir chaque nouvelle ligne dans newRowList
        For Each newRow In newRowList
            duplicate = False
            ' Vérifier les doublons dans rngSource
            For Each Row In rngSource
                If tblSource.DataBodyRange(Row.Row - 1, tblSource.ListColumns("idArticle").index).Value = tblSource.DataBodyRange(newRow, tblSource.ListColumns("idArticle").index).Value And Row.Value = 1 Then
                    ' Marquer la ligne comme doublon et ajouter un commentaire
                    tblSource.DataBodyRange(newRow, tblSource.ListColumns("Commentaire").index).Value = "Article non inséré: Doublon existant"
                    duplicate = True
                End If
            Next
            ' Si pas de doublon, ajouter la ligne à waitingInsertList
            If Not duplicate Then
                ReDim Preserve waitingInsertList(insertSize)
                waitingInsertList(insertSize) = newRow
                insertSize = insertSize + 1
            End If
        Next
    End If

    ' Vérifier si waitingInsertList n'est pas vide
    If Not ArrayIsEmpty(waitingInsertList) Then
        ' Appel de la fonction rootDataing_data pour traiter les données
        rootDataResult = rootDataing_data()
    
    
    'Parcourir chaque ligne à insérer
    For Each Row In waitingInsertList
        'Parcourir la liste des tableaux de destination
        For Each rootData In rootDataResult
            ' Définir la feuille de calcul cible et le tableau cible
            Set targetWks = Worksheets(rootData(1))
            Set targetTable = targetWks.ListObjects(rootData(0))
            Set targetRng = targetTable.ListColumns(rootData(2)(0)).DataBodyRange
            
            
            ' Ajouter une nouvelle ligne au tableau
            Set newInsertRow = targetTable.ListRows.Add
            
            sourceIndex = 0
            printing = "Insertion de la ligne " & Row & " dans la feuille " & rootData(1) & " dans le tableau " & rootData(0) & " "
            printingSource = "avec pour source: "
            printingTarget = "avec pour cible: "
            
            For Each Source In rootData(2)
                printingSource = printingSource & Source & " "
                printingTarget = printingTarget & rootData(3)(sourceIndex) & " "
            
                ' Insérer la valeur dans la nouvelle ligne
                newInsertRow.Range(targetTable.ListColumns(rootData(3)(sourceIndex)).index).Value = tblSource.DataBodyRange(Row, tblSource.ListColumns(Source).index).Value
            
                sourceIndex = sourceIndex + 1
            Next
        Next
                    ' Marquer la ligne source comme insérée
            tblSource.DataBodyRange(Row, tblSource.ListColumns("inserted").index).Value = 1
            tblSource.DataBodyRange(Row, tblSource.ListColumns("Commentaire").index).Value = "Article ajouté"
    Next
    End If
    
End Sub

