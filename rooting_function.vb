Function ArrayIsEmpty(fArray()) As Boolean
    On Error Resume Next
    
    intUpper = UBound(fArray)
    ArrayIsEmpty = Err
    
End Function

Function rooting_data() As Variant
    
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
    rootheaderRow = tblrootData.HeaderRowRange.Row
    
    rootDataSourceIndex = 0
    For Each rootDataSource In rngrootData
        If rootDataSource = "idArticle" Then
            ReDim Preserve sheetList(rootDataSourceIndex)
            ReDim Preserve tableList(rootDataSourceIndex)
            sheetList(rootDataSourceIndex) = tblrootData.DataBodyRange(rootDataSource.Row - rootheaderRow, tblrootData.ListColumns("targetSheet").index).Value
            tableList(rootDataSourceIndex) = tblrootData.DataBodyRange(rootDataSource.Row - rootheaderRow, tblrootData.ListColumns("targetTable").index).Value
            rootDataSourceIndex = rootDataSourceIndex + 1
        End If
    Next
    
    insertRowIndex = 0
    dataIndex = 0
    If ArrayIsEmpty(tableList) Then
        MsgBox ("Table de routage vide")
        Exit Function
    Else
        For Each Table In tableList
            index = 0
            For Each rootDataSource In rngrootData
                If tblrootData.DataBodyRange(rootDataSource.Row - rootheaderRow, tblrootData.ListColumns("targetTable").index).Value = Table Then
                ReDim Preserve sourceList(index)
                ReDim Preserve targetList(index)
                sourceList(index) = rootDataSource
                targetList(index) = tblrootData.DataBodyRange(rootDataSource.Row - rootheaderRow, tblrootData.ListColumns("targetColumn").index).Value
                index = index + 1
                End If
            Next
            ReDim Preserve rootDataList(dataIndex)
            rootDataList(dataIndex) = Array(Table, sheetList(tableIndex), sourceList, targetList)
            tableIndex = tableIndex + 1
            dataIndex = dataIndex + 1
        Next
   

   rooting_data = rootDataList

End If
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
    Dim headerRow As Long
    Dim insertedValue As String
    
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
    Dim printingsize As Integer
    
    Set tblSource = ActiveSheet.ListObjects("MLM_IMPORTATEUR_PRODUITS")
    Set rngSource = tblSource.ListColumns("inserted").DataBodyRange
    headerRow = tblSource.HeaderRowRange.Row
    Set printingsize = 0

   
    
    ' Parcourir chaque ligne dans rngSource
    For Each Row In rngSource
        ' Vérifier si la valeur de la ligne est 0 et que la colonne "référence article MLM" n'est pas vide
        If Row.Value = 0 Or Row.Value = "" And tblSource.DataBodyRange(Row.Row - headerRow, tblSource.ListColumns("référence article MLM").index).Value <> "" Then
            ' Redimensionner le tableau newRowList pour ajouter la nouvelle ligne
            ReDim Preserve newRowList(size)
            newRowList(size) = Row.Row - headerRow
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
                If tblSource.DataBodyRange(Row.Row - headerRow, tblSource.ListColumns("référence article MLM").index).Value = tblSource.DataBodyRange(newRow, tblSource.ListColumns("référence article MLM").index).Value And Row.Value = 1 Then
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
        rootDataResult = rooting_data()
    
    
    'Parcourir chaque ligne à insérer
    printing = "Article ajouté:" & vbCrLf
    For Each Row In waitingInsertList
        'Parcourir la liste des tableaux de destination
        For Each rootData In rootDataResult
            ' Définir la feuille de calcul cible et le tableau cible
            Set targetWks = Worksheets(rootData(1))
            Set targetTable = targetWks.ListObjects(rootData(0))
            Set targetRng = targetTable.ListColumns(rootData(3)(0)).DataBodyRange
            
            
            ' Ajouter une nouvelle ligne au tableau
            Set newInsertRow = targetTable.ListRows.Add
            
            sourceIndex = 0
            
            
            For Each Source In rootData(2)
                
            
                ' Insérer la valeur dans la nouvelle ligne
                newInsertRow.Range(targetTable.ListColumns(rootData(3)(sourceIndex)).index).Value = tblSource.DataBodyRange(Row, tblSource.ListColumns(Source).index).Value
                insertedValue = tblSource.DataBodyRange(Row, tblSource.ListColumns(Source).index).Value
                sourceIndex = sourceIndex + 1
            Next
        Next
                    ' Marquer la ligne source comme insérée
            tblSource.DataBodyRange(Row, tblSource.ListColumns("inserted").index).Value = 1
            tblSource.DataBodyRange(Row, tblSource.ListColumns("Commentaire").index).Value = "Article ajouté: " + Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss")
            
            If printingsize <= 9 Then
                printing = printing & insertedValue & ","
                printingsize = printingsize + 1
            Else
                printing = printing & insertedValue & vbCrLf
                printingsize = 0
            End If
    Next
    MsgBox (printing)
    Else
    MsgBox("Aucun article à ajouter")
    End If
    
End Sub




