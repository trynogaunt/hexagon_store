Function ArrayIsEmpty(fArray()) As Boolean
    On Error Resume Next
    
    intUpper = UBound(fArray)
    ArrayIsEmpty = Err
    
End Function

Function rooting_data(fdataType As String) As Variant
    Dim rootSheet As Worksheet
    Dim rootTbl As ListObject
    Dim rootRng As Range
    Dim targetSheet As String
    Dim targetTbl As String
    Dim rootData() As Variant
    Dim rootTable() As Variant
    Dim size As Integer
    Dim targetCol As String
    Dim rootValue As Range
    size = 0
    Set rootSheet = Worksheets("routing_sheet")
    Set rootTbl = rootSheet.ListObjects("routing_table")
    Set rootRng = rootTbl.ListColumns("source").DataBodyRange

    For Each rootValue In rootRng
        If rootValue.Value = fdataType Then
            targetSheet = rootTbl.DataBodyRange(rootValue.Row - 1, rootTbl.ListColumns("targetSheet").Index).Value
            targetTbl = rootTbl.DataBodyRange(rootValue.Row - 1, rootTbl.ListColumns("targetTable").Index).Value
            targetCol = rootTbl.DataBodyRange(rootValue.Row - 1, rootTbl.ListColumns("targetColumn").Index).Value
            rootTable = Array(targetSheet, targetTbl, targetCol)
            Debug.Print "root liste: " & targetSheet & " / " & targetTbl & " / " & targetCol
            ReDim Preserve rootData(size)
            rootData(size) = rootTable
            size = size + 1
        End If
    Next
   

    rooting_data = rootData

End Function
Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function



Sub AddArticle()
    Dim tbl As ListObject
    Dim rng As Range
    Dim waitingInsert() As Variant
    Dim noInsertIndex() As Integer
    Dim noDuplicateInsert() As Variant
    Dim insertCount As Integer
    Dim duplicate As Boolean
    Dim targetList() As Variant
    Dim targetTbl As ListObject
    Dim targetWks As Worksheet
    Dim targetRng As Range
    Dim lastRow As Long
    Dim idCell As String
    Dim comCell As String
    Dim data_root() As Variant
    

    'Récupération de la ligne et de l'id de l'article à ajouter
    Set tbl = ActiveSheet.ListObjects("Insertion")
    Set rng = tbl.ListColumns("inserted").DataBodyRange
    
    targetList() = Array(Array("Feuil2", "reception", "idArticle"))
    
    idCell = "idArticle"
    comCell = "Commentaire"
    insertCell = "inserted"
    
    size = 0
    insertCount = 0
    
    For Each insertedValue In rng
        If insertedValue.Value = 0 And tbl.DataBodyRange(insertedValue.Row - 1, tbl.ListColumns(idCell).Index).Value <> "" Then 'Vérifie si déjà inséré , si non enregistre l'id et la ligne
            ReDim Preserve waitingInsert(size)
            waitingInsert(size) = Array(tbl.DataBodyRange(insertedValue.Row - 1, tbl.ListColumns(idCell).Index).Value, insertedValue.Row - 1)
            size = size + 1
        End If
    Next
    
    'Vérifie si un id correspondant est déjà inséré
    If Not ArrayIsEmpty(waitingInsert) Then
        For Each article In waitingInsert
            duplicate = False
            For Each insertedValue In rng
                If tbl.DataBodyRange(insertedValue.Row - 1, tbl.ListColumns(idCell).Index).Value = article(0) And insertedValue.Value = 1 Then
                    tbl.DataBodyRange(article(1), tbl.ListColumns(comCell).Index).Value = "Article non inséré: Doublon existant" 'Commente la valeur déjà existante
                    duplicate = True
                End If
            Next
            If duplicate = False Then
                ReDim Preserve noDuplicateInsert(insertCount)
                noDuplicateInsert(insertCount) = article
                insertCount = insertCount + 1
            End If
        Next
    End If
     If Not ArrayIsEmpty(noDuplicateInsert) Then 'Vérifie qu'il y a bien des valeurs a insérer
        For Each Insert In noDuplicateInsert 'Insere les valeurs non dupliquées dans les tableaux correspondant
            With tbl
                For Each Column In .ListColumns
                    data_root() = rooting_data(Column.Name)
                    If Not ArrayIsEmpty(data_root) Then
                        For Each Root In data_root
                                Set targetWks = Worksheets(Root(0))
                                Set targetTbl = targetWks.ListObjects(Root(1))
                                Set targetRng = targetTbl.ListColumns(Root(2)).DataBodyRange
                                lastRow = targetRng.Rows(targetRng.Rows.count).Row
                        Next
                    End If
                Next
                .DataBodyRange(Insert(1), .ListColumns(insertCell).Index).Value = 1 'Mets à jour l'état de l'article
                .DataBodyRange(Insert(1), .ListColumns(comCell).Index).Value = "Article inséré" 'Ajout le commentaire d'insertion
                Debug.Print "Article " & Insert(0) & " inséré"
            End With
        Next
    End If
End Sub
