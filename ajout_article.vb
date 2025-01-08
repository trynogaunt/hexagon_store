Function ArrayIsEmpty(fArray()) As Boolean
    On Error Resume Next
    
    intUpper = UBound(fArray)
    ArrayIsEmpty = Err
    
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
        For Each target In targetList
            Debug.Print "Article " & Insert(0) & " en cours d'insertion dans le tableau " & target(1)
            Set targetWks = Worksheets(target(0))
            Set targetTbl = targetWks.ListObjects(target(1))
            Set targetRng = targetTbl.ListColumns(target(2)).DataBodyRange
            lastRow = targetRng.Rows(targetRng.Rows.count).Row
        
            targetTbl.DataBodyRange(lastRow, targetTbl.ListColumns(target(2)).Index).Value = Insert(0)
        
        Next
        tbl.DataBodyRange(Insert(1), tbl.ListColumns(insertCell).Index).Value = 1 'Mets à jour l'état de l'article
        tbl.DataBodyRange(Insert(1), tbl.ListColumns(comCell).Index).Value = "Article inséré" 'Ajout le commentaire d'insertion
        Debug.Print "Article " & Insert(0) & " inséré"
    Next
End If
End Sub
