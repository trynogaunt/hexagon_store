Function ArrayIsEmpty(fArray()) As Boolean
    On Error Resume Next
    
    intUpper = UBound(fArray)
    ArrayIsEmpty = Err
    
End Function

Function rooting_data(fwaitingInsertRow()) As Variant
    
    Dim tblRoot As ListObject
    Dim rngRoot As Range
    Dim tableList() As Variant
    Dim sheetList() As Variant
    Dim sourceList() As Variant
    Dim targetList() As Variant
 
    Dim rootList() As Variant

  
    Dim insertRowIndex As Integer
    Dim rootSourceIndex As Integer
    Dim index As Integer
    Dim dataIndex As Integer
    
    Set tblRoot = Worksheets("routing_sheet").ListObjects("routing_table")
    Set rngRoot = tblRoot.ListColumns("source").DataBodyRange
    
    rootSourceIndex = 0
    For Each rootSource In rngRoot
        'Debug.Print rootSource & " pour la ligne " & insertRow
        If rootSource = "idArticle" Then
            ReDim Preserve sheetList(rootSourceIndex)
            ReDim Preserve tableList(rootSourceIndex)
            sheetList(rootSourceIndex) = tblRoot.DataBodyRange(rootSource.Row - 1, tblRoot.ListColumns("targetSheet").index).Value
            tableList(rootSourceIndex) = tblRoot.DataBodyRange(rootSource.Row - 1, tblRoot.ListColumns("targetTable").index).Value
            rootSourceIndex = rootSourceIndex + 1
        End If
    Next
    

    
   
    
    insertRowIndex = 0
    dataIndex = 0
    For Each insertRow In fwaitingInsertRow
        tableIndex = 0
        For Each Table In tableList
        Debug.Print dataIndex
            index = 0
            For Each rootSource In rngRoot
                If tblRoot.DataBodyRange(rootSource.Row - 1, tblRoot.ListColumns("targetTable").index).Value = Table Then
                ReDim Preserve sourceList(index)
                ReDim Preserve targetList(index)
                sourceList(index) = rootSource
                targetList(index) = tblRoot.DataBodyRange(rootSource.Row - 1, tblRoot.ListColumns("targetColumn").index).Value
                index = index + 1
                End If
            Next
            ReDim Preserve rootList(dataIndex)
            rootList(dataIndex) = Array(insertRow, Table, sheetList(tableIndex), sourceList, targetList)
            tableIndex = tableIndex + 1
            dataIndex = dataIndex + 1
        Next
        
        insertRowIndex = insertRowIndex + 1
    Next
   
   rooting_data = rootList

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
    
    Dim newRowList() As Variant
    Dim size As Integer
    size = 0
    
    Dim waitingInsertList() As Variant
    Dim insertSize As Integer
    insertSize = 0
    
    
    Dim duplicate As Boolean
    Dim rootResult() As Variant
    Dim printing As String
    Dim printingSource As String
    Dim printingTarget As String
    
    Set tblSource = ActiveSheet.ListObjects("Insertion")
    Set rngSource = tblSource.ListColumns("inserted").DataBodyRange
    
    
    For Each Row In rngSource
        If Row.Value = 0 And tblSource.DataBodyRange(Row.Row - 1, tblSource.ListColumns("idArticle").index).Value <> "" Then
            ReDim Preserve newRowList(size)
            newRowList(size) = Row.Row - 1
            size = size + 1
        End If
    Next
    
    If Not ArrayIsEmpty(newRowList) Then
        For Each newRow In newRowList
            duplicate = False
            For Each Row In rngSource
                If tblSource.DataBodyRange(Row.Row - 1, tblSource.ListColumns("idArticle").index).Value = tblSource.DataBodyRange(newRow, tblSource.ListColumns("idArticle").index) And Row.Value = 1 Then
                    tblSource.DataBodyRange(newRow, tblSource.ListColumns("Commentaire").index).Value = "Article non inséré: Doublon existant"
                    duplicate = True
                End If
            Next
            If Not duplicate Then
                ReDim Preserve waitingInsertList(insertSize)
                waitingInsertList(insertSize) = newRow
                insertSize = insertSize + 1
            End If
        Next
    End If
    
    If Not ArrayIsEmpty(waitingInsertList) Then
      rootResult = rooting_data(waitingInsertList)
    End If
    
    For Each dataRooted In rootResult
        Set targetWks = Worksheets(dataRooted(2))
        Set targetTable = targetWks.ListObjects(dataRooted(1))
        Set targetRng = targetTable.ListColumns("idArticle").DataBodyRange
        lastRow = targetRng.Rows(targetRng.Rows.Count).Row
        sourceIndex = 0
        printing = "Insertion de la ligne " & dataRooted(0) & " dans la feuilles " & dataRooted(2) & " dans le tableau " & dataRooted(1) & " "
        printingSource = "avec pour source: "
        printingTarget = "avec pour cible: "
        For Each Source In dataRooted(3)
            printingSource = printingSource & Source & " "
            printingTarget = printingTarget & dataRooted(4)(sourceIndex) & " "
            targetTable.DataBodyRange(lastRow, targetTable.ListColumns(dataRooted(4)(sourceIndex)).index).Value = tblSource.DataBodyRange(dataRooted(0), tblSource.ListColumns(Source).index).Value
            sourceIndex = sourceIndex + 1
        Next
        printing = printing & printingSource & printingTarget
        Debug.Print targetTable.Name & " " & lastRow
        tblSource.DataBodyRange(dataRooted(0), tblSource.ListColumns("inserted").index).Value = 1
        tblSource.DataBodyRange(dataRooted(0), tblSource.ListColumns("Commentaire").index).Value = "Article ajouté"
    Next
    
    
    
End Sub

