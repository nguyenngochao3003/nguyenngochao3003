Public arrAttribute()
Public i, j As Integer
Sub QuinlanDecisionTree()
    
    ' // load data to array
    Call LoadToArray
    
    ' //vecto dac trung cho cac thuoc tinh
    Call PropertyVector
    
    ' // count unit vector
    Call unitvector
    
    '// add vao decision and remove row having unit vector
    Call addingNodes
    
    
End Sub
    ' // load data to array
    Sub LoadToArray()
        arrAttribute = Sheet1.Range("A1").CurrentRegion
    End Sub
    
    ' //vecto dac trung cho cac thuoc tinh
   
    Function PropertyVector()
        Dim dicAttribute As New Dictionary
        
        '// outcome
        
        
        '//attribute list
            For j = LBound(arrAttribute, 2) To UBound(arrAttribute, 2)
                For i = LBound(arrAttribute, 1) + 1 To UBound(arrAttribute, 1)
                    If Not dicAttribute.Exists(arrAttribute(1, j) & "/" & arrAttribute(i, j)) Then
                        dicAttribute.Add arrAttribute(1, j) & "/" & arrAttribute(i, j), arrAttribute(i, j)
                    End If
                Next i
            Next j
            
            Dim keyAttribute As String
            
            '// calculate property vector
            For j = LBound(arrAttribute, 2) To UBound(arrAttribute, 2)
                For i = LBound(arrAttribute, 1) + 1 To UBound(arrAttribute, 1)
                    
                    
                    
                    
                    Next Item
                Next i
            Next j
            
            
            '// array affter add nodes
            
    End Function
    
    ' // count unit vector
    Sub unitvector()
    
    End Sub
    
    '// add vao decision and remove row having unit vector
    Sub addingNodes()

    End Sub
    
