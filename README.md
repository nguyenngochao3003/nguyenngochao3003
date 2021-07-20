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
