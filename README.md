Enum Built_inConstant

    inc = 1
    evenNumber
    
    E_enum = 5
    
    No
    
    xlEdgeLeft
    xlEdgeTop
    xlEdgeBottom
    xlEdgeRight
    xlInsideVertical
    
    dateOccur = 7
    dateconfirm
    
    lBlue_R = 221
    lBlue_G = 235
    lBlue_B = 247
    
    blue_R = 155
    blue_G = 194
    blue_B = 230
    
End Enum

Sub AddRow()

    Dim rNum As Long, rng_adress As String
    rNum = Cells(No, E_enum).End(xlDown).Row + inc
    rng_adress = "E" + CStr(rNum) + ":I" + CStr(rNum)
    
    Dim No_cell As range
    Set No_cell = range("E" + CStr(rNum))

    Call RuledLine(rng_adress, xlEdgeLeft)
    Call RuledLine(rng_adress, xlEdgeTop)
    Call RuledLine(rng_adress, xlEdgeBottom)
    Call RuledLine(rng_adress, xlEdgeRight)
    Call RuledLine(rng_adress, xlInsideVertical)
    
    With range(rng_adress)
    
        .HorizontalAlignment = xlGeneral
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
            
    End With
    
    No_cell = "=ROW()-6"
    Cells(rNum, dateOccur) = Date
    Cells(rNum, dateconfirm) = Date
    
    If No_cell.Value Mod evenNumber = 0 Then
    
        range(rng_adress).Interior.Color = RGB(lBlue_R, lBlue_G, lBlue_B)
        
    Else
        
        range(rng_adress).Interior.Color = RGB(blue_R, blue_G, blue_B)
    
    End If
      
End Sub

Sub RuledLine(adr As String, bic As Long)
  
    range(adr).Borders(bic).LineStyle = xlContinuous
    
End Sub
