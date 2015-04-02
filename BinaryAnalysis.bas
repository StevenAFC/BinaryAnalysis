Option Explicit

Function Main(target As Worksheet, filePath As String)
    Dim intFileNum As Integer, bytTemp As Byte
    Dim grid(0 To 255, 0 To 255) As Double
    intFileNum = FreeFile
 
    Open filePath For Binary Access Read As intFileNum
    
    Dim x As Integer, y As Integer
    
    x = -1
    
    Do While Not EOF(intFileNum)
        Get intFileNum, , bytTemp
         
        If (x = -1) Then
            x = bytTemp
        Else
            y = bytTemp
            
            If x <> 0 And y <> 0 Then
                grid(x, y) = grid(x, y) + 1
            End If
            
            x = -1
            y = -1
        End If
          
    Loop
    Close intFileNum
    
    Dim min As Double
    Dim max As Double
    
    min = grid(0, 0)
    max = grid(0, 0)
        
    For x = 0 To 255
        
        For y = 0 To 255
            
            If grid(x, y) < min Then
                min = grid(x, y)
            End If
            
            If grid(x, y) > max Then
                max = grid(x, y)
            End If
            
        Next y
    
    Next x
    
    Dim colorValue As Integer

    For x = 0 To 255
        
        For y = 0 To 255
            
            If grid(x, y) > 0 Then

                colorValue = (255 * 0.75) - 1 + ((grid(x, y) - min) / (max - min) * 255) * 0.75
                target.Cells(x + 1, y + 1).Interior.Color = RGB(colorValue, 0, 0)
                
            End If

        Next y
    
    Next x
 
End Function
