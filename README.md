# VBA_Credit_Assignment
Sub Credit():

'Assigning all of my values as needed Last row is long since it is a long value in compairson to double which has percentages
Dim Credit_Name As String
Dim Volume_Total As Double
Dim Firstopenvalue As Double
Dim Last_Row As Long
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double

'Assigning my values to their corresponding colum
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'since we are counting all of the volumes for a given ticker we need to start at zero
Volume_Total = 0


Dim Summary_Row As Long
Summary_Row = 2
Firstopenvalue = 2

Last_Row = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

'starting the for loop at 2 since the first is the title
'the for loop will cyclethrough the entire column becasue fo the i +1
For I = 2 To Last_Row

'first if statment is returnign thecredit names and their total values
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    Credit_Name = Cells(I, 1).Value
    Opening_Price = Cells(Firstopenvalue, 3).Value
    Firstopenvalue = I + 1
    Closing_Price = Cells(I, 6).Value
    Volume_Total = Volume_Total + Cells(I, 7).Value
    
'this is what will be returned and in which column
    Cells(Summary_Row, 9).Value = Credit_Name
    Cells(Summary_Row, 12).Value = Volume_Total
    
'this will return
    Yearly_Change = Closing_Price - Opening_Price
    Cells(Summary_Row, 10).Value = Yearly_Change

'second if statement is returning the percentage of the yearly change
                If Opening_Price > 0 Then
                Range("K" & Summary_Row).Value = (Yearly_Change / Opening_Price)
                Else
                Range("K" & Summary_Row).Value = 0
                End If
              
 'this if statment is used to assign a color to a value that is greater o (green) else its red
                If Yearly_Change >= 0 Then
                Range("j" & Summary_Row).Interior.ColorIndex = 4
                Else
                Range("j" & Summary_Row).Interior.ColorIndex = 3
                End If

'I couldent figure out how to get yellow in an elseif statement so I just made one for itsself that if it equalled 0
                If Yearly_Change = 0 Then
                Range("j" & Summary_Row).Interior.ColorIndex = 6
                
                End If
            
  ' this is to make sure the value is returned as a percentage
                Range("K" & Summary_Row).NumberFormat = "0.00%"

    Summary_Row = Summary_Row + 1
    Volume_Total = 0
                
Else

'this will act as the counter for the volume = 0 above
    Volume_Total = Volume_Total + Cells(I, 7).Value

End If

'next i is before the max and min since we dont need to loop for those
Next I

'these functions are returning the min and max of the percentages as well as the greatest volume #
Range("Q2") = "%" & CStr(WorksheetFunction.max(Range("K2:K" & Last_Row)) * 100)
Range("Q3") = "%" & CStr(WorksheetFunction.Min(Range("K2:K" & Last_Row)) * 100)
Range("Q4") = CStr(WorksheetFunction.max(Range("L2:L" & Last_Row)))

'these functions are matchign the ticker name to the value
Greatest_Value = WorksheetFunction.Match(WorksheetFunction.max(Range("K2:K" & Last_Row)), Range("K2:K" & Last_Row), 0)
Range("P2") = Cells(Greatest_Value + 1, 9)
Lowest_Value = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & Last_Row)), Range("K2:K" & Last_Row), 0)
Range("P3") = Cells(Lowest_Value + 1, 9)
Greatest_Volume = WorksheetFunction.Match(WorksheetFunction.max(Range("L2:L" & Last_Row)), Range("L2:L" & Last_Row), 0)
Range("P4") = Cells(Greatest_Volume + 1, 9)

End Sub

'I want to preface that I had a lot fo help from the TA, Tutor and Chat GPT in order to explain the numerous errors of my code, this had to be one of the hardest assignment ive ever done and hope it meets the requiments!
