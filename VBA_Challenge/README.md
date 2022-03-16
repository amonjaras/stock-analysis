# **Report VBA Stock Analysis Challenge**
## **Overview of the Project**

## **Results**

#### **Writing the code**

###### Step 1a: Create a *tickerIndex* variable equal to zero in order to get the correct index within the four different arrays.

`tickerIndex = 0`

###### Step 1b: Create three output arrays to hold an arbitrary number of variables of the same type.

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

###### Step 2a: Create a loop to initialize the *tickerVolumes* to zero.

```
For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
```

###### Step 2b: Create a second loop that can be executed over all raws in a spreadsheet.

`For i = 2 To RowCount`

###### Step 3a: Inside *Step 2b* we create an script to increase the volume to the current ticker.

`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`

###### Step 3b: We verify if the current row is the first row with the selected tickerIndex

```
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
```
###### Step 3c: We verify if the current row is the last row with the selected tickerIndex

```
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
End If
```
###### Step 3d: finally we increase the tickerIndex

```
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerIndex = tickerIndex + 1
End If
```

###### Step 4: All the outputs are presented to the proper columns using a loop

```
For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

    Next i
```
#### **Result comparison**


## **Summary**


This work belongs to [^1].
Unit [^2].
[^note]:
[^1]: Audrey MONJARAS :mexico: :canada:
[^2]: Unit 2 Excel VBA
