{\rtf1\ansi\ansicpg1252\cocoartf2761
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub MasterScript()\
    Dim ws As Worksheet\
    Dim lastRow As Long\
    Dim i As Long\
    Dim currentTicker As String\
    Dim startRow As Long\
    Dim endRow As Long\
    Dim openPrice As Double\
    Dim closePrice As Double\
    Dim quarterlyChange As Double\
    Dim percentChange As Double\
    Dim totalVolume As Double\
    Dim maxIncrease As Double\
    Dim maxDecrease As Double\
    Dim maxVolume As Double\
    Dim maxIncreaseTicker As String\
    Dim maxDecreaseTicker As String\
    Dim maxVolumeTicker As String\
    Dim startDate As Date\
    Dim endDate As Date\
    Dim dateValue As String\
    Dim yearPart As Integer\
    Dim monthPart As Integer\
    Dim dayPart As Integer\
    Dim outputRow As Long\
\
    ' Initialize variables for tracking the maximum values\
    maxIncrease = -1\
    maxDecrease = 1\
    maxVolume = 0\
\
    ' Loop through each sheet\
    For Each ws In ThisWorkbook.Worksheets\
        If ws.Name <> "Quarterly Summary" Then\
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\
            \
            ' Add headers to the current sheet\
            ws.Cells(1, 9).Value = "Ticker"\
            ws.Cells(1, 10).Value = "Quarterly Change"\
            ws.Cells(1, 11).Value = "Percent Change"\
            ws.Cells(1, 12).Value = "Total Stock Volume"\
            \
            ' Convert text to date format\
            For i = 2 To lastRow\
                dateValue = ws.Cells(i, 2).Value\
                If IsNumeric(dateValue) And Len(dateValue) = 8 Then\
                    yearPart = CInt(Left(dateValue, 4))\
                    monthPart = CInt(Mid(dateValue, 5, 2))\
                    dayPart = CInt(Right(dateValue, 2))\
                    ws.Cells(i, 2).Value = DateSerial(yearPart, monthPart, dayPart)\
                End If\
            Next i\
            \
            ' Initialize output row\
            outputRow = 2\
            \
            ' Loop through the data by ticker and quarter\
            i = 2\
            Do While i <= lastRow\
                currentTicker = ws.Cells(i, 1).Value ' <ticker> column\
                \
                ' Ensure the date value is a valid date by converting the text to a date\
                On Error Resume Next\
                startDate = ws.Cells(i, 2).Value ' <date> column\
                On Error GoTo 0\
                \
                If startDate = 0 Then\
                    Debug.Print "Invalid date at row " & i & ": " & dateValue\
                    i = i + 1\
                    GoTo ContinueLoop\
                End If\
                \
                startRow = i\
                \
                ' Find the end of the current ticker and quarter\
                Do While ws.Cells(i, 1).Value = currentTicker And Month(ws.Cells(i, 2).Value) \\ 4 = Month(startDate) \\ 4 And i <= lastRow\
                    i = i + 1\
                Loop\
                \
                endRow = i - 1\
                \
                ' Ensure the date value is a valid date by converting the text to a date\
                On Error Resume Next\
                endDate = ws.Cells(endRow, 2).Value ' <date> column\
                On Error GoTo 0\
                \
                If endDate = 0 Then\
                    Debug.Print "Invalid date at row " & endRow & ": " & dateValue\
                    i = i + 1\
                    GoTo ContinueLoop\
                End If\
                \
                ' Calculate the quarterly change, percent change, and total stock volume\
                openPrice = ws.Cells(startRow, 3).Value ' <open> column\
                closePrice = ws.Cells(endRow, 6).Value ' <close> column\
                quarterlyChange = closePrice - openPrice\
                If openPrice <> 0 Then\
                    percentChange = (closePrice - openPrice) / openPrice\
                Else\
                    percentChange = 0\
                End If\
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7))) ' <vol> column\
                \
                ' Output the results to the current sheet\
                ws.Cells(outputRow, 9).Value = currentTicker\
                ws.Cells(outputRow, 10).Value = quarterlyChange\
                ws.Cells(outputRow, 11).Value = percentChange\
                ws.Cells(outputRow, 12).Value = totalVolume\
                outputRow = outputRow + 1\
\
                ' Track the maximum values\
                If percentChange > maxIncrease Then\
                    maxIncrease = percentChange\
                    maxIncreaseTicker = currentTicker\
                End If\
                If percentChange < maxDecrease Then\
                    maxDecrease = percentChange\
                    maxDecreaseTicker = currentTicker\
                End If\
                If totalVolume > maxVolume Then\
                    maxVolume = totalVolume\
                    maxVolumeTicker = currentTicker\
                End If\
                \
ContinueLoop:\
            Loop\
            \
            ' Autofit columns in the current sheet\
            ws.Columns("I:O").AutoFit\
            \
            ' Format Percent Change column as percentage with two decimal places\
            ws.Columns("K").NumberFormat = "0.00%"\
            \
            ' Apply conditional formatting for positive changes (green)\
            With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")\
                .Interior.Color = RGB(0, 255, 0) ' Green color\
            End With\
            \
            ' Apply conditional formatting for negative changes (red)\
            With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")\
                .Interior.Color = RGB(255, 0, 0) ' Red color\
            End With\
            \
            ' Insert two blank columns between Total Stock Volume and the summary table\
            ws.Columns("M:N").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove\
            \
            ' Output the summary information to the current sheet\
            ws.Cells(1, 15).Value = "Ticker"\
            ws.Cells(1, 16).Value = "Value"\
            ws.Cells(2, 14).Value = "Greatest % Increase"\
            ws.Cells(2, 15).Value = maxIncreaseTicker\
            ws.Cells(2, 16).Value = Format(maxIncrease, "0.00%")\
            ws.Cells(3, 14).Value = "Greatest % Decrease"\
            ws.Cells(3, 15).Value = maxDecreaseTicker\
            ws.Cells(3, 16).Value = Format(maxDecrease, "0.00%")\
            ws.Cells(4, 14).Value = "Greatest Total Volume"\
            ws.Cells(4, 15).Value = maxVolumeTicker\
            ws.Cells(4, 16).Value = maxVolume\
        End If\
    Next ws\
End Sub\
\
\
}