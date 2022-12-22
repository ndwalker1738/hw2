{\rtf1\ansi\ansicpg1252\cocoartf2706
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub multiple_year_stock()\
\
\
    'Set variables\
    Dim Ws_Count As Integer\
    Dim ticker As String\
    Dim summary_table As Integer\
    Dim yearly_change As Double\
    Dim open_row As Double\
    Dim percent_change As Double\
    Dim stock_value As Double\
    \
\
    \
    'Set the worksheet count variable to the number of worksheets\
    \
    Ws_Count = ActiveWorkbook.Worksheets.Count\
    \
    'Loops through each worksheet and row\
    \
    For w = 1 To Ws_Count\
    \
        Worksheets(w).Activate\
    'Summary Table Column titles\
            Range("M1").Value = "ticker"\
            Range("N1").Value = "yearly change"\
            Range("O1").Value = "percent change"\
            Range("P1").Value = "total volume stock"\
            \
             LastRow = Cells(Rows.Count, 1).End(xlUp).Row\
            \
            summary_table = 2\
            \
            open_row = 2\
            \
            percent_change = 0\
            \
            stock_value = 0\
            \
       For i = 2 To LastRow\
        \
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then\
            \
            open_row = i\
            \
            End If\
            \
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
            \
            stock_value = stock_value + Cells(i, 7).Value\
            \
                    Range("M" & summary_table).Value = Cells(i, 1).Value\
                    \
                    yearly_change = Cells(i, 6).Value - Cells(open_row, 3).Value\
                    \
                    Range("N" & summary_table).Value = yearly_change\
                    \
                    Range("O" & summary_table).Value = (yearly_change / Cells(open_row, 3).Value)\
                \
                    Range("P" & summary_table).Value = (stock_value + Cells(i, 7).Value)\
                    \
                    summary_table = summary_table + 1\
                    \
                    'open_row = (i + 1)\
                \
                'MsgBox (Cells(i, 1).Value)\
                \
              Else\
                \
             stock_value = stock_value + Cells(i, 7).Value\
                \
            End If\
            \
        Next i\
    \
    'Go to the last row of Column O because the percent change is the column\
    \
    Endrow = Cells(Rows.Count, "O").End(xlUp).Row\
    \
    'Set variables for second table which are all percentages\
    \
    \
    greatest_inc = 0\
    \
    greatest_dec = 0\
    \
    greatest_tv = 0\
    \
    \
    greatest_inc = WorksheetFunction.Max(Range("O2:O" & Endrow))\
    \
    Range("T2").Value = greatest_inc\
    \
    sideticker = WorksheetFunction.Match(greatest_inc, Range("O2:O" & Endrow), 0)\
    \
    Range("S2").Value = Cells(sideticker + 1, "M")\
    \
    \
    greatest_dec = WorksheetFunction.Min(Range("O2:O" & Endrow))\
    \
    Range("T3").Value = greatest_dec\
    \
    sideticker = WorksheetFunction.Match(greatest_dec, Range("O2:O" & Endrow), 0)\
    \
    Range("S3").Value = Cells(sideticker + 1, "M")\
    \
    \
    greatest_tv = WorksheetFunction.Max(Range("P2:P" & Endrow))\
    \
    Range("T4").Value = greatest_tv\
    \
    sideticker = WorksheetFunction.Match(greatest_tv, Range("P2:P" & Endrow), 0)\
    \
    Range("S4").Value = Cells(sideticker + 1, "M")\
    \
    \
    ' assign names to summary table\
    \
    Range("R2").Value = " Greatest % Increase"\
    Range("R3").Value = " Greatest % Decrease"\
    Range("R4").Value = "Greatest Total Volume"\
    Range("S1").Value = "Ticker"\
    Range("T1").Value = "Value"\
    \
    'format percentages\
    \
    Range("O2:O" & Endrow).NumberFormat = "0.00%"\
 \
    Range("T2:T3").NumberFormat = "0.00%"\
     \
    \
    'format colors\
    \
    RowStart = 2\
    RowEnd = LastRow\
    \
    For i = RowStart To RowEnd\
    \
        If Cells(i, 14) > 0 Then\
        \
        Cells(i, 14).Interior.Color = vbGreen\
        \
        Else\
        \
        Cells(i, 14).Interior.Color = vbRed\
        \
        End If\
        \
        Next i\
        \
       \
        \
    \
    \
    \
    \
    \
    Next w\
    \
    \
\
\
\
\
\
End Sub\
\
}