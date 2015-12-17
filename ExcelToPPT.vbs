Sub ExceltoPPT()

'Open the PPT File
Dim PTT As PowerPoint.Application
Set PTT = New PowerPoint.Application
PTT.Visible = True
PTT.Presentations.Open Filename:="\\Path\To\The\PPT_Template.pptx"


' Define Variables
' -------------------------------------------------
' i is ptt slide number
Dim i As Integer
' j used for cell reference to pull sheet names
Dim j As Integer
' String to store sheetnames to loop through
Dim Sheetname As String
' Use k and l to select columns for stoplights
Dim k As Integer
Dim l As Integer
' Use irow and icol for table sizing
Dim icol As Integer
Dim irow As Integer
' -------------------------------------------------
' Import the data into the slides and perform formatting adjustments

' ---------------------------------------
' Net Promoter Score by Group
' ---------------------------------------

' Copy Paste Graph

    Sheets("NPS Graph").Select
    If ActiveChart Is Nothing Then
        ActiveSheet.ChartObjects("NPSGraph").Activate
    End If
    ActiveChart.ChartArea.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))

    PTT.ActivePresentation.Slides(1).Select
    PTT.ActivePresentation.Slides(1).Shapes.Paste

    PTT.ActivePresentation.Slides(1).Shapes("NPSGraph").Name = "CopyDeleteChart"
    PTT.ActivePresentation.Slides(1).Shapes("CopyDeleteChart").Select
    
    'Size the chart
    With PTT.ActiveWindow.Selection.ShapeRange
    .Height = 318.95
    .Width = 472.32
    End With

    'Copy and past as picture
    PTT.ActiveWindow.Selection.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))
    PTT.ActivePresentation.Slides(1).Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile

    'Position picture
    With PTT.ActivePresentation.Slides(1).Shapes("Picture 2")
    .Left = 55.44
    .Top = 156.5
    .Name = "NPSPicture"
    End With
    
    'Delete chart
    PTT.ActivePresentation.Slides(1).Shapes("CopyDeleteChart").Delete

'Copy Paste Chart

    Sheets("NPSTables").Select
    Range("Y42:AD56").Select
    Selection.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))
    
    PTT.ActiveWindow.View.PasteSpecial DataType:=ppPasteDefault

    'Set column widths
    With PTT.ActiveWindow.Selection.ShapeRange
    .Table.Columns(1).Width = 36
    .Table.Columns(2).Width = 453.6
    .Table.Columns(3).Width = 43.2
    .Table.Columns(4).Width = 50.4
    .Table.Columns(5).Width = 36
    .Table.Columns(6).Width = 36
    End With
    
    'Set row heights
    With PTT.ActiveWindow.Selection.ShapeRange
    .Table.Rows(1).Height = 29
    For irow = 2 To .Table.Rows.Count
    .Table.Rows(irow).Height = 21.23
    Next
    End With

    'Position Table
    With PTT.ActiveWindow.Selection.ShapeRange
    .Left = 30.24
    .Top = 138.24
    End With

    'Send the table to the back
    PTT.ActiveWindow.Selection.ShapeRange.ZOrder msoSendToBack
    

' ---------------------------------------
' Likelihood to Recommend
' ---------------------------------------

'Copy Paste and Resize Chart

    Sheets("Likelihood to Recommend").Select
    If ActiveChart Is Nothing Then
        ActiveSheet.ChartObjects("LTRGraph").Activate
    End If
    ActiveChart.ChartArea.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))

    PTT.ActivePresentation.Slides(2).Select
    PTT.ActivePresentation.Slides(2).Shapes.Paste

    PTT.ActivePresentation.Slides(2).Shapes("LTRGraph").Name = "CopyDeleteChart"
    PTT.ActivePresentation.Slides(2).Shapes("CopyDeleteChart").Select
    
    'Size the chart
    With PTT.ActiveWindow.Selection.ShapeRange
    .Height = 378
    .Width = 545.75
    End With
   
    'Copy and past as picture
    PTT.ActivePresentation.Slides(2).Shapes("CopyDeleteChart").Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))
    PTT.ActivePresentation.Slides(2).Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile

    'Position picture
    With PTT.ActivePresentation.Slides(2).Shapes("Picture 2")
    .Left = 48.25
    .Top = 108
    .Name = "LTRPicture"
    End With
    
    'Delete chart
    PTT.ActivePresentation.Slides(2).Shapes("CopyDeleteChart").Delete

'Copy and Paste Stoplights

    Sheets("NPSTables").Select
    Range("AL68:AM82").Select
    Selection.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))

    PTT.ActivePresentation.Slides(2).Select
    PTT.ActivePresentation.Slides(2).Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile

    With PTT.ActivePresentation.Slides(2).Shapes("Picture 3")
        .Height = 333.36
        .Left = 586.08
        .Top = 119.52
        .Name = "StoplightsNPS"
    End With


' ---------------------------------------
' Demographics %
' ---------------------------------------

    'Copy Paste Range
    Sheets("NPSTables").Select
    Range("Y2:AH16").Select
    Selection.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))

    PTT.ActivePresentation.Slides(3).Select
    PTT.ActiveWindow.View.PasteSpecial DataType:=ppPasteDefault

    'Make sure font is correct
    With PTT.ActiveWindow.Selection.ShapeRange.Table
        For irow = 1 To .Rows.Count
            For icol = 1 To .Columns.Count
                With .Cell(irow, icol).Shape.TextFrame.TextRange.Font
                .Name = "Arial"
                .Size = "11"
                End With
            Next icol
        Next irow
        For icol = 1 To .Columns.Count
            With .Cell(1, icol).Shape.TextFrame.TextRange.Font
            .Bold = True
            End With
            With .Cell(15, icol).Shape.TextFrame.TextRange.Font
            .Bold = True
            End With
            Next icol
    End With

    'Set column widths
    With PTT.ActiveWindow.Selection.ShapeRange
    .Table.Columns(1).Width = 51.84
    .Table.Columns(2).Width = 95.04
    For icol = 3 To .Table.Columns.Count
    .Table.Columns(icol).Width = 62.64
    Next
    End With
    
    'Set row heights
    With PTT.ActiveWindow.Selection.ShapeRange
    .Table.Rows(1).Height = 33.12
    For irow = 2 To .Table.Rows.Count
    .Table.Rows(irow).Height = 18
    Next
    End With

    'Position Table
    With PTT.ActiveWindow.Selection.ShapeRange
    .Left = 36
    .Top = 146.88
    End With


' ---------------------------------------
' NPS by Demographics
' ---------------------------------------

    'Copy Paste Range
    Sheets("NPSTables").Select
    Range("Y22:AG35").Select
    Selection.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:02"))

    PTT.ActivePresentation.Slides(4).Select
    PTT.ActiveWindow.View.PasteSpecial DataType:=ppPasteDefault

    'Make sure font is correct
    With PTT.ActiveWindow.Selection.ShapeRange.Table
        For irow = 1 To .Rows.Count
            For icol = 1 To .Columns.Count
                With .Cell(irow, icol).Shape.TextFrame.TextRange.Font
                .Name = "Arial"
                .Size = "11"
                End With
            Next icol
        Next irow
        For icol = 1 To .Columns.Count
            With .Cell(1, icol).Shape.TextFrame.TextRange.Font
            .Bold = True
            End With
            Next icol
    End With

    'Set column widths
    With PTT.ActiveWindow.Selection.ShapeRange
    .Table.Columns(1).Width = 51.84
    .Table.Columns(2).Width = 95.04
    For icol = 3 To .Table.Columns.Count
    .Table.Columns(icol).Width = 62.64
    Next
    End With
    
    'Set row heights
    With PTT.ActiveWindow.Selection.ShapeRange
    For irow = 1 To .Table.Rows.Count
    .Table.Rows(irow).Height = 18
    Next
    End With

    'Position Table
    With PTT.ActiveWindow.Selection.ShapeRange
    .Left = 66.24
    .Top = 162
    End With


' ---------------------------------------
' Loop Through Graphs for Slides 5 to 19
' ---------------------------------------

j = 106 'The chart names are stored in column B of a worksheet starting at Row 106

For i = 5 To 19

'Skip Slide 15
If i = 15 Then
    GoTo Skip15
End If

'Copy Paste and Resize Left Chart

    Sheetname = Worksheets("NPStables").Cells(j, 2).Value
    Sheets(Sheetname).Select
    If ActiveChart Is Nothing Then
        ActiveSheet.ChartObjects("Chart1").Activate
    End If
    ActiveChart.ChartArea.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))

    PTT.ActivePresentation.Slides(i).Select
    PTT.ActivePresentation.Slides(i).Shapes.Paste

    PTT.ActivePresentation.Slides(i).Shapes("Chart1").Name = "CopyDeleteChart"
    
    'Size the chart
    With PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart")
    .Height = 308.16
    .Width = 270
    End With
   
    'Copy and past as picture
    PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart").Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))
    PTT.ActivePresentation.Slides(i).Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile

    'Position picture
    With PTT.ActivePresentation.Slides(i).Shapes("Picture 2")
    .Left = 39.6
    .Top = 168.48
    .Name = "Chart1Picture"
    End With
    
    'Delete chart
    PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart").Delete
    
        j = j + 1

'Copy Paste and Resize Right Chart
    
    Sheetname = Worksheets("NPStables").Cells(j, 2).Value
    Sheets(Sheetname).Select
    If ActiveChart Is Nothing Then
        ActiveSheet.ChartObjects("Chart1").Activate
    End If
    ActiveChart.ChartArea.Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))

    PTT.ActivePresentation.Slides(i).Select
    PTT.ActivePresentation.Slides(i).Shapes.Paste

    PTT.ActivePresentation.Slides(i).Shapes("Chart1").Name = "CopyDeleteChart"
    
    'Size the chart
    With PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart")
    .Height = 308.16
    .Width = 270
    End With
   
    'Copy and past as picture
    PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart").Copy
    'Wait to load in clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))
    PTT.ActivePresentation.Slides(i).Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile

    'Position picture
    With PTT.ActivePresentation.Slides(i).Shapes("Picture 3")
    .Left = 372.24
    .Top = 168.48
    .Name = "Chart2Picture"
    End With
    
    'Delete chart
    PTT.ActivePresentation.Slides(i).Shapes("CopyDeleteChart").Delete
    
        j = j + 1
        
Skip15:
Next i


' -------------------------------------------
' Loop Through Stoplights for Slides 5 to 19
' -------------------------------------------

k = 4
l = 5
' As series of two column tables with a conditionally formatted "stoplight" begins at
' D21 on a worksheet and extend to the right


For i = 5 To 19

'Skip Slide 15
If i = 15 Then
    GoTo Skip15Again
End If

'Copy and Paste Left Stoplights

    Sheets("Stoplights").Select
    Range(Cells(21, k), Cells(35, l)).Select
    Selection.Copy
    ' Wait to load to clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))

    PTT.ActivePresentation.Slides(i).Select
    PTT.ActiveWindow.View.PasteSpecial DataType:=ppPasteEnhancedMetafile

    With PTT.ActiveWindow.Selection.ShapeRange
        .Height = 262.08
        .Left = 282.24
        .Top = 182.16
        .Name = "Stoplights1"
    End With

        k = k + 2
        l = l + 2

' Copy and Paste Right Stoplights

    Sheets("Stoplights").Select
    Range(Cells(21, k), Cells(35, l)).Select
    Selection.Copy
    ' Wait to load to clipboard
    Application.Wait (Now() + TimeValue("00:00:01"))

    PTT.ActivePresentation.Slides(i).Select
    PTT.ActiveWindow.View.PasteSpecial DataType:=ppPasteEnhancedMetafile

    With PTT.ActiveWindow.Selection.ShapeRange
        .Height = 262.08
        .Left = 614.16
        .Top = 182.16
        .Name = "Stoplights2"
    End With

        k = k + 2
        l = l + 2

Skip15Again:
Next i

    Application.CutCopyMode = False
    Sheets("NPSTables").Select

' Open Save As Dialog
PTT.FileDialog(msoFileDialogSaveAs).Show

End Sub
