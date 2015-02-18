Sub ExportGSMWacht_Click()
    ' Export file for GSM-Wacht
    
    Dim ws As Worksheet, wsExport As Worksheet, wsList As Worksheet
    
    Set wsList = ThisWorkbook.Sheets("Export GSM-Wacht")
    Set wsExport = ThisWorkbook.Sheets("Export")
    
    ' clear this sheet
    wsExport.Range("A2:F200").ClearContents
    
    ' select sheet for month to treat
    SelectedMonth = Range("Export_Month").Value
    SelectedSheet = Format(Month(SelectedMonth), "00") & "-" & Year(SelectedMonth) - 2000
    Set ws = Sheets(SelectedSheet)
    ws.Activate
    
    Counter = 2
    
    With ws
        ' BI Daily Check
        
        For Each c In Range("BE7:BE37").Cells
            PersInitials = c.Value
            If Len(PersInitials) > 1 Then
                PersMatricule = Application.WorksheetFunction.VLookup(PersInitials, wsList.Range("PRIM_Selection").Value, 2, False)
                PersName = Application.WorksheetFunction.VLookup(PersInitials, wsList.Range("PRIM_Selection").Value, 3, False)
                PersDateGarde = Cells(c.Row, 1)
                wsExport.Cells(Counter, 1).Value = PersMatricule
                wsExport.Cells(Counter, 2).Value = PersDateGarde
                wsExport.Cells(Counter, 3).Value = "06:45"
                wsExport.Cells(Counter, 4).Value = "08:45"
                wsExport.Cells(Counter, 5).Value = PersName
                wsExport.Cells(Counter, 6).Value = "BI Daily Check"
                Counter = Counter + 1
            End If
        Next
    
        ' DS Duty
        
        For Each c In Range("BF7:BF37").Cells
            PersInitials = c.Value
            If Len(PersInitials) > 1 Then
                PersMatricule = Application.WorksheetFunction.VLookup(PersInitials, wsList.Range("PRIM_Selection").Value, 2, False)
                PersName = Application.WorksheetFunction.VLookup(PersInitials, wsList.Range("PRIM_Selection").Value, 3, False)
                PersDateGarde = Cells(c.Row, 1)
                wsExport.Cells(Counter, 1).Value = PersMatricule
                wsExport.Cells(Counter, 2).Value = PersDateGarde
                Select Case Weekday(PersDateGarde)
                    Case vbSunday, vbSaturday
                        wsExport.Cells(Counter, 3).Value = "08:45"
                        wsExport.Cells(Counter, 4).Value = "08:45"
                    Case Else
                        If c.Interior.ColorIndex = 15 Then
                        ' Feestdag
                            wsExport.Cells(Counter, 3).Value = "08:45"
                            wsExport.Cells(Counter, 4).Value = "08:45"
                        Else
                        ' Werkdag
                            wsExport.Cells(Counter, 3).Value = "16:45"
                            wsExport.Cells(Counter, 4).Value = "08:45"
                        End If
                End Select
                wsExport.Cells(Counter, 5).Value = PersName
                wsExport.Cells(Counter, 6).Value = "DS Duty"
                Counter = Counter + 1
            End If
        Next

    End With

    wsExport.Activate
        
    MsgBox "Export Done for " & SelectedMonth & " (" & Counter - 2 & " rows)"
    
End Sub

