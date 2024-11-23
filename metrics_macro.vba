Sub AggregateDataFromWorkbooks()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim totalResolvedYes As Long
    Dim totalResolvedNo As Long
    Dim totalEntries As Long
    Dim totalTime As Double
    Dim resolvedCount As Long
    Dim cell As Range
    Dim resolutionTime As Variant
    Dim entryTime As Variant
    Dim timeTaken As Double
    Dim lastRow As Long
    Dim msg As String

    ' Path to the folder containing all the workbooks
    folderPath = "C:\YourFolderPath\" ' Change this to the path where your files are saved

    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "The folder path is incorrect or does not exist.", vbCritical
        Exit Sub
    End If

    ' Initialize counters
    totalResolvedYes = 0
    totalResolvedNo = 0
    totalEntries = 0
    totalTime = 0
    resolvedCount = 0

    ' Loop through all Excel files in the folder
    fileName = Dir(folderPath & "*.xlsm") ' Adjust file extension if needed

    Do While fileName <> ""
        ' Open each workbook
        Set wb = Workbooks.Open(folderPath & fileName)
        Set ws = wb.Sheets("Sheet1") ' Change "Sheet1" to the actual name of the sheet if different

        ' Find the last row in the sheet
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row ' Assumes data in Column B

        ' Count total entries (rows with data)
        totalEntries = totalEntries + lastRow - 1 ' Subtract header row

        ' Loop through the "Resolved?" column (assumed Column I)
        For Each cell In ws.Range("I2:I" & lastRow) ' Adjust column if needed
            If cell.Value = "Yes" Then
                ' Extract resolution time from Column J and entry time from Columns B and C
                resolutionTime = ws.Cells(cell.Row, 10).Value ' Column J
                entryTime = ws.Cells(cell.Row, 2).Value + ws.Cells(cell.Row, 3).Value ' Columns B (Date) + C (Time)

                ' Calculate time taken for resolution
                If IsDate(resolutionTime) And IsDate(entryTime) Then
                    timeTaken = resolutionTime - entryTime
                    totalTime = totalTime + timeTaken
                    resolvedCount = resolvedCount + 1
                End If

                totalResolvedYes = totalResolvedYes + 1
            ElseIf cell.Value = "No" Then
                totalResolvedNo = totalResolvedNo + 1
            End If
        Next cell

        ' Close the workbook without saving changes
        wb.Close False

        ' Move to the next file
        fileName = Dir
    Loop

    ' Calculate average resolution time
    Dim averageTime As Double
    Dim formattedAvgTime As String
    Dim days As Long, hours As Long, minutes As Long, seconds As Long

    If resolvedCount > 0 Then
        averageTime = totalTime / resolvedCount

        ' Convert average time to days, hours, minutes, and seconds
        days = Int(averageTime)
        hours = Int((averageTime - days) * 24)
        minutes = Int(((averageTime - days) * 24 - hours) * 60)
        seconds = Round((((averageTime - days) * 24 - hours) * 60 - minutes) * 60)

        ' Format the output string
        formattedAvgTime = ""
        If days > 0 Then formattedAvgTime = formattedAvgTime & days & " day(s) "
        If hours > 0 Then formattedAvgTime = formattedAvgTime & hours & " hour(s) "
        If minutes > 0 Then formattedAvgTime = formattedAvgTime & minutes & " minute(s) "
        If seconds > 0 Then formattedAvgTime = formattedAvgTime & seconds & " second(s)"
    Else
        formattedAvgTime = "No Resolved Cases"
    End If

    ' Display results
    msg = "Total Entries: " & totalEntries & vbCrLf & _
          "Resolved (Yes): " & totalResolvedYes & vbCrLf & _
          "Unresolved (No): " & totalResolvedNo & vbCrLf & _
          "Average Resolution Time: " & formattedAvgTime
    MsgBox msg, vbInformation, "Aggregation Results"

End Sub
