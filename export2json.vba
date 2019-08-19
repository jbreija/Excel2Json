'export to JSON
    Dim fs As Object
    Dim jsonfile
    Dim rangetoexport As Range
    Dim rowcounter As Long
    Dim columncounter As Long
    Dim linedata As String
    Dim EmptyFlag As Boolean
    Dim NumFlag As Boolean
    Dim NestObjectFlag As Boolean
    Dim ArrayFlag As Boolean
    Dim NestArrayFlag As Boolean
    Dim ColumnHeader As String
    Dim slashremover As String
    Dim ReplaceChar As String
    Dim CurrentCell As Variant
    Dim NestArrayAttribute As Boolean
    Dim ArrayName As String
    Dim ws As Worksheet
    Dim Tb As String
        
    NestObjectFlag = False
    ArrayFlag = False
    Tb = Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32)
    
'
' SortOriginalOrder
'
    ActiveWorkbook.Worksheets("research").ListObjects("Table3072").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("research").ListObjects("Table3072").Sort.SortFields. _
        Add Key:=Range("Table3072[[#All],[Original Order]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("research").ListObjects("Table3072").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'define the export range here, varies per each sheet
    Set rangetoexport = Sheet011.Range("B2:Y946")
   
    'create a blank file for writing data to
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'define the output directory of the file here
    output_path = ThisWorkbook.Path & "\stats"
    If Dir(output_path, vbDirectory) = vbNullString Then
        VBA.FileSystem.MkDir (output_path)
    End If
    
    Set jsonfile = fs.CreateTextFile(output_path & "\research" & ".json", True)

    'print a { at the very beginning of the json file
    linedata = "{"
    jsonfile.WriteLine linedata
    
    'for every row in the range defined above
    For rowcounter = 2 To rangetoexport.Rows.Count
        linedata = ""
        'On Error Resume Next
        
        'for every cell in the rowcounter row
        For columncounter = 1 To rangetoexport.Columns.Count
            
            CurrentCell = rangetoexport.Cells(rowcounter, columncounter)
            EmptyFlag = IsEmpty(CurrentCell)
            NumFlag = IsNumeric(CurrentCell)
            firstchar = Left(rangetoexport.Cells(1, columncounter), 1)
            
            If NumFlag = False Then
            CurrentCell = """" & CurrentCell & """"
            End If

            'empty cell in last column
            If EmptyFlag And columncounter = rangetoexport.Columns.Count And rowcounter <> rangetoexport.Rows.Count And Len(linedata) - 1 > 0 And ArrayFlag = False Then

            'skip empty cell
            ElseIf EmptyFlag Then
                GoTo NextIteration
                
            '1st cell in a row then start a new object
            ElseIf columncounter = 1 And ArrayFlag = True Then
                If Len(linedata) - 1 > 0 Then
                    linedata = Left(linedata, Len(linedata) - 1)
                End If
            
                linedata = Tb & Tb & Tb & "]" & Chr$(13) & Tb & "}," & Chr$(13) & Tb & Tb & CurrentCell & ": {"
                ArrayFlag = False
            
            ElseIf columncounter = 1 Then
                 linedata = Tb & rangetoexport.Cells(1, columncounter) & CurrentCell & ": {"
                 ArrayFlag = False
   
            'beginning of array
            ElseIf firstchar = "[" And ArrayFlag = False Then
                ArrayFlag = True
                ArrayName = rangetoexport.Cells(1, columncounter)
                ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                
                If IsEmpty(rangetoexport.Cells(rowcounter, 1)) = True And IsEmpty(rangetoexport.Cells(rowcounter - 1, 1)) = True And IsEmpty(rangetoexport.Cells(rowcounter, columncounter - 1)) = False Then
                    linedata = linedata & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": [" & Chr$(13) & Tb & Tb & Tb & CurrentCell & ","
                Else
                
                    linedata = linedata & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": [" & Chr$(13) & Tb & Tb & Tb & CurrentCell & ","
                End If
                
            'middle array attribute
            ElseIf firstchar = "[" And ArrayFlag = True And ArrayName = rangetoexport.Cells(1, columncounter) Then
                If IsEmpty(rangetoexport.Cells(rowcounter + 1, 1)) = False Then
                    linedata = linedata & Tb & Tb & Tb & CurrentCell & Chr$(13) & Tb & Tb & "],"
                    ArrayFlag = False
                Else
                     
                    linedata = linedata & Tb & Tb & Tb & CurrentCell & ","
                End If
                
                
            ElseIf firstchar = "[" And ArrayFlag = True And ArrayName <> rangetoexport.Cells(1, columncounter) Then
                ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                If Len(linedata) - 1 > 0 Then
                    linedata = Left(linedata, Len(linedata) - 1)
                End If
                ArrayName = rangetoexport.Cells(1, columncounter)
                linedata = linedata & Tb & Tb & Chr$(13) & Tb & Tb & "]" & "," & Chr$(13) & Tb & Tb & Tb & """" & ColumnHeader & """" & ": [" & Chr$(13) & Tb & Tb & Tb & CurrentCell & ","
                    
                    
            'end of array
            ElseIf firstchar <> "[" And ArrayFlag = True Then
                ArrayFlag = False
                ArrayName = ""
            
                If Len(linedata) - 1 > 0 Then
                    linedata = Left(linedata, Len(linedata) - 1)
                End If
            
                'single cell array after normal array
                If firstchar = ">" Then
                    ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                    ReplaceChar = Replace(CurrentCell, " | ", """" & ", " & Chr$(13) & Tb & Tb & Tb & """")
                    linedata = linedata & Chr$(13) & Tb & Tb & "]," & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": [" & Chr$(13) & Tb & Tb & Tb & ReplaceChar & Chr$(13) & Tb & Tb & Tb & "], "
                    linedata = Left(linedata, Len(linedata) - 1)
                    ArrayFlag = False
                   
                ElseIf firstchar = "^" Then
                    NestArrayFlag = True
                    ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                    linedata = linedata & Chr$(13) & Tb & Tb & "]," & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": ["
                    ArrayFlag = False

                Else
                    ColumnHeader = rangetoexport.Cells(1, columncounter)
                    linedata = linedata & Chr$(13) & Tb & Tb & "]," & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": " & CurrentCell & ","
                    ArrayFlag = False
                    ArrayName = ""

                End If
            
            
            Else
                ArrayFlag = False
            
                'beginning of nested array
                If firstchar = "^" And NestArrayFlag = False Then
                    NestArrayFlag = True
                    ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                    linedata = linedata & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": ["
            
                ElseIf firstchar = "^" And NestArrayAttribute = True Then
                   ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                   linedata = linedata & Tb & Tb & "{"
            
                'first nestarrayattribute
                ElseIf firstchar = "#" And NestArrayAttribute = False Then
                    NestArrayAttribute = True
                    ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                    linedata = linedata & Chr$(13) & Tb & Tb & Tb & "{" & Chr$(13) & Tb & Tb & Tb & Tb & """" & ColumnHeader & """" & ": " & CurrentCell & ","
                        
                'middle of nested array attribute
                ElseIf firstchar = "#" And NestArrayAttribute = True Then
                    ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                    linedata = linedata & Chr$(13) & Tb & Tb & Tb & Tb & """" & ColumnHeader & """" & ": " & CurrentCell & ","
                    
                'last nestarray attribute
                ElseIf firstchar <> "#" And NestArrayAttribute = True Then
                    NestArrayAttribute = False
                    NestArray = False
                    If Len(linedata) - 1 > 0 Then
                        linedata = Left(linedata, Len(linedata) - 1)
                    End If
                    ColumnHeader = rangetoexport.Cells(1, columncounter)
                    linedata = linedata & Chr$(13) & Tb & Tb & Tb & "}" & Chr$(13) & Tb & Tb & "]," & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": " & CurrentCell & ","
            
            'not an array element
                Else
                    NestArrayFlag = False
                    NestArrayAttribute = False
      
                'Single Cell Array
                    If firstchar = ">" Then
                            ColumnHeader = Right(rangetoexport.Cells(1, columncounter), Len(rangetoexport.Cells(1, columncounter)) - 1)
                            ReplaceChar = Replace(CurrentCell, " | ", """" & ", " & Chr$(13) & Tb & Tb & Tb & """")
                            linedata = linedata & Chr$(13) & Tb & Tb & """" & ColumnHeader & """" & ": [" & Chr$(13) & Tb & Tb & Tb & ReplaceChar & ","
                            linedata = Left(linedata, Len(linedata) - 1) & Chr$(13) & Tb & Tb & "],"
                    Else
 
                        'last cell in row number
                        If columncounter = rangetoexport.Columns.Count And NumFlag = True Then
                        linedata = linedata & Chr$(13) & Tb & Tb & """" & rangetoexport.Cells(1, columncounter) & """" & ": " & CurrentCell
                        
                        'last cell in row string
                        ElseIf columncounter = rangetoexport.Columns.Count Then
                            linedata = linedata & Chr$(13) & Tb & Tb & """" & rangetoexport.Cells(1, columncounter) & """" & ": " & CurrentCell & Chr$(13) & Tb & "}, "
                                            
                        ElseIf rowcounter = rangetoexport.Rows.Count And columncounter = rangetoexport.Columns.Count Then
                            linedata = Left(linedata, Len(linedata) - 1)
                            linedata = linedata & Chr$(13) & Tb & Tb & """" & rangetoexport.Cells(1, columncounter) & """" & ": " & CurrentCell
                        
                        Else
                            linedata = linedata & Chr$(13) & Tb & Tb & """" & rangetoexport.Cells(1, columncounter) & """" & ": " & CurrentCell & ","
                        End If
                    End If
                End If
            End If
            
        If columncounter = rangetoexport.Columns.Count And rowcounter <> rangetoexport.Rows.Count And Len(linedata) - 1 > 0 And ArrayFlag = False Then
            If Right(linedata, 1) = "," Then
                linedata = Left(linedata, Len(linedata) - 1)
            End If
            linedata = linedata & Chr$(13) & Tb & "}, "
                
        ElseIf columncounter = rangetoexport.Columns.Count And rowcounter <> rangetoexport.Rows.Count And Len(linedata) - 1 > 0 And NestArrayAttribute = False And ArrayFlag = True Then
            linedata = Left(linedata, Len(linedata) - 1)
            linedata = linedata & Chr$(13) & Tb & Tb & Tb & "}, "
        
        End If
        
        'used to break out of the loop if the cell is empty
NextIteration:

        If Len(linedata) - 1 > 0 And Right(linedata, 1) = "," And rowcounter = rangetoexport.Rows.Count And columncounter = rangetoexport.Columns.Count Then
                linedata = Left(linedata, Len(linedata) - 1)
       End If

        'next cell in the row
        Next

       If Len(linedata) - 1 > 0 And Right(linedata, 1) = "," And IsEmpty(rangetoexport.Cells(rowcounter + 1, 1)) = False And columncounter <> rangetoexport.Columns.Count Then
            linedata = Left(linedata, Len(linedata) - 1)
       End If

    'print the data to the file
    jsonfile.WriteLine linedata
        
    'next row in the range
    Next

'some punction added at the end of every json file
linedata = Tb & "}" & Chr$(13) & "}"
    
jsonfile.WriteLine linedata
jsonfile.Close
   
Set fs = Nothing

End Sub