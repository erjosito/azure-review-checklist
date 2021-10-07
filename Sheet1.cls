VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 0

Sub export_json()
    On Error GoTo errh
    ' Constants
    Const row1 = 5
    Const row_limit = 2000   'Safety net
    Const line_break = vbCrLf
    Const category_column = 1
    Const subcategory_column = 2
    Const text_column = 3
    Const severity_column = 4
    Const status_column = 5
    Const link_column = 7
    Const export_file_name = "checklist.json"
    ' Variables
    Dim row As Integer
    Dim json As String
    Dim check_category, check_subcategory, check_text, check_severity, check_status, check_link As String
    Dim export_file_path As Variant
    Dim double_quote As String
    ' Initialization
    double_quote = Chr(34) ' double quote as a variable
    row = row1
    json = "{" + line_break
    json = json + "  " + double_quote + "items" + double_quote + ": ["
    ' Loop through all rows as long as there is content
    Do While row < row_limit And Len(Cells(row, category_column)) > 0
        If row > row1 Then json = json + ","
        json = json + line_break
        check_category = Cells(row, category_column)
        check_subcategory = Cells(row, subcategory_column)
        check_text = Cells(row, text_column)
        check_severity = Cells(row, severity_column)
        If Cells(row, link_column).Hyperlinks.Count > 0 Then
            check_link = Cells(row, link_column).Hyperlinks(1).Address
        Else
            check_link = Cells(row, link_column)
        End If
        row = row + 1
        json = json + "    {" + line_break
        json = json + "      " + double_quote + "category" + double_quote + ": " + double_quote + check_category + double_quote + "," + line_break
        json = json + "      " + double_quote + "subcategory" + double_quote + ": " + double_quote + check_subcategory + double_quote + "," + line_break
        json = json + "      " + double_quote + "text" + double_quote + ": " + double_quote + Replace(Replace(check_text, Chr(10), ""), Chr(13), "") + double_quote + "," + line_break
        json = json + "      " + double_quote + "severity" + double_quote + ": " + double_quote + check_severity + double_quote + "," + line_break
        json = json + "      " + double_quote + "link" + double_quote + ": " + double_quote + Replace(check_link, "en-us/", "") + double_quote + line_break
        json = json + "    }"
    Loop
    ' Finish JSON
    json = json + line_break + "  ]" + line_break
    json = json + "}" + line_break
    ' Write JSON to file
    ' MsgBox json
    'export_file_path = ActiveWorkbook.Path + "\" + export_file_name
    export_file_path = ""
    export_file_path = Application.GetSaveAsFilename(FileFilter:="JSON File (*.json), *.json", Title:="Exporting JSON checklist", InitialFileName:=ActiveWorkbook.Path + "\" + export_file_name)
    'checks to make sure the user hasn't canceled the dialog
    If export_file_path <> False Then
        'MsgBox "Exporting to " + export_file_path
        Open export_file_path For Output As #1
        Print #1, json
        Close #1
    End If
    Exit Sub
errh:
    If Err.Number <> 0 Then
        MsgBox "Error while exporting checklist to JSON " & Err.Description
    End If
End Sub


'Parse JSON code using the JsonConverter module
Sub import_checklist()
    On Error GoTo errh
    ' Constants
    Const row1 = 5
    Const row_limit = 2000   'Safety net
    Const category_column = 1
    Const subcategory_column = 2
    Const text_column = 3
    Const severity_column = 4
    Const status_column = 5
    Const link_column = 7
    'Variables
    Dim json_file As Variant
    Dim json_ts As TextStream
    Dim FSO As New FileSystemObject
    Dim textline As String
    Dim json As String
    Dim json_object, json_item As Object
    Dim check_category, check_subcategory, check_text, check_severity, check_status, check_link As String
    Dim double_quote As String
    Dim line_elements() As String
    Dim row, i As Integer
    'Initialization
    row = row1
    double_quote = Chr(34) ' double quote as a variable
    json = ""
    ChDir ActiveWorkbook.Path
    json_file = Application.GetOpenFilename(Title:="Please choose a file to open", FileFilter:="JSON Files *.json* (*.json),")
    If json_file = False Then
        MsgBox "No file selected.", vbExclamation, "Sorry!"
        Exit Sub
    Else
        'Go through the file line by line
        ' Read .json file
        Set json_ts = FSO.OpenTextFile(json_file, ForReading)
        json = json_ts.ReadAll
        json_ts.Close
        json = Replace(json, vbCrLf, "")
        Set json_object = JsonConverter.ParseJson(json)  'This line is not working. Run-time error 10001
        row = row1
        For Each json_item In json_object("items")
            Cells(row, category_column) = json_item("category")
            Cells(row, subcategory_column) = json_item("subcategory")
            Cells(row, text_column) = json_item("text")
            Cells(row, severity_column) = json_item("severity")
            Cells(row, link_column) = json_item("link")
            row = row + 1
        Next json_item
        'Blank the rest of the rows
        Do While row < row_limit And Len(Cells(row, category_column)) > 0
            Cells(row, category_column) = ""
            Cells(row, subcategory_column) = ""
            Cells(row, text_column) = ""
            Cells(row, severity_column) = ""
            Cells(row, link_column) = ""
            row = row + 1
        Loop
    End If
    Exit Sub
errh:
    If Err.Number <> 0 Then
        MsgBox "Error while importing checklist from JSON " & Err.Description
    End If
End Sub

