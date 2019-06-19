Attribute VB_Name = "save_load_geom_file"
Option Explicit

Private Sub export_geoFile()
    'exports the joint numbers, spherical coordinates, and the boxed cells to the data tab
    With ThisWorkbook.Sheets("createGeo")
        lastrow = .Cells(Rows.count, "A").End(xlUp).ROW
        myarray2 = .Range("b2:c14").Value
        myArray3 = .Range("n3:O5").Value
        myArray4 = .Range("n8:o11").Value
        myArray5 = .Range("a19:C" & lastrow).Value
    End With
    Workbooks.Add
    Set wbfordomeapp = ActiveWorkbook
    wbfordomeapp.Sheets.Add.Name = "Data"
    Application.DisplayAlerts = False
    wbfordomeapp.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    With wbfordomeapp.Sheets("Data")
        .Range("b4:C16").Value = myarray2
        .Range("b17:C19").Value = myArray3
        .Range("b20:C23").Value = myArray4
        .Range("F4:h" & lastrow - 15).Value = myArray5
        Erase myArray5
        .Range("b24") = "dome radius at shoe end"
        .Range("b25") = "Hs"
        .Range("b26") = "Vs"
        .Range("b27") = "avg beam length"
        .Range("b28") = "geometry name"
        .Range("b29") = "R/D Ratio"
        .Range("b30") = "support beam offset"
        .Range("b32") = "6095"
        .Range("b33") = "6620"
        .Range("d32") = "LF"
        .Range("d33") = "LF"
        .Range("E32") = "102"" coil"
        .Range("E33") = "96"" coil"
        .Range("b34:b36") = "xxx"
        .Range("c24") = Range("domerad_shoe")
        .Range("C25") = Hs
        .Range("C26") = Vs
        .Range("C27").Value = Range("average_beam_length").Value
        .Range("C28").Value = Range("geometry_name").Value
        .Range("C29").Value = Range("r_over_d_ratio").Value
        .Range("C30").Value = Range("shoe_beam_offset").Value
        .Range("C32") = Range("len_102_coil")
        .Range("C33") = Range("len_96_coil")
        .Range("c38") = Range("max_panel_altitude")
        .Range("b38") = "max panel altitude"
        
    End With
End Sub



Sub save_to_geometry_libary()

    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    'open the msg box
    strPrompt = "Publish for all users in the Geometry Library?"
    iRet = MsgBox(strPrompt, vbYesNo, "warning")
    If iRet = vbYes Then
        FilePath = "X:\HMT_DOME_SOFTWARE\Geometry Library\geometries created\"
    Else
        FilePath = ThisWorkbook.Path & "\"
    End If
    
    Dim wireframeMode As String
    
    Call altitude_beam_node_angle_check
    If count_total <> 0 Then
        strPrompt = "Requirements are not met, are you sure you want to export?"
        iRet = MsgBox(strPrompt, vbYesNo, "warning")
        If iRet = vbYes Then
        Else
            Exit Sub
        End If
    End If
    
    wireframeMode = Sheets("Nodes_code").Range("fab_or_analysis")
    
    Sheets("Nodes_code").Range("fab_or_analysis") = "analysis"
    Call master_nodes
    Call master_beams
    Call master_panels
    Call export_geoFile

    wbfordomeapp.Activate
    With ThisWorkbook.Sheets("createGeo")
        DT = Format(CStr(Now), "mm-dd_hh-mm-ss")
        domedia = Round(.Range("anchor_bolt_diameter"), 3)    'domedia = Replace(Round(.Cells(2, "C"), 3), ".", "_")
        beta_angle = Round(.Range("beta_create"), 3)
        shoe_offset = Round(.Range("shoe_beam_offset"), 3)
        supportbeamlen = Round(.Range("support_beam_length_to_pin"), 3)
        geomname = .Range("geometry_name")
    End With
    fileSaveName = FilePath & domedia & " ft - " & beta_angle & Chr(176) & "- " & _
        geomname & " - " & shoe_offset & " offs X " & supportbeamlen & " lng - " & DT & ".xlsx"
        
    ThisWorkbook.Sheets("Nodes_code").Range("fab_or_analysis") = wireframeMode
    wbfordomeapp.SaveAs Filename:=fileSaveName
 
    wbfordomeapp.Close SaveChanges:=False
    Set wbfordomeapp = Nothing
    
    ThisWorkbook.Sheets("createGeo").Activate
    

    MsgBox "File has been SUCCESSFULLY exported to the following location " & FilePath
  
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description & " error : the file failed to export"
    ThisWorkbook.Sheets("createGeo").Activate
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub load_file()
    Dim xArray1() As Variant


    Dim dataworbook As Workbook
    
    Dim acadApp As Object
    Dim acadDoc As Object
    Dim BBarray() As Variant
    Dim BlkAtts As Variant

    Dim CurrentFile As String
    Dim effName As String
    Dim FilterData() As Variant
    Dim FilterType() As Integer
    Dim i As Long
    Dim lastcolumn As Long
    Dim lastrow As Long
    Dim Nrows As Long
    Dim objInSelect As Variant

    Dim Path As String
    Dim rowBB As Long
    ReDim BBarray(1 To 2000, 1 To 16)
    ReDim FilterType(3)
    ReDim FilterData(3)
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Path = "X:\HMT_DOME_SOFTWARE\Geometry Library\geometries created" & "\"
   
    CurrentFile = Dir(Path + "*.xlsx", vbNormal)
   
    rowBB = 1
    Do While CurrentFile <> ""
    
        Workbooks.Open Filename:=Path + CurrentFile
    
    
        Set Load_sheet = ActiveWorkbook.Sheets("Data")
    
        Set dataworbook = ActiveWorkbook
  
        Load_sheet.Activate
        lastrow2 = Load_sheet.Cells(Rows.count, "f").End(xlUp).ROW
        xArray1 = Load_sheet.Range("G5:H" & lastrow2).Value
    
        With ThisWorkbook.Sheets("createGeo")
            .Range("C3:C4").Value = Load_sheet.Range("C5:C6").Value
            .Range("beta_create").Value = Load_sheet.Range("C9").Value
            .Range("support_beam_length_to_pin") = Load_sheet.Range("c8")
            .Range("shoe_beam_offset").Value = Load_sheet.Range("C30").Value
            .Range("C10:C12").Value = Load_sheet.Range("C12:C14").Value
            .Range("dish_angle").Value = Load_sheet.Range("c7")
        
            If Load_sheet.Range("c38") = "" Then
                .Range("max_panel_altitude").Value = 107.5
            Else
                .Range("max_panel_altitude").Value = Load_sheet.Range("c38")
            End If
            .Calculate
        End With
    
 
        ThisWorkbook.Sheets("createGeo").Activate
  

        ActiveSheet.Unprotect

        Call master_beams_creation
        Call master_nodes_creation
        Call master_panels_creation
        Range("c15").Select
    
        Call find_polar_angle_branches_part1

        With ThisWorkbook.Sheets("createGeo")
            For xrow = 1 To lastrow2 - 5
                If .Cells(xrow + 19, "b").HasFormula Then   ' paste all the values from xarray1 but dont paste where there is a formula
                Else
                    .Cells(xrow + 19, "b") = xArray1(xrow, 1)
                    .Cells(xrow + 19, "c") = xArray1(xrow, 2)
                End If
            Next xrow
            .Calculate
        End With
   
        Call altitude_beam_node_angle_check

        Call find_unique_values_beams_PRE
    
        ThisWorkbook.Sheets("createGeo").Activate

        Call a_reset_worksheet_to_defaults
    
        Calculate


        Load_sheet.Range("c34") = ""
        Load_sheet.Range("b34") = ""
 
 
        Load_sheet.Range("c31") = Range("average_panel_area")
        Load_sheet.Range("b31") = "average_panel_area"
        
        Load_sheet.Range("c38") = Range("max_panel_altitude")
        Load_sheet.Range("b38") = "max_panel_altitude"
    
        Load_sheet.Range("b32") = "PN6095 102"" coil"
        Load_sheet.Range("b33") = "PN6620 96"" coil"
        
        Load_sheet.Range("d32") = "LF"
        Load_sheet.Range("d33") = "LF"
        
         Load_sheet.Range("d30") = "in."
         Load_sheet.Range("d31") = "in^2"
         
          Load_sheet.Range("d34") = "in."
    
        Load_sheet.Range("d33") = ""
        Load_sheet.Range("e33") = ""
        Load_sheet.Range("d32") = ""
        Load_sheet.Range("e32") = ""
    
     
        Load_sheet.Range("b35") = ""
        Load_sheet.Range("b36") = ""
    
        dataworbook.Close SaveChanges:=True
    
        Set Load_sheet = Nothing
    
        CurrentFile = Dir
        
    Loop
    
    
    Application.SreenUpdating = True
    
    Application.Calculation = xlCalculationAutomatic
   

End Sub






