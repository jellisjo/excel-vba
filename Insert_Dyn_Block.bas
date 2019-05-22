Attribute VB_Name = "Insert_Dyn_Block"
Option Explicit
Sub Insert_block()

    Dim acadDoc As Object
    Dim acadApp As Object
    Dim objBlockRef As AcadBlockReference
    Dim FilePath As String
    Dim strName As String
    Dim insertionPoint(0 To 2) As Double
    Dim pntPanel(0 To 2) As Double
    Dim BlkAtts As Variant
    
    'Check if AutoCAD application is open. If it is not opened create a new instance and make it visible.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    acadApp.Visible = True
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application")
        acadApp.Visible = True
    End If
     
    'Check (again) if there is an AutoCAD object.
    If acadApp Is Nothing Then
        MsgBox "Sorry, it was impossible to start AutoCAD!", vbCritical, "AutoCAD Error"
        Exit Sub
    End If
    On Error GoTo 0
    acadApp.WindowState = acMax
    
    'If there is no active drawing create a new one.
    On Error Resume Next
    Set acadDoc = acadApp.Documents.Add
    If Err.Number <> 0 Then
        MsgBox "Could not get the AutoCAD application.  Restart Autocad and try again"
        End
    End If
    On Error GoTo 0
     
    acadDoc.ActiveSpace = acModelSpace
    
    FilePath = ThisWorkbook.Path & "\Support\"
    strName = FilePath & "blocks.dwg"    'all sorts of blocks that you want to use can be added in this file
    Set objBlockRef = acadDoc.ModelSpace.InsertBlock(insertionPoint, strName, 1, 1, 1, 0)
    objBlockRef.Delete
        
    pntPanel(0) = 0
    pntPanel(1) = 0
    Set objBlockRef = acadDoc.ModelSpace.InsertBlock(pntPanel, "Dyn_Rec", 1, 1, 1, 0)
    BlkAtts = objBlockRef.GetDynamicBlockProperties
    
'        Dim i As Long
'        On Error Resume Next
'        For i = 0 To UBound(BlkAtts)
'            Debug.Print BlkAtts(i).Value & "   " & i         'use this to figure out the position of each attribute
'        Next i
'        On Error GoTo 0
    
    BlkAtts(0).Value = 25#   'bottom length
    BlkAtts(2).Value = 30#   'left length
    BlkAtts(4).Value = 20#    'right length
    BlkAtts(6).Value = 0# ' text x position
    BlkAtts(7).Value = 13#  'text y position
    BlkAtts(8).Value = 3.14 / 2  'text rotation
    BlkAtts(10).Value = 2#   'text height
    
    BlkAtts = objBlockRef.GetAttributes
    BlkAtts(0).textString = "BB1"
                
    acadDoc.Regen acAllViewports
    acadApp.ZoomExtents
    
End Sub

