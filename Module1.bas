Attribute VB_Name = "Module1"
Option Explicit

Dim acadDoc As Object
    Dim AcadUcsObject As AcadUCS



Function Cross3D(A As Variant, B As Variant) As Variant
    ' A and B must be dimensioned Double(0 to 2)
    Dim variable_C(0 To 2) As Double
    variable_C(0) = A(1) * B(2) - A(2) * B(1)
    variable_C(1) = -(A(0) * B(2) - A(2) * B(0))
    variable_C(2) = A(0) * B(1) - A(1) * B(0)
    Cross3D = variable_C
End Function


Function Add_UCS_improved(origin As Variant, xAxisPnt _
    As Variant, yAxisPnt As Variant, ucsName As String) As AcadUCS
    Dim xAxisVec(0 To 2) As Double
    Dim yAxisVec(0 To 2) As Double
    Dim perpYaxisPnt(0 To 2) As Double
    Dim xCy As Variant, perpYaxisVec As Variant
    xAxisVec(0) = xAxisPnt(0) - origin(0)
    xAxisVec(1) = xAxisPnt(1) - origin(1)
    xAxisVec(2) = xAxisPnt(2) - origin(2)
    yAxisVec(0) = yAxisPnt(0) - origin(0)
    yAxisVec(1) = yAxisPnt(1) - origin(1)
    yAxisVec(2) = yAxisPnt(2) - origin(2)
    xCy = Cross3D(xAxisVec, yAxisVec)
    perpYaxisVec = Cross3D(xCy, xAxisVec)
    perpYaxisPnt(0) = perpYaxisVec(0) + origin(0)
    perpYaxisPnt(1) = perpYaxisVec(1) + origin(1)
    perpYaxisPnt(2) = perpYaxisVec(2) + origin(2)
    Set AcadUcsObject = acadDoc.UserCoordinateSystems.Add(origin, xAxisPnt, _
        perpYaxisPnt, ucsName)
End Function


Sub draw_CNC_flashing_3Dverify()
    Dim circleObj(1) As AcadCircle
    Dim pointUCS As Variant
    Dim plineObjLW1 As AcadLWPolyline
    Dim solidObjextrus_sides1 As Acad3DSolid
    Dim solidObjextrus_sides2 As Acad3DSolid
    Dim curves_sides(0) As AcadEntity
    Dim curves_panel_cir(1) As AcadEntity
    Dim vertices() As Double

    Dim acadApp As Object

    Dim currUCS As AcadUCS
    Dim PointVariantOrigin(0 To 2) As Double
    Dim PointVariantXAxis(0 To 2) As Double
    Dim PointVariantYAxis(0 To 2) As Double
    Dim plineObjLW As AcadLWPolyline
    Dim TransMatrix As Variant
    Dim regionObj As Variant
    Dim center(0 To 2) As Double
    Dim pointObj1 As AcadPoint
    
    
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
    
   
    'define the original ucs
    With acadDoc
        Set currUCS = .UserCoordinateSystems.Add( _
            .GetVariable("UCSORG"), _
            .Utility.TranslateCoordinates(.GetVariable("UCSXDIR"), acUCS, acWorld, 0), _
            .Utility.TranslateCoordinates(.GetVariable("UCSYDIR"), acUCS, acWorld, 0), _
            "OriginalUCSj")
    End With

   
    PointVariantYAxis(0) = 0
    PointVariantYAxis(1) = 5
    PointVariantYAxis(2) = 50
    PointVariantOrigin(0) = 0
    PointVariantOrigin(1) = 5
    PointVariantOrigin(2) = 0
    PointVariantXAxis(0) = 50
    PointVariantXAxis(1) = 5
    PointVariantXAxis(2) = 0
    Call Add_UCS_improved(PointVariantOrigin, PointVariantYAxis, PointVariantXAxis, "UCSName1")
        

    ReDim vertices(7)
    vertices(0) = -68.5: vertices(1) = -0.19
    vertices(2) = -75.02: vertices(3) = 49.38
    vertices(4) = vertices(2) - 100: vertices(5) = vertices(3)
    vertices(6) = vertices(4): vertices(7) = 0
    Set plineObjLW = acadDoc.ModelSpace.AddLightWeightPolyline(vertices)
    plineObjLW.Closed = True
    TransMatrix = AcadUcsObject.GetUCSMatrix()
    plineObjLW.TransformBy (TransMatrix)
    Set curves_sides(0) = plineObjLW
    regionObj = acadDoc.ModelSpace.AddRegion(curves_sides)
    Set solidObjextrus_sides1 = acadDoc.ModelSpace.AddExtrudedSolid(regionObj(0), -50, 0)
    regionObj(0).Delete
    
    Set plineObjLW1 = plineObjLW.Mirror(PointVariantOrigin, PointVariantXAxis)
    Set curves_sides(0) = plineObjLW1
    regionObj = acadDoc.ModelSpace.AddRegion(curves_sides)
    Set solidObjextrus_sides2 = acadDoc.ModelSpace.AddExtrudedSolid(regionObj(0), -50, 0)
    regionObj(0).Delete
    plineObjLW1.Delete
    plineObjLW.Delete
        
    PointVariantYAxis(0) = 0
    PointVariantYAxis(1) = 0
    PointVariantYAxis(2) = 2
    PointVariantOrigin(0) = 0
    PointVariantOrigin(1) = 0
    PointVariantOrigin(2) = 0
    PointVariantXAxis(0) = 9.9
    PointVariantXAxis(1) = -8.02
    PointVariantXAxis(2) = 0
        
    Call Add_UCS_improved(PointVariantOrigin, PointVariantYAxis, PointVariantXAxis, "UCSName1")
    acadDoc.ActiveUCS = AcadUcsObject

     
    Dim cir_rad As Double
    cir_rad = 16.25
    center(0) = 68.5
    center(1) = -0.25
    center(2) = 0
        
    pointUCS = acadDoc.Utility.TranslateCoordinates(center, acUCS, acWorld, False)
    Set circleObj(0) = acadDoc.ModelSpace.AddCircle(pointUCS, cir_rad)
    Set curves_panel_cir(0) = circleObj(0)
    Set pointObj1 = acadDoc.ModelSpace.AddPoint(pointUCS)
        
        
    cir_rad = 25
    center(0) = -68.5
    center(1) = -0.25
    center(2) = 0
        
    pointUCS = acadDoc.Utility.TranslateCoordinates(center, acUCS, acWorld, False)
    Set circleObj(1) = acadDoc.ModelSpace.AddCircle(pointUCS, cir_rad)
    Set pointObj1 = acadDoc.ModelSpace.AddPoint(pointUCS)
 
         
    acadDoc.ActiveUCS = currUCS
        
    
    acadDoc.Regen acAllViewports
    acadApp.ZoomExtents
   
End Sub




