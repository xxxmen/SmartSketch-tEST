Imports System
Public Class frmMain

    Private Sub btnHello_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHello.Click
        Dim result As DialogResult
        Dim objSmartSketch As Object
        Dim sketchDocument As SmartSketch.Document
        Dim drawnSymbol As SmartSketch.Symbol2d
        Dim symSPID As String
        Dim symType As String
        Dim drawnConnector As SmartSketch.Connector2d
        Dim connSPID As String
        Dim connType As String
        Dim drawnSmartSym As SmartSketch.SmartSymbol2d
        Dim sppidDatasource As New Llama.LMADataSource
        Dim sppidExchanger As Llama.LMExchanger
        Dim sppidPipeRun As Llama.LMPipeRun
        Dim sppidItemTag As String
        Dim sppidPipingMatClass As String
        Dim symTypes As String = Nothing
        Dim connTypes As String = Nothing
        Dim matClasses As String = Nothing
        Dim linePattern As SmartSketch.LinearPattern
        Dim colorConst As SmartSketch.ColorConstants
        'Dim obj As Object
        Dim txtboxString As String = Nothing

        ' Get SPPID INI File
        Using file As New OpenFileDialog
            file.Title = "Open SmartPlant Manager INI File"
            file.InitialDirectory = "\\isihouprod\PTHouEngApps\SPPID\INI Files"
            file.Filter = "SmartPlant Manager INI File|*.ini"
            result = file.ShowDialog

            If result = Windows.Forms.DialogResult.OK Then
                sppidDatasource.SiteNode = file.FileName
            Else
                sppidDatasource = Nothing
                Me.Close()
            End If
            result = Nothing
        End Using

        Using file As New OpenFileDialog
            file.Title = "Open P&ID File"
            file.InitialDirectory = "\\US001SF0020\Technip USA\PT\HOU\DEPT\Engr-MethSys\02 Initiatives\23 PFDs in SPPID"
            file.Filter = "SmartSketch Document|*.igr|SmartPlant P&ID|*.pid|All Files|*.*"
            result = file.ShowDialog

            ' Get SmartSketch file name
            If result = Windows.Forms.DialogResult.OK Then

                ' Start SmartSketch Application
                objSmartSketch = CreateObject("SmartSketch.Application")

                ' Open SmartSketch file
                sketchDocument = objSmartSketch.Documents.Open(file.FileName)

                ' Code to manipulate data from SmartSketch and SPPID
                For index As Integer = 1 To sketchDocument.ActiveSheet.Symbols.Count
                    drawnSymbol = sketchDocument.ActiveSheet.Symbols.Item(index)
                    If Not drawnSymbol.IsAttributeSetPresent("P&IDAttributes") Then Continue For
                    symSPID = drawnSymbol.AttributeSets("P&IDAttributes").Item("ModelID").Value.ToString
                    symType = drawnSymbol.AttributeSets("P&IDAttributes").Item("ModelItemType").Value.ToString
                    sppidDatasource.ProjectNumber = drawnSymbol.AttributeSets("P&IDAttributes").Item("ProjectNumber").Value.ToString
                    If symSPID IsNot Nothing AndAlso symSPID <> "" Then
                        Select Case symType
                            Case "Exchanger"
                                Try
                                    sppidExchanger = sppidDatasource.GetExchanger(symSPID)
                                    sppidItemTag = sppidExchanger.Attributes("ItemTag").Value.ToString
                                    If sppidItemTag IsNot Nothing AndAlso sppidItemTag <> "" Then
                                        MsgBox(sppidItemTag)
                                    End If
                                Catch ex As Exception
                                    MsgBox(sppidDatasource.SiteNode)
                                    MsgBox(sppidDatasource.ProjectNumber)
                                Finally
                                    sppidExchanger = Nothing
                                End Try
                            Case Else
                        End Select
                    End If
                    ReleaseObject(drawnSymbol)
                    If InStr(symTypes, symType, vbTextCompare) > 0 Then Continue For
                    If symTypes Is Nothing Then
                        symTypes = symType
                    Else
                        symTypes = symTypes & ";" & symType
                    End If
                Next
                MsgBox(symTypes)
                symTypes = Nothing

                Dim attrNames As String = Nothing
                For index As Integer = 1 To sketchDocument.ActiveSheet.SmartSymbols2d.Count
                    drawnSmartSym = sketchDocument.ActiveSheet.SmartSymbols2d.Item(index)
                    If Not drawnSmartSym.IsAttributeSetPresent("P&IDAttributes") Then Continue For
                    For i = 1 To drawnSmartSym.AttributeSets("P&IDAttributes").Count
                        Dim pidAttr As SmartSketch.Attribute
                        Dim pidAttrName As String = Nothing
                        pidAttr = drawnSmartSym.AttributeSets("P&IDAttributes").Item(i)
                        pidAttrName = pidAttr.Name
                        If pidAttrName Is Nothing OrElse pidAttrName = "" Then Continue For
                        If InStr(attrNames, pidAttrName, vbTextCompare) > 0 Then Continue For
                        If attrNames Is Nothing Then
                            attrNames = pidAttrName
                        Else
                            attrNames = attrNames & ";" & pidAttrName
                        End If
                        pidAttrName = Nothing
                        ReleaseObject(pidAttr)
                    Next
                    'For ind As Integer = 1 To drawnSmartSym.DrawingObjects.Count
                    '    obj = drawnSmartSym.DrawingObjects.Item(ind)
                    '    symType = obj.Type.ToString
                    '    If symType IsNot Nothing Then
                    '        'If Not obj.IsAttributeSetPresent("P&IDAttributes") Then Continue For
                    '        Select Case CInt(symType)
                    '            Case SmartSketch.ObjectType.igTextBox
                    '                If obj.text Is Nothing OrElse obj.text = "" Then Continue For
                    '                If txtboxString Is Nothing Then
                    '                    txtboxString = obj.Text
                    '                Else
                    '                    txtboxString = txtboxString & "||" & obj.text
                    '                End If
                    '                '    Dim objAssigned As SmartSketch.TextBox
                    '                'Case SmartSketch.ObjectType.igLineString2d
                    '                '    Dim objAssigned As SmartSketch.LineString2d
                    '                'Case SmartSketch.ObjectType.igBoundary2d
                    '                '    Dim objAssigned As SmartSketch.Boundary2d
                    '                'Case SmartSketch.ObjectType.igLine2d
                    '                '    Dim objAssigned As SmartSketch.Line2d
                    '                'Case SmartSketch.ObjectType.igPoint2d
                    '                '    Dim objAssigned As SmartSketch.Point2d
                    '                'Case SmartSketch.ObjectType.igCircle2d
                    '                '    Dim objAssigned As SmartSketch.Circle2d
                    '                'Case SmartSketch.ObjectType.igArc2d
                    '                '    Dim objAssigned As SmartSketch.Arc2d
                    '        End Select
                    '    End If
                    '    If InStr(symTypes, symType, vbTextCompare) > 0 Then Continue For
                    '    If symTypes Is Nothing Then
                    '        symTypes = symType
                    '    Else
                    '        symTypes = symTypes & ";" & symType
                    '    End If
                    '    ReleaseObject(obj)
                    'Next
                    'If txtboxString IsNot Nothing AndAlso txtboxString <> "" Then
                    '    MsgBox(txtboxString)
                    'End If
                    'txtboxString = Nothing
                    'ReleaseObject(drawnSmartSym)
                Next
                'MsgBox(symTypes)
                MsgBox(attrNames)

                For index As Integer = 1 To sketchDocument.ActiveSheet.Connectors2d.Count
                    drawnConnector = sketchDocument.ActiveSheet.Connectors2d.Item(index)
                    If Not drawnConnector.IsAttributeSetPresent("P&IDAttributes") Then Continue For
                    connSPID = drawnConnector.AttributeSets("P&IDAttributes").Item("ModelID").Value.ToString
                    connType = drawnConnector.AttributeSets("P&IDAttributes").Item("ModelItemType").Value.ToString
                    linePattern = drawnConnector.LinearStyle.Pattern
                    sppidDatasource.ProjectNumber = drawnConnector.AttributeSets("P&IDAttributes").Item("ProjectNumber").Value.ToString

                    If connSPID IsNot Nothing AndAlso connSPID <> "" Then
                        Select Case connType
                            Case "PipeRun"
                                Try
                                    sppidPipeRun = sppidDatasource.GetPipeRun(connSPID)
                                    sppidItemTag = sppidPipeRun.Attributes("ItemTag").Value.ToString
                                    sppidPipingMatClass = sppidPipeRun.Attributes("PipingMaterialsClass").Value.ToString
                                    If sppidPipingMatClass IsNot Nothing AndAlso sppidPipingMatClass <> "" Then
                                        Select Case sppidPipingMatClass
                                            Case "A0B"
                                                colorConst = SmartSketch.ColorConstants.igBlueColor
                                                linePattern = sketchDocument.LinearPatterns.Item("Data Link")
                                            Case "A0H"
                                                colorConst = SmartSketch.ColorConstants.igMagentaColor
                                                linePattern = sketchDocument.LinearPatterns.Item("Pneumatic")
                                        End Select
                                        drawnConnector.LinearStyle.Color = colorConst
                                        drawnConnector.LinearStyle.Pattern = linePattern

                                        If InStr(matClasses, sppidPipingMatClass, vbTextCompare) = 0 Then
                                            If matClasses Is Nothing Then
                                                matClasses = sppidPipingMatClass
                                            Else
                                                matClasses = matClasses & ";" & sppidPipingMatClass
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    MsgBox(sppidDatasource.SiteNode)
                                    MsgBox(sppidDatasource.ProjectNumber)
                                Finally
                                    sppidPipeRun = Nothing
                                End Try

                            Case Else
                        End Select
                    End If
                    ReleaseObject(colorConst)
                    ReleaseObject(drawnConnector)
                    If InStr(connTypes, connType, vbTextCompare) > 0 Then Continue For
                    If connTypes Is Nothing Then
                        connTypes = connType
                    Else
                        connTypes = connTypes & ";" & connType
                    End If
                Next
                MsgBox(connTypes)
                MsgBox(matClasses)

                objSmartSketch.Visible = True
                'ReleaseObject(linePattern)
                'sketchDocument.Close(False)
                'objSmartSketch.Quit()
                'ReleaseObject(sketchDocument)
                'ReleaseObject(objSmartSketch)
            Else
                Me.Close()
            End If
        End Using
        sppidDatasource = Nothing
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class

' *********
' JUNK CODE
' *********
'Dim objSmartSketch As SmartSketch.Document
'Dim objLine1 As SmartSketch.Line2d
'Dim objLine2 As SmartSketch.Line2d

'objSmartSketch = GetObject(objFile.FileName, "SmartSketch.Document")
'objSmartSketch.Visible = True

'objSmartSketch.Documents.Add("SmartSketch.Document", "normal.igr")
'With objSmartSketch.ActiveDocument.ActiveSheet.Lines2d
'    objLine1 = .AddBy2Points(0, 0, 1, 1)
'    objLine2 = .AddBy2Points(1, 1, 2, 2)
'End With

' More code that manipulates objLin1 and objLine2

'objSmartSketch.ActiveDocument.SaveAs("X:\Test.dwg")
'objSmartSketch.ActiveDocument.Close(False)
'objLine1 = Nothing
'objLine2 = Nothing
'objSmartSketch = Nothing

'MsgBox(drawnSymbol.AttributeSets("P&IDAttributes").Item("ProjectNumber").Value.ToString)

'If sketchDocument.ActiveSheet.IsAttributeSetPresent("P&IDAttributes") Then
'    MsgBox(sketchDocument.ActiveSheet.AttributeSets("P&IDAttributes").Item("ProjectNumber").Value.ToString)
'End If

'For index As Integer = 1 To sketchDocument.LinearPatterns.Count
'    linePattern = sketchDocument.LinearPatterns.Item(index)
'    MsgBox(linePattern.InferredName & ";" & linePattern.StyleName)
'Next
