Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry


Public Class Class1
    <CommandMethod("AdskGreeting")>
    Public Sub AdskGreeting()
        '' Get the current document and database, and start a transaction
        ''Dim acDocc As Documentt = Autodesk.AutoCAD.ApplicationServices.Application
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table record for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                         OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                            OpenMode.ForWrite)

            '' Creates a new MText object and assigns it a location,
            '' text value and text style
            Using objText As MText = New MText

                '' Specify the insertion point of the MText object
                objText.Location = New Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0)

                '' Set the text string for the MText object
                objText.Contents = "Greetings, Welcome to AutoCAD .NET"

                '' Set the text style for the MText object
                objText.TextStyleId = acCurDb.Textstyle

                '' Appends the new MText object to model space
                acBlkTblRec.AppendEntity(objText)

                '' Appends to new MText object to the active transaction
                acTrans.AddNewlyCreatedDBObject(objText, True)
            End Using

            '' Saves the changes to the database and closes the transaction
            acTrans.Commit()
        End Using
    End Sub

    <CommandMethod("GaborGabor")>
    Public Sub Gabor()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf &
                                                                     "Enter your name: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)

        Application.ShowAlertDialog("The name entered was: " &
                                    pStrRes.StringResult)


    End Sub

    <CommandMethod("Gaborpoint")>
    Public Sub Gaborpoint()
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Enter the start point of the line: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Prompt for the end point
        pPtOpts.Message = vbLf & "Enter the end point of the line: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptEnd As Point3d = pPtRes.Value

        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acBlkTbl As BlockTable
            Dim acBlkTblRec As BlockTableRecord

            '' Open Model space for write
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                         OpenMode.ForRead)

            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                            OpenMode.ForWrite)

            '' Define the new line
            Using acLine As Line = New Line(ptStart, ptEnd)
                '' Add the line to the drawing
                acBlkTblRec.AppendEntity(acLine)
                acTrans.AddNewlyCreatedDBObject(acLine, True)
            End Using

            '' Zoom to the extents or limits of the drawing
            acDoc.SendStringToExecute("._zoom _all ", True, False, False)

            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using

        Dim ed As Editor = acDoc.Editor
        Dim length As Double
        length = ptStart.DistanceTo(ptEnd)
        ed.WriteMessage(vbLf & "Line length:" & System.Math.Round(length, 2))
        ed.WriteMessage(vbLf & "Line length:" & System.Math.Round(length, 5))
        ed.WriteMessage(vbLf & "Line length:" & System.Math.Round(length, 0))

    End Sub

    <CommandMethod("queryy")>
    Public Sub queryy()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            For Each acObjId As ObjectId In acBlkTblRec
                acDoc.Editor.WriteMessage(vbLf & "DXF name: " & acObjId.ObjectClass().DxfName)
                acDoc.Editor.WriteMessage(vbLf & "ObjectID: " & acObjId.ToString())
                acDoc.Editor.WriteMessage(vbLf & "Handle: " & acObjId.Handle.ToString())

            Next

        End Using

    End Sub

    <CommandMethod("addcircle")>
    Public Sub addcircle()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartOpenCloseTransaction()
        Using acTrans
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim acCirc As Circle = New Circle()
            Using acCirc
                acCirc.Center = New Point3d(5, 5, 0)
                acCirc.Radius = 3

                acBlkTblRec.AppendEntity(acCirc)
                acTrans.AddNewlyCreatedDBObject(acCirc, True)
            End Using

            acTrans.Commit()
            ''acDoc.Editor.WriteMessage(vbLf & "D: " & acCirc.Diameter)
        End Using

    End Sub

    <CommandMethod("SelectObjectOnscreen")>
    Public Sub Selectobject()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Dim acSSPrompt As PromptSelectionResult
            acSSPrompt = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet
                acSSet = acSSPrompt.Value

                Dim acSSObj As SelectedObject
                For Each acSSObj In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Entity
                        acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            acEnt.ColorIndex = 3

                        End If

                    End If
                Next
                acTrans.Commit()
            End If

        End Using

    End Sub
    Public Function GenzabTyp1Calc(ByVal data As Dictionary(Of String, Double), ByVal o As Double, ByVal uh As Double) As Dictionary(Of String, Double)
        Dim results As Dictionary(Of String, Double) = New Dictionary(Of String, Double)

        ''clearance of columns
        results.Add("zvvsv", o - 2 * (data.Item("oss") + data.Item("s")))

        ''axis
        results.Add("osLSy", 0)
        results.Add("osLVy", data.Item("v"))
        results.Add("osPSy", o * uh)
        results.Add("osPVy", results.Item("osPSy") + data.Item("v"))

        ''baseplate nad anchor
        results.Add("pposL", data.Item("ospp") + data.Item("pps") / 2)
        results.Add("podlLV", results.Item("pposL") * uh)
        results.Add("ppLS", results.Item("podlLV") + data.Item("podl"))
        results.Add("ppLV", results.Item("ppLS") + data.Item("pph"))
        results.Add("ppLL", data.Item("ospp"))
        results.Add("ppLP", results.Item("ppLL") + data.Item("pps"))
        results.Add("kL", data.Item("ospp") + data.Item("k"))
        results.Add("kLV", results.Item("ppLV"))

        results.Add("pposP", o - data.Item("ospp") - data.Item("pps") / 2)
        results.Add("podlPV", results.Item("pposP") * uh)
        results.Add("ppPS", results.Item("podlPV") + data.Item("podl"))
        results.Add("ppPV", results.Item("ppPS") + data.Item("pph"))
        results.Add("ppPL", o - data.Item("ospp") - data.Item("pps"))
        results.Add("ppPP", o - data.Item("ospp"))
        results.Add("kP", o - data.Item("ospp") - data.Item("k"))
        results.Add("kPV", results.Item("ppPV"))

        ''columns
        results.Add("sLLx", data.Item("oss"))
        results.Add("sLPx", data.Item("oss") + data.Item("s"))
        results.Add("sLSy", results.Item("ppLV"))
        results.Add("sLLy", data.Item("oss") * uh + data.Item("v"))
        results.Add("sLPy", (data.Item("oss") + data.Item("s")) * uh + data.Item("v"))

        results.Add("sPPx", o - data.Item("oss"))
        results.Add("sPLx", o - data.Item("oss") - data.Item("s"))
        results.Add("sPSy", results.Item("ppPV"))
        results.Add("sPPy", (o - data.Item("oss")) * uh + data.Item("v"))
        results.Add("sPLy", (o - data.Item("oss") - data.Item("s")) * uh + data.Item("v"))

        ''grab hanlde
        results.Add("mLVx", data.Item("oss"))
        results.Add("mLVy", results.Item("sLLy"))
        results.Add("mPVx", o - data.Item("oss"))
        results.Add("mPVy", results.Item("sPLy"))
        results.Add("mLSx", data.Item("oss") + data.Item("s"))
        results.Add("mLSy", results.Item("sLLy") - data.Item("mh"))
        results.Add("mPSx", o - data.Item("oss") - data.Item("s"))
        results.Add("mPSy", results.Item("sPLy") - data.Item("mh"))

        '' horizonal infill bottom
        results.Add("vvdLSx", data.Item("oss") + data.Item("s"))
        results.Add("vvdLSy", (data.Item("oss") + data.Item("s")) * uh + data.Item("rhv"))
        results.Add("vvdLVx", results.Item("vvdLSx"))
        results.Add("vvdLVy", (data.Item("oss") + data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd"))

        results.Add("vvdPSx", o - data.Item("oss") - data.Item("s"))
        results.Add("vvdPSy", (o - data.Item("oss") - data.Item("s")) * uh + data.Item("rhv"))
        results.Add("vvdPVx", results.Item("vvdPSx"))
        results.Add("vvdPVy", (o - data.Item("oss") - data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd"))

        '' horizonal infill top
        results.Add("vvhLSx", data.Item("oss") + data.Item("s"))
        results.Add("vvhLSy", (data.Item("oss") + data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd") + data.Item("hvsv"))
        results.Add("vvhLVx", results.Item("vvhLSx"))
        results.Add("vvhLVy", (data.Item("oss") + data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd") + data.Item("hvsv") + +data.Item("hvh"))

        results.Add("vvhPSx", o - data.Item("oss") - data.Item("s"))
        results.Add("vvhPSy", (o - data.Item("oss") - data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd") + data.Item("hvsv"))
        results.Add("vvhPVx", results.Item("vvhPSx"))
        results.Add("vvhPVy", (o - data.Item("oss") - data.Item("s")) * uh + data.Item("rhv") + data.Item("hvd") + data.Item("hvsv") + +data.Item("hvh"))

        ''vertical infill data
        results.Add("minPocet", ((results.Item("zvvsv") - data.Item("omin")) / (data.Item("zvv") + data.Item("omin"))) + 1)
        results.Add("pocet", Math.Ceiling(results.Item("minPocet")))
        results.Add("zvvo", (results.Item("zvvsv") - (results.Item("pocet") - 1) * data.Item("zvv")) / results.Item("pocet"))

        ''vertical infill first line
        results.Add("zvv_x", (data.Item("oss") + data.Item("s") + results.Item("zvvo")))
        results.Add("zvvSy", (results.Item("zvv_x") * uh + data.Item("rhv") + data.Item("hvd")))
        results.Add("zvvVy", (results.Item("zvv_x") * uh + data.Item("rhv") + data.Item("hvd") + +data.Item("hvsv")))

        ''vertical infill second line
        results.Add("zvv2_x", (data.Item("oss") + data.Item("s") + results.Item("zvvo") + data.Item("zvv")))
        results.Add("zvv2Sy", (results.Item("zvv2_x") * uh + data.Item("rhv") + data.Item("hvd")))
        results.Add("zvv2Vy", (results.Item("zvv2_x") * uh + data.Item("rhv") + data.Item("hvd") + +data.Item("hvsv")))



        Return results

    End Function
    Public Function GenzabTyp1CalcCoords(ByVal results As Dictionary(Of String, Double), ByVal ptStart As Point3d, ByVal ptEnd As Point3d, ByVal o As Double)
        Dim points As Dictionary(Of Integer, Point2d) = New Dictionary(Of Integer, Point2d)
        Dim tmpPoint = New Point2d()
        '' axis Left bottom:1
        tmpPoint = New Point2d(ptStart.X, ptStart.Y)
        points.Add(1, tmpPoint)
        '' axis Left top:2
        tmpPoint = New Point2d(ptStart.X, ptStart.Y + results.Item("osLVy"))
        points.Add(2, tmpPoint)
        '' axis Right bottom:3
        tmpPoint = New Point2d(ptStart.X + o, ptStart.Y + results.Item("osPSy"))
        points.Add(3, tmpPoint)
        '' axis Right top:4
        tmpPoint = New Point2d(ptStart.X + o, ptStart.Y + results.Item("osPVy"))
        points.Add(4, tmpPoint)

        ''base plate left
        '' point 10
        tmpPoint = New Point2d(ptStart.X + results.Item("ppLL"), ptStart.Y + results.Item("ppLS"))
        points.Add(10, tmpPoint)
        '' point 11
        tmpPoint = New Point2d(ptStart.X + results.Item("ppLP"), ptStart.Y + results.Item("ppLV"))
        points.Add(11, tmpPoint)

        ''column left
        '' point 12
        tmpPoint = New Point2d(ptStart.X + results.Item("sLLx"), ptStart.Y + results.Item("sLSy"))
        points.Add(12, tmpPoint)
        '' point 13
        tmpPoint = New Point2d(ptStart.X + results.Item("sLLx"), ptStart.Y + results.Item("sLLy"))
        points.Add(13, tmpPoint)
        '' point 14
        tmpPoint = New Point2d(ptStart.X + results.Item("sLPx"), ptStart.Y + results.Item("sLSy"))
        points.Add(14, tmpPoint)
        '' point 15
        tmpPoint = New Point2d(ptStart.X + results.Item("sLPx"), ptStart.Y + results.Item("sLPy"))
        points.Add(15, tmpPoint)

        ''base plate right
        '' point 16
        tmpPoint = New Point2d(ptStart.X + results.Item("ppPL"), ptStart.Y + results.Item("ppPS"))
        points.Add(16, tmpPoint)
        '' point 17
        tmpPoint = New Point2d(ptStart.X + results.Item("ppPP"), ptStart.Y + results.Item("ppPV"))
        points.Add(17, tmpPoint)

        ''column right
        '' point 18
        tmpPoint = New Point2d(ptStart.X + results.Item("sPLx"), ptStart.Y + results.Item("sPSy"))
        points.Add(18, tmpPoint)
        '' point 19
        tmpPoint = New Point2d(ptStart.X + results.Item("sPLx"), ptStart.Y + results.Item("sPLy"))
        points.Add(19, tmpPoint)
        '' point 20
        tmpPoint = New Point2d(ptStart.X + results.Item("sPPx"), ptStart.Y + results.Item("sPSy"))
        points.Add(20, tmpPoint)
        '' point 21
        tmpPoint = New Point2d(ptStart.X + results.Item("sPPx"), ptStart.Y + results.Item("sPPy"))
        points.Add(21, tmpPoint)

        '' anchor insert points for block
        '' point 22
        tmpPoint = New Point2d(ptStart.X + results.Item("kL"), ptStart.Y + results.Item("kLV"))
        points.Add(22, tmpPoint)
        '' point 23
        tmpPoint = New Point2d(ptStart.X + results.Item("kP"), ptStart.Y + results.Item("kPV"))
        points.Add(23, tmpPoint)

        ''grab hanlde bottom
        '' point 24
        tmpPoint = New Point2d(ptStart.X + results.Item("mLSx"), ptStart.Y + results.Item("mLSy"))
        points.Add(24, tmpPoint)
        '' point 25
        tmpPoint = New Point2d(ptStart.X + results.Item("mPSx"), ptStart.Y + results.Item("mPSy"))
        points.Add(25, tmpPoint)

        '' horizontal infill top
        '' point 26
        tmpPoint = New Point2d(ptStart.X + results.Item("vvhLVx"), ptStart.Y + results.Item("vvhLVy"))
        points.Add(26, tmpPoint)
        '' point 27
        tmpPoint = New Point2d(ptStart.X + results.Item("vvhPVx"), ptStart.Y + results.Item("vvhPVy"))
        points.Add(27, tmpPoint)
        '' point 28
        tmpPoint = New Point2d(ptStart.X + results.Item("vvhLSx"), ptStart.Y + results.Item("vvhLSy"))
        points.Add(28, tmpPoint)
        '' point 29
        tmpPoint = New Point2d(ptStart.X + results.Item("vvhPSx"), ptStart.Y + results.Item("vvhPSy"))
        points.Add(29, tmpPoint)

        '' horizontal infill bottom
        '' point 30
        tmpPoint = New Point2d(ptStart.X + results.Item("vvdLVx"), ptStart.Y + results.Item("vvdLVy"))
        points.Add(30, tmpPoint)
        '' point 31
        tmpPoint = New Point2d(ptStart.X + results.Item("vvdPVx"), ptStart.Y + results.Item("vvdPVy"))
        points.Add(31, tmpPoint)
        '' point 32
        tmpPoint = New Point2d(ptStart.X + results.Item("vvdLSx"), ptStart.Y + results.Item("vvdLSy"))
        points.Add(32, tmpPoint)
        '' point 33
        tmpPoint = New Point2d(ptStart.X + results.Item("vvdPSx"), ptStart.Y + results.Item("vvdPSy"))
        points.Add(33, tmpPoint)

        '' vertical infill - first line
        '' point 34
        tmpPoint = New Point2d(ptStart.X + results.Item("zvv_x"), ptStart.Y + results.Item("zvvSy"))
        points.Add(34, tmpPoint)
        '' point 35
        tmpPoint = New Point2d(ptStart.X + results.Item("zvv_x"), ptStart.Y + results.Item("zvvVy"))
        points.Add(35, tmpPoint)

        '' vertical infill - second line
        '' point 36
        tmpPoint = New Point2d(ptStart.X + results.Item("zvv2_x"), ptStart.Y + results.Item("zvv2Sy"))
        points.Add(36, tmpPoint)
        '' point 37
        tmpPoint = New Point2d(ptStart.X + results.Item("zvv2_x"), ptStart.Y + results.Item("zvv2Vy"))
        points.Add(37, tmpPoint)


        Return points

    End Function

    <CommandMethod("GenZabTyp1")>
    Public Sub GenZabTyp1()
        '' dictionary for store block data
        Dim dynBlkDimensions As Dictionary(Of String, Double) = New Dictionary(Of String, Double)
        Dim dimensions As Dictionary(Of String, Double) = New Dictionary(Of String, Double)
        '' input data
        '' change it withinterface input
        dimensions.Add("v", 1.13)
        dimensions.Add("ospp", 0.032)
        dimensions.Add("k", 0.08)
        dimensions.Add("pps", 0.11)
        dimensions.Add("podl", 0.01)
        dimensions.Add("pph", 0.02)
        dimensions.Add("rhv", 0.144)
        dimensions.Add("hvd", 0.01)
        dimensions.Add("hvsv", 0.806)
        dimensions.Add("hvh", 0.01)
        dimensions.Add("mh", 0.05)
        dimensions.Add("oss", 0.057)
        dimensions.Add("s", 0.016)
        dimensions.Add("zvv", 0.01)
        dimensions.Add("omin", 0.12)

        Dim blockname As String
        blockname = "kotva"



        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Enter the start point of the line: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Prompt for the end point
        pPtOpts.Message = vbLf & "Enter the end point of the line: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptEnd As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' check if inserted points is in right direction from left to right, if not switch the two points
        If ptStart.X > ptEnd.X Then
            Dim ptStart_temp As Point3d
            Dim ptEnd_temp As Point3d
            ptStart_temp = ptStart
            ptEnd_temp = ptEnd
            ptStart = ptEnd_temp
            ptEnd = ptStart_temp
        End If

        '' calculate angle
        ''Dim angle As Double = Math.Atan((ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X))
        Dim angle As Double = (ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X)
        acDoc.Editor.WriteMessage(vbLf & "y/x: " & angle)
        If angle > 0.12 Then
            acDoc.Editor.WriteMessage(vbLf & "sklon > 0.12")
        End If

        '' calculate length
        Dim length As Double = Math.Abs(ptEnd.X - ptStart.X)
        acDoc.Editor.WriteMessage(vbLf & "length: " & length)
        If length < (dimensions.Item("ospp") + dimensions.Item("pps")) * 2 Then
            acDoc.Editor.WriteMessage(vbLf & "nedostatocna dlzka pre vykreslenie zabradlia. Min. dlzka musi byt viac ako: " & (dimensions.Item("ospp") + dimensions.Item("pps")) * 2 & "m. Koniec")
            Exit Sub
        End If
        ''TODO: check if length is enough to draw the railing

        ''calc data
        Dim results As Object
        results = GenzabTyp1Calc(dimensions, length, angle)
        ''results = GenzabTyp1Calc(dimensions, length, angle)

        '' calc point coordinates
        Dim calcCoords As Object
        calcCoords = GenzabTyp1CalcCoords(results, ptStart, ptEnd, length)


        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Dim acBlkTbl As BlockTable
            Dim acBlkTblRec As BlockTableRecord
            Dim blkRecID As ObjectId = ObjectId.Null

            '' Open Model space for write
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            '' insert anchor blocks
            ''check if the block is present
            If Not acBlkTbl.Has(blockname) Then
                acDoc.Editor.WriteMessage(vbLf & "block: " & blockname & " nie je pritomny vo vykrese, Koniec")
                Exit Sub
            Else
                blkRecID = acBlkTbl(blockname)
            End If
            '' insert blocks
            If blkRecID <> ObjectId.Null Then
                Dim tmpPoint = New Point3d(calcCoords.Item(22).X, calcCoords.Item(22).Y, ptStart.Z)
                Dim acBlkRefSt As New BlockReference(tmpPoint, blkRecID)
                Using acBlkRefSt
                    Dim acCurSpaceBlkTblRec As BlockTableRecord
                    acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                    acCurSpaceBlkTblRec.AppendEntity(acBlkRefSt)
                    acTrans.AddNewlyCreatedDBObject(acBlkRefSt, True)
                End Using
                tmpPoint = New Point3d(calcCoords.Item(23).X, calcCoords.Item(23).Y, ptStart.Z)
                Dim acBlkRefEnd As New BlockReference(tmpPoint, blkRecID)
                Using acBlkRefEnd
                    Dim acCurSpaceBlkTblRec As BlockTableRecord
                    acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                    acCurSpaceBlkTblRec.AppendEntity(acBlkRefEnd)
                    acTrans.AddNewlyCreatedDBObject(acBlkRefEnd, True)
                End Using


                '' read block reference for start point
                ''Dim acBlkRefSt As New BlockReference(ptStart, blkRecID)

                ''    '' read dynamic block parameters
                ''    Dim pc As DynamicBlockReferencePropertyCollection
                ''    pc = acBlkRefSt.DynamicBlockReferencePropertyCollection
                ''    Dim DBRP As DynamicBlockReferenceProperty
                ''    For Each DBRP In pc
                ''        Try
                ''            acDoc.Editor.WriteMessage(vbLf & "parametername: " & DBRP.PropertyName & " value: " & DBRP.Value)
                ''            '' store dynamic block parameters to local variable - dictionary
                ''            dynBlkDimensions.Add(DBRP.PropertyName, DBRP.Value)
                ''        Catch
                ''            acDoc.Editor.WriteMessage(vbLf & "parametername: " & DBRP.PropertyName)
                ''        End Try
                ''    Next

                ''    ' insert block to start point
                ''    Using acBlkRefSt
                ''        Dim acCurSpaceBlkTblRec As BlockTableRecord
                ''        acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                ''        acCurSpaceBlkTblRec.AppendEntity(acBlkRefSt)
                ''        acTrans.AddNewlyCreatedDBObject(acBlkRefSt, True)
                ''    End Using
                ''    ' insert block to end point
                ''    Dim acBlkRefEnd As New BlockReference(ptEnd, blkRecID)
                ''    Dim tmpScale As Scale3d = New Scale3d(-1, 1, 1)
                ''    acBlkRefEnd.ScaleFactors = tmpScale
                ''    Using acBlkRefEnd
                ''        Dim acCurSpaceBlkTblRec As BlockTableRecord
                ''        acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                ''        acCurSpaceBlkTblRec.AppendEntity(acBlkRefEnd)
                ''        acTrans.AddNewlyCreatedDBObject(acBlkRefEnd, True)
                ''    End Using

            End If



            '' Define the new line
            '' base line
            CreateLine(ptStart, ptEnd, acBlkTblRec, acTrans)
            '' axis Left
            Dim tmpPoint1 = New Point3d(calcCoords.Item(1).X, calcCoords.Item(1).Y, ptStart.Z)
            Dim tmpPoint2 = New Point3d(calcCoords.Item(2).X, calcCoords.Item(2).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            '' axis Right
            tmpPoint1 = New Point3d(calcCoords.Item(3).X, calcCoords.Item(3).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(4).X, calcCoords.Item(4).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            '' base plate left
            tmpPoint1 = New Point3d(calcCoords.Item(10).X, calcCoords.Item(10).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(11).X, calcCoords.Item(11).Y, ptStart.Z)
            CreateRectangle(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            ''left column left 
            tmpPoint1 = New Point3d(calcCoords.Item(12).X, calcCoords.Item(12).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(13).X, calcCoords.Item(13).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            ''left column rigth 
            tmpPoint1 = New Point3d(calcCoords.Item(14).X, calcCoords.Item(14).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(15).X, calcCoords.Item(15).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            '' base plate right
            tmpPoint1 = New Point3d(calcCoords.Item(16).X, calcCoords.Item(16).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(17).X, calcCoords.Item(17).Y, ptStart.Z)
            CreateRectangle(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            ''right column right 
            tmpPoint1 = New Point3d(calcCoords.Item(18).X, calcCoords.Item(18).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(19).X, calcCoords.Item(19).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            ''right column left 
            tmpPoint1 = New Point3d(calcCoords.Item(20).X, calcCoords.Item(20).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(21).X, calcCoords.Item(21).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''grab hanlde top
            tmpPoint1 = New Point3d(calcCoords.Item(13).X, calcCoords.Item(13).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(21).X, calcCoords.Item(21).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''grab hanlde bottom
            tmpPoint1 = New Point3d(calcCoords.Item(24).X, calcCoords.Item(24).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(25).X, calcCoords.Item(25).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''horizontal infill top - top
            tmpPoint1 = New Point3d(calcCoords.Item(26).X, calcCoords.Item(26).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(27).X, calcCoords.Item(27).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''horizontal infill top - bottom
            tmpPoint1 = New Point3d(calcCoords.Item(28).X, calcCoords.Item(28).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(29).X, calcCoords.Item(29).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''horizontal infill top - top
            tmpPoint1 = New Point3d(calcCoords.Item(30).X, calcCoords.Item(30).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(31).X, calcCoords.Item(31).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''horizontal infill top - bottom
            tmpPoint1 = New Point3d(calcCoords.Item(32).X, calcCoords.Item(32).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(33).X, calcCoords.Item(33).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            ''vertical infill-first line
            tmpPoint1 = New Point3d(calcCoords.Item(34).X, calcCoords.Item(34).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(35).X, calcCoords.Item(35).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            For i As Integer = 1 To results.Item("pocet") - 2
                tmpPoint1 = New Point3d(calcCoords.Item(34).X + i * (results.Item("zvvo") + dimensions.Item("zvv")), calcCoords.Item(34).Y + angle * i * (results.Item("zvvo") + dimensions.Item("zvv")), ptStart.Z)
                tmpPoint2 = New Point3d(calcCoords.Item(35).X + i * (results.Item("zvvo") + dimensions.Item("zvv")), calcCoords.Item(35).Y + angle * i * (results.Item("zvvo") + dimensions.Item("zvv")), ptStart.Z)
                CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            Next

            ''vertical infill-second line
            tmpPoint1 = New Point3d(calcCoords.Item(36).X, calcCoords.Item(36).Y, ptStart.Z)
            tmpPoint2 = New Point3d(calcCoords.Item(37).X, calcCoords.Item(37).Y, ptStart.Z)
            CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)

            For i As Integer = 1 To results.Item("pocet") - 2
                tmpPoint1 = New Point3d(calcCoords.Item(36).X + i * (results.Item("zvvo") + dimensions.Item("zvv")), calcCoords.Item(36).Y + angle * i * (results.Item("zvvo") + dimensions.Item("zvv")), ptStart.Z)
                tmpPoint2 = New Point3d(calcCoords.Item(37).X + i * (results.Item("zvvo") + dimensions.Item("zvv")), calcCoords.Item(37).Y + angle * i * (results.Item("zvvo") + dimensions.Item("zvv")), ptStart.Z)
                CreateLine(tmpPoint1, tmpPoint2, acBlkTblRec, acTrans)
            Next

            acTrans.Commit()

        End Using
    End Sub
    Private Sub CreateLine(ptStart, ptEnd, acBlkTblRec, acTrans)
        Using acLine As Line = New Line(ptStart, ptEnd)
            '' Add the line to the drawing
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
        End Using
    End Sub
    Private Sub CreateRectangle(leftBottomPoint, rightTopPoint, acBlkTblRec, acTrans)
        '' using left bottom point of rectangele and right top point of rectangel to create it
        Dim dx As Double
        Dim dy As Double
        dx = rightTopPoint.X - leftBottomPoint.X
        dy = rightTopPoint.Y - leftBottomPoint.Y
        Using acPoly As Polyline = New Polyline()
            acPoly.AddVertexAt(0, New Point2d(leftBottomPoint.X, leftBottomPoint.Y), 0, 0, 0)
            acPoly.AddVertexAt(1, New Point2d(leftBottomPoint.X + dx, leftBottomPoint.Y), 0, 0, 0)
            acPoly.AddVertexAt(2, New Point2d(leftBottomPoint.X + dx, leftBottomPoint.Y + dy), 0, 0, 0)
            acPoly.AddVertexAt(3, New Point2d(leftBottomPoint.X, leftBottomPoint.Y + dy), 0, 0, 0)
            acPoly.Closed = True
            '' Add the polyline to the drawing
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
        End Using
    End Sub
End Class