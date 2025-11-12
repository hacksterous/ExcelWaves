Sub ShrinkTextToFitShape(s As Shape)
	Dim fs As Single
	fs = 14 'Start with a large font size

	With s.TextFrame2
		.AutoSize = msoAutoSizeNone
		.HorizontalAnchor = msoAnchorCenter
		.VerticalAnchor = msoAnchorMiddle
		.WordWrap = msoTrue
	End With

	Do While fs > 6
		s.TextFrame2.TextRange.Font.Size = fs
		If s.TextFrame2.TextRange.BoundWidth <= s.Width And _
		   s.TextFrame2.TextRange.BoundHeight <= s.Height Then
			Exit Do
		End If
		fs = fs - 1
	Loop
End Sub

Function addLine(ws As Worksheet, x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Shape
	Dim shp As Shape
	Set shp = ws.Shapes.addLine(x1, y1, x2, y2)
	shp.Name = "VBA_" & ActiveSheet.Shapes.Count
	Set addLine = shp
End Function

Sub DrawPolygonFromPoints(ByRef polygonShape As Shape, _
	ByVal x1 As Single, ByVal y1 As Single, _
	x2 As Single, y2 As Single, _
	x3 As Single, y3 As Single, _
	x4 As Single, y4 As Single, _
	x5 As Single, y5 As Single, _
	ByVal x6 As Single, ByVal y6 As Single, _
	label As String, leftMargin As Single, _
	doingD As Boolean)

	Dim ws As Worksheet
	Set ws = ActiveSheet

	'Create a freeform shape starting at the first point
	Dim shapeBuilder As FreeformBuilder
	Set shapeBuilder = ws.Shapes.BuildFreeform(msoEditingAuto, x1, y1)

	'Add remaining points
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x2, y2
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x3, y3
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x4, y4
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x5, y5
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x6, y6
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x1, y1

	'Convert to shape and fill
	Set polygonShape = shapeBuilder.ConvertToShape

	With polygonShape.Fill
		.Visible = msoTrue
		.Solid
	End With
	
	polygonShape.Line.Weight = 1.5
	
	If doingD Then
		polygonShape.Fill.ForeColor.RGB = RGB(226, 243, 250) 'Light blue fill
		polygonShape.Fill.Transparency = 0.5 '0 = opaque, 1 = fully transparent
		polygonShape.Line.ForeColor.RGB = RGB(21, 96, 130)
	Else
		polygonShape.Fill.ForeColor.RGB = RGB(66, 66, 66) 'Light gray fill
		polygonShape.Fill.Patterned msoPatternDownwardDiagonal
		polygonShape.Fill.Transparency = 0 '0 = opaque, 1 = fully transparent
		polygonShape.Line.ForeColor.RGB = RGB(66, 66, 66)
	End If
	
	polygonShape.Name = "VBA_" & ActiveSheet.Shapes.Count
	
	With polygonShape.TextFrame2
		.TextRange.Text = label
		.HorizontalAnchor = msoAnchorCenter
		.VerticalAnchor = msoAnchorMiddle
		.MarginTop = 30
		.MarginBottom = 0
		.MarginLeft = leftMargin
		.MarginRight = 0
	End With
	
	With polygonShape.TextFrame2.TextRange.Font
		.Name = "Cambria"
		.Size = 14
		.Bold = True
		.Fill.ForeColor.RGB = RGB(0, 0, 0)
	End With

	polygonShape.ZOrder msoSendToBack

	ShrinkTextToFitShape polygonShape
End Sub

Sub DrawGap(ByRef polygonShape As Shape, _
	ByVal x1 As Single, ByVal y1 As Single, _
	x2 As Single, y2 As Single, _
	x3 As Single, y3 As Single, _
	x4 As Single, y4 As Single, _
	x5 As Single, y5 As Single, _
	ByVal x6 As Single, ByVal y6 As Single)

	Dim ws As Worksheet
	Set ws = ActiveSheet

	'Create a freeform shape starting at the first point
	Dim shapeBuilder As FreeformBuilder
	Set shapeBuilder = ws.Shapes.BuildFreeform(msoEditingAuto, x1, y1)

	'Add remaining points
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x2, y2
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x3, y3
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x4, y4
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x5, y5
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x6, y6
	shapeBuilder.AddNodes msoSegmentLine, msoEditingAuto, x1, y1

	'Convert to shape and fill
	Set polygonShape = shapeBuilder.ConvertToShape

	With polygonShape.Fill
		.Visible = msoTrue
		.Solid
	End With
	
	polygonShape.Line.Weight = 0

	polygonShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
	polygonShape.Fill.Transparency = 0 '0 = opaque, 1 = fully transparent
	polygonShape.Line.ForeColor.RGB = RGB(255, 255, 255)

	Dim edge As Shape

	Set edge = addLine(ws, x1 - 2, y1, x2 - 2, y2)
	edge.Line.ForeColor.RGB = RGB(180, 180, 180)
	edge.Line.Weight = 1

	Set edge = addLine(ws, x6 - 2, y6, x1 - 2, y1)
	edge.Line.ForeColor.RGB = RGB(180, 180, 180)
	edge.Line.Weight = 1

	Set edge = addLine(ws, x3 + 1, y3, x4 + 1, y4)
	edge.Line.ForeColor.RGB = RGB(180, 180, 180)
	edge.Line.Weight = 0.5

	Set edge = addLine(ws, x4 + 1, y4, x5 + 1, y5)
	edge.Line.ForeColor.RGB = RGB(180, 180, 180)
	edge.Line.Weight = 0.5

	polygonShape.Name = "VBA_" & ActiveSheet.Shapes.Count

	polygonShape.ZOrder msoSendToFront
End Sub

Sub DrawWaves()
	Dim ws As Worksheet
	Set ws = ActiveSheet

	Dim r As Integer, c As Integer
	Dim cell As Range, targetCell As Range
	Dim leftPos As Single, bottomPos As Single
	Dim cellWidth As Single, cellHeight As Single
	Dim clkToQDelay As Integer
	Dim firstChar As String, secondChar As String

	Dim x1 As Single, y1 As Single
	Dim x2 As Single, y2 As Single
	Dim x3 As Single, y3 As Single
	Dim x4 As Single, y4 As Single
	Dim x5 As Single, y5 As Single
	Dim x6 As Single, y6 As Single

	Dim polygonShape As Shape

	Dim colLetter As String
	Dim targetColLetter As String

	clkToQDelay = 4

	Columns("C:ZZ").ColumnWidth = 6
	Rows("1:100").RowHeight = 30

	Dim shp As Shape
	'Count = 0
	For Each shp In ws.Shapes
		If shp.Name Like "VBA_*" Then
			shp.Delete
			'Count = Count + 1
		End If
	Next shp
	'MsgBox Count & " shapes deleted."

	Dim row As Integer
	row = 1
	' Loop through rows 1 to 100 and columns A to Z
	For r = 1 To 100
		For c = 1 To 200

			If c < 27 Then
				colLetter = Chr(64 + c)
			Else
				firstChar = Chr(64 + (c \ 26))
				secondChar = Chr(65 + (c Mod 26))
				colLetter = firstChar & secondChar
			End If
	
			Set cell = ws.Range(colLetter & r)
			Set cellBelow = ws.Range(colLetter & (r + 1)) 'for labels
	
			If colLetter = "A" Then
				row = cell.value
			Else
				If cell.value > Chr(31) And cell.value < Chr(127) And c >= 3 And Left$(cell.value, 1) <> "#" Then
					If c < 27 Then
						targetColLetter = Chr(64 + c)
					Else
						firstChar = Chr(64 + (c \ 26))
						secondChar = Chr(65 + (c Mod 26))
						targetColLetter = firstChar & secondChar
					End If
					If row <= 100 And row > 0 Then
						Set targetCell = ws.Range(targetColLetter & row + 1)
						leftPos = targetCell.Left
						bottomPos = targetCell.Top
						cellWidth = targetCell.Width
						cellHeight = targetCell.Height

						If cell.value = "zh" Then
							addLine ws, leftPos, bottomPos - cellHeight / 2, leftPos + clkToQDelay, bottomPos - cellHeight
							addLine ws, leftPos + clkToQDelay, bottomPos - cellHeight, leftPos + cellWidth, bottomPos - cellHeight
						ElseIf cell.value = "zl" Then
							addLine ws, leftPos, bottomPos - cellHeight / 2, leftPos + clkToQDelay, bottomPos
							addLine ws, leftPos + clkToQDelay, bottomPos, leftPos + cellWidth, bottomPos
						ElseIf cell.value = "lz" Then
							addLine ws, leftPos, bottomPos, leftPos + clkToQDelay, bottomPos - cellHeight / 2
							addLine ws, leftPos + clkToQDelay, bottomPos - cellHeight / 2, leftPos + cellWidth, bottomPos - cellHeight / 2
						ElseIf cell.value = "hz" Then
							addLine ws, leftPos, bottomPos - cellHeight, leftPos + clkToQDelay, bottomPos - cellHeight / 2
							addLine ws, leftPos + clkToQDelay, bottomPos - cellHeight / 2, leftPos + cellWidth, bottomPos - cellHeight / 2
						ElseIf cell.value = "l" Then
							addLine ws, leftPos, bottomPos, leftPos + cellWidth, bottomPos
						ElseIf cell.value = "h" Then
							addLine ws, leftPos, bottomPos - cellHeight, leftPos + cellWidth, bottomPos - cellHeight
						ElseIf cell.value = "c" Then
							addLine ws, leftPos, bottomPos - cellHeight, leftPos + cellWidth, bottomPos - cellHeight
							addLine ws, leftPos, bottomPos, leftPos, bottomPos - cellHeight
						ElseIf cell.value = "k" Then
							addLine ws, leftPos, bottomPos - cellHeight, leftPos, bottomPos
							addLine ws, leftPos, bottomPos, leftPos + cellWidth, bottomPos
						ElseIf cell.value = "r" Then
							addLine ws, leftPos, bottomPos, leftPos + clkToQDelay, bottomPos - cellHeight 'up
							addLine ws, leftPos + clkToQDelay, bottomPos - cellHeight, leftPos + cellWidth, bottomPos - cellHeight 'across
						ElseIf cell.value = "f" Then
							addLine ws, leftPos, bottomPos - cellHeight, leftPos + clkToQDelay, bottomPos
							addLine ws, leftPos + clkToQDelay, bottomPos, leftPos + cellWidth, bottomPos
						ElseIf cell.value = "z" Then
							addLine ws, leftPos, bottomPos - cellHeight / 2, leftPos + cellWidth, bottomPos - cellHeight / 2
						ElseIf cell.value = "D" Or cell.value = "X" Then
							Dim doingD As Boolean
							If cell.value = "D" Then
								doingD = True
							Else
								doingD = False
							End If
							dataCount = 1
							localCount = c
						
							Do
								localCount = localCount + 1
								If localCount < 27 Then
									colLetter = Chr(64 + localCount)
								Else
									firstChar = Chr(64 + (localCount \ 26))
									secondChar = Chr(65 + (localCount Mod 26))
									colLetter = firstChar & secondChar
								End If
								Set cell = ws.Range(colLetter & r)
								If (cell.value <> "d" And doingD) Or (cell.value <> "x" And Not doingD) Then
									Exit Do
								End If
								dataCount = dataCount + 1
							Loop

							leftPos = targetCell.Left
							bottomPos = targetCell.Top
							cellWidth = targetCell.Width
							cellHeight = targetCell.Height
							
							'MsgBox "dataCount = " & dataCount
							
							x1 = leftPos
							y1 = bottomPos - cellHeight / 2
							x2 = leftPos + clkToQDelay
							y2 = bottomPos - cellHeight

							x3 = leftPos + dataCount * cellWidth - clkToQDelay
							y3 = bottomPos - cellHeight
							x4 = x3 + clkToQDelay
							y4 = bottomPos - cellHeight / 2
							x5 = x3
							y5 = bottomPos
							x6 = leftPos + clkToQDelay
							y6 = bottomPos
							'x4 - x1 is the width of the data
							Dim label As String
							label = ""
							If cellBelow.value <> "//" Then
								label = cellBelow.value
							End If
							Call DrawPolygonFromPoints(polygonShape, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6, label, x4 - x1, doingD)
							c = c + dataCount - 1
						End If
					End If 'row
				End If 'cell.Value
			End If 'colLetter
		Next c
	Next r

	'repeat for gap
	row = 1
	' Loop through rows 1 to 100 and columns A to Z
	For r = 1 To 100
		For c = 1 To 200
			If c < 27 Then
				colLetter = Chr(64 + c)
			Else
				firstChar = Chr(64 + (c \ 26))
				secondChar = Chr(65 + (c Mod 26))
				colLetter = firstChar & secondChar
			End If
	
			Set cell = ws.Range(colLetter & r)
			Set cellBelow = ws.Range(colLetter & (r + 1)) 'for labels
	
			If colLetter = "A" Then
				row = cell.value
			Else
				If cell.value > Chr(31) And cell.value < Chr(127) And c >= 3 And Left$(cell.value, 1) <> "#" Then
					If c < 27 Then
						targetColLetter = Chr(64 + c)
					Else
						firstChar = Chr(64 + (c \ 26))
						secondChar = Chr(65 + (c Mod 26))
						targetColLetter = firstChar & secondChar
					End If
					If row <= 100 And row > 0 Then
						Set targetCell = ws.Range(targetColLetter & row + 1)
	
						leftPos = targetCell.Left
						bottomPos = targetCell.Top
						cellWidth = targetCell.Width
						cellHeight = targetCell.Height
	
						'draw the gap
						If cellBelow.value = "//" Then
							x1 = leftPos + 0.1 * cellWidth
							y1 = bottomPos - cellHeight
							x2 = leftPos + 0.3 * cellWidth
							y2 = bottomPos - 1.5 * cellHeight
							x3 = leftPos + 0.9 * cellWidth
							y3 = bottomPos - 1.5 * cellHeight
							x4 = leftPos + 0.7 * cellWidth
							y4 = bottomPos - cellHeight
							x5 = leftPos + 0.9 * cellWidth
							y5 = bottomPos - 0.5 * cellHeight
							x6 = leftPos + 0.3 * cellWidth
							y6 = bottomPos - 0.5 * cellHeight
							Call DrawGap(polygonShape, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6)
							Call DrawGap(polygonShape, x1, cellHeight + y1, x2, cellHeight + y2, x3, cellHeight + y3, x4, cellHeight + y4, x5, cellHeight + y5, x6, cellHeight + y6)
						End If
					End If
				End If
			End If
		Next c
	Next r

	If Err.Number <> 0 Then
		MsgBox "Error: " & Err.Number & ": " & Err.Description
	End If
	On Error GoTo 0
End Sub
