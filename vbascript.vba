Sub RestorePicturesAndFitRows_Min32cm()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim targetRow As Long
    Dim rowTop As Double
    Dim offsetInRow As Double
    Dim neededHeight As Double
    Dim minPts As Double
    Dim padding As Double
    Dim oldPlacement As Variant

    ' 3.2 cm minimum for both pictures and rows
    minPts = Application.CentimetersToPoints(3.2)

    ' Small safety margin
    padding = Application.CentimetersToPoints(0.05)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo SafeExit

    For Each ws In ActiveWorkbook.Worksheets
        For Each shp In ws.Shapes

            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then

                oldPlacement = shp.Placement

                ' Prevent row resize from resizing the image
                shp.Placement = xlMove

                ' Keep correct proportions
                shp.LockAspectRatio = msoTrue

                ' Restore original size, anchored from top-left
                On Error Resume Next
                shp.ScaleWidth 1, msoTrue, msoScaleFromTopLeft
                shp.ScaleHeight 1, msoTrue, msoScaleFromTopLeft
                On Error GoTo SafeExit

                ' Minimum picture height = 3.2 cm
                If shp.Height < minPts Then
                    shp.Height = minPts
                End If

                ' Row where the image starts
                targetRow = shp.TopLeftCell.Row
                rowTop = ws.Rows(targetRow).Top

                ' Vertical offset of image inside that row
                offsetInRow = shp.Top - rowTop
                If offsetInRow < 0 Then offsetInRow = 0

                ' Row must fully contain the image
                neededHeight = offsetInRow + shp.Height + padding

                ' Minimum row height = 3.2 cm
                If neededHeight < minPts Then neededHeight = minPts

                ' Increase row height if needed
                If ws.Rows(targetRow).RowHeight < neededHeight Then
                    ws.Rows(targetRow).RowHeight = neededHeight
                ElseIf ws.Rows(targetRow).RowHeight < minPts Then
                    ws.Rows(targetRow).RowHeight = minPts
                End If

                ' Restore original placement setting
                shp.Placement = oldPlacement

            End If
        Next shp
    Next ws

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
