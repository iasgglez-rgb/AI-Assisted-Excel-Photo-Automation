Attribute VB_Name = "Módulo1"
Option Explicit

' ============================================================
' CONFIGURATION
' ============================================================

Private Const PHOTO_FOLDER As String = "BF"
Private Const WORKSHEET_NAME As String = "Hoja4"

' Photo box dimensions:
' Rows 8-19
' Columns B-W
'
' The same dimensions are used for every photo position.

Private Const BOX_FIRST_ROW As Long = 8
Private Const BOX_LAST_ROW As Long = 19

Private Const BOX_FIRST_COLUMN As Long = 3      ' C
Private Const BOX_LAST_COLUMN As Long = 24      ' X


' ============================================================
' MAIN PROCEDURE
' ============================================================

Public Sub InsertarFotos()

    Dim ws As Worksheet
    Dim photoFolder As String
    Dim missingFiles As String

    Dim i As Long
    Dim fileName As String
    Dim photoPath As String

    Dim photoBox As Range
    Dim pic As Shape

    Dim insertedCount As Long
    Dim missingCount As Long


    ' --------------------------------------------------------
    ' Get worksheet
    ' --------------------------------------------------------

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(WORKSHEET_NAME)

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "The worksheet '" & WORKSHEET_NAME & "' was not found.", _
            vbCritical, _
            "Worksheet Not Found"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Locate BF folder
    ' --------------------------------------------------------

    photoFolder = ThisWorkbook.Path & _
                  Application.PathSeparator & _
                  PHOTO_FOLDER


    If Dir(photoFolder, vbDirectory) = "" Then

        MsgBox _
            "The folder 'BF' was not found." & vbCrLf & vbCrLf & _
            "Expected location:" & vbCrLf & _
            photoFolder, _
            vbExclamation, _
            "Photo Folder Not Found"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Remove previously inserted photographs
    ' --------------------------------------------------------

    DeleteExistingPhotos ws


    ' ========================================================
    ' INSTALLED PHOTOS
    ' ========================================================

    For i = 1 To 6

        fileName = CStr(i) & ".png"

        photoPath = photoFolder & _
                    Application.PathSeparator & _
                    fileName

        Set photoBox = GetPhotoBox(ws, i, False)


        If Dir(photoPath) <> "" Then

            Set pic = InsertPhoto(ws, photoPath, photoBox)

            insertedCount = insertedCount + 1

        Else

            missingFiles = missingFiles & _
                           fileName & _
                           "  (Installed)" & _
                           vbCrLf

            missingCount = missingCount + 1

        End If

    Next i


    ' ========================================================
    ' DISMANTLED PHOTOS
    ' ========================================================

    For i = 1 To 6

        fileName = CStr(i) & "B.png"

        photoPath = photoFolder & _
                    Application.PathSeparator & _
                    fileName

        Set photoBox = GetPhotoBox(ws, i, True)


        If Dir(photoPath) <> "" Then

            Set pic = InsertPhoto(ws, photoPath, photoBox)

            insertedCount = insertedCount + 1

        Else

            missingFiles = missingFiles & _
                           fileName & _
                           "  (Dismantled)" & _
                           vbCrLf

            missingCount = missingCount + 1

        End If

    Next i


    ' ========================================================
    ' FINAL REPORT
    ' ========================================================

    If missingCount = 0 Then

        MsgBox _
            "Photo insertion completed successfully." & vbCrLf & vbCrLf & _
            "Photographs inserted: " & insertedCount, _
            vbInformation, _
            "Completed"

    Else

        MsgBox _
            "Photo insertion completed with warnings." & vbCrLf & vbCrLf & _
            "Photographs inserted: " & insertedCount & vbCrLf & _
            "Photographs missing: " & missingCount & vbCrLf & vbCrLf & _
            "Missing files:" & vbCrLf & _
            missingFiles, _
            vbExclamation, _
            "Completed With Missing Photos"

    End If

End Sub


' ============================================================
' DESTINATION BOX
' ============================================================

Private Function GetPhotoBox( _
    ByVal ws As Worksheet, _
    ByVal position As Long, _
    ByVal dismantled As Boolean) As Range


    Dim startRow As Long
    Dim startColumn As Long


    ' Each photo occupies 12 rows.
    ' Four rows separate consecutive photo boxes.

    startRow = 8 + ((position - 1) * 16)


    If dismantled Then

        startColumn = 29       ' AC

    Else

        startColumn = 3        ' C

    End If


    Set GetPhotoBox = ws.Range( _
        ws.Cells(startRow, startColumn), _
        ws.Cells( _
            startRow + 11, _
            startColumn + 21 _
        ) _
    )

End Function


' ============================================================
' INSERT PHOTO
' ============================================================
'
' FIT-TO-BOX:
'
' The photograph is resized to the largest possible size
' that fits inside the destination box while preserving
' its original aspect ratio.
'
' The photograph is then centered inside the box.
'
' ============================================================

Private Function InsertPhoto( _
    ByVal ws As Worksheet, _
    ByVal photoPath As String, _
    ByVal photoBox As Range) As Shape


    Dim pic As Shape

    Dim boxLeft As Double
    Dim boxTop As Double
    Dim boxWidth As Double
    Dim boxHeight As Double

    Dim scaleFactor As Double


    ' --------------------------------------------------------
    ' Destination dimensions
    ' --------------------------------------------------------

    boxLeft = photoBox.Left
    boxTop = photoBox.Top

    boxWidth = photoBox.Width
    boxHeight = photoBox.Height


    ' --------------------------------------------------------
    ' Insert image
    ' --------------------------------------------------------

    Set pic = ws.Shapes.AddPicture( _
        fileName:=photoPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=boxLeft, _
        Top:=boxTop, _
        Width:=-1, _
        Height:=-1)


    ' --------------------------------------------------------
    ' Identify the image as one created by this macro
    ' --------------------------------------------------------

    pic.Name = "PhotoBot_" & _
               Format(Timer * 1000, "0") & "_" & _
               ws.Shapes.Count


    ' --------------------------------------------------------
    ' Preserve aspect ratio
    ' --------------------------------------------------------

    pic.LockAspectRatio = msoTrue


    ' --------------------------------------------------------
    ' Calculate FIT-TO-BOX scale
    ' --------------------------------------------------------

    scaleFactor = WorksheetFunction.Min( _
        boxWidth / pic.Width, _
        boxHeight / pic.Height)


    ' --------------------------------------------------------
    ' Resize
    ' --------------------------------------------------------

    pic.Width = pic.Width * scaleFactor


    ' --------------------------------------------------------
    ' Center inside destination box
    ' --------------------------------------------------------

    pic.Left = boxLeft + _
               ((boxWidth - pic.Width) / 2)

    pic.Top = boxTop + _
              ((boxHeight - pic.Height) / 2)


    ' --------------------------------------------------------
    ' Make image follow the cells
    ' --------------------------------------------------------

    pic.Placement = xlMoveAndSize


    Set InsertPhoto = pic

End Function


' ============================================================
' DELETE EXISTING PHOTOS
' ============================================================
'
' Only photographs previously created by this macro are
' deleted. Other shapes or objects in the worksheet remain.
'
' ============================================================

Private Sub DeleteExistingPhotos(ByVal ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1

        If Left(ws.Shapes(i).Name, 9) = "PhotoBot_" Then

            ws.Shapes(i).Delete

        End If

    Next i

End Sub

