' ============================================================
' BULK IMAGE INSERTER FOR EXCEL - ENGLISH VERSION
' Supported formats: PNG, JPG, JPEG, GIF, BMP, TIFF, WMF, EMF, WEBP, SVG, ICO, HEIC
' Version: 1.1 - Security reviewed and bug-fixed
' ============================================================

Option Explicit

' ---------- CONFIGURATION ----------
Private Const COL_IMAGE     As Integer = 3      ' Column containing image filenames (C = 3)
Private Const START_ROW     As Long    = 2      ' First data row (skip header)
Private Const ROW_HEIGHT_PX As Double  = 80     ' Row height in points
Private Const IMG_PADDING   As Double  = 8      ' Padding around image inside cell
Private Const MAX_SHAPES    As Long    = 500    ' Safety limit on total images inserted
' -----------------------------------

' ============================================================
' Main subroutine: Insert images in bulk
' ============================================================
Sub InsertImagesBulk()

    Dim ws          As Worksheet
    Dim sFolder     As String
    Dim imgPath     As String
    Dim r           As Long
    Dim lastRow     As Long
    Dim countOK     As Long
    Dim notFound    As String
    Dim imgName     As String

    ' --- Folder picker dialog ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing your images"
        .InitialFileName = Environ("USERPROFILE") & "\Pictures\"
        If .Show = False Then Exit Sub
        sFolder = .SelectedItems(1)
    End With

    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"

    ' --- Verify folder exists ---
    If Dir(sFolder, vbDirectory) = "" Then
        MsgBox "Folder not found: " & sFolder, vbCritical, "Error"
        Exit Sub
    End If

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_IMAGE).End(xlUp).Row

    If lastRow < START_ROW Then
        MsgBox "No data found from row " & START_ROW & " onwards!", vbInformation
        Exit Sub
    End If

    ' --- Supported image extensions ---
    Dim exts As Variant
    exts = Array( _
        ".png", ".PNG", _
        ".jpg", ".JPG", _
        ".jpeg", ".JPEG", _
        ".gif", ".GIF", _
        ".bmp", ".BMP", _
        ".tif", ".TIF", _
        ".tiff", ".TIFF", _
        ".wmf", ".WMF", _
        ".emf", ".EMF", _
        ".webp", ".WEBP", _
        ".svg", ".SVG", _
        ".ico", ".ICO", _
        ".heic", ".HEIC", _
        ".heif", ".HEIF" _
    )

    ws.Columns(COL_IMAGE).ColumnWidth = 14
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    countOK = 0
    notFound = ""

    Dim oShp As Shape

    For r = START_ROW To lastRow

        ' Safety: enforce maximum shape limit
        If countOK >= MAX_SHAPES Then
            notFound = notFound & "  - Reached limit of " & MAX_SHAPES & " images, stopped at row " & r & vbNewLine
            Exit For
        End If

        imgName = Trim(CStr(ws.Cells(r, COL_IMAGE).Value))

        ' Skip empty cells or non-numeric values
        If imgName = "" Or imgName = "0" Or Not IsNumeric(imgName) Then GoTo NextRow

        ' Search for image file with all supported extensions
        imgPath = ""
        Dim e As Integer
        For e = 0 To UBound(exts)
            If Dir(sFolder & imgName & exts(e)) <> "" Then
                imgPath = sFolder & imgName & exts(e)
                Exit For
            End If
        Next e

        If imgPath = "" Then
            notFound = notFound & "  - Row " & r & ": " & imgName & ".* (file not found)" & vbNewLine
            GoTo NextRow
        End If

        ' Set row height
        ws.Rows(r).RowHeight = ROW_HEIGHT_PX

        ' Remove any existing image in this cell
        Dim shpOld As Shape
        For Each shpOld In ws.Shapes
            If shpOld.TopLeftCell.Row = r And shpOld.TopLeftCell.Column = COL_IMAGE Then
                shpOld.Delete
            End If
        Next shpOld
        Set shpOld = Nothing

        ' Get cell dimensions
        Dim cLeft As Double: cLeft = ws.Cells(r, COL_IMAGE).Left
        Dim cTop  As Double: cTop  = ws.Cells(r, COL_IMAGE).Top
        Dim cW    As Double: cW    = ws.Columns(COL_IMAGE).Width
        Dim cH    As Double: cH    = ws.Rows(r).RowHeight

        ' Insert the picture
        Set oShp = Nothing
        On Error Resume Next
        Set oShp = ws.Shapes.AddPicture( _
            Filename:=imgPath, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=cLeft, Top:=cTop, _
            Width:=cW, Height:=cH _
        )
        On Error GoTo 0

        If oShp Is Nothing Then
            Dim ext As String: ext = LCase(Mid(imgPath, InStrRev(imgPath, ".")))
            If ext = ".heic" Or ext = ".heif" Then
                notFound = notFound & "  - Row " & r & ": " & imgName & ext & " (requires HEIC codec for Windows)" & vbNewLine
            ElseIf ext = ".svg" Then
                notFound = notFound & "  - Row " & r & ": " & imgName & ext & " (SVG requires Excel 2016 or later)" & vbNewLine
            Else
                notFound = notFound & "  - Row " & r & ": " & imgName & ext & " (insert error)" & vbNewLine
            End If
            GoTo NextRow
        End If

        ' Resize image: maintain aspect ratio, center in cell
        With oShp
            .ScaleWidth 1, msoTrue
            .ScaleHeight 1, msoTrue

            Dim ratio As Double: ratio = .Width / .Height
            Dim nW As Double, nH As Double

            If ratio >= (cW - IMG_PADDING) / (cH - IMG_PADDING) Then
                nW = cW - IMG_PADDING: nH = nW / ratio
            Else
                nH = cH - IMG_PADDING: nW = nH * ratio
            End If

            .Width = nW
            .Height = nH
            .Left = cLeft + (cW - nW) / 2
            .Top = cTop + (cH - nH) / 2
            .Placement = xlMoveAndSize
            .LockAspectRatio = msoTrue
            .Name = "Img_R" & r
        End With
        Set oShp = Nothing

        countOK = countOK + 1

NextRow:
    Next r

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Summary message
    Dim msg As String
    msg = "Successfully inserted: " & countOK & " image(s)!" & vbNewLine
    msg = msg & "File is safe to save and share."
    If notFound <> "" Then
        msg = msg & vbNewLine & vbNewLine & "Warnings - could not process:" & vbNewLine & notFound
    End If
    MsgBox msg, vbInformation, "Done"

End Sub

' ============================================================
' Delete all images on the active sheet
' ============================================================
Sub DeleteAllImages()

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Are you sure you want to delete ALL images on this sheet?", _
                    vbQuestion + vbYesNo, "Confirm")
    If answer = vbNo Then Exit Sub

    Dim shp As Shape
    Dim n   As Long
    n = 0

    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            shp.Delete
            n = n + 1
        End If
    Next shp

    MsgBox "Deleted " & n & " image(s).", vbInformation, "Done"

End Sub
