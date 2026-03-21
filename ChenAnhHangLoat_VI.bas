' ============================================================
' CHEN ANH HANG LOAT VAO EXCEL - PHIEN BAN TIENG VIET
' Ho tro: PNG, JPG, JPEG, JFIF, GIF, BMP, TIFF, WMF, EMF, WEBP, SVG, ICO, HEIC, HEIF
' Phien ban: 1.1 - Da kiem tra bao mat va sua loi
' ============================================================

Option Explicit

' ---------- CAU HINH CHUNG ----------
Private Const COL_IMAGE     As Integer = 3      ' Cot chua ten file anh (C = 3)
Private Const START_ROW     As Long    = 2      ' Hang bat dau doc du lieu
Private Const ROW_HEIGHT_PX As Double  = 80     ' Chieu cao hang (points)
Private Const IMG_PADDING   As Double  = 8      ' Khoang trang xung quanh anh
Private Const MAX_SHAPES    As Long    = 500    ' Gioi han so luong anh (bao ve hieu nang)
' ------------------------------------

' ============================================================
' Ham chinh: Chen anh hang loat
' ============================================================
Sub ChenAnhHangLoat()

    Dim ws          As Worksheet
    Dim sFolder     As String
    Dim imgPath     As String
    Dim r           As Long
    Dim lastRow     As Long
    Dim countOK     As Long
    Dim notFound    As String
    Dim imgName     As String

    ' --- Chon thu muc chua anh ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Chon thu muc chua anh"
        .InitialFileName = Environ("USERPROFILE") & "\Pictures\"
        If .Show = False Then Exit Sub
        sFolder = .SelectedItems(1)
    End With

    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"

    ' --- Kiem tra thu muc ton tai ---
    If Dir(sFolder, vbDirectory) = "" Then
        MsgBox "Khong tim thay thu muc: " & sFolder, vbCritical, "Loi"
        Exit Sub
    End If

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_IMAGE).End(xlUp).Row

    If lastRow < START_ROW Then
        MsgBox "Khong co du lieu tu hang " & START_ROW & " tro xuong!", vbInformation
        Exit Sub
    End If

    ' --- Danh sach dinh dang ho tro ---
    Dim exts As Variant
    exts = Array( _
        ".png", ".PNG", _
        ".jpg", ".JPG", _
        ".jpeg", ".JPEG", _
        ".jfif", ".JFIF", _
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

        ' Kiem tra gioi han so luong anh
        If countOK >= MAX_SHAPES Then
            notFound = notFound & "  - Da dat gioi han " & MAX_SHAPES & " anh, dung lai tai hang " & r & vbNewLine
            Exit For
        End If

        imgName = Trim(CStr(ws.Cells(r, COL_IMAGE).Value))

        ' Bo qua o trong hoac gia tri khong phai so
        If imgName = "" Or imgName = "0" Or Not IsNumeric(imgName) Then GoTo NextRow

        ' Tim file anh voi tung dinh dang
        imgPath = ""
        Dim e As Integer
        For e = 0 To UBound(exts)
            If Dir(sFolder & imgName & exts(e)) <> "" Then
                imgPath = sFolder & imgName & exts(e)
                Exit For
            End If
        Next e

        If imgPath = "" Then
            notFound = notFound & "  - Hang " & r & ": " & imgName & ".* (khong tim thay file)" & vbNewLine
            GoTo NextRow
        End If

        ' Dat chieu cao hang
        ws.Rows(r).RowHeight = ROW_HEIGHT_PX

        ' Xoa anh cu trong o (neu co)
        Dim shpOld As Shape
        For Each shpOld In ws.Shapes
            If shpOld.TopLeftCell.Row = r And shpOld.TopLeftCell.Column = COL_IMAGE Then
                shpOld.Delete
            End If
        Next shpOld
        Set shpOld = Nothing

        ' Lay kich thuoc o hien tai
        Dim cLeft As Double: cLeft = ws.Cells(r, COL_IMAGE).Left
        Dim cTop  As Double: cTop  = ws.Cells(r, COL_IMAGE).Top
        Dim cW    As Double: cW    = ws.Columns(COL_IMAGE).Width
        Dim cH    As Double: cH    = ws.Rows(r).RowHeight

        ' Chen anh vao Excel
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
            ' Xu ly truong hop dac biet (HEIC, SVG can codec rieng)
            Dim ext As String: ext = LCase(Mid(imgPath, InStrRev(imgPath, ".")))
            If ext = ".heic" Or ext = ".heif" Then
                notFound = notFound & "  - Hang " & r & ": " & imgName & ext & " (can cai HEIC codec cho Windows)" & vbNewLine
            ElseIf ext = ".svg" Then
                notFound = notFound & "  - Hang " & r & ": " & imgName & ext & " (SVG can Excel 2016 tro len)" & vbNewLine
            Else
                notFound = notFound & "  - Hang " & r & ": " & imgName & ext & " (loi khi chen anh)" & vbNewLine
            End If
            GoTo NextRow
        End If

        ' Can chinh anh: giu ti le, can giua trong o
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
            .Name = "Img_R" & r ' Dat ten de de quan ly sau nay
        End With
        Set oShp = Nothing

        countOK = countOK + 1

NextRow:
    Next r

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Thong bao ket qua
    Dim msg As String
    msg = "Da chen thanh cong: " & countOK & " anh!" & vbNewLine
    msg = msg & "File an toan de luu va gui."
    If notFound <> "" Then
        msg = msg & vbNewLine & vbNewLine & "Canh bao - Khong xu ly duoc:" & vbNewLine & notFound
    End If
    MsgBox msg, vbInformation, "Hoan thanh"

End Sub

' ============================================================
' Xoa toan bo anh tren sheet hien tai
' ============================================================
Sub XoaAnhHangLoat()

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Ban co chac muon xoa TAT CA anh tren sheet nay?", _
                    vbQuestion + vbYesNo, "Xac nhan")
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

    MsgBox "Da xoa " & n & " anh.", vbInformation, "Hoan thanh"

End Sub
