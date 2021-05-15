Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports QRCoder

Public Class ADDQR
    <CommandMethod("ADDQR")>
    Public Sub ADDQR()
        On Error GoTo errzone
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim sqlcn As ADODB.Connection = New ADODB.Connection
        Dim sqlrst As ADODB.Recordset = New ADODB.Recordset

        'Получение имени чертежа для поиска в базе данных о ней
        Dim fileName As String = InputBox("Введите наименование чертежа", "Поиск данных в базе")
        If (fileName = "") Then
            Exit Sub
        End If

        'Параметры строки подключения
        sqlcn.ConnectionString = "Provider=SQLOLEDB;" _
            & "Data Source=HOME-PC\SQLEXPRESS;" _
            & "Initial Catalog=SCHEMES_BD;" _
            & "Integrated Security=SSPI"
        'Открываем подключение
        sqlcn.Open()

        'Формируем запрос  пар-мом
        Dim cmdP = New ADODB.Command
        With cmdP
            .ActiveConnection = sqlcn
            .CommandText = "SELECT * FROM [SCHEMES] WHERE SCHEME_NAME =?"
            Dim prm = .CreateParameter("SCHEME_NAME", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100)
            .Parameters.Append(prm)
            .CommandType = ADODB.CommandTypeEnum.adCmdText
        End With
        cmdP.Parameters("SCHEME_NAME").Value = fileName

        'Отправка запроса 
        sqlrst.Open(cmdP)

        'Присвоение переменной data полученных данных
        Dim sqlResult As String = "МАРКИРОВКА: " & sqlrst.Fields(2).Value & vbCrLf & "ИНФОРМАЦИЯ: " & sqlrst.Fields(3).Value

        'Закрываем Recordset, подключение и стираем их из памяти
        sqlrst.Close()
        sqlcn.Close()
        sqlrst = Nothing
        sqlcn = Nothing



        'Запрос места вставки на чертеже
        Dim ptOpts As PromptPointOptions = New PromptPointOptions("")
        ptOpts.Message = vbLf & "Выберите место вставки: "
        Dim ptRes As PromptPointResult = doc.Editor.GetPoint(ptOpts)
        Dim startPt As Point3d = ptRes.Value
        If ptRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If



        'Формирование QR из данных
        Dim gen As New QRCodeGenerator
        Dim qrData = gen.CreateQrCode(sqlResult, QRCodeGenerator.ECCLevel.Q)
        Dim qrCode As New QRCode(qrData)
        Dim qrResult = qrCode.GetGraphic(1)



        'Блок определения размера QR
        'Первая точка рамки информации
        Dim ptInfoBlockBottomLeft As PromptPointOptions = New PromptPointOptions("")
        ptInfoBlockBottomLeft.Message = vbLf & "Выберите нижний левый угол рамки 'Основной надписи':"
        Dim ptInfoBlockBottomLeftRes As PromptPointResult = doc.Editor.GetPoint(ptInfoBlockBottomLeft)
        Dim infoBlockBottomLeftPt As Point3d = ptInfoBlockBottomLeftRes.Value
        If ptInfoBlockBottomLeftRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If

        'Вторая точка рамки информации
        Dim ptInfoBlockTopRight As PromptPointOptions = New PromptPointOptions("")
        ptInfoBlockTopRight.Message = vbLf & "Выберите верхний правый угол рамки 'Основной надписи':"
        Dim ptInfoBlockTopRightRes As PromptPointResult = doc.Editor.GetPoint(ptInfoBlockTopRight)
        Dim infoBlockTopRightPt As Point3d = ptInfoBlockTopRightRes.Value
        If ptInfoBlockTopRightRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If

        'Длина и высота рамки 
        '(185 * ??? = lengthInfoBlock) 
        '(55  * ??? = heightInfoBlock)
        Dim lengthInfoBlock As Double = infoBlockTopRightPt.X - infoBlockBottomLeftPt.X
        Dim heightInfoBlock As Double = infoBlockTopRightPt.Y - infoBlockBottomLeftPt.Y

        'Получение масштаба по сравнению со стандартными значениями (185x55)
        Dim scaleInfoBlock As Double = (185 * 55) / (lengthInfoBlock * heightInfoBlock)

        'Получение размера 1-ого сектора QR
        Dim sizeFullQr As Double = 30 * scaleInfoBlock
        Dim sizeOneBlockQr As Double = (sizeFullQr / qrResult.Height)



        'Блок создания массива по обьектам Bitmap и прорисовка QR по его черным/белым квадратам 
        '(обрабатывает каждый пиксель битмапа и обрабатывает его color значение)
        Dim db As Database = doc.Database
        Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            'Создание блока
            Dim block As BlockTableRecord = New BlockTableRecord()
            block.Name = "QRCODEBLOCK"
            bt.UpgradeOpen()
            Dim blockId As ObjectId = bt.Add(block)
            trans.AddNewlyCreatedDBObject(block, True)

            For y As Integer = 0 To qrResult.Height - 1
                For x As Integer = 0 To qrResult.Width - 1
                    Dim color As Drawing.Color = qrResult.GetPixel(x, y)

                    'Построение контура квадрата
                    Dim Plbox As Polyline = New Polyline()
                    Dim pt As Point2d = New Point2d(startPt.X + (x * sizeOneBlockQr), startPt.Y + (y * sizeOneBlockQr))
                    Plbox.AddVertexAt(0, New Point2d(pt.X, pt.Y), 0.0, -1.0, -1.0)
                    Plbox.AddVertexAt(1, New Point2d(pt.X + sizeOneBlockQr, pt.Y), 0.0, -1.0, -1.0)
                    Plbox.AddVertexAt(2, New Point2d(pt.X + sizeOneBlockQr, pt.Y + sizeOneBlockQr), 0.0, -1.0, -1.0)
                    Plbox.AddVertexAt(3, New Point2d(pt.X, pt.Y + sizeOneBlockQr), 0.0, -1.0, -1.0)
                    Plbox.Closed = True

                    Dim pLineId As ObjectId = block.AppendEntity(Plbox)
                    trans.AddNewlyCreatedDBObject(Plbox, True)

                    Dim ObjIds As ObjectIdCollection = New ObjectIdCollection()
                    ObjIds.Add(pLineId)

                    'Добавдение штриховки на квадрат
                    Dim oHatch As Hatch = New Hatch()
                    Dim normal As Vector3d = New Vector3d(0.0, 0.0, 1.0)
                    oHatch.Normal = normal
                    oHatch.Elevation = 0.0
                    oHatch.PatternScale = 2.0
                    oHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID")
                    If (color.ToString = "Color [A=255, R=255, G=255, B=255]") Then
                        oHatch.ColorIndex = 255 'Белый квадрат
                    Else
                        oHatch.ColorIndex = 250 'Черный квадрат
                    End If

                    block.AppendEntity(oHatch)
                    trans.AddNewlyCreatedDBObject(oHatch, True)

                    oHatch.Associative = True
                    oHatch.AppendLoop(HatchLoopTypes.Default, ObjIds)
                    oHatch.EvaluateHatch(True)
                Next
            Next

            'Добавляем ссылку на блок
            Dim br As BlockReference = New BlockReference(Point3d.Origin, blockId)
            btr.AppendEntity(br)
            trans.AddNewlyCreatedDBObject(br, True)

            'Отправка добавленных объектов на чертеж
            trans.Commit()
        End Using
        Exit Sub



        'Зона обработки ошибок 
        '(для коррекции отображения пользовательских ошибок ввода и прочего)
errzone:
        Select Case Err.Number
            Case 3021
                MsgBox("Объект не найден в базе!" & vbNewLine _
                    & "________________________" & vbNewLine _
                    & "Code: " & Err.Number, vbCritical, "[ОШИБКА]")

            Case -2147467259
                MsgBox("Не удалось подключиться к базе!" & vbNewLine _
                    & "________________________" & vbNewLine _
                    & "Code: " & Err.Number, vbCritical, "[ОШИБКА]")

            Case 5
                MsgBox("Невозможно разметсить блок!" & vbNewLine _
                    & "(блок с именем 'QRCODEBLOCK', уже существует)" & vbNewLine _
                    & "________________________" _
                    & vbNewLine & "Code: " & Err.Number, vbCritical, "[ОШИБКА]")

            Case 13
                MsgBox("В поле ввода введены недопустимые значения!" & vbNewLine _
                    & "________________________" & vbNewLine _
                    & "Code: " & Err.Number, vbCritical, "[ОШИБКА]")

            Case Else
                MsgBox(Err.Description & vbNewLine _
                    & "________________________" & vbNewLine _
                    & "Code: " & Err.Number, vbCritical, "[ОШИБКА]")
        End Select
    End Sub



    'Тестовая функция для проверки размеров QR 
    '(не является частью основной программы, только для тестирования)
    <CommandMethod("TESTSIZE")>
    Public Sub TESTSIZE()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ptOpts As PromptPointOptions = New PromptPointOptions("")
        Dim ptOpts2 As PromptPointOptions = New PromptPointOptions("")

        ptOpts.Message = vbLf & "1 точка: "
        Dim ptRes As PromptPointResult = doc.Editor.GetPoint(ptOpts)
        Dim startPt As Point3d = ptRes.Value
        If ptRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If



        ptOpts2.Message = vbLf & "2 точка: "
        Dim ptRes2 As PromptPointResult = doc.Editor.GetPoint(ptOpts2)
        Dim startPt2 As Point3d = ptRes2.Value
        If ptRes2.Status = PromptStatus.Cancel Then
            Exit Sub
        End If


        MsgBox("ПЕРВАЯ ТОЧКА: " & vbNewLine _
             & "X1: " & startPt.X & vbNewLine _
             & "Y1: " & startPt.Y & vbNewLine _
             & "ВТОРАЯ ТОЧКА: " & vbNewLine _
             & "X2: " & startPt2.X & vbNewLine _
             & "Y2: " & startPt2.Y & vbNewLine _
             & "РАЗРЕШЕНИЕ: " & vbNewLine _
             & (startPt2.X - startPt.X) & vbNewLine _
             & "x" & vbNewLine _
             & (startPt2.Y - startPt.Y) & vbNewLine _
             , vbCritical, "[ТЕСТ]")
    End Sub
End Class