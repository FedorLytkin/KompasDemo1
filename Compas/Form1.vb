Imports Kompas6Constants
Imports Kompas6API5
Imports ksConstants
Imports KompasAPI7
Public Class Form1

    Public iKompasObject As KompasAPI7._Application


    '  Определения для функций GetObjParam и SetObjParam
    '  '+'  отмечены объекты, для которых реализованы  GetObjParam и SetObjParam
    '  '+-'  отмечены объекты, для которых реализован только GetObjParam
    Public Const ALLPARAM As Short = -1 ' все параметры объекта в системе координат владельца
    Public Const SHEET_ALLPARAM As Short = -2 ' тоже что и  ALLPARAM  но параметры объекта в СК листа
    Public Const NURBS_CLAMPED_PARAM As Short = -5 ' параметры нурбса, преобразовать узловой вектор в зажатый
    Public Const NURBS_CLAMPED_SHEETPARAM As Short = -6 ' параметры нурбса в СК листа, преобразовать узловой вектор в зажатый
    Public Const VIEW_ALLPARAM As Short = -7 ' все параметры объекта в СК вида


    Public Const stACTIVE As Short = 0 ' состояние для вида, слоя, документа
    Public Const stREADONLY As Short = 1 ' состояние для вида, слоя
    Public Const stINVISIBLE As Short = 2 ' состояние для вида, слоя
    Public Const stCURRENT As Short = 3 ' состояние для вида, слоя
    Public Const stPASSIVE As Short = 1 ' состояние для документа


    Public Kompas As Kompas6API5.Application ' Интерфейс KompasObject
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        iKompasObject = CreateObject("Kompas.Application.7")
        iKompasObject.Visible = 1
        Dim api7 As KompasAPI7.IApplication
        Dim Doc As KompasAPI7.AssemblyDocument
        Dim Docs As KompasAPI7.Documents = iKompasObject.Documents
        Doc = iKompasObject.Documents.Open("C:\Users\aidarhanov.n.VEZA-SPB\Downloads\Telegram Desktop\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\Модели\ЕЛГ 10.003.00.300 Крышка.a3d", 1, 0)
        Doc.Close(0)

    End Sub
    Sub createDraw()
        Dim doc As Kompas6API5.Document2D
        Kompas = CreateObject("Kompas.Application.5")
        Dim docPar As Kompas6API5.DocumentParam ' Интерфейс ksDocumentParam
        ' Структура параметров документа
        'docPar = iKompasObject.GetParamStruct(StructType2DEnum.ko_DocumentParam)
        docPar = Kompas.GetParamStruct(StructType2DEnum.ko_DocumentParam)

        Dim sheet As Kompas6API5.SheetPar
        Dim standart As Kompas6API5.StandartSheet
        Dim view As Kompas6API5.ViewParam
        Dim number As Integer ' Интерфейс ksViewParam ' Интерфейс ksStandartSheet ' Интерфейс ksSheetPar
        If Not docPar Is Nothing Then ' Интерфейс создан

            docPar.Init() ' Инициализация
            docPar.fileName = "C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\2.cdw" ' Имя  файла документа
            docPar.comment = "create document" ' Комментарий к документу
            docPar.author = "user" ' Автор документа
            docPar.regime = 0 ' Режим ( 0 - видимый, 1 - слепой )
            docPar.type = 1 ' Тип документа ( 0 - нестандартный, 1 - стандартный чертеж )

            sheet = docPar.GetLayoutParam ' Интерфейс параметров оформления

            If Not sheet Is Nothing Then ' Интерфейс создан

                sheet.shtType = 1 ' Тип штампа из указанной библиотеки для спецификации ( номер стиля из указанной библиотеки )
                sheet.layoutName = "" ' Имя библиотеки оформления,

                standart = sheet.GetSheetParam() ' Интерфейс параметров стандартного листа

                If Not standart Is Nothing Then ' Интерфейс создан

                    standart.format = 3 ' Формат листа 0( А0 ) ... 4( А4 )
                    standart.multiply = 1 ' Кратность формата
                    standart.direct = 0 ' Расположение штампа ( 0 - вдоль короткой стороны, 1 - вдоль длинной )

                    ' Создаем документ: лист, формат А3, горизонтально расположенный и с системным штампом 1
                    doc.kscre
                    If doc.ksCreateDocument(docPar) Then

                        ' Структура параметров вида
                        view = Kompas.GetParamStruct(Kompas6Constants.StructType2DEnum.ko_ViewParam)

                        If Not view Is Nothing Then ' Интерфейс создан

                            view.x = 10 ' Точка привязки вида
                            view.y = 20
                            view.angle = 45 ' Угол поворота вида
                            view.scale_ = 0.5 ' Масштаб вида
                            view.color = RGB(10, 20, 10) ' Цвет вида в активном состоянии
                            view.state = stACTIVE ' Состояние вида
                            view.name = "user view" ' Имя вида

                            number = 2

                            ' У документа создадим вид с номером 2, масштабом 0.5, под углом 45 гр
                            doc.ksCreateSheetView(view, number)

                            doc.ksLayer(5) ' Создадим слой с номером 5

                            doc.ksLineSeg(20, 10, 40, 10, 1) ' Отрисовка отрезков
                            doc.ksLineSeg(40, 10, 40, 30, 1)
                            doc.ksLineSeg(40, 30, 20, 30, 1)
                            doc.ksLineSeg(20, 30, 20, 10, 1)

                            Kompas.ksMessage("нарисовали")

                            ' Получить параметры документа
                            doc.ksGetObjParam(doc.reference, docPar, ALLPARAM)

                            Kompas.ksMessage("type = " & docPar.type & " f = " & standart.format & " m = " & standart.multiply & " d = " & standart.direct)

                            Kompas.ksMessage("Имя файла : " & docPar.fileName)
                            Kompas.ksMessage("Комментарий : " & docPar.comment)
                            Kompas.ksMessage("Автор : " & docPar.author)

                            doc.ksSaveDocument("") ' Сохраним документ
                            doc.ksCloseDocument() ' Закрыть документ
                        End If
                    End If
                End If
            End If
        End If

    End Sub
    Function is_running() As Boolean
        Dim p As Process = New Process
        p.StartInfo.UseShellExecute = 0
        p.StartInfo.RedirectStandardOutput = 1
        p.StartInfo.FileName = "KOMPAS (64)"
        p.Start()
        Dim Output As String = p.StandardOutput.ReadToEnd
        p.WaitForExit()
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        createDraw()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        inii()
    End Sub

    Sub d2add()

        Dim OpenFile As String = "C:\Users\aidarhanov.n.VEZA-SPB\Downloads\Telegram Desktop\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\Модели\ЕЛГ 10.003.00.101 Стенка.m3d"
        Dim newKompasAPI As IApplication
        'Dim disp As System.Runtime.InteropServices.UnmanagedType.IDispatch
        Dim pDocuments As IDocuments

        Dim pDocument As IKompasDocument
        Dim pKompasDocument2D As IKompasDocument2D
        Dim pViewsAndLayersManager As IViewsAndLayersManager
        Dim pViews As IViews
        Dim pView As IView
        Dim air(2) As Integer
        air(0) = 1
        air(0) = 3
        air(0) = 5
        Dim ks As Object = CreateObject("Kompas.Application.7")

        newKompasAPI = ks
        newKompasAPI.Visible = 1
        pDocuments = newKompasAPI.Documents
        pDocument = pDocuments.Add(1, True)
        pKompasDocument2D = pDocument
        pViewsAndLayersManager = pKompasDocument2D.ViewsAndLayersManager
        pViews = pViewsAndLayersManager.Views
        pViews.AddStandartViews(OpenFile, "ff", air, 0, 0, 1, 10, 10)
        pKompasDocument2D.SaveAs("C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\sadfasdfs.dxf")

    End Sub

    Sub inii()
        Dim OpenFile As String = "C:\Users\aidarhanov.n.VEZA-SPB\Downloads\Telegram Desktop\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\Модели\ЕЛГ 10.003.00.101 Стенка.m3d"

        Dim kmp As KompasObject
        Dim doc2d As ksDocument2D
        Dim docparametr As ksDocumentParam
        Dim odoc As Object
        kmp = CreateObject("Kompas.Application.5")
        If kmp Is Nothing Then Exit Sub
        doc2d = kmp.Document2D
        docparametr = kmp.GetParamStruct(StructType2DEnum.ko_DocumentParam)
        If Not docparametr Is Nothing Then
            docparametr.comment = "Фрагмент"
            docparametr.author = "Автор"
            docparametr.fileName = "filename"
            docparametr.regime = 0
            docparametr.type = 2
            doc2d.ksCreateDocument(docparametr)

            Dim view As Kompas6API5.ViewParam = kmp.GetParamStruct(StructType2DEnum.ko_ViewParam)
            If Not view Is Nothing Then ' Интерфейс создан
                'вот тут нужно вставить вид с модели
                view.x = 10 ' Точка привязки вида
                view.y = 20
                view.angle = 0 ' Угол поворота вида
                view.scale_ = 0.5 ' Масштаб вида
                view.color = RGB(10, 20, 10) ' Цвет вида в активном состоянии
                view.state = stACTIVE ' Состояние вида
                view.name = "user view" ' Имя вида

                Dim Number As Integer = 2

                ' У документа создадим вид с номером 2, масштабом 0.5, под углом 45 гр
                doc2d.ksCreateSheetView(view, Number)

                doc2d.ksLayer(5) ' Создадим слой с номером 5
                view.add
                doc2d.ksCreateSheetStandartViews(OpenFile, 1, 0, 0)

                'doc2d.ksLineSeg(20, 10, 40, 10, 1) ' Отрисовка отрезков
                'doc2d.ksLineSeg(40, 10, 40, 30, 1)
                'doc2d.ksLineSeg(40, 30, 20, 30, 1)
                'doc2d.ksLineSeg(20, 30, 20, 10, 1)


                ' Получить параметры документа
                doc2d.ksGetObjParam(doc2d.reference, doc2d, ALLPARAM)
            End If
            doc2d.ksSaveDocument("C:\Users\aidarhanov.n.VEZA-SPB\Desktop\test 1\test1.dxf")
            odoc = kmp.ActiveDocument2D()
        End If
    End Sub
    Sub open3d()
        Dim OpenFile As String = "C:\Users\aidarhanov.n.VEZA-SPB\Downloads\Telegram Desktop\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\ЕЛГ 10.003.00.000 Фильтр для пайки SF-300\Модели\ЕЛГ 10.003.00.101 Стенка.m3d"
        Dim newKompasAPI As KompasAPI7.IApplication
        'Dim disp As System.Runtime.InteropServices.UnmanagedType.IDispatch
        Dim pDocuments As KompasAPI7.IDocuments

        Dim pDocument As KompasAPI7.IKompasDocument
        Dim Doc3D As KompasAPI7.IKompasDocument3D
        Dim pPart7 As KompasAPI7.IPart7

        Dim pSheetMetalContainer As KompasAPI7.ISheetMetalContainer
        Dim pSheetMetalBodies As KompasAPI7.ISheetMetalBodies
        Dim pSheetMetalBody As KompasAPI7.ISheetMetalBody

        Dim ks As Object = CreateObject("Kompas.Application.7")
        'iKompasObject.Visible = 1

        newKompasAPI = ks

        pDocuments = newKompasAPI.Documents
        pDocument = pDocuments.Open(OpenFile, True, False)
        Doc3D = pDocument

        pPart7 = Doc3D.TopPart

        pSheetMetalContainer = pPart7
        pSheetMetalBodies = pSheetMetalContainer.SheetMetalBodies
        pSheetMetalBody = pSheetMetalBodies.SheetMetalBody(0)
        If newKompasAPI.IsKompasCommandCheck(40794) = 0 Then newKompasAPI.ExecuteKompasCommand(40794, False)
        'pSheetMetalBody.Straighten = True

        Doc3D.RebuildDocument()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        d2add()
    End Sub
    Sub createDrView()

    End Sub
End Class
