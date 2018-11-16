

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Thu Mar 09 18:01:24 2017
 */
/* Compiler settings for libConsts.odl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __ksConstants_h__
#define __ksConstants_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __Kompas6Constants_LIBRARY_DEFINED__
#define __Kompas6Constants_LIBRARY_DEFINED__

/* library Kompas6Constants */
/* [helpfile][version][uuid] */ 

//-------------------------------------------------------------------------------
// _1_ Константы событий
// ---
  // _1_00_ ksNotifyType                      - Перечень интерфейсов событий
  // _1_01_ ksKompasObjectNotifyEnum          - 1  События для KompasObject 12
  // _1_02_ ksDocumentFileNotifyEnum          - 2  События для документов : Работа с файлом 10
  // _1_03_ ksStampNotifyEnum                 - 3  События для основной надписи 4
  // _1_04_ ksObject2DNotifyEnum              - 4  События для объектов 2Д документов 21
  // _1_05_ ksSelectionMngNotifyEnum          - 5  События для менеджера селектирования 3
  // _1_06_ ksSpcObjectNotifyEnum             - 6  События для объекта спецификации 16
  // _1_07_ ksSpcDocumentNotifyEnum           - 7  События для документа спецификации 6
  // _1_08_ ksSpecificationNotifyEnum         - 8  События для спецификации 11
  // _1_09_ ksDocument3DNotifyEnum            - 9  События для 3D документа 8
  // _1_10_ ksObject3DNotifyEnum              - 10 События для объекта 3D документа 12
  // _1_11_ ksDocument2DNotifyEnum            - 11 События для 2D документа 6
  // _1_12_ ksPropertyManagerNotifyEnum       - 12 События панели свойств или процессных параметров. 10
  // _1_13_ ksPropertyUserControlNotifyEnum   - 13 События пользовательского элемента управления 2.
  // _1_14_ ksDocumentFrameNotifyEnum         - 14 События для окна документа : клавиатура, мышь, события по отрисовке документа. 13
  // _1_15_ ksViewsAndLayersManagerNotifyEnum - 15 События для менеджера видов и слоев  2
  // _1_16_ ksLibraryManagerNotifyEnum        - 16 События для менеджера библиотек 13
  // _1_17_ ksLayoutSheetsNotifyEnum          - 17 События для листов оформления 3
  // _1_18_ ksProcess2DNotifyEnum             - 19 События процесса 2D. 2
  // _1_19_ ksProcess3DNotifyEnum             - 20 События процесса 3D. 3
  // Всего событий - 167
  //
  // _1_108_ ksMaterialPropertyTypeEnum       - Тип выбора материала.

//-------------------------------------------------------------------------------
// _2_ Основные константы
// ---
  // _2_01_ KompasAPIObjectTypeEnum	- Типы объектов КОМПАС API
  // _2_02_ ksTypeLookStyle				  - Тип отрисовки визуальной части
  // _2_03_ DocumentTypeEnum				  - Типы документов КОМПАС
  // _2_04_ FrameRegimeEnum				  - Pежим отображения окна
  // _2_05_ ZoomTypeEnum					    - Тип изменения масштаба отображения документа в окне.
  // _2_06_ StructType2DEnum				  - Константы интерфейсов параметров создаваемых методом GetParamStruct объекта KompasObject
  // _2_07_ ProcessTypeEnum				  - Типы процессов в КОМПАС
  // _2_08_ DocumentCloseOptions			- Действия при закрытии документа КОМПАС
  // _2_09_ ConvertCoordTypeEnum			- Типы преобразования логических координат в координаты документа
  // _2_10_ DrawingObjectTypeEnum		- Типы графических объектов
  // _2_11_ DocumentFormatEnum			  - Форматы листа
  // _2_12_ ksAlignEnum					    - Выравнивание
  // _2_13_ TabulatorFillingEnum			- Заполнение табулятора
  // _2_14_ ErrorType					      - Ошибки API, кроме 3D
  // _2_15_ ksKOMPASConverterEnum		- Типы внутренних конвертеров КОМПАС 3D
  // _2_16_ ksLengthUnitEnum				  - Единицы измерения длины
  // _2_17_ ksStampEnum					    - Идентификаторы ячеек штампа
  // _2_18_ ksAPITypeEnum				    - Тип API интерфейса
  // _2_19_ ksTreeTypeEnum				    - Типы дерева построения 3D документа
  // _2_20_ ksKompasCommandEnum			- Команды компаса.
  // _2_21_ ksHideMessageEnum			  - Скрывать выдачу сообщений.
  // _2_22_ ksConstraintTypeEnum			- Типы параметрических ограничений.
  // _2_23_ ksBmpSizeEnum				    - Размеры иконок.
  // _2_24_ ksRequestFilesTypeEnum   - Тип процесса запрашивающего файл или список файлов.
  // _2_25_ ksAngleEnum              - Углы поворотов кратные 90 гр.
  // _2_26_ ksChooseManagerTypeEnum  - Тип менеджера выбора объектов.
  // _2_27_ ksRecoverErrorEnum       - Признак ошибок при открытии документа с восстановлением.
  // _2_28_ ksExternalFilesTypesEnum - Тип внешнего файла.
  // _2_29_ ksBooleanType            - Типы булевых операций над твердыми телами и замкнутыми кривыми
  // _2_30_ ksSnapTypeEnum           - Тип привязки к объектам
	 // _2_31_ ksNewDocumentSettingsTypeEnum - Тип настроек новых документов.
  // _2_32_ ksPrinterTypeEnum        - Текущий принтер.
  // _2_33_ ksSheetType              - Тип листа
  // _2_34_ ksDrawInScreenPlaneEnum  - Типы отображения объектов внешней триангуляции.
  // _2_35_ ksTextAlignEnum          - Выравнивание текста.
  // _2_36_ ksSlaveDocumentTypeEnum  - Типы подчиненных режимов редактирования документов КОМПАС
  // _2_37_ ksKompasVariantEnum      - Константы вариантов реализаций Компас.
  // _2_38_ ksCheckSumVersionEnum    - Версии котрольных сумм ICheckSum


//-------------------------------------------------------------------------------
// _3_ Константы панели свойств и процессов
// ---
  // _3_01_ ControlTypeEnum                - Типы контролов панели свойств
  // _3_02_ SpecPropertyButtonEnum         - Предопределенные кнопки панели Kompas.
  // _3_03_ SpecPropertyToolBarEnum        - Предопределенные спецпанели для панели свойств
  // _3_04_ SeparatorTypeEnum              - Типы элемента панели свойств - разделитель (сепаратор).
  // _3_05_ CheckStateEnum                 - Состояние чекбокса
  // _3_06_ PropertyControlNameVisibility  - Видимость имени контрола на панели свойств
  // _3_07_ ButtonTypeEnum                 - Типы кнопок для элемента панели свойств - Набор кнопок
  // _3_08_ SlideTypeEnum                  - Тип отображения слайда в окне.
  // _3_09_ DefaultFixTypeEnum             - Тип фиксированности для умолчательных элеменов управления панели свойств
  // _3_10_ PropertyManagerLayout          - Положение панели свойств (вверху, внизу, слева, справа, плавает)
  // _3_11_ ksEditListTypeEnum             - Типы кнопок для элемента панели свойств - Набор кнопок
  // _3_12_ ksEditListCommandEnum          - Идентификаторы стандартных команд для элемента панели свойств - Список
  // _3_13_ ksCheckBoxVisualStyleEnum       - Визуальный стиль чекбокса
  // _3_14_ ksPhantomTypeEnum              - Типы фантома
  // _3_15_ ksProcess2DTypeEnum            - Типы процессов 2D
  // _3_16_ ksProcess3DTypeEnum            - Типы процессов 3D
  // _3_17_ ksTwinSwitcherValueEnum        - Значения переключателя
  // _3_18_ ksProcessContextMenuType       - Тип процессного меню
  // _3_19_ ksEnterButtonIconTypeEnum      - Тип иконки для кнопки создать


//-------------------------------------------------------------------------------
// _4_ Константы 2D документа
// ---
  // _4_01_ FilterCondotionStateEnum       - Состояние параметра в условии фильтрации слоев
  // _4_02_ LtViewType                     - Типы видов.
  // _4_03_ LayersGroupWayEnum             - Способ группировки слоев.
  // _4_04_ ksTextNumberingEnum            - Тип нумерации абзаца.
  // _4_05_ ksTextSizeEnum                 - Размерный коэффиициент текста.
  // _4_06_ ksTextItemEnum                 - Тип компоненты текста.
  // _4_07_ ksTextHorizontalFormatEnum     - Признак горизонтального форматирования.
  // _4_08_ ksAllocationEnum               - Размещенеие объекта.
  // _4_09_ ksMarkInsideFormEnum           - Тип формы для марки (без линии-выноски).
  // _4_10_ ksMarkOnLinePosTypeEnum	      - Положение марки относительно линии.
  // _4_11_ ksArrowEnum                    - Тип стрелки линии-выноски.
  // _4_12_ ksShelfDirectionEnum           - Направление полки.
  // _4_13_ ksDrawingObjecParamTypeEnum    - Тип параметров объекта.
  // _4_15_ ksMarkNodeEnum                 - Тип узла марки.
  // _4_16_ ksCurveStyleEnum               - Системные стили линий
  // _4_17_ ksTextStyleEnum                - Системные стили текста
  // _4_18_ ksContourFormEnum              - Форма контура Обозначения узла.
  // _4_19_ ksAlignmentTypeEnum            - Тип ориентации объекта.
  // _4_20_ ksLengthUnitsEnum              - Единицы измерения длины.
  // _4_21_ ksMassUnitsEnum                - Единицы измерения массы.
  // _4_22_ ksMassSettingModeEnum          - Варианты задания МЦХ.
  // _4_23_ ksAnnotationSymbolEnum         - Аннотационные символы.
  // _4_24_ ksDimensionTextAlignEnum       - Выравнивание размерной надписи.
  // _4_25_ ksDimensionTextBracketsEnum    - Размер в скобках.
  // _4_26_ ksDimensionDeviationEnum       - Отклонения номинального значения размера.
  // _4_27_ ksDimensionTextTypeEnum        - Тип размерной надписи.
  // _4_28_ ksDimensionBaseEnum            - Параметр отрисовки текста.
  // _4_29_ ksDimensionTextPosEnum         - Положение размерной надписи отноcительно выносной линии.
  // _4_30_ ksDimensionArrowPosEnum        - Размещение стрелок относительно выносной линии.
  // _4_31_ ksLineDimensionOrientationEnum - Тип ориентации линейного размера.
  // _4_32_ ksHeightDimTypeEnum            - Тип размеров высоты.
  // _4_33_ ksAngleDimTypeEnum             - Тип углового размера.
  // _4_34_ ksRoughSignEnum                - Тип значка шероховатости.
  // _4_35_ ksLeaderSignEnum               - Тип значка для линии выноски.
  // _4_36_ ksPositionLederFormEnum        - Тип формы для позиционной линии выноски.
  // _4_37_ ksChangeLeaderSignEnum         - Тип значка для обозначения изменения.
  // _4_38_ ksToleranceSuffixSignEnum      - Знак в обозначении базы допуска.
  // _4_39_ ksTablePointEnum               - Тип расположения точки на таблице.
  // _4_40_ ksTableTileLayoutEnum          - Расположение заголовка таблицы.
  // _4_41_ ksCellBoundariesEnum           - Типы границ таблицы.
  // _4_42_ ksTextLineType                 - Тип строки текста.
  // _4_43_ ksMlVertexLimiterEnum				  - Типы ограничений в вершинах мультилинии.
  // _4_44_ ksMlEndLimiterEnum				      - Типы ограничений на концах мультилинии.
  // _4_45_ ksMlVertexTrackingEnum				  - Типы обхода вершин мультилинии.
  // _4_46_ ksContourSegmentEnum				    - Типы сегментов контура.
  // _4_47_ ksAnnotativeSymbolEnum         - Типы специальных символов для аннатационных объектов.
  // _4_48_ ksCornerTypeEnum               - Тип угла объекта.
  // _4_49_ ksInsertionTypeEnum            - Тип вставки фрагмента или вида.
  // _4_50_ ksEquidistantTypeEnum          - Тип построения эквидистанты.
  // _4_51_ ksCentreMarkerEnum             - Тип обозначения центра.
  // _4_52_ ksSemiAxisTypeEnum             - Тип полуоси для обозначения центра.
  // _4_53_ ksAttributeTypeEnum            - Тип данных для типа аттрибута.
  // _4_54_ ksRelativeProjectionTypeEnum   - Тип проекции для ассоциативного вида.
  // _4_55_ ksBreakLineTypeEnum            - Тип линии разрыва.
  // _4_56_ ksHatchStyleEnum               - Системные стили штриховки.
  // _4_57_ ksRegionTypeEnum               - Тип региона.
  // _4_58_ ksColouringTypeEnum            - Тип заливки.
  // _4_59_ ksSheetsRangeEnum              - Тип диапазона страниц.
  // _4_60_ ksPrintJob_OutputColorTypeEnum - Цвет вывода на печать.
  // _4_61_ ksHypertextTypeEnum            - Тип ссылки на текст.
  // _4_62_ ksHyperLinkTypeEnum            - Тип гиперссылки.
  // _4_63_ ksBisectorVariant              - Вариант решения биссектрисы для двух прямых.
  // _4_64_ ksNurbsByPointsBuildingTypeEnum     - Способ формирования точек сплайна.
  // _4_65_ ksNurbsByPointsPointConstraintsEnum - Вариант управления точкой сплайна.
  // _4_66_ ksNurbsByPointsAproximationTypeEnum - Способ вычисления шага аппроксимации для точек сплайна
  // _4_67_ ksToleranceArrowType           - Тип стрелки ответвления у допуска формы
  // _4_68_ ksShowHideTmpObjTypeEnum       - Тип отображения временного объекта в документе.


//------------------------------------------------------------------------------- 
// _5_ Константы менеджера библиотек
// ---
  // _5_01_ LibraryTypeEnum          - Типы библиотек
  // _5_02_ LibraryStyleEnum         - Стили отображения прикладных библиотек.
  // _5_03_ ksSystemControlStartEnum - Результаты передачи управления системе КОМПАС


//-------------------------------------------------------------------------------
// _6_ Константы спецификации и атрибутов
// ---
  // _6_01_ SpecificationVariantEnum     - Варианты оформления спецификации.
  // _6_02_ SpecificationLinkTypeEnum    - Режимы связи сборки или чертежа со спецификацией.
  // _6_03_ SpecificationColumnTypeEnum  - Типы колонок спецификации.
  // _6_04_ ValueTypeEnum                - Типы значения атрибута, его колонок и колонок спецификации.
  // _6_05_ SortTypeEnum                 - Типы сортировки.
  // _6_06_ SpecificationObjectTypeEnum  - Типы объектов для спецификации.
  // _6_07_ SpecificationObjectStateEnum - Состояние объекта спецификации.
  // _6_08_ ksSpecificationStyleDifferenceTypeEnum - Отличие стиля СП от библиотечного.


//-------------------------------------------------------------------------------
// _7_ Константы дополнительных свойств объектов
// ---
  // _7_01_ ksPropertyTypeEnum - Типы свойств.
  // _7_02_ ksReportFiltersTypeEnum - Типы фильтров в команде создать отчет.
	 // _7_03_ ksVariantMarkingTypeEnum - Параметры формирования обозначения.
  // _7_04_ ksReportTypeEnum - Тип отчета
  // _7_05_ ksValueFormatEnum - Формат формирования текста для свойства.
  // _7_06_ ksGroupeTypeEnum  - Тип группировки колонки отчета.
  // _7_07_ ksNumberingTypeEnum - Формат нумерации.
  // _7_08_ ksRowsNumberingTypeEnum - Формат нумерации строк отчета.
  // _7_09_ ksReportStyleInitEnum - Способ инициализации стиля отчета.
  // _7_10_ ksReportBuildingTypeEnum - Способ выбора объектов для отчета.
  // _7_11_ ksPageLayoutTypeEnum - Тиn компоновки таблиц отчета.
  // _7_12_ ksObjectPropertyControlTypeEnum - Типы контролов для отображения свойств в окне свойств.


  //-------------------------------------------------------------------------------
  // _1_00_ Перечень интерфейсов событий
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("27ED970D-21D0-4a9c-9B83-B8CA54DC908E") 
enum _ksNotifyType
    {
        ntKompasObjectNotify	= 1,
        ntDocumentFileNotify	= 2,
        ntStampNotify	= 3,
        ntObject2DNotify	= 4,
        ntSelectionMngNotify	= 5,
        ntSpcObjectNotify	= 6,
        ntSpcDocumentNotify	= 7,
        ntSpecificationNotify	= 8,
        ntDocument3DNotify	= 9,
        ntObject3DNotify	= 10,
        ntDocument2DNotify	= 11,
        ntPropertyManagerNotify	= 12,
        ntPropertyUserControlNotifyEnum	= 13,
        ntDocumentFrameNotify	= 14,
        ntViewsAndLayersManagerNotify	= 15,
        ntLibraryManagerNotify	= 16,
        ntLayoutSheetsNotify	= 17,
        ntProcess2DNotify	= 18,
        ntProcess3DNotify	= 19
    } 	ksNotifyType;

  //-------------------------------------------------------------------------------
  // _1_01_ События для KompasObject и IApplication
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("5CD8F199-625F-4ad2-B490-ABEFF7A8FBD9") 
enum _ksKompasObjectNotifyEnum
    {
        koCreateDocument	= 1,
        koBeginOpenDocument	= 2,
        koOpenDocument	= 3,
        koChangeActiveDocument	= 4,
        koApplicatinDestroy	= 5,
        koBeginCreate	= 6,
        koBeginOpenFile	= 7,
        koBeginCloseAllDocument	= 8,
        koKeyDown	= 9,
        koKeyUp	= 10,
        koKeyPress	= 11,
        koBeginRequestFiles	= 12,
        koBeginChoiceMaterial	= 13,
        koChoiceMaterial	= 14
    } 	ksKompasObjectNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_02_ События для документов : Работа с файлом
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("C3652DD1-6086-4425-9A0B-A65D00E8F885") 
enum _ksDocumentFileNotifyEnum
    {
        kdBeginCloseDocument	= 1,
        kdCloseDocument	= 2,
        kdBeginSaveDocument	= 3,
        kdSaveDocument	= 4,
        kdActiveDocument	= 5,
        kdDeactiveDocument	= 6,
        kdBeginSaveAsDocument	= 7,
        kdDocumentFrameOpen	= 8,
        kdProcessActivate	= 9,
        kdProcessDeactivate	= 10,
        kdBeginProcess	= 11,
        kdEndProcess	= 12
    } 	ksDocumentFileNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_03_ События для основной надписи
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("F9121A45-E1DC-4dc9-853D-2FEA0DF86A04") 
enum _ksStampNotifyEnum
    {
        kdBeginEditStamp	= 1,
        kdEndEditStamp	= 2,
        kdStampCellDblClick	= 3,
        kdStampCellBeginEdit	= 4,
        kdStampBeginClearCells	= 5
    } 	ksStampNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_04_ События для объектов 2Д документов
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("0BF956B1-706D-4934-ACD7-D697507861CE") 
enum _ksObject2DNotifyEnum
    {
        koChangeActive	= 1,
        koBeginDelete	= 2,
        koDelete	= 3,
        koBeginMove	= 4,
        koMove	= 5,
        koBeginRotate	= 6,
        koRotate	= 7,
        koBeginScale	= 8,
        koScale	= 9,
        koBeginTransform	= 10,
        koTransform	= 11,
        koBeginCopy	= 12,
        koCopy	= 13,
        koBeginSymmetry	= 14,
        koSymmetry	= 15,
        koBeginProcess	= 16,
        koEndProcess	= 17,
        koCreateObject	= 18,
        koUpdateObject	= 19,
        koBeginDestroyObject	= 20,
        koDestroyObject	= 21,
        koBeginPropertyChanged	= 22,
        koPropertyChanged	= 23
    } 	ksObject2DNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_05_ События для менеджера селектирования
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("31C97EDA-5954-4e22-9097-778F7996690F") 
enum _ksSelectionMngNotifyEnum
    {
        ksmSelect	= 1,
        ksmUnselect	= 2,
        ksmUnselectAll	= 3
    } 	ksSelectionMngNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_06_ События для объекта спецификации
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("A8F27F92-9C2C-44bf-9175-EAB3B5DD8CA9") 
enum _ksSpcObjectNotifyEnum
    {
        soBeginDelete	= 1,
        soDelete	= 2,
        soCellDblClick	= 3,
        soCellBeginEdit	= 4,
        soChangeCurrent	= 5,
        soDocumentBeginAdd	= 6,
        soDocumentAdd	= 7,
        soDocumentRemove	= 8,
        soBeginGeomChange	= 9,
        soGeomChange	= 10,
        soBeginProcess	= 11,
        soEndProcess	= 12,
        soCreateObject	= 13,
        soUpdateObject	= 14,
        soBeginCopy	= 15,
        soCopy	= 16
    } 	ksSpcObjectNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_07_ События для документа спецификации
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6D3E732D-86AC-4d20-8BF1-6B041DB989B1") 
enum _ksSpcDocumentNotifyEnum
    {
        sdDocumentBeginAdd	= 1,
        sdDocumentAdd	= 2,
        sdDocumentBeginRemove	= 3,
        sdDocumentRemove	= 4,
        sdSpcStyleBeginChange	= 5,
        sdSpcStyleChange	= 6
    } 	ksSpcDocumentNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_08_ События для спецификации
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("72F076AF-7103-4718-A9AC-4A5324C6B076") 
enum _ksSpecificationNotifyEnum
    {
        ssTuningSpcStyleBeginChange	= 1,
        ssTuningSpcStyleChange	= 2,
        ssChangeCurrentSpcDescription	= 3,
        ssSpcDescriptionAdd	= 4,
        ssSpcDescriptionRemove	= 5,
        ssSpcDescriptionBeginEdit	= 6,
        ssSpcDescriptionEdit	= 7,
        ssSynchronizationBegin	= 8,
        ssSynchronization	= 9,
        ssBeginCalcPositions	= 10,
        ssCalcPositions	= 11,
        ssBeginCreateObject	= 12
    } 	ksSpecificationNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_09_ События для 3D документа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("CFB5B381-6180-4c14-9CAB-188C8CBAF865") 
enum _ksDocument3DNotifyEnum
    {
        d3BeginRebuild	= 1,
        d3Rebuild	= 2,
        d3BeginChoiceMaterial	= 3,
        d3ChoiceMaterial	= 4,
        d3BeginChoiceMarking	= 5,
        d3ChoiceMarking	= 6,
        d3BeginSetPartFromFile	= 7,
        d3BeginCreatePartFromFile	= 8,
        d3CreateEmbodiment	= 9,
        d3DeleteEmbodiment	= 10,
        d3ChangeCurrentEmbodiment	= 11,
        d3BeginChoiceProperty	= 12,
        d3ChoiceProperty	= 13
    } 	ksDocument3DNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_10_ События для объекта 3D документа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B4BCC48A-5F67-48e6-8412-78E244E37FB3") 
enum _ksObject3DNotifyEnum
    {
        o3BeginDelete	= 1,
        o3Delete	= 2,
        o3Excluded	= 3,
        o3Hidden	= 4,
        o3BeginPropertyChanged	= 5,
        o3PropertyChanged	= 6,
        o3BeginPlacementChanged	= 7,
        o3PlacementChanged	= 8,
        o3BeginProcess	= 9,
        o3EndProcess	= 10,
        o3CreateObject	= 11,
        o3UpdateObject	= 12,
        o3ChangeCurrent	= 13
    } 	ksObject3DNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_11_ События для 2D документа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("ABD2E40C-5964-4b06-A6A8-281BA3A172E9") 
enum _ksDocument2DNotifyEnum
    {
        d2BeginRebuild	= 1,
        d2Rebuild	= 2,
        d2BeginChoiceMaterial	= 3,
        d2ChoiceMaterial	= 4,
        d2BeginInsertFragment	= 5,
        d2LocalFragmentEdit	= 6,
        d2BeginChoiceProperty	= 7,
        d2ChoiceProperty	= 8
    } 	ksDocument2DNotifyEnum;

  //-------------------------------------------------------------------------------
  // _2_02_ Тип отрисовки визуальной части
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("1C48894B-7109-4975-AD3F-9FB0F819C151") 
enum _ksTypeLookStyle
    {
        tls_VisualStudio_97	= 0,
        tls_VisualStudio_NET	= 1,
        tls_Office_2003	= 2,
        tls_VisualStudio2005	= 3,
        tls_WindowsXP	= 4,
        tls_Office_2007	= 5,
        tls_Office_2007_LunaBlue	= 5,
        tls_Office_2007_ObsidianBlack	= 6,
        tls_Office_2007_Aqua	= 7,
        tls_Office_2007_Silver	= 8,
        tls_VisualStudio2008	= 9,
        tls_VisualStudio2010	= 10,
        tls_Office_2010_Blue	= 11,
        tls_Office_2010_Dark	= 12,
        tls_Office_2010_White	= 13,
        tls_Carbon	= 14
    } 	ksTypeLookStyle;

  //-------------------------------------------------------------------------------
  // _1_12_ События панели свойств или процессных параметров.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B3AF1092-5D34-46f2-B63A-46F51FAC400F") 
enum _ksPropertyManagerNotifyEnum
    {
        prButtonClick	= 1,
        prChangeControlValue	= 2,
        prControlCommand	= 3,
        prButtonUpdate	= 4,
        prProcessActivate	= 5,
        prProcessDeactivate	= 6,
        prCommandHelp	= 7,
        prSelectItem	= 8,
        prCheckItem	= 9,
        prChangeActiveTab	= 10,
        prEditFocus	= 11,
        prUserMenuCommand	= 12,
        prLayoutChanged	= 13,
        prGetContextMenuType	= 14,
        prFillContextPanel	= 15,
        prFillContextIconMenu	= 16,
        prEndEditItem	= 17,
        prChangeTabExpanded	= 18
    } 	ksPropertyManagerNotifyEnum;

  //-------------------------------------------------------------------------------
  //  _3_02_ Предопределенные кнопки  панели свойств
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("0EDBCD89-5DC7-4d27-B2AC-94D7C911706C") 
enum _SpecPropertyButtonEnum
    {
        pbEnter	= 1,
        pbEsc	= 2,
        pbAutoCreate	= 3,
        pbSaveState	= 4,
        pbNewSearch	= 5,
        pbPrevObj	= 6,
        pbNextObj	= 7,
        pbHelp	= 8,
        pbCopyProps	= 9
    } 	SpecPropertyButtonEnum;

  //-------------------------------------------------------------------------------
  // _3_03_ Предопределенные спецпанели для панели свойств
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("0C773864-1BB5-4fd0-8BB9-2072234C8D8D") 
enum _SpecPropertyToolBarEnum
    {
        pnUnknown	= 0,
        pnEmpty	= 1,
        pnEscHelp	= 2,
        pnEnterEscHelp	= 3,
        pnEnterEscCreateHelp	= 4,
        pnEnterEscCreateSaveHelp	= 5,
        pnEnterEscCreateSaveSearchHelp	= 6,
        pnEnterEscSaveSearchPrevNextHelp	= 7,
        pnEnterEscSearchHelp	= 8,
        pnEscSaveSearchHelp	= 9,
        pnEnterEscCreateSearchHelp	= 10,
        pnEnterEscSaveSearchHelp	= 11,
        pnEscSaveStateHelp	= 12,
        pnEnterEscSearchPrevNextHelp	= 13
    } 	SpecPropertyToolBarEnum;

//-------------------------------------------------------------------------------
// _2_01_ Типы объектов КОМПАС API 7.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("1EBC1F32-2CD5-407d-993A-61AAC88D7F80") 
enum _KompasAPIObjectTypeEnum
    {
        ksObjectUnknown	= 0,
        ksObjectApplication	= 10001,
        ksObjectDocuments	= 10002,
        ksObjectKompasError	= 10003,
        ksObjectProcessParam	= 10004,
        ksObjectPropertyTabs	= 10005,
        ksObjectPropertyTab	= 10006,
        ksObjectPropertyControls	= 10007,
        ksObjectPropertySeparator	= 10008,
        ksObjectPropertyEdit	= 10009,
        ksObjectPropertyList	= 10010,
        ksObjectPropertyCheckBox	= 10011,
        ksObjectPropertyMultiButton	= 10012,
        ksObjectPropertyUserControl	= 10013,
        ksObjectPropertyBmpList	= 10014,
        ksObjectPropertySlideBox	= 10016,
        ksObjectPropertyGrid	= 10017,
        ksObjectDocumentFrame	= 10018,
        ksObjectDocumentFrames	= 10019,
        ksObjectPropertyManager	= 10020,
        ksObjectDrawingDocument	= 10021,
        ksObjectFragmentDocument	= 10022,
        ksObjectSpecificationDocument	= 10023,
        ksObjectPartDocument	= 10024,
        ksObjectAssemblyDocument	= 10025,
        ksObjectTextDocument	= 10026,
        ksObjectPropertyTextButton	= 10027,
        ksObjectPropertySpinEdit	= 10028,
        ksObjectViewsAndLayersManager	= 10029,
        ksObjectViews	= 10030,
        ksObjectView	= 10031,
        ksObjectAssociationView	= 10032,
        ksObjectLayerGroups	= 10033,
        ksObjectLayerGroup	= 10035,
        ksObjectLayers	= 10036,
        ksObjectLayer	= 10037,
        ksObjectLayerFilterCondition	= 10038,
        ksObjectLayerFilterConditions	= 10039,
        ksObjectDocumentSettings	= 10040,
        ksObjectDocument2DSettings	= 10041,
        ksObjectDrawingDocumentSettings	= 10042,
        ksObjectFragmentDocumentSettings	= 10043,
        ksObjectSystemSettings	= 10044,
        ksObjectDocument3DSettings	= 10045,
        ksObjectLibraryManager	= 10050,
        ksObjectProcedure	= 10051,
        ksObjectProceduresLibraries	= 10052,
        ksObjectProceduresLibrary	= 10053,
        ksObjectProcedures	= 10054,
        ksObjectInsertsLibraries	= 10055,
        ksObjectInsertsLibrary	= 10056,
        ksObjectInserts	= 10057,
        ksObjectInsert	= 10058,
        ksObjectSpecificationDescriptions	= 10059,
        ksObjectSpecificationDescription	= 10060,
        ksObjectSpecificationStyle	= 10061,
        ksObjectSpecificationColumnStyles	= 10062,
        ksObjectSpecificationColumnStyle	= 10063,
        ksObjectSpecificationSectionStyles	= 10064,
        ksObjectSpecificationSectionStyle	= 10065,
        ksObjectSpecificationTuning	= 10066,
        ksObjectSpecificationTuningSections	= 10067,
        ksObjectSpecificationTuningSection	= 10068,
        ksObjectSpecificationSubsections	= 10069,
        ksObjectSpecificationSubsection	= 10070,
        ksObjectAdditionalBlockStyles	= 10071,
        ksObjectAdditionalBlockStyle	= 10072,
        ksObjectAdditionalBlockTunings	= 10073,
        ksObjectAdditionalBlockTuning	= 10074,
        ksObjectAdditionalBlockSectionTunings	= 10075,
        ksObjectAdditionalBlockSectionTuning	= 10076,
        ksObjectSheetFormat	= 10077,
        ksObjectTextStyle	= 10078,
        ksObjectFont	= 10079,
        ksObjectTabulators	= 10080,
        ksObjectTabulator	= 10081,
        ksObjectSpecificationBaseObjects	= 10083,
        ksObjectSpecificationCommentObjects	= 10084,
        ksObjectSpecificationObject	= 10085,
        ksObjectSpecificationBaseObject	= 10086,
        ksObjectSpecificationCommentObject	= 10087,
        ksObjectSpecificationColumns	= 10088,
        ksObjectSpecificationColumn	= 10089,
        ksObjectSpecificationColumnItems	= 10090,
        ksObjectSpecificationColumnItem	= 10091,
        ksObjectAttachedDocuments	= 10092,
        ksObjectAttachedDocument	= 10093,
        ksObjectPropertyFileName	= 10094,
        ksObjectPropertyColor	= 10095,
        ksObjectLayoutSheets	= 10096,
        ksObjectLayoutSheet	= 10097,
        ksObjectConverter	= 10098,
        ksObjectCheckSum	= 10099,
        ksObjectProgressBarIndicator	= 10100,
        ksObjectPropertyEditList	= 10101,
        ksObjectPropertyLibExplorer	= 10102,
        ksObjectVariable7	= 10103,
        ksObjectInsertionParameters	= 10104,
        ksObjectMath2D	= 10105,
        ksObjectSelectionManager	= 10106,
        ksObjectChooseManager	= 10107,
        ksObjectStamp	= 10108,
        ksObjectPropertyStyleList	= 10109,
        ksObjectPrintJob	= 10110,
        ksObjectPrintJob_Sheet	= 10111,
        ksObjectPropertyOpticalProps	= 10112,
        ksObjectPropertyEditCheckBox	= 10113,
        ksObjectNewPartDocumentSettings	= 10114,
        ksObjectNewAssemblyDocumentSettings	= 10115,
        ksObjectPropertyGroupBegin	= 10116,
        ksObjectPropertyGroupEnd	= 10117,
        ksObjectThreadPattern	= 10118,
        ksObjectThreadDialogParam	= 10119,
        ksObjectPropertyPreviewText	= 10120,
        ksObjectPropertyAggregateControl	= 10121,
        ksObjectPropertyBasePoint	= 10121,
        ksObjectPropertyLinkButton	= 10122,
        ksObjectPropertyMarking	= 10123,
        ksObjectUserDataStoragesMng	= 10500,
        ksObjectUserDataStorages	= 10501,
        ksObjectUserDataStorage	= 10502,
        ksObjectParametriticConstraint	= 10503,
        ksObjectAttribute	= 10504,
        ksObjectColumnInfo	= 10505,
        ksObjectAttributeType	= 10506,
        ksObjectAttrTypeMng	= 10507,
        ksObjectProperty	= 10508,
        ksObjectPropertyMng	= 10509,
        ksObjectReportProcess	= 10510,
        ksObjectReport	= 10511,
        ksObjectReportStyle	= 10512,
        ksObjectReportStyleColumn	= 10513,
        ksObjectText	= 10700,
        ksObjectTextLine	= 10701,
        ksObjectTextItem	= 10702,
        ksObjectTableCell	= 10703,
        ksObjectExternalTessellation	= 10704,
        ksObjectContour	= 10705,
        ksObjectContourSegment	= 10706,
        ksObjectContourLineSegment	= 10707,
        ksObjectContourArc	= 10708,
        ksObjectExternalGDI	= 10709,
        ksObjectLibArraySettings	= 10710,
        ksObjectTextTable	= 10711,
        ksObjectSnapInfo	= 10712,
        ksObjectPropertyTwinSwitcher	= 10713,
        ksObjectPropertyPoint3D	= 10714,
        ksObjectPart7	= 11000,
        ksObjectModelObject	= 11001,
        ks3dMateConstraint	= 11002,
        ksObjectParts7	= 11003,
        ksObjectVariableTable	= 11004,
        ksObjectSketchs	= 11005,
        ksObjectSketch	= 11006,
        ksObjectExtrussions	= 11007,
        ksObjectExtrusion	= 11008,
        ksObjectBody7	= 11009,
        ksObjectCutExtrusion	= 11010,
        ksObjectSheetMetalBodies	= 11011,
        ksObjectSheetMetalBody	= 11012,
        ksObjectPoints3D	= 11013,
        ksObjectPoint3D	= 11014,
        ksObjectPoint3DParamCoord	= 11015,
        ksObjectPoint3DParamDisplace	= 11016,
        ksObjectPoint3DParamIntersect	= 11017,
        ksObjectPoint3DParamCenter	= 11018,
        ksObjectPoint3DParamCurve	= 11019,
        ksObjectPoint3DParamSurface	= 11020,
        ksObjectPoint3DParamProjection	= 11021,
        ksObjectSheetMetalBend	= 11022,
        ksObjectSheetMetalBends	= 11023,
        ksObjectSheetMetalLineBend	= 11024,
        ksObjectSheetMetalLineBends	= 11025,
        ksObjectSheetMetalHole	= 11026,
        ksObjectSheetMetalHoles	= 11027,
        ksObjectSheetMetalCut	= 11028,
        ksObjectSheetMetalCuts	= 11029,
        ksObjectLineDimensions3D	= 11030,
        ksObjectBaseLineDimension3D	= 11031,
        ksObjectLineDimension3D	= 11032,
        ksObjectRadialDimension3D	= 11033,
        ksObjectDiametralDimension3D	= 11034,
        ksObjectRadialDimensions3D	= 11035,
        ksObjectDiametralDimensions3D	= 11036,
        ksObjectAngleDimension3D	= 11037,
        ksObjectAngleDimensions3D	= 11038,
        ksObjectLocalCoordinateSystems	= 11039,
        ksObjectLocalCoordinateSystem	= 11040,
        ksObjectLocalCSAxesDirectionParam	= 11041,
        ksObjectLocalCSOrientByObjectParam	= 11042,
        ksObjectLocalCSEulerParam	= 11043,
        ksObjectSpline3D	= 11044,
        ksObjectSplines3D	= 11045,
        ksObjectCurveVertexParam	= 11046,
        ksObjectPolyLines	= 11047,
        ksObjectPolyLine	= 11048,
        ksObjectLeaders3D	= 11049,
        ksObjectLeader3D	= 11050,
        ksObjectMarkLeader3D	= 11051,
        ksObjectRough3D	= 11052,
        ksObjectRoughs3D	= 11053,
        ksObjectPositionLeader3D	= 11054,
        ksObjectBrandLeader3D	= 11055,
        ksObjectBase3D	= 11056,
        ksObjectBases3D	= 11057,
        ksObjectTolerances3D	= 11058,
        ksObjectTolerance3D	= 11059,
        ksObjectControlPoints	= 11060,
        ksObjectControlPoint	= 11061,
        ksObjectConjunctivePoints	= 11062,
        ksObjectConjunctivePoint	= 11063,
        ksObjectSplitLines	= 11064,
        ksObjectSplitLine	= 11065,
        ksObjectSurfacePatches	= 11066,
        ksObjectSurfacePatch	= 11067,
        ksObjectFaceRemovers	= 11068,
        ksObjectFaceRemover	= 11069,
        ksObjectSurfaceSewers	= 11070,
        ksObjectSurfaceSewer	= 11071,
        ksObjectNurbsSurfaces	= 11072,
        ksObjectNurbsSurface	= 11073,
        ksObjectSurfacesIntersectionCurves	= 11074,
        ksObjectSurfacesIntersectionCurve	= 11075,
        ksObjectEquidistants3D	= 11076,
        ksObjectEquidistant3D	= 11077,
        ksObjectTrimmedCurves	= 11078,
        ksObjectTrimmedCurve	= 11079,
        ksObjectFeaturePatterns	= 11080,
        ksObjectLinearPattern	= 11081,
        ksObjectCircularPattern	= 11082,
        ksObjectPathPattern	= 11083,
        ksObjectPartsLinearPattern	= 11084,
        ksObjectPartsCircularPattern	= 11085,
        ksObjectPartsPathPattern	= 11086,
        ksObjectAuxLinearPattern	= 11087,
        ksObjectAuxCircularPattern	= 11088,
        ksObjectAuxPathPattern	= 11089,
        ksObjectPointDrivenPattern	= 11090,
        ksObjectPartsPointDrivenPattern	= 11091,
        ksObjectDerivedPattern	= 11092,
        ksObjectMirrorPattern	= 11093,
        ksObjectShellMirrorPattern	= 11094,
        ksObjectAuxMirrorPattern	= 11095,
        ksObjectRuledSurfaces	= 11096,
        ksObjectRuledSurface	= 11097,
        ksObjectExtensionSurfaces	= 11098,
        ksObjectExtensionSurface	= 11099,
        ksObjectEquidistantSurfaces	= 11100,
        ksObjectEquidistantSurface	= 11101,
        ksObjectTrimmedSurfaces	= 11102,
        ksObjectTrimmedSurface	= 11103,
        ksObjectVector3D	= 11104,
        ksObjectVector3DBy2VertexesParameters	= 11105,
        ksObjectVector3DByCoefficientsParameters	= 11106,
        ksObjectVector3DBy2AnglesParameters	= 11107,
        ksObjectVector3DByObjectParameters	= 11108,
        ksObjectVector3DAlongSurfaceNormalParameters	= 11109,
        ksObjectVector3DByCurveParameters	= 11110,
        ksObjectVector3DByScreenNormalParameters	= 11111,
        ksObjectVector3DByLocalCSParameters	= 11112,
        ksObjectConnectCurves	= 11113,
        ksObjectConnectCurve	= 11114,
        ksObjectFilletCurves	= 11115,
        ksObjectFilletCurve	= 11116,
        ksObjectSurfaceThickenings	= 11117,
        ksObjectSurfaceThickening	= 11118,
        ksObjectArcs3D	= 11119,
        ksObjectArc3D	= 11120,
        ksObjectAuxPointDrivenPattern	= 11121,
        ksObjectBodiesPointDrivenPattern	= 11122,
        ksObjectTablePattern	= 11123,
        ksObjectPartsTablePattern	= 11124,
        ksObjectAuxTablePattern	= 11125,
        ksObjectBodiesTablePattern	= 11126,
        ksObjectRotateds	= 11127,
        ksObjectRotated	= 11128,
        ksObjectCutRotated	= 11129,
        ksObjectExtrusionSurfaces	= 11130,
        ksObjectExtrusionSurface	= 11131,
        ksObjectRotatedSurfaces	= 11132,
        ksObjectRotatedSurface	= 11133,
        ksObjectPoint3DParamBySphere	= 11134,
        ksObjectPoint3DParamByCylinder	= 11135,
        ksObjectMeshPointsSurfaces	= 11136,
        ksObjectMeshPointsSurface	= 11137,
        ksObjectCloudPointsSurfaces	= 11138,
        ksObjectCloudPointsSurface	= 11139,
        ksObjectImportedSurfaces	= 11140,
        ksObjectImportedSurface	= 11141,
        ksObjectBodiesLinearPattern	= 11142,
        ksObjectBodiesCircularPattern	= 11143,
        ksObjectBodiesPathPattern	= 11144,
        ksObjectScalings3D	= 11145,
        ksObjectScaling3D	= 11146,
        ksObjectCurveOutLine	= 11147,
        ksObjectCurveOutLines	= 11148,
        ksObjectCurveByLaw	= 11149,
        ksObjectCurveByLaws	= 11150,
        ksObjectIsoparametricCurve	= 11151,
        ksObjectIsoparametricCurves	= 11152,
        ksObjectIsoparametricCurvesSet	= 11153,
        ksObjectIsoparametricCurvesSets	= 11154,
        ksObjectSplineOnSurface	= 11155,
        ksObjectSplinesOnSurfaces	= 11156,
        ksObjectProjectionCurve	= 11157,
        ksObjectProjectionCurves	= 11158,
        ksObjectCurveBy2Projections	= 11159,
        ksObjectCurvesBy2Projectionses	= 11160,
        ksObjectContour3D	= 11161,
        ksObjectContours3D	= 11162,
        ksObjectLineSegment3D	= 11163,
        ksObjectLineSegments3D	= 11164,
        ksObjectUnhistoredCurve3D	= 11165,
        ksObjectUnhistoredCurves3D	= 11166,
        ksObjectTechnicalDemand3D	= 11167,
        ksObjectEmbodiment	= 11168,
        ksObjectOperationLinearDimension	= 11169,
        ksObjectOperationAngularDimension	= 11170,
        ksObjectOperationRadialDimension	= 11171,
        ksObjectOperationDiametralDimension	= 11172,
        ksObjectSketchLinearDimension	= 11173,
        ksObjectSketchAngularDimension	= 11174,
        ksObjectSketchBreakAngularDimension	= 11175,
        ksObjectSketchRadialDimension	= 11176,
        ksObjectSketchBreakRadialDimension	= 11177,
        ksObjectSketchDiametralDimension	= 11178,
        ksObjectConicSpiral3D	= 11179,
        ksObjectCylindricSpiral3D	= 11180,
        ksObjectSpirals3D	= 11181,
        ksObjectPointsArrOnCurve	= 11182,
        ksObjectPointsArrsOnCurves	= 11183,
        ksObjectPointsArrOnSurface	= 11184,
        ksObjectPointsArrsOnSurfaces	= 11185,
        ksObjectPointsArrFromFile	= 11186,
        ksObjectPointsArrsFromFiles	= 11187,
        ksObjectAxis3D	= 11188,
        ksObjectAxes3D	= 11189,
        ksObjectAxis3DBy2Points	= 11190,
        ksObjectAxis3DBy2Planes	= 11191,
        ksObjectAxis3DByConeface	= 11192,
        ksObjectAxis3DByEdge	= 11193,
        ksObjectAxis3DByPointAndObject	= 11194,
        ksObjectAxis3DByOperation	= 11195,
        ksObjectPlanes3D	= 11196,
        ksObjectPlane3D	= 11197,
        ksObjectPlane3DByPlaneCurve	= 11198,
        ksObjectPlane3DTangentToFaceInPoint	= 11199,
        ksObjectPlane3DByOffset	= 11200,
        ksObjectPlane3DBy3Points	= 11201,
        ksObjectPlane3DByAngle	= 11202,
        ksObjectPlane3DByEdgeAndPoint	= 11203,
        ksObjectPlane3DParallelByPoint	= 11204,
        ksObjectPlane3DPerpendicularByEdge	= 11205,
        ksObjectPlane3DNormalToSurface	= 11206,
        ksObjectPlane3DMiddle	= 11207,
        ksObjectPlane3DByEdgeAndPlane	= 11208,
        ksObjectPlane3DBy2Edge	= 11209,
        ksObjectPlane3DTangentToFace	= 11210,
        ksObjectUserObject3D	= 11211,
        ksObjectUserObjects3D	= 11212,
        ksObjectLinearPatternAnyCopy	= 11213,
        ksObjectCircularPatternAnyCopy	= 11214,
        ksObjectPathPatternAnyCopy	= 11215,
        ksObjectPointDrivenPatternAnyCopy	= 11217,
        ksObjectTablePatternAnyCopy	= 11218,
        ksObjectLinearUnhistoriedDimension	= 11219,
        ksObjectAngularUnhistoriedDimension	= 11220,
        ksObjectRadialUnhistoriedDimension	= 11221,
        ksObjectDiametralUnhistoriedDimension	= 11222,
        ksObjectPlacement3D	= 11223,
        ksObjectLayers3D	= 11224,
        ksObjectLayer3D	= 11225,
        ksObjectLayerGroups3D	= 11226,
        ksObjectLayerGroup3D	= 11227,
        ksObjectDocument3DManager	= 11228,
        ksObjectToleranceRecalc	= 11229,
        ksObjectSpecRough3D	= 11230,
        ksObjectSketchBreakLinearDimension	= 11231,
        ksObjectMathCurve3D	= 11232,
        ksObjectMathSurface3D	= 11233,
        ksObjectBilletObsolete	= 11234,
        ksObjectBilletsObsoletes	= 11235,
        ksObjectCopyGeometry	= 11236,
        ksObjectCopiesGeometry	= 11237,
        ksObjectCollectionGeometry	= 11238,
        ksObjectCollectionsGeometry	= 11239,
        ksObjectUserWireFrame3D	= 11240,
        ksObjectThreads	= 11241,
        ksObjectThread	= 11242,
        ks3dMateConstraints	= 11243,
        ksMate3DByAngle	= 11244,
        ksMate3DByTangent	= 11245,
        ksMate3DSymmetry	= 11246,
        ksMate3DTransmission	= 11247,
        ksMate3DCamGear	= 11248,
        ksMate3DDependentPosition	= 11249,
        ksObjectHoles3D	= 11250,
        ksObjectHole3D	= 11251,
        ksObjectCountersinkHoleParameters	= 11252,
        ksObjectSpotfacingHoleParameters	= 11253,
        ksObjectCountersinkSpotfacingHoleParameters	= 11254,
        ksObjectConicHoleParameters	= 11255,
        ksObjectDrawingObject	= 13000,
        ksObjectDrawingText	= 13001,
        ksObjectDrawingTexts	= 13002,
        ksObjectStraightAxis	= 13003,
        ksObjectCircleAxis	= 13004,
        ksObjectArcAxis	= 13005,
        ksObjectBuildingAxes	= 13006,
        ksObjectAxisJut	= 13007,
        ksObjectMarkNodes	= 13008,
        ksObjectMarkNode	= 13009,
        ksObjectMarkOnLeader	= 13010,
        ksObjectMarkOnLine	= 13011,
        ksObjectMarkInsideForm	= 13012,
        ksObjectMarks	= 13013,
        ksObjectCutUnitMarking	= 13014,
        ksObjectCutUnitMarkings	= 13015,
        ksObjectUnitMarking	= 13016,
        ksObjectUnitMarkings	= 13017,
        ksObjectUnitNumber	= 13018,
        ksObjectUnitNumbers	= 13019,
        ksObjectMultiTextLeader	= 13020,
        ksObjectMultiTextLeaders	= 13021,
        ksObjectBrace	= 13022,
        ksObjectBraces	= 13023,
        ksObjectLineSegments	= 13024,
        ksObjectLineSegment	= 13025,
        ksObjectArcs	= 13026,
        ksObjectArc	= 13027,
        ksObjectRadialDimension	= 13028,
        ksObjectDiametralDimension	= 13029,
        ksObjectBreakRadialDimension	= 13030,
        ksObjectRadialDimensions	= 13031,
        ksObjectBreakRadialDimensions	= 13032,
        ksObjectDiametralDimensions	= 13033,
        ksObjectLineDimension	= 13041,
        ksObjectLineDimensions	= 13042,
        ksObjectBreakLineDimension	= 13043,
        ksObjectBreakLineDimensions	= 13044,
        ksObjectHeightDimension	= 13045,
        ksObjectHeightDimensions	= 13046,
        ksObjectAngleDimension	= 13047,
        ksObjectBreakAngleDimension	= 13048,
        ksObjectAngleDimensions	= 13049,
        ksObjectArcDimension	= 13050,
        ksObjectArcDimensions	= 13051,
        ksObjectLeader	= 13052,
        ksObjectLeaders	= 13053,
        ksObjectRough	= 13054,
        ksObjectRoughs	= 13055,
        ksObjectMarkLeader	= 13056,
        ksObjectBrandLeader	= 13057,
        ksObjectPositionLeader	= 13058,
        ksObjectChangeLeader	= 13059,
        ksObjectBase	= 13060,
        ksObjectBases	= 13061,
        ksObjectDrawingTable	= 13062,
        ksObjectDrawingTables	= 13063,
        ksObjectTolerance	= 13064,
        ksObjectTolerances	= 13065,
        ksObjectCutLines	= 13066,
        ksObjectCutLine	= 13067,
        ksObjectViewPointer	= 13068,
        ksObjectViewPointers	= 13069,
        ksObjectMultiline	= 13070,
        ksObjectMultilines	= 13071,
        ksObjectDrawingContour	= 13072,
        ksObjectDrawingContours	= 13073,
        ksObjectCircles	= 13074,
        ksObjectCircle	= 13075,
        ksObjectPoints	= 13076,
        ksObjectPoint	= 13077,
        ksObjectBeziers	= 13078,
        ksObjectBezier	= 13079,
        ksObjectMacroObjects	= 13080,
        ksObjectMacroObject	= 13081,
        ksObjectLines	= 13082,
        ksObjectLine	= 13083,
        ksObjectPolyLines2D	= 13084,
        ksObjectPolyLine2D	= 13085,
        ksObjectNurbses	= 13086,
        ksObjectNurbs	= 13087,
        ksObjectRasters	= 13088,
        ksObjectRaster	= 13089,
        ksObjectOleDrawingObjects	= 13090,
        ksObjectOleDrawingObject	= 13091,
        ksObjectEllipses	= 13092,
        ksObjectEllipse	= 13093,
        ksObjectEllipseArcs	= 13094,
        ksObjectEllipseArc	= 13095,
        ksObjectRectangles	= 13096,
        ksObjectRectangle	= 13097,
        ksObjectRegularPolygons	= 13098,
        ksObjectRegularPolygon	= 13099,
        ksObjectEquidistants	= 13100,
        ksObjectEquidistant	= 13101,
        ksObjectInsertionFragment	= 13102,
        ksObjectInsertionView	= 13103,
        ksObjectInsertionObjects	= 13104,
        ksObjectInsertionDefinition	= 13105,
        ksObjectCentreMarkers	= 13106,
        ksObjectCentreMarker	= 13107,
        ksObjectRemoteElements	= 13108,
        ksObjectRemoteElement	= 13109,
        ksObjectAxisLines	= 13110,
        ksObjectAxisLine	= 13111,
        ksObjectHatchParam	= 13112,
        ksObjectDrawingGroup	= 13113,
        ksObjectDrawingGroups	= 13114,
        ksObjectCurve2D	= 13115,
        ksObjectHatches	= 13116,
        ksObjectHatch	= 13117,
        ksObjectColourings	= 13118,
        ksObjectColouring	= 13119,
        ksObjectSpecRough	= 13120,
        ksObjectTechnicalDemand	= 13121,
        ksObjectAnnotativeLineSegment	= 13122,
        ksObjectAnnotativeCircle	= 13123,
        ksObjectAnnotativeEllipse	= 13124,
        ksObjectAnnotativeArc	= 13125,
        ksObjectAnnotativeEllipseArc	= 13126,
        ksObjectAnnotativePolyLine	= 13127,
        ksObjectAnnotativePoint	= 13128,
        ksObjectAnnotativeText	= 13129,
        ksObjectBuildingCutLines	= 13131,
        ksObjectBuildingCutLine	= 13132,
        ksObjectWaveLines	= 13133,
        ksObjectWaveLine	= 13134,
        ksObjectBrokenLines	= 13135,
        ksObjectBrokenLine	= 13136,
        ksObjectCopyObjectParam	= 13137,
        ksObjectCurveCopyObjectParam	= 13138,
        ksObjectCircleCopyObjectParam	= 13139,
        ksObjectCircularCopyObjectParam	= 13140,
        ksObjectMeshCopyObjectParam	= 13141,
        ksObjectLocalCoordinateSystem2D	= 13142,
        ksObjectLocalCoordinateSystems2D	= 13143,
        ksObjectAttachedLeader	= 13144,
        ksObjectAttachedLeaders	= 13145,
        ksObjectLoadCombinationsParam	= 13146,
        ksObjectOpenDocumentParam	= 13147,
        ksObjectAssociationTables	= 13148,
        ksObjectAssociationTable	= 13149,
        ksObjectNurbsesByPoints	= 13150,
        ksObjectNurbsByPoints	= 13151,
        ksObjectProcess2D	= 13152,
        ksObjectPhantom2D	= 13153,
        ksObjectProcess3D	= 13154
    } 	KompasAPIObjectTypeEnum;

//-------------------------------------------------------------------------------
// _3_01_ Типы контролов панели свойств
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("8406BD38-D629-4fe3-975C-EE266499C913") 
enum ControlTypeEnum
    {
        ksControlUnknown	= 0,
        ksControlSeparator	= 1,
        ksControlEditInt	= 2,
        ksControlEditReal	= 3,
        ksControlEditStr	= 4,
        ksControlListInt	= 5,
        ksControlListReal	= 6,
        ksControlListStr	= 7,
        ksControlCheckBox	= 8,
        ksControlMultiButton	= 9,
        ksControlGrid	= 10,
        ksControlSlideBox	= 11,
        ksControlUser	= 12,
        ksControlTextButton	= 13,
        ksControlSpinInt	= 14,
        ksControlSpinReal	= 15,
        ksControlFileName	= 16,
        ksControlColor	= 17,
        ksControlEditList	= 18,
        ksControlEditLength	= 19,
        ksControlEditAngle	= 20,
        ksControlEditPoint	= 21,
        ksControlListLength	= 22,
        ksControlListAngle	= 23,
        ksControlBmpList	= 24,
        ksControlLibExplorer	= 25,
        ksControlListScale	= 26,
        ksControlLineStyle	= 27,
        ksControlOpticalProps	= 28,
        ksControlEditCheckBox	= 29,
        ksControlPointStyle	= 30,
        ksControlPointStyle3D	= 31,
        ksControlLineStyle3D	= 32,
        ksControlHatchStyle	= 33,
        ksControlGroupBegin	= 34,
        ksControlGroupEnd	= 35,
        ksControlTwinSwitcher	= 36,
        ksControlPoint3D	= 37,
        ksControlPreviewText	= 38,
        ksControlAggregateTriple	= 40,
        ksControlBasePoint	= 41,
        ksControlLinkButton	= 42,
        ksControlMarking	= 43
    } 	ControlTypeEnum;

//-------------------------------------------------------------------------------
// _2_03_ Типы документов КОМПАС
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("60BA0C60-7328-11d6-9B72-00A0D2A435C0") 
enum _DocumentTypeEnum
    {
        ksDocumentUnknown	= 0,
        ksDocumentDrawing	= 1,
        ksDocumentFragment	= 2,
        ksDocumentSpecification	= 3,
        ksDocumentPart	= 4,
        ksDocumentAssembly	= 5,
        ksDocumentTextual	= 6,
        ksDocumentTechnologyAssembly	= 7
    } 	DocumentTypeEnum;

//-------------------------------------------------------------------------------
// _3_04_ Типы разделителя (сепаратора)
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("8406BD38-D629-4fe3-975C-EE266499C914") 
enum _SeparatorTypeEnum
    {
        ksSeparatorDownName	= 0,
        ksSeparatorUpName	= 1,
        ksSeparatorWithoutLine	= 2,
        ksSeparatorBMPLeftName	= 3,
        ksSeparatorBMPRightName	= 4
    } 	SeparatorTypeEnum;

//-------------------------------------------------------------------------------
// _3_05_ Состояние чекбокса  
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("8406BD38-D629-4fe3-975C-EE266499C915") 
enum _CheckStateEnum
    {
        ksCheckUndefined	= 0,
        ksCheckCurrent	= 1,
        ksCheckFixed	= 2,
        ksCheckVariable	= 3
    } 	CheckStateEnum;

//-------------------------------------------------------------------------------
// _3_06_ Видимость имени контрола на панели свойств
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("BC3EDB2C-EB01-4365-A7EB-F833CDA41DFC") 
enum _PropertyControlNameVisibility
    {
        ksNameAlwaysVisible	= 0,
        ksNameHorizontalVisible	= 1,
        ksNameVerticalVisible	= 2,
        ksNameNoVisible	= 3
    } 	PropertyControlNameVisibility;

//-------------------------------------------------------------------------------
// _3_07_ Типы кнопок для элемента панели свойств - Набор кнопок
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("B8F97560-B866-44e7-A741-8F8B93EFDF2C") 
enum _ButtonTypeEnum
    {
        ksPushButton	= 0,
        ksCheckButton	= 1,
        ksRadioButton	= 2
    } 	ButtonTypeEnum;

  //-------------------------------------------------------------------------------
  // _1_13_ События пользовательского элемента управления.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("A583E322-F0AC-42db-AC9A-69823D7277A5") 
enum _ksPropertyUserControlNotifyEnum
    {
        puCreateOCX	= 1,
        puDestroyOCX	= 2
    } 	ksPropertyUserControlNotifyEnum;

//-------------------------------------------------------------------------------
// _3_08_ Тип отображения слайда в окне.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("2BD06688-6524-4c8c-ABE5-827C6A119CF2") 
enum _SlideTypeEnum
    {
        ksSlide	= -1,
        ksBitmap	= 1,
        ksGroup	= 2,
        ksKompasDocument	= 3,
        ksKompasText	= 4
    } 	SlideTypeEnum;

//-------------------------------------------------------------------------------
//  _3_09_ Тип фиксированности для умолчательных элементов управления панели свойств.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("3E0245A7-9FF9-4ee5-8C09-0F1FFDDE88F9") 
enum _DefaultFixTypeEnum
    {
        ksAllFixOff	= -1,
        ksAllFix	= 0,
        ksPointFix	= 1,
        ksAngleFix	= 2
    } 	DefaultFixTypeEnum;

//-------------------------------------------------------------------------------
//  _2_04_ Pежим отображения окна ( мaximize, minimize, нормальное ).
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("F34D2579-798F-4d4e-BAB6-8968DE7F7B37") 
enum FrameRegimeEnum
    {
        ksFrameMinimize	= 0,
        ksFrameMaximize	= 1,
        ksFrameRestore	= 2
    } 	FrameRegimeEnum;

//-------------------------------------------------------------------------------
//  _2_05_ Тип изменения масштаба отображения документа в окне.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("DFC4947C-917D-4266-AD2A-4948BD1DE889") 
enum _ZoomTypeEnum
    {
        ksZoomNext	= 0,
        ksZoomPrevious	= 1,
        ksZoomAll	= 2
    } 	ZoomTypeEnum;

  //-------------------------------------------------------------------------------
  // _1_14_ События для окна документа : клавиатура, мышь, события по отрисовке документа
  // ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("9637E1BB-CF0F-4dc7-8425-3F3957EB313A") 
enum _ksDocumentFrameNotifyEnum
    {
        frBeginPaint	= 1,
        frClosePaint	= 2,
        frMouseDown	= 3,
        frMouseUp	= 4,
        frMouseDblClick	= 5,
        frBeginPaintGL	= 6,
        frClosePaintGL	= 7,
        frAddGabarit	= 8,
        frActivate	= 9,
        frDeactivate	= 10,
        frCloseFrame	= 11,
        frMouseMove	= 12,
        frShowOcxTree	= 13,
        frBeginPaintTmpObjects	= 14,
        frClosePaintTmpObjects	= 15
    } 	ksDocumentFrameNotifyEnum;

  //-------------------------------------------------------------------------------
  // _2_06_ StructType2DEnum - Типы интерфейсов параметров создаваемых методом GetParamStruct объекта KompasObject
  // ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("9637E1BB-CF0F-4dc7-8425-3F3957EB313B") 
enum _StructType2DEnum
    {
        ko_Type1	= 1,
        ko_Type2	= 2,
        ko_Type3	= 3,
        ko_Type5	= 4,
        ko_Type6	= 5,
        ko_Phantom	= 6,
        ko_PlacementParam	= 7,
        ko_ViewParam	= 8,
        ko_LayerParam	= 9,
        ko_RequestInfo	= 10,
        ko_LineSegParam	= 11,
        ko_ArcByAngleParam	= 12,
        ko_ArcByPointParam	= 13,
        ko_MathPointParam	= 14,
        ko_RectParam	= 15,
        ko_PointParam	= 16,
        ko_BezierPointParam	= 17,
        ko_NurbsPointParam	= 18,
        ko_BezierParam	= 19,
        ko_CircleParam	= 20,
        ko_LineParam	= 21,
        ko_EllipseParam	= 22,
        ko_EllipsArcParam	= 23,
        ko_EllipsArcParam1	= 24,
        ko_EquidParam	= 25,
        ko_HatchParam	= 26,
        ko_ParagraphParam	= 27,
        ko_TextParam	= 28,
        ko_TextLineParam	= 29,
        ko_TextItemFont	= 30,
        ko_TextItemParam	= 31,
        ko_StandartSheet	= 32,
        ko_SheetSize	= 33,
        ko_SheetPar	= 34,
        ko_DocumentParam	= 35,
        ko_ColumnInfoParam	= 36,
        ko_AttributeType	= 37,
        ko_Attribute	= 38,
        ko_LibraryAttrTypeParam	= 39,
        ko_TAN	= 40,
        ko_CON	= 41,
        ko_DimText	= 42,
        ko_LDimSource	= 43,
        ko_DimDrawing	= 44,
        ko_LDimParam	= 45,
        ko_LBreakDimSource	= 46,
        ko_BreakDimDrawing	= 47,
        ko_LBreakDimParam	= 48,
        ko_ADimSource	= 49,
        ko_ADimParam	= 50,
        ko_ABreakDimParam	= 51,
        ko_RDimSource	= 52,
        ko_RDimDrawing	= 53,
        ko_RDimParam	= 54,
        ko_RBreakDrawing	= 55,
        ko_RBreakDimParam	= 56,
        ko_RoughPar	= 57,
        ko_ShelfPar	= 58,
        ko_RoughParam	= 59,
        ko_LeaderParam	= 60,
        ko_PosLeaderParam	= 61,
        ko_BrandLeaderParam	= 62,
        ko_MarkerLeaderParam	= 63,
        ko_BaseParam	= 64,
        ko_CutLineParam	= 65,
        ko_ViewPointerParam	= 66,
        ko_ToleranceBranch	= 67,
        ko_ToleranceParam	= 68,
        ko_CurvePattern	= 69,
        ko_CurvePicture	= 70,
        ko_CurvePatternEx	= 71,
        ko_CurveStyleParam	= 72,
        ko_DimensionPartsParam	= 73,
        ko_TextStyleParam	= 74,
        ko_ConicArcParam	= 75,
        ko_PolylineParam	= 76,
        ko_LibStyle	= 77,
        ko_TechnicalDemandParam	= 78,
        ko_SpecRoughParam	= 79,
        ko_DimensionOptions	= 80,
        ko_SpcColumnParam	= 81,
        ko_LibraryStyleParam	= 82,
        ko_InertiaParam	= 83,
        ko_MassInertiaParam	= 84,
        ko_VariableParam	= 85,
        ko_SnapOptions	= 86,
        ko_NurbsParam	= 87,
        ko_InsertFragmentParam	= 88,
        ko_ConstraintParam	= 89,
        ko_CornerParam	= 90,
        ko_RectangleParam	= 91,
        ko_RegularPolygonParam	= 92,
        ko_CentreParam	= 93,
        ko_DocAttachSpcParam	= 94,
        ko_SpcObjParam	= 95,
        ko_RasterParam	= 96,
        ko_RecordTypeAttrParam	= 97,
        ko_NumberTypeAttrParam	= 98,
        ko_SpcStyleColumnParam	= 99,
        ko_SpcStyleSectionParam	= 100,
        ko_SpcSubSectionParam	= 101,
        ko_SpcTuningSectionParam	= 102,
        ko_SpcTuningStyleParam	= 103,
        ko_SpcStyleParam	= 104,
        ko_SpcDescrParam	= 105,
        ko_QualityItemParam	= 106,
        ko_QualityContensParam	= 107,
        ko_LtVariant	= 108,
        ko_ContourParam	= 109,
        ko_DoubleValue	= 110,
        ko_Char255	= 111,
        ko_UserParam	= 112,
        ko_HatchLineParam	= 113,
        ko_HatchStyleParam	= 114,
        ko_OrdinatedSourceParam	= 115,
        ko_OrdinatedDrawingParam	= 116,
        ko_OrdinatedDimParam	= 117,
        ko_SheetOptions	= 118,
        ko_InsertFragmentParamEx	= 119,
        ko_TreeNodeParam	= 120,
        ko_ViewColorParam	= 121,
        ko_AssociationViewParam	= 122,
        ko_AxisLineParam	= 123,
        ko_TextDocumentParam	= 124,
        ko_RemoteElementParam	= 125,
        ko_CopyObjectParam	= 126,
        ko_OverlapObjectOptions	= 127,
        ko_ChangeLeaderParam	= 128,
        ko_ParametrisationParam	= 9000
    } 	StructType2DEnum;

//-------------------------------------------------------------------------------
// _3_10_ Положение панели свойств (вверху, внизу, слева, справа, плавает)
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("2B49AD89-5992-460c-ADCD-3EF087FABF18") 
enum _PropertyManagerLayout
    {
        pmFloating	= 0,
        pmAlignTop	= 1,
        pmAlignBottom	= 2,
        pmAlignRight	= 3,
        pmAlignLeft	= 4
    } 	PropertyManagerLayout;

//-------------------------------------------------------------------------------
// _2_07_ Типы процессов в КОМПАС 3D
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("BFA9DCE2-973A-4f80-89FC-F2966509665D") 
enum _ProcessTypeEnum
    {
        prUnknown	= 0,
        prPoint	= 10000,
        prPointAlong	= 10001,
        prIntersectPoint	= 10002,
        prAllIntersectPoint	= 10003,
        prPointOnDistance	= 10004,
        prLineSeg	= 10005,
        prParallelLineSeg	= 10006,
        prPerpendLineSeg	= 10007,
        prTanLineSegByOutsidePnt	= 10008,
        prTanLineSegByPntOn	= 10009,
        prTangent2LineSeg	= 10010,
        prContourLineSeg	= 10011,
        prContourParallelLineSeg	= 10012,
        prContourPerpendLineSeg	= 10013,
        prContourTanLineSegByOutsidePnt	= 10014,
        prLine	= 10015,
        prVerticalLine	= 10016,
        prHorizontalLine	= 10017,
        prPerpendLine	= 10018,
        prParallelLine	= 10019,
        prTangent2Line	= 10020,
        prTanLineByPntOn	= 10021,
        prTanLineByOutsidePnt	= 10022,
        prBisectorLine	= 10023,
        prCircle	= 10024,
        prCircle3Points	= 10025,
        prCircleCentreOnEl	= 10026,
        prCircleTangent	= 10027,
        prCircleTangent2	= 10028,
        prCircleTangent3	= 10029,
        prCircle2Points	= 10030,
        prCircleArc	= 10031,
        prArc3Points	= 10032,
        prArc2PointsAngle	= 10033,
        prArc2Points	= 10034,
        prArcTangent	= 10035,
        prContourArc	= 10036,
        prContourConArc	= 10037,
        prEllipse	= 10038,
        prEllipseGabDiagonal	= 10039,
        prEllipseTangent2	= 10040,
        prEllipseCentre3Points	= 10041,
        prEllipseParallel3Points	= 10042,
        prEllipseParallelCentre2Points	= 10043,
        prEllipseGabCentrePoint	= 10044,
        prBezier	= 10045,
        prContourBezier	= 10046,
        prPolyline	= 10047,
        prNurbs	= 10048,
        prContourNurbs	= 10049,
        prRectangle	= 10050,
        prRectangleCentrePoint	= 10051,
        prPolygon	= 10052,
        prEquidToObj	= 10053,
        prAssemblyEquid	= 10054,
        prLineDimension	= 10055,
        prCommonBaseLineDim	= 10056,
        prChainLineDim	= 10057,
        prCommonLineLineDim	= 10058,
        pr2ObjectsLineDim	= 10059,
        prCutLineDimension	= 10060,
        prAngleDimension	= 10061,
        prCommonBaseAngleDim	= 10062,
        prChainAngleDim	= 10063,
        prCommonLineAngleDim	= 10064,
        prCutAngleDimension	= 10065,
        prRadialDimension	= 10066,
        prRadialDimensionWithBreak	= 10067,
        prDiametralDimension	= 10068,
        prArcDimension	= 10069,
        prOrdinateDimension	= 10070,
        prLeader	= 10071,
        prBrandLeader	= 10072,
        prMarkLeader	= 10073,
        prPositionLeader	= 10074,
        prChangeLeader	= 10075,
        prHatch	= 10076,
        prText	= 10077,
        prTable	= 10078,
        prRough	= 10079,
        prBase	= 10080,
        prCutLine	= 10081,
        prViewPointer	= 10082,
        prRemoteElement	= 10083,
        prAxedLineSegment	= 10084,
        prCentreMarker	= 10085,
        prAssemblyContour	= 10086,
        prFormTolerance	= 10087,
        prInsertRaster	= 10088,
        prMakeMacro	= 10089,
        prInsertFragment	= 10090,
        prCreateSheetView	= 10091,
        prInsertOLEObject	= 10092,
        prCreateStandartSheetView	= 10093,
        prCreateSectionSheetView	= 10094,
        prCreateArbitrarySheetView	= 10095,
        prCreateProjectionSheetView	= 10096,
        prCreateArrowSheetView	= 10097,
        prCreateRemoteSheetView	= 10098,
        prCreateLocalSheetView	= 10099,
        prCreateLocalSectionSheetView	= 10100,
        prCreateBrokenSheetView	= 10101,
        prContour	= 10102,
        prChamfer	= 10103,
        prChamferPolyContour	= 10104,
        prFillet	= 10105,
        prFilletPolyContour	= 10106,
        prProjectionObject	= 10107,
        prSmartDimension	= 10108,
        prSmartAxedLineSegment	= 10109,
        prMeasurePointProperties	= 10110,
        prMeasureDistance2Points	= 10111,
        prMeasureDistance2PointsByCurve	= 10112,
        prMeasureDistancePointObject	= 10113,
        prMeasureDistance2Curves	= 10114,
        prMeasureAngle2Lines	= 10115,
        prMeasureAngle3Points	= 10116,
        prPerimeter	= 10117,
        prMeasureArea	= 10118,
        prMix	= 10119,
        prMix3DRevolution	= 10120,
        prMix3DExtrision	= 10121,
        prMix3DExtrusion	= 10121,
        prObjectShift	= 10122,
        prObjectShiftAngleLen	= 10123,
        prObjectRotate	= 10124,
        prObjectScale	= 10125,
        prObjectSymmetry	= 10126,
        prObjectMultiply	= 10127,
        prObjectMultiplyByCurve	= 10128,
        prObjectMultiplyByCircle	= 10129,
        prObjectMultiplyByRing	= 10130,
        prObjectMultiplyByMesh	= 10131,
        prMoveDeformation	= 10132,
        prRotateDeformation	= 10133,
        prScaleDeformation	= 10134,
        prCutObjectPart	= 10135,
        prCutObjectPartBy2Points	= 10136,
        prJustify	= 10137,
        prRemoveChamfer	= 10138,
        prBreakCurve	= 10139,
        prBreakCurveNParts	= 10140,
        prBlackBox	= 10141,
        prConvertToNurbs	= 10142,
        prParametricHorizontal	= 10143,
        prParametricVertical	= 10144,
        prParametricXAlign	= 10145,
        prParametricYAlign	= 10146,
        prParametricMergePoints	= 10147,
        prParametricPointOnCurve	= 10148,
        prParametricPointSymmetry	= 10149,
        prParametricParallel	= 10150,
        prParametricNormal	= 10151,
        prParametricColinear	= 10152,
        prParametricTangent	= 10153,
        prParametricFixPoint	= 10154,
        prParametricEqualRadiuses	= 10155,
        prParametricEqualLength	= 10156,
        prParametricFixDimension	= 10157,
        prParametricChangeDimension	= 10158,
        prParametricSelected	= 10159,
        prParametricDeleteObjConstraints	= 10160,
        prParametricDeleteAllConstraints	= 10161,
        prSelectObject	= 10162,
        prSelectLayer	= 10163,
        prSelectSheetView	= 10164,
        prSelectWithRect	= 10165,
        prSelectOutSideRect	= 10166,
        prSelectWithCutRect	= 10167,
        prSelectWithCutPolyline	= 10168,
        prExcludeObject	= 10169,
        prExcludeLayer	= 10170,
        prExcludeSheetView	= 10171,
        prExcludeWithRect	= 10172,
        prExcludeOutSideRect	= 10173,
        prExcludeWithCutRect	= 10174,
        prExcludeWithCutPolyline	= 10175,
        prSmartLine	= 10176,
        prBrace	= 10177,
        prAutoDimL	= 10178,
        prAutoDimA	= 10179,
        prAutoDimD	= 10180,
        prAutoDimR	= 10181,
        prAutoDimLToPoint	= 10182,
        prAutoDimLBreak	= 10183,
        prAutoDimABreak	= 10184,
        prTechnicalDemandPlacement	= 10185,
        prDirectAxis	= 10186,
        prArcAxis	= 10187,
        prCircleAxis	= 10188,
        prWaveLine	= 10189,
        prMarkOnLDRPosNum	= 10190,
        prMarkWoLDRPosNum	= 10191,
        prKnotNumber	= 10192,
        prUnitMarking	= 10193,
        prCutUnitMarking	= 10194,
        prMultiTextLeader	= 10195,
        prColouring	= 10196,
        prMultiLine	= 10197,
        prBuildingCutLine	= 10198,
        prBrokenLine	= 10199,
        prCreateReport	= 10200,
        prCreateAttachedLeaders	= 10201,
        prEditProperties	= 10202,
        prParametricBisector	= 10203,
        prParametricFixedLenght	= 10204,
        prParametricFixedAngle	= 10205,
        prParametricPointOnCurveMiddle	= 10206,
        prTechnicalDemand	= 10207,
        prSpecRough	= 10208,
        prMoveSpecRough	= 10209,
        prDeleteHistory	= 10210,
        prUndo	= 10211,
        prRedo	= 10212,
        prEmbodimentsReport	= 10213,
        prArrayParamReport	= 10214,
        prConicCurve	= 10215,
        prConicCurve4Or5Point	= 10216,
        prBaseExtrusion	= 20000,
        prBossExtrusion	= 20001,
        prCutExtrusion	= 20002,
        prExtrusionSurface	= 20003,
        prBaseRotated	= 20004,
        prBossRotated	= 20005,
        prCutRotated	= 20006,
        prRotatedSurface	= 20007,
        prBaseEvolution	= 20008,
        prBossEvolution	= 20009,
        prCutEvolution	= 20010,
        prEvolutionSurface	= 20011,
        prBaseLoft	= 20012,
        prBossLoft	= 20013,
        prCutLoft	= 20014,
        prLoftSurface	= 20015,
        prFillet3D	= 20016,
        prChamfer3D	= 20017,
        prCutByPlane	= 20018,
        prCutBySketch	= 20019,
        prMeshCopy	= 20020,
        prCircularCopy	= 20021,
        prCurveCopy	= 20022,
        prMirrorCopy	= 20023,
        prMirrorAllCopy	= 20024,
        prDerivativePartArray	= 20025,
        prMeshPartArray	= 20026,
        prCircularPartArray	= 20027,
        prCurvePartArray	= 20028,
        prIncline	= 20029,
        prShell	= 20030,
        prRib	= 20031,
        prHole	= 20032,
        prThread	= 20033,
        prCPlaneOffset	= 20034,
        prCPlane3Points	= 20035,
        prCPlaneAngle	= 20036,
        prCPlaneEdgePoint	= 20037,
        prCPlaneParallel	= 20038,
        prCPlanePerpendicular	= 20039,
        prCPlaneNormalToSurface	= 20040,
        prCPlaneTangentToSurface	= 20041,
        prCPlaneLineToEdge	= 20042,
        prCAxis2Points	= 20043,
        prCAxis2Planes	= 20044,
        prCAxisConeface	= 20045,
        prCAxisEdge	= 20046,
        prCAxisOperation	= 20047,
        prPolyline3D	= 20048,
        prSpline3D	= 20049,
        prCylindricSpiral	= 20050,
        prConicSpiral	= 20051,
        prImportedSurface	= 20052,
        prInsertScetch	= 20053,
        prEditScetch	= 20054,
        prOrientationScetch	= 20055,
        prInPlacePartEdit	= 20056,
        prOutPlacePartEdit	= 20057,
        prAddDetail	= 20058,
        prAddAssembly	= 20059,
        prMateCoincident	= 20060,
        prMateConcentric	= 20061,
        prMateParallel	= 20062,
        prMatePerpendicular	= 20063,
        prMateOnDistance	= 20064,
        prMateOnAngle	= 20065,
        prMateTangent	= 20066,
        prPartVariables	= 20067,
        prCopyBilletPart	= 20068,
        prMakeMoldCavity	= 20069,
        prMakeUnionComps	= 20070,
        prAddPartFromFile	= 20071,
        prMovePart	= 20072,
        prRotatePartWC	= 20073,
        prRotatePartAxis	= 20074,
        prRotatePartPoint	= 20075,
        prMakeSplitLine	= 20076,
        prMeasureDistance3D	= 20077,
        prMeasurePerimeter3D	= 20078,
        prMeasureArea3D	= 20079,
        prMeasureMix3D	= 20080,
        prMeasureInterferenceVolumes	= 20081,
        prBaseShMtSolid	= 20082,
        prShMtBend	= 20083,
        prShMtCombinedBend	= 20084,
        prShMtBendLine	= 20085,
        prShMtBendHook	= 20086,
        prShMtHole	= 20087,
        prShMtCut	= 20088,
        prBaseShMtPlate	= 20089,
        prShMtClosedCorner	= 20090,
        prShMtBendStraighten	= 20091,
        prShMtBendBended	= 20092,
        prShMtBendParamUnfold	= 20093,
        prPatchSurface	= 20094,
        prSewSurface	= 20095,
        prMakeFaceRemover	= 20096,
        prCPlaneMiddle	= 20097,
        prCPointControl	= 20098,
        prCPointConjunctive	= 20099,
        prCAggregateOper	= 20100,
        prCPlaneLineToFlat	= 20101,
        prPoint3D	= 20103,
        prLocalCoordinateSystem	= 20104,
        prLineDimention3DPlane	= 20105,
        prLineDimention3D	= 20106,
        prAngleDimention3D	= 20107,
        prRough3D	= 20108,
        prTolerance3D	= 20109,
        prBrandLeader3D	= 20110,
        prMarkerLeader3D	= 20111,
        prPositionLeader3D	= 20112,
        prBase3D	= 20113,
        prLeader3D	= 20114,
        prSaveBody	= 20115,
        prCreateSketch	= 20116,
        prMeasureInformation	= 20117,
        prEquidistant3D	= 20118,
        prChoiceOperationResult	= 20119,
        prChoiceBodyUnit	= 20120,
        prSelectCurrentCS	= 20121,
        prShmtRuledOperation	= 20122,
        prArc3D	= 20123,
        prConnectCurve	= 20124,
        prTrimCurve	= 20125,
        prFilletCurve	= 20126,
        prSwithOwnCS	= 20127,
        prScalingOperation	= 20128,
        prPointDrivenPattern	= 20129,
        prNurbs3DByObjects	= 20130,
        prCurveOperationCrossing	= 20131,
        prConvertToNurbsSurface3D	= 20133,
        prNurbsSurface3DByPoints	= 20134,
        prNurbsSurface3DByCurves	= 20135,
        prArrayPointsFromFile	= 20136,
        prArrayPointsOnCurve	= 20137,
        prArrayPointsByCloud	= 20138,
        prOffsetSurface	= 20139,
        prAuxObjectMultiplyByMesh	= 20140,
        prAuxObjectMultiplyByRing	= 20141,
        prAuxObjectMultiplyByCurve	= 20142,
        prTrimmedSurface	= 20143,
        prSurfaceToBody	= 20144,
        prAxisByDirection	= 20145,
        prRuledSurface	= 20146,
        prExtensionSurface	= 20147,
        prCPlaneTangentAtPoint	= 20148,
        prCPlaneAtCurve	= 20149,
        prArrayPintsOnSyrface	= 20150,
        prAuxObjectMultiplyMirror	= 20151,
        prOutlineCurve	= 20152,
        prSplineOnSurface	= 20153,
        prPartsPointDrivenPattern	= 20154,
        prChooseLinearPattern	= 20155,
        prChooseCircularPattern	= 20156,
        prChooseCurvePattern	= 20157,
        prChoosePointDrivenPattern	= 20158,
        prChooseTablePattern	= 20159,
        prChooseMirrorPattern	= 20160,
        prTablePattern	= 20161,
        prAuxTablePattern	= 20162,
        prPartsTablePattern	= 20163,
        prAuxPointDrivenPattern	= 20164,
        prBodiesLinearPattern	= 20165,
        prBodiesCircularPattern	= 20166,
        prBodiesCurvePattern	= 20167,
        prBodiesPointDrivenPattern	= 20168,
        prBodiesTablePattern	= 20169,
        prContour3D	= 20170,
        prCurveOper2Projection	= 20171,
        prCurveByLaw	= 20172,
        prBodyReposition	= 20173,
        prIsoparamCurve	= 20174,
        prIsoparamCurveArr	= 20175,
        prBindingMesh	= 20176,
        prSaveBodyAs	= 20177,
        prBlendSurface	= 20178,
        prLineSegment3D	= 20179,
        prEmbodiment	= 20180,
        prCreateSpecificationObjects	= 20181,
        prDeleteSpecificationObjects	= 20182,
        prCreateSpecificationFromAssembly	= 20183,
        prSpecRough3D	= 20184,
        prShmtRuledCowling	= 20185,
        prRepositionPart	= 20186,
        prAddLocalPartFromFile	= 20187,
        prAddLayoutGeometryFromFile	= 20188,
        prAddBilletPartFromFile	= 20189,
        prMateSymmetry	= 20190,
        prMateDependent	= 20191,
        prMateCamGear	= 20192,
        prMateRotation	= 20193,
        prMateRotationTransfer	= 20194,
        prTechnicalDemand3D	= 20195,
        prAlignShelfs	= 36075,
        prZoomWindow	= 32411,
        prMoveView	= 32418,
        prPanoramaView	= 32419,
        prRotateView	= 32420,
        prEditSelectedObject	= 35736,
        prEditCopy	= 0xe122,
        prEditCut	= 0xe123,
        prEditPaste	= 0xe125
    } 	ProcessTypeEnum;

//-------------------------------------------------------------------------------
// _2_08_ Действия при закрытии документа КОМПАС
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("37009011-39C7-434c-8099-8B48B6E36353") 
enum _DocumentCloseOptions
    {
        kdDoNotSaveChanges	= 0,
        kdSaveChanges	= 1,
        kdPromptToSaveChanges	= 2
    } 	DocumentCloseOptions;

  //-------------------------------------------------------------------------------
  // _2_09_ Типы преобразования логических координат в координаты документа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B3B4A2A8-4342-4f7e-B928-2134F3EB5D66") 
enum _ConvertCoordTypeEnum
    {
        kcDocument	= 1,
        kcGeoView	= 2,
        kcCurrentPlane	= 3
    } 	ConvertCoordTypeEnum;

//-------------------------------------------------------------------------------
// _4_01_  Состояние параметра в условии фильтрации слоев.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("468BC140-6F10-4413-B73A-F7D341251CC9") 
enum _FilterCondotionStateEnum
    {
        ksStateUndefined	= -1,
        ksStateFALSE	= 0,
        ksStateTRUE	= 1
    } 	FilterCondotionStateEnum;

//-------------------------------------------------------------------------------
// _4_02_  Типы видов.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("33DA907C-B060-4aa0-8EEB-FBDD0E7FD326") 
enum _LtViewType
    {
        vtUnknown	= -1,
        vt_System	= 0,
        vt_Normal	= 1,
        vt_Arbitrary	= 2,
        vt_Standart	= 3,
        vt_Projected	= 4,
        vt_Arrow	= 5,
        vt_Remote	= 6,
        vt_Section	= 7
    } 	LtViewType;

//-------------------------------------------------------------------------------
//  _4_03_ Способ группировки слоев.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("2A68C0E0-ACF1-47d2-A3BC-964F58EBEBC0") 
enum _LayersGroupWayEnum
    {
        wgLayers	= 0,
        wgLayersCharacteristics	= 1
    } 	LayersGroupWayEnum;

//-------------------------------------------------------------------------------
//_2_10_  Типы графических объектов.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("8E233AAB-F840-4c9f-AA02-7D2B1C6D6813") 
enum _DrawingObjectTypeEnum
    {
        ksUnknown	= -1,
        ksAllObj	= 0,
        ksDrLineSeg	= 1,
        ksDrCircle	= 2,
        ksDrArc	= 3,
        ksDrDrawText	= 4,
        ksDrPoint	= 5,
        ksDrHatch	= 7,
        ksDrBezier	= 8,
        ksDrLDimension	= 9,
        ksDrADimension	= 10,
        ksDrDDimension	= 13,
        ksDrRDimension	= 14,
        ksDrRBreakDimension	= 15,
        ksDrRough	= 16,
        ksDrBase	= 17,
        ksDrWPointer	= 18,
        ksDrCut	= 19,
        ksDrLeader	= 20,
        ksDrPosLeader	= 21,
        ksDrBrandLeader	= 22,
        ksDrMarkerLeader	= 23,
        ksDrTolerance	= 24,
        ksDrTable	= 25,
        ksDrContour	= 26,
        ksDrMacro	= 27,
        ksDrLine	= 28,
        ksLayer	= 29,
        ksDrFragment	= 30,
        ksDrPolyline	= 31,
        ksDrEllipse	= 32,
        ksDrNurbs	= 33,
        ksDrEllipseArc	= 34,
        ksDrRectangle	= 35,
        ksDrRegularPolygon	= 36,
        ksDrEquid	= 37,
        ksDrLBreakDimension	= 38,
        ksDrABreakDimension	= 39,
        ksDrOrdinateDimension	= 40,
        ksDrColorFill	= 41,
        ksDrCentreMarker	= 42,
        ksDrArcDimension	= 43,
        ksDrRaster	= 45,
        ksDrChangeLeader	= 46,
        ksDrRemoteElement	= 47,
        ksDrAxisLine	= 48,
        ksDrOLEObject	= 49,
        ksDrUnitNumber	= 50,
        ksDrBrace	= 51,
        ksDrMarkOnLeader	= 52,
        ksDrMarkOnLine	= 53,
        ksDrMarkInsideForm	= 54,
        ksDrWaveLine	= 55,
        ksDrStraightAxis	= 56,
        ksDrBrokenLine	= 57,
        ksDrCircleAxis	= 58,
        ksDrArcAxis	= 59,
        ksDrCutUnitMarking	= 60,
        ksDrUnitMarking	= 61,
        ksDrMultiTextLeader	= 62,
        ksDrExternalView	= 63,
        ksDrAnnLineSeg	= 64,
        ksDrAnnCircle	= 65,
        ksDrAnnEllipse	= 66,
        ksDrAnnArc	= 67,
        ksDrAnnEllipseArc	= 68,
        ksDrAnnPolyline	= 69,
        ksDrAnnPoint	= 70,
        ksDrAnnText	= 71,
        ksDrMultiLine	= 72,
        ksDrBuildingCutLine	= 73,
        ksDrAttachedLeader	= 74,
        ksDrConditionCrossing	= 75,
        ksReportTable	= 76,
        ksEmbodimentsTable	= 77,
        ksArrayParamTable	= 79,
        ksDrNurbsByPoints	= 80,
        ksDrConicCurve	= 81,
        ksDrConicCurve4Point	= 82,
        ksDrConicCurve5Point	= 82,
        ksView	= 123
    } 	DrawingObjectTypeEnum;

  //-------------------------------------------------------------------------------
  // _1_15_ События для менеджера видов и слоев
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("254F906A-E026-4f5e-AF11-86664919D9F5") 
enum _ksViewsAndLayersManagerNotifyEnum
    {
        vmBeginEdit	= 1,
        vmEndEdit	= 2
    } 	ksViewsAndLayersManagerNotifyEnum;

  //------------------------------------------------------------------------------
  // _5_01_ Типы библиотек.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("96D179F1-6D16-4358-8C9B-06664635C624") 
enum _LibraryTypeEnum
    {
        ksLibraryUnknown	= 0,
        ksLibraryProcedure	= 1,
        ksLibraryFragment	= 2,
        ksLibraryModel	= 3,
        ksLibraryDocuments	= 4
    } 	ksLibraryTypeEnum;

  //------------------------------------------------------------------------------
  // _5_02_ Стили отображения прикладных библиотек.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("75FE9898-4CB3-44c6-B3FB-D6CF6D881DDD") 
enum _LibraryStyleEnum
    {
        ksLibraryStyleUnknown	= 0,
        ksLibraryStyleMenu	= 1,
        ksLibraryStyleDialog	= 2,
        ksLibraryStyleWindow	= 3,
        ksLibraryStyleBar	= 4,
        ksLibraryStyleInvisible	= 5
    } 	ksLibraryStyleEnum;

  //-------------------------------------------------------------------------------
  // _1_16_ События для менеджера библиотек
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("273C6046-7F28-4ef6-95B8-D629C34AC501") 
enum _ksLibraryManagerNotifyEnum
    {
        ksLMBeginAttach	= 1,
        ksLMAttach	= 2,
        ksLMBeginDetach	= 3,
        ksLMDetach	= 4,
        ksLMBeginExecute	= 5,
        ksLMEndExecute	= 6,
        ksLMSystemControlStop	= 7,
        ksLMSystemControlStart	= 8,
        ksLMAddLibraryDescription	= 9,
        ksLMDeleteLibraryDescription	= 10,
        ksLMAddInsert	= 11,
        ksLMDeleteInsert	= 12,
        ksLMEditInsert	= 13,
        ksLMTryExecute	= 14
    } 	ksLibraryManagerNotifyEnum;

  //-------------------------------------------------------------------------------
  // _5_03_ Результаты передачи управления системе КОМПАС
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("DEE6790C-115B-40c6-A70E-82F7E186EB95") 
enum _ksSystemControlStartEnum
    {
        ksSCStoppedByMenuCommand	= 1,
        ksSCCloseApplication	= 0,
        ksSCStopItself	= -1,
        ksSCAlreadyStarted	= -2,
        ksSCStartedByAnotherLibrary	= -3,
        ksSCError	= -4
    } 	ksSystemControlStartEnum;

  //------------------------------------------------------------------------------
  // _6_01_ SpecificationVariantEnum - Варианты оформления спецификации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B6A0558F-5DD8-4dff-8505-5ED41A199EBF") 
enum _SpecificationVariantEnum
    {
        ksSpecificationSimple	= 0,
        ksSpecificationVariantA	= 1,
        ksSpecificationVariantB	= 2,
        ksSpecificationVariantV	= 3,
        ksSpecificationVariantG	= 4
    } 	ksSpecificationVariantEnum;

  //------------------------------------------------------------------------------
  // _6_02_ SpecificationLinkTypeEnum - Режимы связи сборки или чертежа со спецификацией.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4BEFECBE-42CF-4839-9642-70309FD4DCAD") 
enum _SpecificationLinkTypeEnum
    {
        ksLinkNone	= 0,
        ksLinkOnlyObjects	= 1,
        ksLinkWithPositionCalculate	= 2
    } 	ksSpecificationLinkTypeEnum;

  //------------------------------------------------------------------------------
  // _6_03_ SpecificationColumnTypeEnum - Типы колонок спецификации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("5A85CA29-8EA1-4a50-8F80-C146A69F704F") 
enum _SpecificationColumnTypeEnum
    {
        ksSColumnUnknown	= 0,
        ksSColumnFormat	= 1,
        ksSColumnZone	= 2,
        ksSColumnPosition	= 3,
        ksSColumnMark	= 4,
        ksSColumnName	= 5,
        ksSColumnCount	= 6,
        ksSColumnNote	= 7,
        ksSColumnMass	= 8,
        ksSColumnMaterial	= 9,
        ksSColumnUser	= 10,
        ksSColumnCode	= 11,
        ksSColumnFactory	= 12,
        ksSColumnDocumentNumber	= 13,
        ksSColumnDocumentName	= 14,
        ksSColumnDocumentCode	= 15,
        ksSColumnCodeOKP	= 16
    } 	ksSpecificationColumnTypeEnum;

  //------------------------------------------------------------------------------
  // _6_04_ ValueTypeEnum - Типы значения атрибута, его колонок и колонок спецификации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("8923B50C-2FC3-4ead-BBF5-FED84D327909") 
enum _ValueTypeEnum
    {
        ksValueTypeUnknown	= 0,
        ksValueTypeInteger	= 1,
        ksValueTypeFloat	= 2,
        ksValueTypeString	= 3,
        ksValueTypeRecord	= 4
    } 	ksValueTypeEnum;

  //------------------------------------------------------------------------------
  // _6_05_ SortTypeEnum - Типы сортировки
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("62EC772D-D447-4c98-8CA9-5C813177894A") 
enum _SortTypeEnum
    {
        ksSortTypeNone	= 0,
        ksSortTypeCompositeUp	= 1,
        ksSortTypeUp	= 3,
        ksSortTypeDocument	= 4,
        ksSortTypeDown	= 5,
        ksSortTypeCompositeDown	= 6
    } 	ksSortTypeEnum;

  //------------------------------------------------------------------------------
  // _2_11_ DocumentFormatEnum - Форматы листа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("841AB2BC-4098-4761-93A8-5CC9BEE9B2B3") 
enum _DocumentFormatEnum
    {
        ksFormatA0	= 0,
        ksFormatA1	= 1,
        ksFormatA2	= 2,
        ksFormatA3	= 3,
        ksFormatA4	= 4,
        ksFormatA5	= 5,
        ksFormatUser	= 6
    } 	ksDocumentFormatEnum;

  //------------------------------------------------------------------------------
  // _2_12_ AlignEnum - Выравнивание
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E81E5CCE-B748-4f43-858E-4621569D2C9A") 
enum _AlignEnum
    {
        ksAlignDefault	= -1,
        ksAlignLeft	= 0,
        ksAlignCenter	= 1,
        ksAlignRight	= 2,
        ksAlignAllWidth	= 3,
        ksAlignDecimal	= 3
    } 	ksAlignEnum;

  //------------------------------------------------------------------------------
  // _2_13_ TabulatorFillingEnum - Заполнение табулятора
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("646816D3-95BE-4652-9C34-4B48248E5372") 
enum _TabulatorFillingEnum
    {
        ksTabulatorFillingNone	= 0,
        ksTabulatorFillingBaseLine	= 1,
        ksTabulatorFillingCenterLine	= 2,
        ksTabulatorFillingBaseDot	= 3,
        ksTabulatorFillingCenterDot	= 4,
        ksTabulatorFillingBaseDash	= 5,
        ksTabulatorFillingCenterDash	= 6
    } 	ksTabulatorFillingEnum;

  //------------------------------------------------------------------------------
  // _2_14_ ErrorType - Ошибки API, кроме 3D
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6008B376-7E42-464c-8DC6-75B99E7D071F") 
enum _ErrorType
    {
        etNoTXTDocument	= -8,
        etNo3dDocument	= -7,
        etNoAllDocument	= -6,
        etNoSPCDocument	= -5,
        etLibraryClose	= -4,
        etNoPreView	= -3,
        etNoDocument	= -2,
        etAbort	= -1,
        etSuccess	= 0,
        etError1	= 1,
        etError2	= 2,
        etError3	= 3,
        etError4	= 4,
        etError5	= 5,
        etError6	= 6,
        etError7	= 7,
        etError8	= 8,
        etError9	= 9,
        etError10	= 10,
        etError11	= 11,
        etError12	= 12,
        etError13	= 13,
        etError14	= 14,
        etError15	= 15,
        etError16	= 16,
        etError17	= 17,
        etError18	= 18,
        etError19	= 19,
        etError20	= 20,
        etError21	= 21,
        etError22	= 22,
        etError23	= 23,
        etError24	= 24,
        etError25	= 25,
        etError26	= 26,
        etError27	= 27,
        etError28	= 28,
        etError29	= 29,
        etError30	= 30,
        etError31	= 31,
        etError32	= 32,
        etError33	= 33,
        etError34	= 34,
        etError35	= 35,
        etError36	= 36,
        etError37	= 37,
        etError38	= 38,
        etError39	= 39,
        etError40	= 40,
        etError41	= 41,
        etError42	= 42,
        etError43	= 43,
        etError44	= 44,
        etError45	= 45,
        etError46	= 46,
        etError47	= 47,
        etError48	= 48,
        etError49	= 49,
        etError50	= 50,
        etError51	= 51,
        etError52	= 52,
        etError53	= 53,
        etError54	= 54,
        etError55	= 55,
        etError56	= 56,
        etError57	= 57,
        etError58	= 58,
        etError59	= 59,
        etError60	= 60,
        etError61	= 61,
        etError62	= 62,
        etError63	= 63,
        etError64	= 64,
        etError65	= 65,
        etError66	= 66,
        etError67	= 67,
        etError68	= 68,
        etError69	= 69,
        etError70	= 70,
        etError71	= 71,
        etError72	= 72,
        etError73	= 73,
        etError74	= 74,
        etError75	= 75,
        etError76	= 76,
        etError77	= 77,
        etError78	= 78,
        etError79	= 79,
        etError80	= 80,
        etError81	= 81,
        etError82	= 82,
        etError83	= 83,
        etError84	= 84,
        etError85	= 85,
        etError86	= 86,
        etError87	= 87,
        etError88	= 88,
        etError89	= 89,
        etError90	= 90,
        etError91	= 91,
        etError92	= 92,
        etError93	= 93,
        etError94	= 94,
        etError95	= 95,
        etError96	= 96,
        etError97	= 97,
        etError98	= 98,
        etError99	= 99,
        etError100	= 100,
        etError101	= 101,
        etError102	= 102,
        etError103	= 103,
        etError104	= 104,
        etError105	= 105,
        etError106	= 106,
        etError107	= 107,
        etError108	= 108,
        etError109	= 109,
        etError110	= 110,
        etError111	= 111,
        etError112	= 112,
        etError113	= 113,
        etError114	= 114,
        etError115	= 115,
        etError116	= 116,
        etError117	= 117,
        etError118	= 118,
        etError119	= 119,
        etError120	= 120,
        etError121	= 121,
        etError122	= 122,
        etError123	= 123,
        etError124	= 124,
        etError125	= 125,
        etError126	= 126,
        etError127	= 127,
        etError128	= 128,
        etError129	= 129,
        etError130	= 130,
        etError131	= 131,
        etError132	= 132,
        etError133	= 133,
        etError134	= 134,
        etError135	= 135,
        etError136	= 136,
        etError137	= 137,
        etError138	= 138,
        etError139	= 139,
        etError140	= 140,
        etError141	= 141,
        etError142	= 142,
        etError143	= 143,
        etError144	= 144,
        etError145	= 145,
        etError146	= 146,
        etError147	= 147,
        etError148	= 148,
        etError149	= 149,
        etError150	= 150,
        etError151	= 151,
        etError152	= 152,
        etError153	= 153,
        etError154	= 154,
        etError155	= 155,
        etError156	= 156,
        etError157	= 157,
        etError158	= 158,
        etError159	= 159,
        etError160	= 160,
        etError161	= 161,
        etError162	= 162,
        etError163	= 163,
        etError164	= 164,
        etError165	= 165,
        etError166	= 166,
        etError167	= 167,
        etError168	= 168,
        etError169	= 169,
        etError170	= 170,
        etError171	= 171,
        etError172	= 172,
        etError173	= 173,
        etError174	= 174,
        etError175	= 175,
        etError176	= 176,
        etError177	= 177,
        etError178	= 178,
        etError179	= 179,
        etError180	= 180,
        etError181	= 181,
        etError182	= 182,
        etError183	= 183,
        etError184	= 184,
        etError185	= 185,
        etError186	= 186,
        etError187	= 187,
        etError188	= 188,
        etError189	= 189,
        etError190	= 190,
        etError191	= 191,
        etError192	= 192,
        etError193	= 193,
        etError194	= 194,
        etError195	= 195,
        etError196	= 196,
        etError197	= 197,
        etError198	= 198,
        etError199	= 199,
        etError200	= 200,
        etError201	= 201,
        etError202	= 202,
        etError203	= 203,
        etError204	= 204,
        etError205	= 205,
        etError206	= 206,
        etError207	= 207,
        etError208	= 208,
        etError209	= 209,
        etError210	= 210,
        etError211	= 211,
        etError212	= 212,
        etError213	= 213,
        etError214	= 214,
        etError215	= 215,
        etError216	= 216
    } 	ErrorType;

  //------------------------------------------------------------------------------
  // _6_06_ SpecificationObjectTypeEnum - Типы объектов для спецификации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7D842385-B7AA-4296-87FA-A1E4199401F9") 
enum _SpecificationObjectTypeEnum
    {
        ksSpecificationUnknownObject	= 0,
        ksSpecificationBaseObject	= 1,
        ksSpecificationComment	= 2,
        ksSpecificationSectionName	= 3,
        ksSpecificationBlock	= 4,
        ksSpecificationReserveString	= 5,
        ksSpecificationEmptyString	= 6
    } 	ksSpecificationObjectTypeEnum;

  //------------------------------------------------------------------------------
  // _6_07_ SpecificationObjectStateEnum - Состояние объекта спецификации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("DF2C6BD8-A9EF-45d3-BF3B-9823587B9C79") 
enum _SpecificationObjectStateEnum
    {
        ksObjectStateIndependent	= 0,
        ksObjectStateFromInsert	= 1,
        ksObjectStateEdit	= 2,
        ksObjectStateUserSetNotEdit	= 3
    } 	ksSpecificationObjectStateEnum;

  //-------------------------------------------------------------------------------
  // _1_17_  События для листов оформления
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("604BA20F-A38F-4dcd-A961-F6320FDF8C94") 
enum _ksLayoutSheetsNotifyEnum
    {
        ksLayoutAdd	= 1,
        ksLayoutDelete	= 2,
        ksLayoutUpdate	= 3
    } 	ksLayoutSheetsNotifyEnum;

  //-------------------------------------------------------------------------------
  // _2_15_  Типы внутренних конвертеров КОМПАС 3D
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("2EAE1E94-4D8B-4df0-8548-672A31EDC36D") 
enum _ksKOMPASConverterEnum
    {
        ksConverterToRaster	= 0,
        ksConverterToSAT	= 1,
        ksConverterToXT	= 2,
        ksConverterToSTEP	= 3,
        ksConverterToIGES	= 4,
        ksConverterToVRML	= 5,
        ksConverterToSTL	= 6
    } 	ksKOMPASConverterEnum;

  //-------------------------------------------------------------------------------
  // _2_16_ Единицы измерения длины
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("DE3BFC6A-9040-4f10-A46E-BCBAE8204FD9") 
enum _ksLengthUnitEnum
    {
        ksLUMillimetres	= 0,
        ksLUCentimetres	= 1,
        ksLUDecimetres	= 2,
        ksLUMetres	= 3,
        ksLUKilometres	= 4
    } 	ksLengthUnitEnum;

  //-------------------------------------------------------------------------------
  // _2_17_ - Идентификаторы ячеек штампа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6861E06D-DF13-4e3b-ABD0-B41D2C024101") 
enum _ksStampEnum
    {
        ksStPartNumber	= 1,
        ksStDescription	= 2,
        ksStMaterial	= 3,
        ksStMass	= 5,
        ksStScale	= 6,
        ksStSheetNumber	= 7,
        ksStNumberOfSheets	= 8,
        ksStCompany	= 9,
        ksStTypeOfWork	= 10,
        ksStDocumentLetter1	= 40,
        ksStDocumentLetter2	= 41,
        ksStDocumentLetter3	= 42,
        ksStFullFileName	= 43,
        ksStShortFileName	= 44,
        ksStMarkingLine	= 45,
        ksStDocumentName	= 51,
        ksStDocumentCode	= 52,
        ksStOKPCode	= 53,
        ksStAuthor	= 110,
        ksStCheckedBy	= 111,
        ksStMfgApprovedBy	= 112,
        ksStDesigner	= 113,
        ksStRateOfInspection	= 114,
        ksStApprovedBy	= 115,
        ksStEndDesignDate	= 130,
        ksStCheckedDate	= 131,
        ksStMfgApprovedDate	= 132,
        ksStExecutionDate	= 133,
        ksStRateOfInspectionDate	= 134,
        ksStApprovedDate	= 135
    } 	ksStampEnum;

//-------------------------------------------------------------------------------
// _3_11_ Тип элемента панели свойств - Список
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("2503363E-1889-4767-9D51-A10F49795B13") 
enum _ksEditListTypeEnum
    {
        ksEditList	= 0,
        ksCheckList	= 1,
        ksRadioList	= 2
    } 	ksEditListTypeEnum;

//-------------------------------------------------------------------------------
// _3_12_ Идентификаторы стандартных команд для элемента панели свойств - Список
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("D38BF483-6C4B-47f5-8815-12AB81602C3D") 
enum _ksEditListCommandEnum
    {
        ksListItemNew	= 1,
        ksListItemDelete	= 2,
        ksListItemMoveUp	= 3,
        ksListItemMoveDown	= 4,
        ksListItemEdit	= 5
    } 	ksEditListCommandEnum;

  //-------------------------------------------------------------------------------
  // _2_18_ - Тип API интерфейса
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("EE887909-DE17-471b-AB3C-B8FB2AF892FA") 
enum _ksAPITypeEnum
    {
        ksAPIUndef	= 0,
        ksAPI5Auto	= 1,
        ksAPI7Dual	= 2,
        ksAPI3DCom	= 3
    } 	ksAPITypeEnum;

//-------------------------------------------------------------------------------
//_2_19_ ksTreeTypeEnum - Типы дерева построения 3D документа.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("71250E67-3E9B-4d17-907F-7AC88D761108") 
enum _ksTreeTypeEnum
    {
        ksOperTree	= 0,
        ksMultiTree	= 1
    } 	ksTreeTypeEnum;

//-------------------------------------------------------------------------------
//_2_20_ ksKompasCommandEnum - Команды Компаса.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("72B27F9D-FFBE-418f-A311-A892948CFC9B") 
enum _ksKompasCommandEnum
    {
        ksCMViewFullScreen	= 32403,
        ksCMSaveAll	= 32404,
        ksCMSaveTechnicalDemand	= 32405,
        ksCMSaveTechnicalDemandToTxt	= 32406,
        ksCMCloseTechnicalDemand	= 32407,
        ksCMCloseSpcSlave	= 32408,
        ksCMDocumentSetup	= 32410,
        ksCMZoomWindow	= 32411,
        ksCMZoomIn	= 32412,
        ksCMZoomOut	= 32413,
        ksCMScaleView	= 32414,
        ksCMZoomEntireDocument	= 32415,
        ksCMZoomSelected	= 32416,
        ksCMRefresh	= 32417,
        ksCMMoveView	= 32418,
        ksCMPanoramaView	= 32419,
        ksCMRotateView	= 32420,
        ksCMSetAttributeTypes	= 32421,
        ksCMSetCurveStyles	= 32422,
        ksCMSetTextStyles	= 32423,
        ksCMSetStampStyles	= 32424,
        ksCMSetTextShape	= 32425,
        ksCMSetGraphicShape	= 32426,
        ksCMSetHatchStyles	= 32427,
        ksCMSetSpcStyles	= 32448,
        ksCMSummaryInfo	= 32440,
        ksCMMoveDocumentEX	= 32522,
        ksCMMoveDocumentFirst	= 32523,
        ksCMMoveDocumentLast	= 32524,
        ksCMMoveDocumentPrev	= 32525,
        ksCMMoveDocumentNext	= 32526,
        ksCMRetryCommand	= 32534,
        ksCMTutor3D	= 32535,
        ksCMTutor2D	= 32540,
        ksCMSetProperties	= 32541,
        ksCMSetReportStyles	= 32542,
        ksCMZoom1	= 32545,
        ksCMCloseAll	= 33008,
        ksCMRepeatFind	= 33207,
        ksCMStop	= 33206,
        ksCMEscape	= 33211,
        ksCMCreateObject	= 33700,
        ksCMCansel	= 33701,
        ksCMSpcObjectsSort	= 33796,
        ksCMSpcRebuild	= 33797,
        ksCMFullPageHeight	= 33800,
        ksCMFullPageWidth	= 33801,
        ksCMSpcMakePosition	= 33891,
        ksCMSpcObjectDelete	= 33892,
        ksCMSpcObjectInsert	= 33893,
        ksCMSpcSynchronize	= 33896,
        ksCMSpcInsertLine	= 33898,
        ksCMSpcShowAll	= 33900,
        ksCMSpcTuningSetup	= 33901,
        ksCMSpcCopyObject	= 33902,
        ksCMSpcObjectMoveUp	= 33903,
        ksCMSpcObjectMoveDovn	= 33904,
        ksCMSpcObjectEdit	= 33905,
        ksCMSpcInsertIspoln	= 33906,
        ksCMSpcOpenGeometryDocs	= 33907,
        ksCMFragmentManager	= 35704,
        ksCMEditFragment	= 35705,
        ksCMCreateLocalFragment	= 35706,
        ksCMEditFragmentSource	= 35707,
        ksCMSelectGroup	= 35711,
        ksCMSelectLayer	= 35712,
        ksCMSelectByType	= 35713,
        ksCMStoreGroup	= 35715,
        ksCMSelectView	= 35716,
        ksCMSelectByAttr	= 35717,
        ksCMSelectByCurveStyle	= 35718,
        ksCMSelectPrevList	= 35719,
        ksCMUnSelectSroup	= 35720,
        ksCMUnSelectLayer	= 35721,
        ksCMUnSelectByType	= 35722,
        ksCMUnSelectView	= 35723,
        ksCMUnSelectByAttr	= 35724,
        ksCMUnSelectByCurveStyle	= 35725,
        ksCMUnSelectPrevList	= 35726,
        ksCMMakeMacro	= 35727,
        ksCMDestroyMacro	= 35728,
        ksCMViewObjAttr	= 35729,
        ksCMChangeSelObjStyles	= 35730,
        ksCMChangeSelObjLayer	= 35731,
        ksCMEditObject	= 35736,
        ksCMToggleToObjLayer	= 35737,
        ksCMObjStreamline	= 35738,
        ksCMDelAuxCurves	= 35739,
        ksCMDelAuxCurvesInCurentView	= 35740,
        ksCMDelStamp	= 35741,
        ksCMDelTechnicalDemand	= 35742,
        ksCMDelSpecRough	= 35743,
        ksCMSheetViewStates	= 35744,
        ksCMViewLayerStates	= 35745,
        ksCMGridOnOf	= 35746,
        ksCMSnapSuspend	= 35748,
        ksCMSnapSetup	= 35749,
        ksCMOrthoModeOnOff	= 35750,
        ksCMDiscreteModeOnOff	= 35747,
        ksCMRebuildSheet	= 35751,
        ksCMSheetViewParams	= 35752,
        ksCMTechnicalDemand	= 35753,
        ksCMSpecRough	= 35754,
        ksCMSlaveSpc	= 35755,
        ksCMSlaveChangeList	= 35756,
        ksCMAddEditSpcObject	= 35760,
        ksCMAddEditChangeListObject	= 35761,
        ksCMSpcSinhronize	= 35763,
        ksCMSheetSpc	= 35764,
        ksCMLayoutManager	= 35765,
        ksCMAddPage	= 35766,
        ksCMGoto00	= 36028,
        ksCMToggleCursor	= 36029,
        ksCMRegulateLeaderLineX	= 36075,
        ksCMRegulateLeaderLineY	= 36076,
        ksCMEditSpcObject	= 36077,
        ksCMEditSpcObjectForGeom	= 36086,
        ksCMGridSetup	= 36090,
        ksCMSlaveSpcDelegate	= 36094,
        ksCMAddSpcDelegate	= 36095,
        ksCMAutoCreateSpcObj	= 36097,
        ksCMBuildTreeView	= 37003,
        ksSheetViewParams	= 37005,
        ksEditDocumentProperties	= 37171,
        ksEditInserionFragmentProperties	= 37172,
        ksEditInserionViewProperties	= 37173,
        ksEditMacroObjectProperties	= 37174,
        ksClearAttachedLeaders	= 37186,
        ksCMParameters	= 37608,
        ksCMZoomUndo	= 38530,
        ksCMZoomRedo	= 38531,
        ksViewShowBreakups	= 39344,
        ksCMHideCPlaces	= 40360,
        ksCMHideCPlanes	= 40361,
        ksCMHideCAxies	= 40362,
        ksCMHideSketches	= 40363,
        ksCMHideSurfaces	= 40364,
        ksCMHideThreads	= 40365,
        ksCMHideCurves	= 40366,
        ksCMHidePoints	= 40367,
        ksCMHideAllObjects	= 40368,
        ksCMHideDimensions	= 40369,
        ksCMHideDesignations	= 40370,
        ksCMCreateSheetFromModel	= 40373,
        ksCMDeleteRollbackObjects	= 40375,
        ksCMSelectedObjectProperties	= 40461,
        ksCMSelectedObjectOwnerProperties	= 40462,
        ksCMSelectedCompanentProperties	= 40463,
        ksCMSelectedCompanentInstanceProperties	= 40464,
        ksCMEditBilletPart	= 40621,
        ksCMChangeBilletPart	= 40622,
        ksCMViewFeatureInTree	= 40524,
        ksCMLODOn	= 40610,
        ksCM3DArrayDestroy	= 40615,
        ksCMEmbodimentManager	= 40710,
        ksCMAdditionNumberberManager	= 40711,
        ksCMHideInCompCPlaces	= 40728,
        ksCMHideInCompCPlanes	= 40729,
        ksCMHideInCompCAxies	= 40730,
        ksCMHideInCompSketches	= 40731,
        ksCMHideInCompSurfaces	= 40732,
        ksCMHideInCompThreads	= 40733,
        ksCMHideInCompCurves	= 40734,
        ksCMHideInCompPoints	= 40735,
        ksCMHideInCompAllObjects	= 40736,
        ksCMHideInCompDimensions	= 40737,
        ksCMHideInCompDesignations	= 40738,
        ksCM3DSavePartAs	= 40744,
        ksCM3DUnitParts	= 40745,
        ksCM3DAssemblyDestroy	= 40746,
        ksCMZoomSketch	= 40872,
        ksCMWireframeMode	= 41882,
        ksCMHiddenRemovedMode	= 41883,
        ksCMHiddenThinMode	= 41884,
        ksCMShadedMode	= 41885,
        ksCMPerspective	= 41886,
        ksCMShadedWireframeMode	= 41893,
        ksCMRotateCCW	= 41887,
        ksCMRotateCC	= 41888,
        ksCMRotate90CCW	= 41889,
        ksCMRotate90CC	= 41890,
        ksCMFastLines	= 41891,
        ksCMTreeStructure	= 41904,
        ksCMPropertyEditor	= 45171,
        ksCMFileNew	= 0xe100,
        ksCMFileOpen	= 0xe101,
        ksCMFileClose	= 0xe102,
        ksCMFileSave	= 0xe103,
        ksCMFileSaveAs	= 0xe104,
        ksCMFilePrintSetup	= 0xe106,
        ksCMFilePrint	= 0xe107,
        ksCMFilePrintPreview	= 0xe109,
        ksCMFileSendMail	= 0xe10c,
        ksCMEditClear	= 0xe120,
        ksCMEditClearAll	= 0xe121,
        ksCMEditCopy	= 0xe122,
        ksCMEditCut	= 0xe123,
        ksCMEditFind	= 0xe124,
        ksCMEditPaste	= 0xe125,
        ksCMEditRepeat	= 0xe128,
        ksCMEditReplace	= 0xe129,
        ksCMEditSelectAll	= 0xe12a,
        ksCMEditUndo	= 0xe12b,
        ksCMEditRedo	= 0xe12c,
        ksCMWindowNew	= 0xe130,
        ksCMAbout	= 0xe140,
        ksCMHelpIndex	= 0xe142,
        ksCMHelpFinder	= 0xe143,
        ksCMContextHelp	= 0xe145,
        ksCMHelp	= 0xe146
    } 	ksKompasCommandEnum;

//-------------------------------------------------------------------------------
// _2_21_ ksHideMessageEnum - Скрывать выдачу сообщений.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E243B5C3-7545-4a3c-9F74-9E3C4BF63D9E") 
enum _ksHideMessageEnum
    {
        ksShowMessage	= 0,
        ksHideMessageYes	= 1,
        ksHideMessageNo	= 2
    } 	ksHideMessageEnum;

//-------------------------------------------------------------------------------
//  _4_04_ ksTextNumberingEnum         - Тип нумерации абзаца.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("32AB7F7B-E979-4233-9D10-47F040D227EE") 
enum _ksTextNumberingEnum
    {
        ksTNumbUnknown	= -1,
        ksTNumbNoNumber	= 0,
        ksTNumbNumber	= 1,
        ksTNumbNewNumber	= 2,
        ksTNumbDisableNumber	= 3
    } 	ksTextNumberingEnum;

//-------------------------------------------------------------------------------
// _4_05_ ksTextSizeEnum - Размерный коэффиициент текста.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("C6AB1395-8B39-4b26-A7AA-32F0F824AB3D") 
enum _ksTextSizeEnum
    {
        ksTextDefault	= 0,
        ksTextNormal	= 1,
        ksTextMiddle	= 2,
        ksTextSmall	= 3,
        ksTextBig	= 2
    } 	ksTextSizeEnum;

//-------------------------------------------------------------------------------
// _4_06_ ksTextItemEnum - Тип компоненты текста.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("C09D6A4E-F9B4-492f-8F84-B572312F4A40") 
enum _ksTextItemEnum
    {
        ksTItString	= 0,
        ksTItNumerator	= 0x1,
        ksTItDenominator	= 0x2,
        ksTItFractionEnd	= 0x3,
        ksTItUpperDeviation	= 0x4,
        ksTItLowerDeviation	= 0x5,
        ksTItDeviationEnd	= 0x6,
        ksTItSBase	= 0x7,
        ksTItSUpperIndex	= 0x8,
        ksTItSLowerIndex	= 0x9,
        ksTItSEnd	= 0x10,
        ksTItSpecialSymbol	= 0x11,
        ksTItSpecialSymbolEnd	= 0x12,
        ksTItSpecialSymbolNext	= 0x13,
        ksTItSpecialSymbolDown	= 0x14,
        ksTItSpecialSymbolRight	= 0x15,
        ksTItTab	= 0x16,
        ksTItFontSymbol	= 0x17,
        ksTItSeparator	= 0x18,
        ksTItHyperTextEnd	= 0x19,
        ksTItFontSymbolW	= 0x2017,
        ksTItHyperText	= 0x2000
    } 	ksTextItemEnum;

//-------------------------------------------------------------------------------
// _4_07_ ksTextHorizontalFormatEnum - Признак горизонтального форматирования.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("FA4FA75F-9E15-4e31-BF75-74761BDB8D89") 
enum _ksTextHorizontalFormatEnum
    {
        ksHFormatNot	= 0,
        ksHFormatStrNarrowing	= 1,
        ksHFormatDivision	= 2
    } 	ksTextHorizontalFormatEnum;

//-------------------------------------------------------------------------------
// _4_08_ ksAllocationEnum - Размещенеие объекта.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("10FAF942-8474-4d5c-A840-9C2AC4D0879D") 
enum _ksAllocationEnum
    {
        ksAlLeft	= 0,
        ksAlCentre	= 1,
        ksAlRight	= 2
    } 	ksAllocationEnum;

//-------------------------------------------------------------------------------
// _4_09_ ksMarkInsideFormEnum - Тип формы для марки (без линии-выноски).
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("8C51566B-7098-44b5-A41C-42C5D4DBD73C") 
enum _ksMarkInsideFormEnum
    {
        ksMFormEmpty	= 0,
        ksMFormCirle	= 1,
        ksMFormCircle	= 1,
        ksMFormRectangle	= 2,
        ksMFormSquare	= 3,
        ksMFormRhomb1	= 4,
        ksMFormRhomb2	= 5,
        ksMFormHexagon	= 6,
        ksMFormTriangle1	= 7,
        ksMFormTriangle2	= 8,
        ksMFormChamferedRectangle	= 9,
        ksMFormCircleWidthVerticalDelimer	= 10,
        ksMFormDoubleCircle	= 11
    } 	ksMarkInsideFormEnum;

//-------------------------------------------------------------------------------
// _4_10_ ksMarkOnLinePosTypeEnum	- Положение марки относительно линии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("1467B4FE-2FD1-4dad-A5B4-F1D22CD891B8") 
enum _ksMarkOnLinePosTypeEnum
    {
        ksMTextAboveLine	= 0,
        ksMTextOnLine	= 1,
        ksMTextUnderLine	= 2
    } 	ksMarkOnLinePosTypeEnum;

//-------------------------------------------------------------------------------
// _4_11_ ksArrowEnum - Тип стрелки линии-выноски.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4906870A-5BEC-4059-A138-F68F0833E039") 
enum _ksArrowEnum
    {
        ksLeaderWithoutArrow	= 0,
        ksLeaderPoint	= 1,
        ksLeaderArrow	= 2,
        ksWithoutArrow	= 0,
        ksPoint	= 1,
        ksArrow	= 2,
        ksUpHalfArrow	= 3,
        ksDownHalfArrow	= 4,
        ksNotch	= 5,
        ksLeftNotch	= 6,
        ksRightAngle	= 7,
        ksClosedArrow	= 8,
        ksOpenArrow	= 9,
        ksDot	= 10,
        ksSmallDot	= 11,
        ksTriangle60	= 12,
        ksTriangle90	= 13
    } 	ksArrowEnum;

//-------------------------------------------------------------------------------
// _4_12_ ksShelfDirectionEnum - Направление полки.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("AD207EF3-2367-414b-9DCA-02298979C5B5") 
enum _ksShelfDirectionEnum
    {
        ksLSLeft	= -1,
        ksLSNone	= 0,
        ksLSRight	= 1,
        ksLSUp	= 2,
        ksLSDown	= 3
    } 	ksShelfDirectionEnum;

//-------------------------------------------------------------------------------
// _4_13_ ksDrawingObjecParamTypeEnum - Тип параметров объекта.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("69E71A09-D5D5-4e39-9448-2C9D74218B44") 
enum _ksDrawingObjectParamTypeEnum
    {
        ksAllParam	= -1,
        ksSheetAllParam	= -2,
        ksViewAllParam	= -7
    } 	ksDrawingObjectParamTypeEnum;

//-------------------------------------------------------------------------------
// _4_15_ ksMarkNodeEnum - тип узла марки.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("642CF99F-9839-4a22-B4F4-D069B2E23A68") 
enum _ksMarkNodeEnum
    {
        ksMarkCircle	= 0,
        ksMarkRefCircle	= 1,
        ksMarkText	= 2
    } 	ksMarkNodeEnum;

//-------------------------------------------------------------------------------
// _4_16_ ksCurveStyleEnum - Системные стили линий.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("639E6A83-A881-472f-A827-1822E2F0217E") 
enum _ksCurveStyleEnum
    {
        ksCSHidden	= -1,
        ksCSUnvisible	= 0,
        ksCSNormal	= 1,
        ksCSThin	= 2,
        ksCSAxial	= 3,
        ksCSDashed	= 4,
        ksCSBrokenLine	= 5,
        ksCSConstruction	= 6,
        ksCSThick	= 7,
        ksCSDash2Dots	= 8,
        ksCSDashedNormal	= 9,
        ksCSNormalDashDot	= 10,
        ksCSThinForHatch	= 11,
        ksCSISO02Dashed	= 12,
        ksCSISO03DashedLSpace	= 13,
        ksCSISO04DashDotLDash	= 14,
        ksCSISO05DashDotLDash2Dots	= 15,
        ksCSISO06DashDotLDash3Dots	= 16,
        ksCSISO07Dotted	= 17,
        ksCSISO08DashDotLShDashes	= 18,
        ksCSISO09DashDot1L2ShDashes	= 19,
        ksCSISO10DashDot	= 20,
        ksCSISO11DashDot2Dashes	= 21,
        ksCSISO12DashDot2Dots	= 22,
        ksCSISO13DashDot3Dots	= 23,
        ksCSISO14DashDot2Dashes2Dots	= 24,
        ksCSISO15DashDot2Dashes3Dots	= 25
    } 	ksCurveStyleEnum;

//-------------------------------------------------------------------------------
// _4_17_ ksTextStyleEnum - Системные стили текста.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("AA46EAFD-5375-447e-933E-CFBBF644BAD7") 
enum _ksTextStyleEnum
    {
        ksTSDefault	= 0,
        ksTSDrawingAnnotation	= 1,
        ksTSSpecifications	= 2,
        ksTSDimensionText	= 3,
        ksTSSurfaceFinish	= 4,
        ksTSLeader1	= 5,
        ksTSLeader2	= 6,
        ksTSLeader3	= 7,
        ksTSShapeDeviations	= 8,
        ksTSTableHeader	= 9,
        ksTSTableCell	= 10,
        ksTSSectionLine	= 11,
        ksTSDirectionArrow	= 12,
        ksTSUnspecifiedSurfaceFinish	= 13,
        ksTSModificationSymbol	= 14,
        ksTSBrace	= 15,
        ksTSUnitNumber	= 16,
        ksTSMultiTextLeader	= 17,
        ksTSUnitMarking	= 18,
        ksTSAxisMark	= 19,
        ksTSMarkOnLeader	= 20,
        ksTSMarkOnLine	= 21,
        ksTSMarkInsideForm	= 22,
        ksTSBOMTableName	= 23,
        ksTSBuildingCutLine	= 24,
        ksTSRprtTableHeader	= 25,
        ksTSRprtTableCell	= 26
    } 	ksTextStyleEnum;

//-------------------------------------------------------------------------------
// _4_18_ ksContourFormEnum - форма контура Обозначения узла.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("1A12DDB4-F266-4b65-9D12-DCF636204AB3") 
enum _ksContourFormEnum
    {
        ksUFormCirle	= 0,
        ksUFormCircle	= 0,
        ksUFormRectangle	= 1,
        ksUFormCRectangle	= 2
    } 	ksContourFormEnum;

//-------------------------------------------------------------------------------
// _4_19_ ksAlignmentTypeEnum - Тип ориентации объекта.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B0784C12-3992-470c-BF46-517FD6FCDA96") 
enum _ksAlignmentTypeEnum
    {
        ksATArbitrary	= 0,
        ksATHorizontal	= 1,
        ksATVertical	= 2
    } 	ksAlignmentTypeEnum;

//-------------------------------------------------------------------------------
// _4_20_ ksLengthUnitsEnum - Единицы измерения длины.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6570E442-6815-46f1-857C-EC807FAE3059") 
enum _ksLengthUnitsEnum
    {
        ksLUnSM	= 0,
        ksLUnMM	= 1,
        ksLUnDM	= 2,
        ksLUnM	= 3,
        ksLUnDocument	= 4
    } 	ksLengthUnitsEnum;

//-------------------------------------------------------------------------------
// _4_21_ ksMassUnitsEnum - Единицы измерения массы.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("033FDCDA-1A12-40a0-BA91-3236DDD1CA7A") 
enum _ksMassUnitsEnum
    {
        ksMUnGR	= 0,
        ksMUnKG	= 1,
        ksMUnDocument	= 4
    } 	ksMassUnitsEnum;

//-------------------------------------------------------------------------------
// _4_22_ ksMassSettingModeEnum - Варианты задания МЦХ.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("164C3650-C186-49ca-9A1C-A8B65DE2A444") 
enum _ksMassSettingModeEnum
    {
        ksCalculateParam	= 0,
        ksManualMass	= 1
    } 	ksMassSettingModeEnum;

//-------------------------------------------------------------------------------
// _2_22_ ksConstraintTypeEnum - Типы параметрических ограничений.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("8972F4A7-4766-46d5-B64D-C5EAAEC1C944") 
enum _ksConstraintTypeEnum
    {
        ksCUnknown	= 0,
        ksCFixedPoint	= 1,
        ksCPointOnCurve	= 2,
        ksCHorizontal	= 3,
        ksCVertical	= 4,
        ksCParallel	= 5,
        ksCPerpendicular	= 6,
        ksCEqualLength	= 7,
        ksCEqualRadius	= 8,
        ksCHAlignPoints	= 9,
        ksCVAlignPoints	= 10,
        ksCMergePoints	= 11,
        ksCAssociation	= 12,
        ksCDimWithVariable	= 13,
        ksCFixedDim	= 14,
        ksCTangentTwoCurves	= 15,
        ksCSymmetryTwoPoints	= 16,
        ksCCollinear	= 17,
        ksCFixedAngle	= 18,
        ksCFixedLenght	= 19,
        ksCPointOnCurveMiddle	= 20,
        ksCBisector	= 21
    } 	ksConstraintTypeEnum;

//-------------------------------------------------------------------------------
// _4_23_ ksAnnotationSymbolEnum - Аннотационные символы.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("1D93B3F4-1823-4898-A3FA-D5EA37423FF9") 
enum _ksAnnotationSymbolEnum
    {
        ksUnknownSymbol	= 0,
        ksDotPoint	= 1,
        ksPlusPoint	= 2,
        ksXPoint	= 3,
        ksSquarePoint	= 4,
        ksTrianglePoint	= 5,
        ksCirclePoint	= 6,
        ksAsteriskPoint	= 7,
        ksStrikeSquarePoint	= 8,
        ksPlusPointTwo	= 9
    } 	ksAnnotationSymbolEnum;

//-------------------------------------------------------------------------------
// _4_24_ ksDimensionTextAlignEnum - Выравнивание размерной надписи.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7818F73C-2C33-446c-802F-A59D3A015151") 
enum _ksDimensionTextAlignEnum
    {
        ksDimACentreLowFont	= 0,
        ksDimAUpperBoundary	= 1,
        ksDimACentre	= 2,
        ksDimALowerBoundary	= 3
    } 	ksDimensionTextAlignEnum;

//-------------------------------------------------------------------------------
// _4_25_ ksDimensionTextBracketsEnum - Размер в скобках.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("535F9A75-7E36-4927-82F8-090FA7869B8C") 
enum _ksDimensionTextBracketsEnum
    {
        ksDimBracketsOff	= 0,
        ksDimBrackets	= 1,
        ksDimSquareBrackets	= 2
    } 	ksDimensionTextBracketsEnum;

//-------------------------------------------------------------------------------
// _4_26_ ksDimensionDeviationEnum - Отклонения номинального значения размера.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("C4547645-30E1-4337-8406-F4D589ED2ED2") 
enum _ksDimensionDeviationEnum
    {
        ksDimDeviation	= 0,
        ksDimLimits	= 1,
        ksDimLineLimits	= 2
    } 	ksDimensionDeviationEnum;

//-------------------------------------------------------------------------------
// _4_27_ ksDimensionTextTypeEnum - Тип размерной надписи.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("99C3E9FD-04D0-4c30-AA44-E07D79A9C12B") 
enum _ksDimensionTextTypeEnum
    {
        ksDimTAuto	= 0,
        ksDimTManual	= 1,
        ksDTPOnShelf	= 2
    } 	ksDimensionTextTypeEnum;

//-------------------------------------------------------------------------------
// _4_28_ ksDimensionBaseEnum - Параметр отрисовки текста.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("2F89CC11-6BF8-40b5-B19A-5884BA43A4BC") 
enum _ksDimensionBaseEnum
    {
        ksDimBaseCenter	= 0,
        ksDimBaseP1	= 1,
        ksDimBaseP2	= 2,
        ksDimCommonBase	= 3,
        ksDimFirstCommonBase	= 4
    } 	ksDimensionBaseEnum;

//-------------------------------------------------------------------------------
// _4_29_ ksDimensionTextPosEnum - Положение размерной надписи отноcительно выносной линии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("D9B90501-2DC1-433f-9EB2-06488DDEB3F2") 
enum _ksDimensionTextPosEnum
    {
        ksDimTextParallelOnLine	= 0,
        ksDimTextParallelInCut	= 1,
        ksDimTextHorizontalInCut	= 2
    } 	ksDimensionTextPosEnum;

//-------------------------------------------------------------------------------
// _4_30_ ksDimensionArrowPosEnum - Размещение стрелок относительно выносной линии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("D3A13BB3-638C-451a-9C17-59473F015B88") 
enum _ksDimensionArrowPosEnum
    {
        ksDimArrowInside	= 0,
        ksDimArrowOutside	= 1,
        ksDimArrowAuto	= 2
    } 	ksDimensionArrowPosEnum;

//-------------------------------------------------------------------------------
// _4_31_ ksLineDimensionOrientationEnum - Тип ориентации линейного размера.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("5376D845-0CDD-4488-BB18-3837D775A5CA") 
enum _ksLineDimensionOrientationEnum
    {
        ksLinDParallel	= 0,
        ksLinDHorizontal	= 1,
        ksLinDVertical	= 2
    } 	ksLineDimensionOrientationEnum;

//-------------------------------------------------------------------------------
// _4_32_ ksHeightDimTypeEnum - Тип размеров высоты.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("5E8F5BD3-4DE4-4223-953F-14806D7248D8") 
enum _ksHeightDimTypeEnum
    {
        ksHDFrontView	= 0,
        ksHDTopView	= 1,
        ksHDTopViewLeader	= 2
    } 	ksHeightDimTypeEnum;

//-------------------------------------------------------------------------------
// _4_33_ ksAngleDimTypeEnum - Тип углового размера.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4A4FEB78-1C67-46e8-B639-79F2C93452DA") 
enum _ksAngleDimTypeEnum
    {
        ksADMinAngle	= 0,
        ksADMaxAngle	= 1,
        ksADMoreAngle	= 2
    } 	ksAngleDimTypeEnum;

//-------------------------------------------------------------------------------
// _4_33_ ksDimTextFormatEnum - Формат отображения размерной надписи.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7632395C-E6E9-4709-9366-C2EC90451BDD") 
enum _ksDimTextFormatEnum
    {
        ksDimTextFormatGMS	= 0,
        ksDimTextFormatGDD	= 1
    } 	ksDimTextFormatEnum;

//-------------------------------------------------------------------------------
// _4_34_ ksRoughSignEnum - Тип значка шероховатости.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("325C3FD7-9095-4e8c-BBDE-0F7157B7B2E2") 
enum _ksRoughSignEnum
    {
        ksNoProcessingType	= 0,
        ksDeleteMaterial	= 1,
        ksWithoutDeleteMaterial	= 2
    } 	ksRoughSignEnum;

//-------------------------------------------------------------------------------
// _4_35_ ksLeaderSignEnum - Тип значка для линии выноски.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B0C3EEB5-464E-42f2-8010-C7E627B44867") 
enum _ksLeaderSignEnum
    {
        ksLSignNone	= 0,
        ksLGlueSign	= 1,
        ksLSolderingSign	= 2,
        ksLSewingSign	= 3,
        ksLCrampSign	= 4,
        ksLcornerCrampSign	= 5,
        ksLMontageJointSign	= 6
    } 	ksLeaderSignEnum;

//-------------------------------------------------------------------------------
// _4_36_ ksPositionLederFormEnum - Тип формы для позиционной линии выноски.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("FBC15446-0C5E-4d09-9DB4-B55E75F61FA2") 
enum _ksPositionLederFormEnum
    {
        ksPLSingleText	= 0,
        ksPLOpenText	= 1,
        ksPLCircle	= 2,
        ksPLHexagon	= 3,
        ksPLCircleWithSeparator	= 4
    } 	ksPositionLederFormEnum;

//-------------------------------------------------------------------------------
// _4_37_ ksChangeLeaderSignEnum - Тип значка для обозначения изменения.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("FDD4B663-EBD9-464e-B4D7-4DA4BB885300") 
enum _ksChangeLeaderSignEnum
    {
        ksCLSSquare	= 0,
        ksCLSCircle	= 1,
        ksCLSBracketSquare	= 2,
        ksCLSBracketCircle	= 3,
        ksCLSBracketCorner	= 4
    } 	ksChangeLeaderSignEnum;

//-------------------------------------------------------------------------------
// _4_37_ ksTolerancePrefixSignEnum - Знак в обозначении допуска.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4A6A4FEE-7731-40e5-AE3A-BE31D8CD6DC6") 
enum _ksTolerancePrefixSignEnum
    {
        ksTPSNone	= 0,
        ksTPSRadius	= 1,
        ksTPSDiametr	= 2,
        ksTPSToleranseT	= 3,
        ksTPSToleranseT2	= 4
    } 	ksTolerancePrefixSignEnum;

//-------------------------------------------------------------------------------
// _4_38_ ksToleranceSuffixSignEnum - Знак в обозначении базы допуска.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("776549B9-1E13-43f0-9284-19CD01157313") 
enum _ksToleranceSuffixSignEnum
    {
        ksTSNone	= 0,
        ksTSToleranseM	= 1,
        ksTSToleranseS	= 2,
        ksTSToleranseP	= 3
    } 	ksToleranceSuffixSignEnum;

//-------------------------------------------------------------------------------
// _4_39_ ksTablePointEnum - Тип расположения точки на таблице.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("96D077FC-CDA7-44d2-B005-35E1F05156AB") 
enum _ksTablePointEnum
    {
        ksTPUndefined	= 0,
        ksTPLeftBottom	= 1,
        ksTPLeftCenter	= 2,
        ksTPLeftUp	= 3,
        ksTPUpCenter	= 4,
        ksTPRightUp	= 5,
        ksTPRightCenter	= 6,
        ksTPRightBottom	= 7,
        ksTPBottomCenter	= 8,
        ksTPCenter	= 9
    } 	ksTablePointEnum;

//-------------------------------------------------------------------------------
// _4_40_ ksTableTileLayoutEnum - Расположение заголовка таблицы.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("15467B11-003D-4166-A80A-132EFCF9ABC9") 
enum _ksTableTileLayoutEnum
    {
        ksTTLFirstRow	= 0,
        ksTTLFirstColumn	= 1,
        ksTTLNotCreate	= 2
    } 	ksTableTileLayoutEnum;

//-------------------------------------------------------------------------------
// _4_41_ ksCellBoundariesEnum - Типы границ таблицы.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7BFFE5C0-CAE7-490f-8CC6-A97E85614120") 
enum _ksCellBoundariesEnum
    {
        ksCBLeftBorder	= 0,
        ksCBRightBorder	= 1,
        ksCBTopBorder	= 2,
        ksCBBottomBorder	= 3,
        ksCBHorisontMidleBorder	= 4,
        ksCBVerticalMidleBorder	= 5,
        ksCBExternalBorders	= 6,
        ksCBAllBorders	= 7,
        ksCBInternalBorders	= 8,
        ksCBNoBorders	= 9
    } 	ksCellBoundariesEnum;

//-------------------------------------------------------------------------------
// _4_42_  ksTextLineType - Тип строки текста.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("FA02990D-B1D1-4ada-B58F-C28772BC5123") 
enum _ksTextLineType
    {
        ksTLError	= 0,
        ksTLText	= 1,
        ksTLVerticalText	= 2,
        ksTLFragment	= 3,
        ksTLRaster	= 4,
        ksTLTable	= 5
    } 	ksTextLineType;

//-------------------------------------------------------------------------------
// _2_38_ ksCheckSumVersionEnum - Версии котрольных сумм (ICheckSum).
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("EF9835A2-F22D-4549-B5F5-C5C7A7AF60CF") 
enum _ksCheckSumVersionEnum
    {
        ksCsrCurrent	= 0,
        ksCsrKompas10SP2	= -1,
        ksCsrKompas11	= -5,
        ksCsrKompas11SP1	= -6,
        ksCsrKompas12	= -10,
        ksCsrKompas12SP1	= -11,
        ksCsrKompas13	= -15,
        ksCsrKompas13SP1	= -16,
        ksCsrKompas13SP2	= -17,
        ksCsrKompas14	= -20,
        ksCsrKompas14SP1	= -21,
        ksCsrKompas14SP2	= -22,
        ksCsrKompas15	= -25,
        ksCsrKompas15Sp1	= -26,
        ksCsrKompas15Sp2	= -27,
        ksCsrKompas16	= -30,
        ksCsrKompas16Sp1	= -31,
        ksCsrKompas17	= -35
    } 	ksCheckSumVersionEnum;

//-------------------------------------------------------------------------------
// _2_23_ ksBmpSizeEnum				    - Размеры иконок.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("C50E703F-6D99-4a83-B4AF-1D8836C8968B") 
enum _ksBmpSizeEnum
    {
        ksBmp1616	= 0,
        ksBmp2424	= 1,
        ksBmp3232	= 2,
        ksBmp4848	= 3
    } 	ksBmpSizeEnum;

  //-------------------------------------------------------------------------------
  // _2_24_ Тип процесса запрашивающщего файл или список файлов.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E8397E6E-FA85-4955-952C-B7CB480E43E2") 
enum _ksRequestFilesTypeEnum
    {
        ksRFUnknown	= 0,
        ksRFSaveBody	= 1,
        koRFUnitParts	= 2,
        koRFCopyBilletPart	= 3,
        koRFSavePartAs	= 4,
        koRFAddDetail	= 5,
        koRFAddAssembly	= 6,
        koRFAddPartFromFile	= 7,
        koRFChangeDetailFile	= 8,
        koRFChangeAssemblyFile	= 9,
        koRFChangeBilletPartFile	= 10,
        koRFSpcObjAddDocument	= 11,
        koRFSpcAssemblyAddDocument	= 12,
        koRFAddLocalDetail	= 13,
        koRFAddLayoutGeometry	= 14
    } 	ksRequestFilesTypeEnum;

//-------------------------------------------------------------------------------
// _4_43_ ksMlVertexLimiterEnum				    - Типы ограничений в вершинах мультилинии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("CC68C0AA-CF0F-4f03-B4A3-C48C2506C3C3") 
enum _ksMlVertexLimiterEnum
    {
        ksMVeUnknown	= 0,
        ksMVeArc	= 1,
        ksMVeAngle	= 2,
        ksMVeTangent	= 3
    } 	ksMlVertexLimiterEnum;

//-------------------------------------------------------------------------------
// _4_44_ ksMlEndLimiterEnum				    - Типы ограничений на концах мультилинии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("A923E285-76F6-45f6-80F9-9D927C50EF2B") 
enum _ksMlEndLimiterEnum
    {
        ksMEndUnknown	= 0,
        ksMEndLine	= 1,
        ksMEndArc	= 2,
        ksMEndPolyline	= 3
    } 	ksMlEndLimiterEnum;

//-------------------------------------------------------------------------------
// _4_45_ ksMlVertexTrackingEnum				    - Типы обхода вершин мультилинии.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("BB476F4B-ECC0-4055-8571-BDA127DA9062") 
enum _ksMlVertexTrackingEnum
    {
        ksMTrShear	= 0,
        ksMTrFillet	= 1,
        ksMTrEquaFillet	= 2
    } 	ksMlVertexTrackingEnum;

//-------------------------------------------------------------------------------
// _4_46_ ksContourSegmentEnum				    - Типы сегментов контура.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4E89BEF0-3351-40af-BC92-5C9E9AF7E7F2") 
enum _ksContourSegmentEnum
    {
        ksCSUnknown	= -1,
        ksCSLineSeg	= 1,
        ksCSCircle	= 2,
        ksCSArc	= 3,
        ksCSBezier	= 8,
        ksCSEllipse	= 32,
        ksCSNurbs	= 33,
        ksCSEllipseArc	= 34
    } 	ksContourSegmentEnum;

//-------------------------------------------------------------------------------
// _4_47_ ksAnnotativeSymbolEnum - Типы специальных символов для аннатационных объектов.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("F3291AD5-863D-474c-BD1A-C3EC884B45B6") 
enum _ksAnnotativeTerminatorSignEnum
    {
        ksASUnknown	= 0,
        ksASArrowInside	= 101,
        ksASArrowOutside	= 102,
        ksASNotchTail	= 103,
        ksASUpHalfArrow	= 104,
        ksASDownHalfArrow	= 105,
        ksASBigArrowInside	= 106,
        ksASArrowOrdinate	= 107,
        ksASTriangle	= 108,
        ksAScircleRad2	= 109,
        ksASCentreMarker	= 110,
        ksASGlueSign	= 111,
        ksASSolderingSign	= 112,
        ksASSewingSign	= 113,
        ksASCrampSign	= 114,
        ksASCornerCrampSign	= 115,
        ksASMontageJointSign	= 116,
        ksASNotch	= 117,
        ksASBaseTriangle	= 118,
        ksASClosedArrowInside	= 119,
        ksASClosedArrowOutside	= 120,
        ksASOpenArrowInside	= 121,
        ksASOpenArrowOutside	= 122,
        ksASRightAngleInside	= 123,
        ksASRightAngleOutside	= 124,
        ksASDot	= 125,
        ksASSmallDot	= 126,
        ksASPoint	= 127,
        ksASLeftNotch	= 128
    } 	ksAnnotativeTerminatorSignEnum;

//-------------------------------------------------------------------------------
// _4_48_ ksCornerTypeEnum - Тип угла объекта.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("CC445886-254C-485b-A5DB-252EA88E11B4") 
enum _ksCornerTypeEnum
    {
        ksCTNoProcess	= 0,
        ksCTChamfer	= 1,
        ksCTFillet	= 2
    } 	ksCornerTypeEnum;

//-------------------------------------------------------------------------------
// _4_49_ ksInsertionTypeEnum - Тип вставки фрагмента или вида.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("54AA5AF9-FB7C-400f-96A8-D3BC49D03985") 
enum _ksInsertionTypeEnum
    {
        ksTUnknown	= -1,
        ksTBodyFragment	= 0,
        ksTReferenceFragment	= 1,
        ksTLocalFragment	= 3,
        ksTBodyView	= 4,
        ksTReferenceView	= 5
    } 	ksInsertionTypeEnum;

//-------------------------------------------------------------------------------
// _4_50_ ksEquidistantTypeEnum  - Тип построения эквидистанты.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("EBDCE31B-939E-4a7a-8007-94E6A680BEC1") 
enum _ksEquidistantTypeEnum
    {
        ksETUnknown	= -1,
        ksETLeft	= 0,
        ksETRight	= 1,
        ksETBoth	= 2
    } 	ksEquidistantTypeEnum;

//-------------------------------------------------------------------------------
// _4_51_ ksCentreMarkerEnum  - Тип обозначения центра.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6E5C0557-2927-401a-BC13-475D68EB8601") 
enum _ksCentreMarkerEnum
    {
        ksCMUnknown	= -1,
        ksCMPlusPoint	= 0,
        ksCMOneAxis	= 1,
        ksCMTwoAxis	= 2
    } 	ksCentreMarkerEnum;

//-------------------------------------------------------------------------------
// _4_52_ ksSemiAxisTypeEnum  - Тип полуоси для обозначения центра.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E8659F22-93CF-4836-8FE6-FBBE15CA1463") 
enum _ksSemiAxisTypeEnum
    {
        ksAxisUnknown	= -1,
        ksAxisXPlus	= 0,
        ksAxisXMinus	= 1,
        ksAxisYPlus	= 2,
        ksAxisYMinus	= 3
    } 	ksSemiAxisTypeEnum;

//-------------------------------------------------------------------------------
// _4_53_ ksAttributeTypeEnum  - Тип данных для типа аттрибута.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("9CB8488E-D11F-445e-A670-B7B00710E4A7") 
enum _ksAttributeTypeEnum
    {
        ksATUnknown	= -1,
        ksATString	= 0,
        ksATDouble	= 1,
        ksATFixedTable	= 2,
        ksATVariableTable	= 3
    } 	ksAttributeTypeEnum;

//-------------------------------------------------------------------------------
// _4_54_ ksRelativeProjectionTypeEnum - Тип проекции стандартного вида относительно главного вида.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4AB63D61-B238-4db1-B8E9-3FB17EC00647") 
enum _ksRelativeProjectionTypeEnum
    {
        ksPtNone	= -1,
        ksPtFront	= 1,
        ksPtRear	= 2,
        ksPtUp	= 3,
        ksPtDown	= 4,
        ksPtLeft	= 5,
        ksPtRight	= 6,
        ksPtIsoXYZ	= 7
    } 	ksRelativeProjectionTypeEnum;

//-------------------------------------------------------------------------------
// _4_55_ ksBreakLineTypeEnum - Тип линии разрыва.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("261421E7-E77C-48cb-BD88-F44873F45FD8") 
enum _ksBreakLineTypeEnum
    {
        ksBLNotImage	= 0,
        ksBLStraight	= 1,
        ksBLBreak	= 2,
        ksBLVawe	= 3
    } 	ksBreakLineTypeEnum;

//-------------------------------------------------------------------------------
// _2_25_ ksAngleEnum - Углы поворотов кратные 90 гр.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("9B1A532C-0B9A-4540-BAC7-747340563F4A") 
enum _ksAngleEnum
    {
        ksAngle0	= 0,
        ksAngle90	= 1,
        ksAngle180	= 2,
        ksAngle270	= 3
    } 	ksAngleEnum;

//-------------------------------------------------------------------------------
// _4_56_ ksHatchStyleEnum - Системные стили штриховки.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("4E8DB7B5-5EF7-41f1-ACDF-EEA801DC8DE5") 
enum _ksHatchStyleEnum
    {
        ksHatchMetal	= 0,
        ksHatchNonMetal	= 1,
        ksHatchTimber	= 2,
        ksHatchNaturalStone	= 3,
        ksHatchCeramics	= 4,
        ksHatchConcrete	= 5,
        ksHatchGlass	= 6,
        ksHatchLiquid	= 7,
        ksHatchNaturallyGround	= 8,
        ksHatchSpreadGround	= 9,
        ksHatchArtificialStone	= 10,
        ksHatchReinforcedConcrete	= 11,
        ksHatchTenseReinforcedConcrete	= 12,
        ksHatchLongitudalTimber	= 13,
        ksHatchSand	= 14
    } 	ksHatchStyleEnum;

//-------------------------------------------------------------------------------
// _4_57_ ksRegionTypeEnum - Тип региона.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("CA52DD30-F138-43ba-B091-5226A22A2702") 
enum _ksRegionTypeEnum
    {
        ksRTInside	= 1,
        ksRTOutside	= 2,
        ksRTCutFrame	= 3
    } 	ksRegionTypeEnum;

//-----------------------------------------------------------------------------
// _4_58_ ksColourTypeEnum - Тип заливки.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("89A01C69-CA0A-42a9-8760-B7A8563FEAAF") 
enum _ksColouringTypeEnum
    {
        ksColouringSolid	= 0,
        ksColouringLinear	= 1,
        ksColouringAngle	= 2,
        ksColouringCone	= 3,
        ksColouringCircle	= 4,
        ksColouringSquare	= 5,
        ksColouringCylinder	= 6
    } 	ksColouringTypeEnum;

//-------------------------------------------------------------------------------
// _4_59_ ksSheetsRangeEnum - Тип диапазона страниц.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("3AA2DDD5-77F9-4ff8-8806-F78E1B179254") 
enum _ksSheetsRangeEnum
    {
        ksAllSheets	= 0,
        ksUnevenSheets	= 1,
        ksEvenSheets	= 2
    } 	ksSheetsRangeEnum;

//-------------------------------------------------------------------------------
// _4_60_ ksPrintJob_OutputColorTypeEnum - Цвет вывода на печать.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("74C95B1C-1E64-441e-A082-7B23B01FC6F4") 
enum _ksOutputColorTypeEnum
    {
        ksPJ_OCBlack	= 0,
        ksPJ_OCByView	= 1,
        ksPJ_OCByLayer	= 2,
        ksPJ_OCByObject	= 3
    } 	ksOutputColorTypeEnum;

	//------------------------------------------------------------------------------
	// _7_01_ ksPropertyTypeEnum - Типы свойств.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6B0CECF7-3587-4aae-85F3-94486B3ECC9B") 
enum _PropertyTypeEnum
    {
        ksPropertyDataTypeUnknown	= 0,
        ksPropertyDataTypeLong	= 1,
        ksPropertyDataTypeDouble	= 2,
        ksPropertyDataTypeString	= 3,
        ksPropertyDataTypeBoolean	= 4,
        ksPropertyDataTypeColorRGB	= 5,
        ksPropertyDataTypeHatchStyle	= 6,
        ksPropertyDataTypeGroup	= 7
    } 	ksPropertyTypeEnum;

  //------------------------------------------------------------------------------
  // _7_02_ ksReportFiltersTypeEnum - Типы фильтров в команде создать отчет.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("19CD8BB7-D814-4675-9D46-4B24CEE753A6") 
enum _ReportFiltersTypeEnum
    {
        ksFilterConditionUnknown	= 0,
        ksFilterConditionEqual	= 1,
        ksFilterConditionSmaller	= 2,
        ksFilterConditionLarger	= 3,
        ksFilterConditionEqualOrSmaller	= 4,
        ksFilterConditionEqualOrLarger	= 5,
        ksFilterConditionContain	= 6,
        ksFilterConditionNotContain	= 7
    } 	ksReportFiltersTypeEnum;

  //------------------------------------------------------------------------------
  // _4_61_ ksHypertextTypeEnum - Тип ссылки на текст.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("02EFD240-5A41-48d0-BDD6-E28D49F09E41") 
enum _ksHypertextTypeEnum
    {
        ksHTViewScale	= 1,
        ksHTViewUnfold	= 3,
        ksHTViewTurn	= 4,
        ksHTViewTurnAngle	= 5,
        ksHTViewName	= 6,
        ksHTObjectText	= 0,
        ksHTObjectText1	= 1,
        ksHTObjectText2	= 2,
        ksHTObjectText3	= 3,
        ksHTObjectText4	= 4,
        ksHTVariableValue	= 0,
        ksHTVariableName	= 1,
        ksHTVariableExpression	= 2,
        ksHTVariableComment	= 3,
        ksHTPositionNumber	= 0x10,
        ksHTSpecificationObjectMarking	= 0x80,
        ksHTSpecificationObjectName	= 0x81,
        ksHTSpecificationObjectCount	= 0x82,
        ksHTObjectProperty	= 0x80,
        ksHTObjectZone	= -1,
        ksHTObjectSheet	= -2,
        ksHTViewTitleZone	= -1,
        ksHTViewTitleSheet	= -2,
        ksHTViewBaseObjectZone	= -3,
        ksHTViewBaseObjectSheet	= -4
    } 	ksHypertextTypeEnum;

	//------------------------------------------------------------------------------
	// _4_62_ ksHyperLinkTypeEnum - Тип гиперссылки.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E0F56898-D122-4a78-BCF9-481B93C2AD12") 
enum _ksHyperLinkTypeEnum
    {
        ksHLUnknown	= 0,
        ksHLFile	= 1,
        ksHLObject	= 2,
        ksHLMail	= 3
    } 	ksHyperLinkTypeEnum;

// -------------------------------------------------------------------------------
// _2_26_ ksChooseManagerTypeEnum - Тип менеджера выбора объектов.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("AB0E173C-8F52-4cd2-9072-BF9A5397A2EE") 
enum _ksChooseManagerTypeEnum
    {
        ksChMUnknown	= -2,
        ksChMAllColors	= -1,
        ksChMLevel1ColorBase	= 0,
        ksChMLevel1Color1	= 1,
        ksChMLevel1Color2	= 2,
        ksChMLevel2ColorBase	= 100,
        ksChMLevel2Color1	= 101,
        ksChMLevel2Color2	= 102
    } 	ksChooseManagerTypeEnum;

// -------------------------------------------------------------------------------
// _2_27_ ksRecoverErrorEnum - Признак ошибок при открытии документа с восстановлением.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("40C81AAB-217E-4e06-9BBB-5EC6DF1A87FB") 
enum _ksRecoverErrorEnum
    {
        ksRNoError	= 0,
        ksRecover	= 1,
        ksRNoOpen	= 2,
        ksRAlreadyOpen	= 3,
        ksRProtected	= 4
    } 	ksRecoverErrorEnum;

	//------------------------------------------------------------------------------
	// _4_63_ ksBisectorVariant - Вариант решения биссектрисы для двух прямых.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("39B8D459-327D-4f76-8957-ADA09F06F0F7") 
enum _ksBisectorVariant
    {
        ksBVNone	= 0,
        ksBVNormalSum	= 1,
        ksBVNormalDiff	= 2
    } 	ksBisectorVariant;

// -------------------------------------------------------------------------------
// _2_28_ ksExternalFilesTypesEnum - Тип внешнего файла.
// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("9EF7CACA-16BC-4106-A2BF-DCE50E7FDF2B") 
enum _ksExternalFilesTypesEnum
    {
        ksEFTUnknown	= -1,
        ksEFTDocumentFile	= 0,
        ksEFTCurveStyleLibrary	= 1,
        ksEFTTextStyleLibrary	= 2,
        ksEFTHatchStyleLibrary	= 3,
        ksEFTAttributeTypesLibrary	= 4,
        ksEFTLayoutsLibrary	= 5,
        ksEFTFragmentFile	= 6,
        ksEFTFragmentsLibrary	= 7,
        ksEFTSheetConnectedToSpc	= 8,
        ksEFTSystemFile	= 9,
        ksEFTIngotsLibrary	= 10,
        ksEFTPartFile	= 11,
        ksEFTRasterFile	= 12,
        ksEFTDocConnectedToSpcObj	= 13,
        ksEFTSpcConnectedToSheet	= 14,
        ksEFTDocConnectedToSpcObjInOtherDocs	= 15,
        ksEFTAssemblyFile	= 16,
        ksEFTModelsLibrary	= 17,
        ksEFTTemporaryFile	= 18,
        ksEFTSourceFileForVariable	= 19,
        ksEFTServiceFile	= 20,
        ksEFTDraftFile	= 21,
        ksEFTHyperLink	= 22,
        ksEFTDataFile	= 23,
        ksEFTReadingRegimFile	= 24,
        ksEFTAssRprtTable	= 25,
        ksEFTEexternalHyperText	= 26,
        ksEFTCopyExternalGeometry	= 27,
        ksEFExternalBilletPart	= 28,
        ksEFTExternalLayoutGeometry	= 29
    } 	ksExternalFilesTypesEnum;

  //------------------------------------------------------------------------------
  // _2_29_ Типы булевых операций над твердыми телами и замкнутыми кривыми
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("58EB5E98-755F-477b-8EAE-40356D8235D3") 
enum _BooleanType
    {
        ksBooleanUnknown	= 0,
        ksIntersect	= 1,
        ksDifference	= 2,
        ksUnion	= 3
    } 	ksBooleanType;

  //------------------------------------------------------------------------------
  // _2_30_ Тип привязки к объектам
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("F9E57A4F-F515-459b-A4E3-ED0515FD8DA3") 
enum _ksSnapTypeEnum
    {
        ksSTUnknown	= -1,
        ksSTUndefine	= 0,
        ksSTNearestPoint	= 1,
        ksSTNearestMiddle	= 2,
        ksSTObjectCentre	= 3,
        ksSTIntersect	= 4,
        ksSTGrid	= 5,
        ksSTXYAlign	= 6,
        ksSTAngleSnap	= 7,
        ksSTPointOnCurve	= 8,
        ksSTNormalToCurve	= 9,
        ksSTTangentToCurve	= 10
    } 	ksSnapTypeEnum;

	//------------------------------------------------------------------------------
	// _2_31_ ksNewDocumentSettingsTypeEnum - Тип настроек новых документов.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("10662ADF-FE7C-4ba2-B734-EB87ED98B4B6") 
enum _ksNewDocumentSettingsTypeEnum
    {
        ksNewPartDocumentSettings	= 1
    } 	ksNewDocumentSettingsTypeEnum;

	//------------------------------------------------------------------------------
	// _1_108_ ksMaterialPropertyTypeEnum - Тип события выбора материала.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("FCBE2BE9-E4B8-4ae2-8DDE-0EE21CE92BCB") 
enum _ksMaterialPropertyTypeEnum
    {
        ksMPNewPartDocumentSettings	= 1
    } 	ksMaterialPropertyTypeEnum;

	//------------------------------------------------------------------------------
	// _7_03_ ksVariantMarkingTypeEnum - Параметры формирования обозначения.
	// ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("8CE4C7AC-EA8C-4027-8726-BF0518A5DEDB") 
enum _ksVariantMarkingTypeEnum
    {
        ksVMFullMarking	= -1,
        ksVMBaseMarking	= 0x1,
        ksVMEmbodimentNumber	= 0x2,
        ksVMAdditionalNumber	= 0x4,
        ksVMCode	= 0x8
    } 	ksVariantMarkingTypeEnum;

  //------------------------------------------------------------------------------
  // _7_04_ ksReportTypeEnum - Тип отчета.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("2E435FAE-AC8C-4f4f-9D9A-15B6698CFEA9") 
enum _ksReportTypeEnum
    {
        ksRTPropertiesReport	= 0,
        ksRTEmbodimentsReport	= 1,
        ksRTPatternWithVariablesTableReport	= 2,
        ksRTAdditionNumbersReport	= 3
    } 	ksReportTypeEnum;

  //------------------------------------------------------------------------------
  // _7_05_ ksValueFormatEnum - Формат формирования текста для свойства
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("6C94797E-068B-4b88-98B5-080C59FF7F4F") 
enum _ksValueFormatEnum
    {
        ksVFVariableValue	= ( 0x1 | 0x8 ) ,
        ksVFvvfToleranceDiv	= ( ( 0x2 | 0x4 )  | 0x8 ) ,
        ksVFvvfFullValue	= ( ( ( 0x1 | 0x2 )  | 0x4 )  | 0x8 ) ,
        ksVFvvfGridFullValue	= ( ( 0x1 | 0x2 )  | 0x4 ) ,
        ksVFBooleanValue	= 0x1,
        ksVFBooleanYesNo	= 0x2,
        ksVFBooleanPlusMinus	= 0x3,
        ksVFFullMarking	= 0x1,
        ksVFOnlyAdditionMarking	= 0x2
    } 	ksValueFormatEnum;

  //------------------------------------------------------------------------------
  // _7_06_ ksGroupeTypeEnum - Тип группировки колонки отчета
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("AB5E7492-47B9-4491-874C-32883AF18A58") 
enum _ksGroupeTypeEnum
    {
        ksGTNone	= 0,
        ksGTMatch	= 1,
        ksGTSum	= 2,
        ksGTRange	= 3,
        ksGTEnum	= 4
    } 	ksGroupeTypeEnum;

  //------------------------------------------------------------------------------
  // _7_07_ ksNumberingTypeEnum - Формат нумерации.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E7DFA4A0-CDF0-4974-B360-AE2E9CF5409C") 
enum _ksNumberingTypeEnum
    {
        ksNTArabicNumerals	= 0,
        ksNTRomanNumerals	= 1,
        ksNTUpperRegEnglish	= 2,
        ksNTLowerRegEnglish	= 3,
        ksNTUpperRegRussian	= 4,
        ksNTLowerRegRussian	= 5
    } 	ksNumberingTypeEnum;

  //------------------------------------------------------------------------------
  // _7_08_ ksRowsNumberingTypeEnum - Формат нумерации строк отчета.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7BFE72BD-889B-4ad4-ACF7-F1DDE53CCC16") 
enum _ksRowsNumberingTypeEnum
    {
        ksRNTNone	= 0,
        ksRNTSimple	= 1,
        ksRNTMultiLevel	= 2
    } 	ksRowsNumberingTypeEnum;

  //------------------------------------------------------------------------------
  // _7_09_ ksReportStyleInitEnum - Способ инициализации стиля отчета.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("73033645-CB59-4096-8FAB-C431CF97BD5D") 
enum _ksReportStyleInitEnum
    {
        ksRSIDefault	= 0
    } 	ksReportStyleInitEnum;

  //------------------------------------------------------------------------------
  // _2_33_ ksSheetType - Тип листа
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("A81E89B8-2FE3-49c8-ADA6-AF7C63F89176") 
enum _SheetTypeEnum
    {
        ksDocumentSheet	= 0,
        ksFrontAdditionalSheet	= 1,
        ksLastAdditionalSheet	= 2
    } 	ksSheetTypeEnum;

  //------------------------------------------------------------------------------
  // _2_34_ ksDrawInScreenPlaneEnum - Способ преобразования координат внешней триангуляции.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("7EFFC011-CDEE-4B7B-B9BE-51BFB117405F") 
enum _ksDrawInScreenPlaneEnum
    {
        ksDrawNone	= 0,
        ksDrawInScreenPlane	= 1,
        ksDrawProjectToScreen	= 2,
        ksDrawProjectFromScreen	= 3
    } 	ksDrawInScreenPlaneEnum;

  //------------------------------------------------------------------------------
  // _2_35_ ksTextAlignEnum - Выравнивание текста.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("0DD9E984-BDC1-41B4-B327-3E21264947DB") 
enum _ksTextAlignEnum
    {
        ksTALeft	= 0,
        ksTARight	= 2,
        ksTAHCenter	= 6,
        ksTATop	= 0,
        ksTABottom	= 8,
        ksTABaseline	= 24,
        ksTAVCenter	= 56
    } 	ksTextAlignEnum;

  //------------------------------------------------------------------------------
  // _7_10_ ksReportBuildingTypeEnum - Способ выбора объектов для отчета.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("978AC4AB-2D1E-4d53-9058-4C01D1FD9FD6") 
enum _ksReportBuildingTypeEnum
    {
        ksRBAllObjects	= 0,
        ksRBChooseObjects	= 1,
        ksRBCurrentView	= 2,
        ksRBChoiceToLevel	= 3
    } 	ksReportBuildingTypeEnum;

  //------------------------------------------------------------------------------
  // _7_11_ ksPageLayoutTypeEnum - Тип компоновки таблиц отчета.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("E9209423-8B0B-4320-81CB-7BC36E342090") 
enum _ksPageLayoutTypeEnum
    {
        ksRPLayoutDefault	= 0,
        ksRPLayout1	= 1
    } 	ksPageLayoutTypeEnum;

  //------------------------------------------------------------------------------
  // _7_12_ ksObjectPropertyControlTypeEnum  - Типы контролов для отображения свойств в окне свойств.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("2142D68B-FF3A-4e66-9304-EBFF6DB2E6D6") 
enum _ksObjectPropertyControlTypeEnum
    {
        ksOPControlGroup	= 0,
        ksOPControlPointCoord	= 1,
        ksOPControlPointStyle	= 2,
        ksOPControlLineStyle	= 4,
        ksOPControlHatchStyle	= 5,
        ksOPControlColor	= 6,
        ksOPControlEditDouble	= 7,
        ksOPControlEditInt	= 10,
        ksOPControlListDouble	= 11,
        ksOPControlListInt	= 14,
        ksOPControlListCheck	= 15,
        ksOPControlListBmp	= 16,
        ksOPControlListString	= 17,
        ksOPControlSimpleText	= 19,
        ksOPControlText	= 25,
        ksOPControlEdit	= 31,
        ksOPControlExternalEdit	= 34,
        ksOPControlExternalStringEdit	= 35,
        ksOPControlFontImageList	= 38,
        ksOPControlExternalText	= 44,
        ksOPControlExternalMultyText	= 45
    } 	ksObjectPropertyControlTypeEnum;

  //------------------------------------------------------------------------------
  // _2_32_ ksPrinterTypeEnum - Текущий принтер.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("747F9EBC-E7EE-45d9-8EA2-471E2A68025B") 
enum _ksPrinterTypeEnum
    {
        ksPTPrintPreviewPrinter	= 0,
        ksPTSpecialPrinter	= 1
    } 	ksPrinterTypeEnum;

  //------------------------------------------------------------------------------
  // _6_08_ ksSpecificationStyleDifferenceTypeEnum - Отличие стиля СП от библиотечного.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("767EC7B7-FE71-4FDA-A0D3-A8BC0651CC1F") 
enum _ksSpecificationStyleDifferenceTypeEnum
    {
        ksSpcStyleEqual	= 0,
        ksSpcStyleDistinguish	= 1,
        ksSpcStyleNotFound	= -1
    } 	ksSpecificationStyleDifferenceTypeEnum;

//-------------------------------------------------------------------------------
// _2_36_ ksSlaveDocumentTypeEnum - Типы подчиненных режимов редактирования документов КОМПАС
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("18901E0A-EC8F-4FD4-9326-D6DBCB58B222") 
enum _ksSlaveDocumentTypeEnum
    {
        ksSDSketchMode	= 1000,
        ksSDSpecificationSlave	= 1001
    } 	ksSlaveDocumentTypeEnum;

//-------------------------------------------------------------------------------
// _4_64_ ksNurbsByPointsBuildingTypeEnum - Способ формирования точек сплайна
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("0C495A54-ADFE-4842-ABF3-AACB63B82445") 
enum _ksNurbsByPointsBuildingTypeEnum
    {
        ksNByPBTUndefined	= 0,
        ksNByPBLinear	= 1,
        ksNByPBChordLength	= 2,
        ksNByPBCentripetal	= 3
    } 	ksNurbsByPointsBuildingTypeEnum;

//-------------------------------------------------------------------------------
// _4_65_ ksNurbsByPointsPointConstraintsEnum - Вариант управления точкой сплайна
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("8DA9E5C9-F17D-44C8-86BE-14E30A8E2DCA") 
enum _ksNurbsByPointsPointConstraintsEnum
    {
        ksNByPBPCNone	= 0,
        ksNByPBPCTangent	= 1,
        ksNByPBPCSmoothG2	= 2
    } 	ksNurbsByPointsPointConstraintsEnum;

//-------------------------------------------------------------------------------
// _4_66_ ksNurbsByPointsAproximationTypeEnum - Способ вычисления шага аппроксимации для точек сплайна
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("F3112C93-D4A7-4BAD-BD87-E4C33B076FA7") 
enum _ksNurbsByPointsAproximationTypeEnum
    {
        ksNBPAproximationStepByCurvature	= 0,
        ksNBPAproximationStepByDeviation	= 1
    } 	ksNurbsByPointsAproximationTypeEnum;

  //------------------------------------------------------------------------------
  // _2_37_ ksKompasVariantEnum - Константы вариантов реализаций Компас.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("B14ED0B9-8C84-4987-8EB2-0ACFFCB13246") 
enum _ksKompasVariantEnum
    {
        ksKompasPro	= 0,
        ksKompasHome	= 0x1,
        ksKompasViewer	= 0x2,
        ksKompasSpds	= 0x4,
        ksKompasGraphic	= 0x100,
        ksKompasInvisible	= 0x200,
        ksKompasLatin	= 0x400,
        ksKompasConsumer	= 0x1000,
        ksKompasStudy	= 0x2000
    } 	ksKompasVariantEnum;

//-------------------------------------------------------------------------------
// _3_13_ ksCheckBoxVisualStyleEnum - Визуальный стиль чекбокса
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("98244AE9-D5D5-4435-9F6C-CA5B92A3CB42") 
enum _ksCheckBoxVisualStyleEnum
    {
        ksCheckBoxDefault	= 0,
        ksCheckBoxSwitcher	= 1
    } 	ksCheckBoxVisualStyleEnum;

//-------------------------------------------------------------------------------
// _3_14_ ksPhantomTypeEnum - Типы фантома
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("53DE8B9A-8806-4ECE-9619-C6A799AAC6FC") 
enum _ksPhantomTypeEnum
    {
        ksUnknownPhantom	= 0,
        ksMoveGroupPhantom	= 1,
        ksLineSegPhantom	= 2,
        ksRectanglePhantom	= 3,
        ksAngleLineSegPhantom	= 4,
        ksHalfRectanglePhantom	= 5,
        ksUserPhantom	= 6,
        ksCirclePhantom	= 7
    } 	ksPhantomTypeEnum;

//-------------------------------------------------------------------------------
// _3_15_ ksProcess2DTypeEnum - Типы процессов 2D
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("BFEFECA5-1EE1-4937-A12E-A39827DCEC6A") 
enum _ksProcess2DTypeEnum
    {
        ksProcess2DCursor	= 1,
        ksProcess2DPlacement	= 2
    } 	ksProcess2DTypeEnum;

  //-------------------------------------------------------------------------------
  // _1_19_ События процесса 2D.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("9EBD4C17-1A7C-4225-AFAF-9BC98C3BAEE0") 
enum _ksProcess2DNotifyEnum
    {
        ksProcess2DPlacementChanged	= 1,
        ksProcess2DExecuteCommand	= 2,
        ksProcess2DRun	= 3,
        ksProcess2DStop	= 4,
        ksProcess2DActivate	= 5,
        ksProcess2DDeactivate	= 6
    } 	ksProcess2DNotifyEnum;

  //-------------------------------------------------------------------------------
  // _1_20_ События процесса 3D.
  // ---
typedef /* [v1_enum][helpcontext][helpstring][uuid] */  DECLSPEC_UUID("A696A867-6D59-4E0A-8DDD-DFC30FF2A774") 
enum _ksProcess3DNotifyEnum
    {
        ksProcess3DPlacementChanged	= 1,
        ksProcess3DExecuteCommand	= 2,
        ksProcess3DRun	= 3,
        ksProcess3DStop	= 4,
        ksProcess3DActivate	= 5,
        ksProcess3DDeactivate	= 6,
        ksProcess3DFilterObjects	= 7,
        ksProcess3DCreateTakeObject	= 8
    } 	ksProcess3DNotifyEnum;

//-------------------------------------------------------------------------------
// _3_16_ ksProcess3DTypeEnum                 - Типы процессов 3D
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("3FEF9E54-8B04-49F0-BD06-F4B0F131D036") 
enum _ksProcess3DTypeEnum
    {
        ksProcess3DPlacementAndEntity	= 1,
        ksProcess3DSelectEntity	= 2
    } 	ksProcess3DTypeEnum;

//-------------------------------------------------------------------------------
// _3_17_ ksTwinSwitcherValueEnum                 - Значения переключателя
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("73B0D849-7E97-4653-A8A5-CB5C3AC51384") 
enum _ksTwinSwitcherValueEnum
    {
        ksTwinSwitcherPos1	= 1,
        ksTwinSwitcherPos2	= 2
    } 	ksTwinSwitcherValueEnum;

//-------------------------------------------------------------------------------
// _4_67_ ksToleranceArrowType - Тип стрелки ответвления у допуска формы
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("273FADE7-D399-456D-9AE6-4C34CF99704C") 
enum _ksToleranceArrowType
    {
        ksTANone	= 0,
        ksTAArrow	= 1,
        ksTATriangle	= 2
    } 	ksToleranceArrowType;

//-------------------------------------------------------------------------------
// _3_18_ ksProcessContextMenuType - Тип процессного меню
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("3FD254BD-E52B-482D-A146-8471512A8F58") 
enum _ksProcessContextMenuType
    {
        ksProcessDefaultContextMenu	= 0,
        ksProcessContextMenuHide	= -2,
        ksProcessContextPanel	= -3,
        ksProcessContextIconMenu	= -4
    } 	ksProcessContextMenuType;

//-------------------------------------------------------------------------------
// _3_19_ ksEnterButtonIconTypeEnum - Тип иконки для кнопки создать
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("4A5AF40B-E744-411E-8BE2-82BC79088FEB") 
enum _ksEnterButtonIconTypeEnum
    {
        ksEnterCheckIcon	= 0,
        ksEnterFloppyIcon	= 1,
        ksEnterNewInputIcon	= 2,
        ksEnterApplyIcon	= 3
    } 	ksEnterButtonIconTypeEnum;

//-------------------------------------------------------------------------------
// _4_68_ ksShowHideTmpObjTypeEnum - Тип отображения временного объекта в документе.
// ---
typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("F3F85250-64AE-46FC-8861-1C5625986D06") 
enum _ksShowHideTmpObjTypeEnum
    {
        ksTmpObjHide	= 0,
        ksTmpObjShow	= 1,
        ksTmpObjShowPhantom	= 2
    } 	ksShowHideTmpObjTypeEnum;

typedef /* [helpcontext][helpstring][v1_enum][uuid] */  DECLSPEC_UUID("BACBE7ED-0424-40F2-85AB-E4B05860495D") 
enum _ksDocumentsLibraryInsertionTypeEnum
    {
        ksInsertionUnknown	= 0,
        ksInsertionFragment	= 2,
        ksInsertionPart	= 4,
        ksInsertionAssembly	= 5,
        ksInsertionTextual	= 6,
        ksInsertionRaster	= 20,
        ksInsertionTable	= 21
    } 	ksDocumentsLibraryInsertionTypeEnum;


DEFINE_GUID(LIBID_Kompas6Constants,0x75C9F5D0,0xB5B8,0x4526,0x86,0x81,0x99,0x03,0xC5,0x67,0xD2,0xED);
#endif /* __Kompas6Constants_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


