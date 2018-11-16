////////////////////////////////////////////////////////////////////////////////
//
//  ldefin2d.h - типы интерфейсов для 2d структур
//
////////////////////////////////////////////////////////////////////////////////
#ifndef __LDEFIN2D_H
#define __LDEFIN2D_H

#ifndef _REFERENCE_DEFINED
typedef long reference;
#define _REFERENCE_DEFINED
#endif


//   result >-1  успешное завершение или не фатальные ошибки
//   result <0   фатальные ошибки - работа прекращается
#define MAXERROR  216

// типы документов
enum DocType {
  lt_DocUnknown       = 0,       //                  - нет активного документа 
  lt_DocSheetStandart = 1,       // тип документа 1  - чертеж стандартный
  lt_DocSheetUser,               //               2  - чертеж нестандартный
  lt_DocFragment,                //               3  - фрагмент
  lt_DocSpc,                     //               4  - спецификация
  lt_DocPart3D,                  //               5  - 3d-документ модель
  lt_DocAssemble3D,              //               6  - 3d-документ сборка
  lt_DocTxtStandart,             //               7  - текстовый документ стандартный
  lt_DocTxtUser,                 //               8  - текстовый документ нестандартный
  lt_DocSpcUser,                 //               9  - спецификация нестандартный формат
  lt_DocTechnologyAssemble3D,    //               10 - 3d-документ технологическая сборка
};

// система квалитета
enum LtQualSystem {
  lt_qsShaft = 1,    // 1 - вала
  lt_qsHole,         // 2 - отверстия
};

// квалитеты
enum LtQualDir {
  lt_qdPreferable = 1, // 1 - предпочтительные
  lt_qdBasic,          // 2 - основные
  lt_qdAdditional,     // 3 - дополнительные
};

// типы данных для LtVariant
enum LtVariantType {
  ltv_Char = 1, // 1 - символ
  ltv_UChar,    // 2 - байт
  ltv_Int,      // 3 - целое
  ltv_UInt,     // 4 - беззнаковое целое
  ltv_Long,     // 5 - длинное целое
  ltv_Float,    // 6 - вешественное
  ltv_Double,   // 7 - двойное вешественное
  ltv_Str,      // 8 - строка 255 символов char[255]
  ltv_NoUsed,   // 9 - пока не используется
  ltv_Short,    // 10 - короткое целое
  ltv_WStr,     // 11 - Строка 255 символов whar_t[255] 
};

// типы точек привязки текста
enum TextAlign {
  txta_Left = 0, // точка привязки слева
  txta_Center,   // точка привязки вцентре
  txta_Right     // точка привязки справа
};

// типы структур создаваемых методом GetParamStruct объекта KompasObject
// enum StructType2D Перенесен в ksConstants.tlb 
// enum LtViewType перенесен в ksConstants.tlb 


// перечисление возможных типов узла дерева библиотеки документов
enum LtNodeType {
  tn_root,  // корень дерева
  tn_dir,   // папка (директория)
  tn_file   // документ (файл)
};


#define TEXT_LENGTH     128
#define MAX_TEXT_LENGTH 255

#define ODBC_DB 0
#define TXT_DB 1

#define TXT_CHAR   1
#define TXT_USHORT 2
#define TXT_SSHORT 3
#define TXT_SLONG  4
#define TXT_ULONG  5
#define TXT_LONG   6
#define TXT_FLOAT  7
#define TXT_DOUBLE 8
#define TXT_INT    9
#define TXT_ALL    0
#define TXT_INDEX  _T("Index1000")
#define TXT_WCHAR 10

#define stACTIVE    0  //состояние для вида, слоя, документа
#define stREADONLY  1  //состояние для вида, слоя
#define stINVISIBLE 2  //состояние для вида, слоя
#define stCURRENT   3  //состояние для вида, слоя
#define stPASSIVE   1  //состояние для документа


// Определения для функции ksSystemPath
#define sptSYSTEM_FILES           0  // Выдать путь на каталог системных файлов
#define sptLIBS_FILES             1  // Выдать путь на каталог файлов приложений
#define sptAPPS_FILES             1  // Выдать путь на каталог файлов приложений
#define sptTEMP_FILES             2  // Выдать путь на каталог сохранения временных файлов
#define sptCONFIG_FILES           3  // Выдать путь на каталог сохранения конфигурации системы
#define sptINI_FILE               4  // Выдать полное имя INI-файла системы
#define sptBIN_FILE               5  // Выдать путь на каталог исполняемых файлов системы
#define sptPROJECT_FILES          6  // Выдать путь на каталог сохранения kompas.prj
#define sptDESKTOP_FILES          7  // Выдать путь на каталог сохранения kompas.dsk
#define sptTEMPLATES_FILES        8  // Выдать путь на каталог шаблонов Компас-документов
#define sptPROFILES_FILES         9  // Выдать путь на каталог сохранения профилей пользователя
#define sptWORK_FILES             10 // Выдать путь на каталог "Мои документы"
#define sptSHEETMETAL_FILES       11 // Выдать путь на каталог таблиц сгибов
#define sptPARTLIB_FILES          12 // Выдать путь на каталог PartLib
#define sptMULTILINE_FILES        13 // Выдать путь к каталогу шаблонов мультилинии
#define sptPRINTDEVICE_FILES      14 // Выдать путь к каталогу конфигураций плоттеров/принтеров
#define sptCURR_WORK_FILES        15 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptCURR_LIBS_FILES        16 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptCURR_APPS_FILES        16 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptCURR_SYSTEM_FILES      17 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptCURR_PROFILES_FILES    18 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptCURR_SHEETMETAL_FILES  19 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptMULTY_APPS_FILES       20 // Выдать список каталогов файлов приложений
#define sptDOC_LIBS_FILES         21 // Выдать путь на каталог файлов библиотек документов
#define sptMULTY_DOC_LIBS_FILES   22 // Выдать список каталогов файлов библиотек документов
#define sptCUR_DOC_LIBS_FILES     23 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save
#define sptUTILS_FILES            24 // Выдать путь на каталог утилит
#define sptMULTY_UTILS_FILES      25 // Выдать список каталогов утилит
#define sptCURR_UTILS_FILES       26 // запоминание последних директориев, с которых выполнилось открытие/сохранение файла в диалоге Open|Save


// Определения для результата функции SystemControlStart
// 27.01.2005 Данные определения устарели 
// см. перечисление ksSystemControlStartEnum в ksConstants.h или ksConstants.tlb
#define scsSTOPPED_FOR_MENU_COMMAND      ( 1) // Выполнена команда меню "Остановить работу библиотеки"
#define scsSTOPPED_FOR_SYSTEM_STOP       ( 0) // Идет закрытие задачи
#define scsSTOPPED_FOR_ITSELF            (-1) // Вызов функции SystemControlStop из-под библиотеки
#define scsSTOPPED_FOR_START_THIS_LIB    (-2) // Принудительный останов при запуске той же библиотеки
#define scsSTOPPED_FOR_START_ANOTHER_LIB (-3) // Принудительный останов при запуске другой библиотеки

//  Определения для функций GetObjParam и SetObjParam
//  '+'   отмечены объекты, для которых реализованы  GetObjParam и SetObjParam
//  '+-'  отмечены объекты, для которых реализован только GetObjParam
#define ALLPARAM                     -1   // все параметры объекта в СК объекта владельца
#define SHEET_ALLPARAM               -2   // тоже что и  ALLPARAM  но параметры объекта в СК листа
#define NURBS_CLAMPED_ALLPARAM       -5   // параметры нурбса, преобразовать узловой вектор в зажатый  
#define NURBS_CLAMPED_SHEET_ALLPARAM -6   // параметры нурбса в СК листа, преобразовать узловой вектор в зажатый
#define VIEW_ALLPARAM                -7   // тоже что и  ALLPARAM  но параметры объекта в СК вида

#define ALLPARAM_W                     -20  // все параметры объекта в СК объекта владельца для структур со строками wchar_t
#define SHEET_ALLPARAM_W               -21  // тоже что и  ALLPARAM  но параметры объекта в СК листа для структур со строками wchar_t
#define VIEW_ALLPARAM_W                -22  // тоже что и  ALLPARAM  но параметры объекта в СК вида для структур со строками wchar_t
#define ASSOCIATION_VIEW_PARAM_W       -23  // параметры ассоциативного вида для структуры со строками wchar_t

#ifdef _UNICODE
#define ALLPARAM_T  ALLPARAM_W
#else
#define ALLPARAM_T  ALLPARAM
#endif // !UNICODE

#ifdef _UNICODE
#define SHEET_ALLPARAM_T  SHEET_ALLPARAM_W
#else
#define SHEET_ALLPARAM_T  SHEET_ALLPARAM
#endif // !UNICODE

#ifdef _UNICODE
#define VIEW_ALLPARAM_T  VIEW_ALLPARAM_W
#else
#define VIEW_ALLPARAM_T  VIEW_ALLPARAM
#endif // !UNICODE

#ifdef _UNICODE
#define ASSOCIATION_VIEW_PARAM_T  ASSOCIATION_VIEW_PARAM_W
#else
#define ASSOCIATION_VIEW_PARAM_T  ASSOCIATION_VIEW_PARAM
#endif // !UNICODE

#define ANGLE_ARC_PARAM          0   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК объекта владельца
#define POINT_ARC_PARAM          1   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК объекта владельца
#define ANGLE_ARC_SHEET_PARAM    2   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК листа
#define POINT_ARC_SHEET_PARAM    3   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК листа
#define ANGLE_ARC_VIEW_PARAM     4   // параметры дуги по углам ( для дуги и эллиптической дуги ) в СК вида
#define POINT_ARC_VIEW_PARAM     5   // параметры дуги по точкам ( для дуги и эллиптической дуги ) в СК вида

#define VIEW_LAYER_STATE         1   // состояние слоя ,вида
#define DOCUMENT_STATE           1   // состояние документа
#define DOCUMENT_SIZE            0   // размер листа
#define DIM_TEXT_PARAM           0   // параметры текста для размеров
#define DIM_TEXT_PARAM_W         6   // параметры текста для размеров для структуры со строками wchar_t
#ifdef _UNICODE
#define DIM_TEXT_PARAM_T  DIM_TEXT_PARAM_W
#else
#define DIM_TEXT_PARAM_T  DIM_TEXT_PARAM
#endif // !UNICODE

#define DIM_SOURSE_PARAM         1   // параметры привязки размера
#define DIM_DRAW_PARAM           2   // параметры отрисовки размера
#define DIM_VALUE                3   // значение размера - double
#define DIM_PARTS                4   // составляющие части для размеров struct DimensionPartsParam
#define SHEET_DIM_PARTS          5   // составляющие части для размеров struct DimensionPartsParam в СК листа
#define TECHNICAL_DEMAND_PAR     -1  // параметры технических требований -
#define TT_FIRST_STR             1000// начало отсчета для получения или замены текста ТТ по строкам
#define CONIC_PARAM              2   // параметры для построения конического сечения ( для эллипса и эллтптической дуги )
#define SPC_TUNING_PARAM         0   // параметры настроек для стиля СП
#define HATCH_PARAM_EX           0   // параметры штриховки расширенные
#define ASSOCIATION_VIEW_PARAM   0   // параметры ассоциативного вида
#define DIM_SOURSE_VIEWPARAM     7   // параметры привязки размера в ситеме координат вида
#define DIM_DRAW_VIEWPARAM       8   // параметры отрисовки размера в ситеме координат вида
#define DIM_SOURSE_SHEETPARAM    9   // параметры привязки размера в ситеме координат листа
#define DIM_DRAW_SHEETPARAM      10  // параметры отрисовки размера в ситеме координат листа

#define ALL_OBJ                0         // все объекты,кроме вспомогательных, которые могут входить в вид                    -
#define LINESEG_OBJ            1         // отрезок                        +
#define CIRCLE_OBJ             2         // окружность                     +
#define ARC_OBJ                3         // дуга                           +
#define TEXT_OBJ               4         // текст                          +
#define POINT_OBJ              5         // точка                          +
#define HATCH_OBJ              7         // штриховка                      +
#define BEZIER_OBJ             8         // bezier сплайн                  +
#define LDIMENSION_OBJ         9         // линейный размер                +
#define ADIMENSION_OBJ         10        // угловой размер                 +
#define DDIMENSION_OBJ         13        // диаметральный размер           +
#define RDIMENSION_OBJ         14        // радиальный размер              +
#define RBREAKDIMENSION_OBJ    15        // радиальный размер с изломом    +
#define ROUGH_OBJ              16        // шероховатость                  +
#define BASE_OBJ               17        // база                           +
#define WPOINTER_OBJ           18        // стрелка вида                   +
#define CUT_OBJ                19        // линия разреза                  +
#define LEADER_OBJ             20        // простая линия выноски                      +
#define POSLEADER_OBJ          21        // линия выноски для обозначения позиции      +
#define BRANDLEADER_OBJ        22        // линия выноски для обозначения клеймения    +
#define MARKERLEADER_OBJ       23        // линия выноски для обозначения маркирования +
#define TOLERANCE_OBJ          24        // допуск формы                   +
#define TABLE_OBJ              25        // таблица                        -     //тексты
#define CONTOUR_OBJ            26        // контур                         +     //стиль
#define MACRO_OBJ              27        // нетипизированный макроэлемент  -
#define LINE_OBJ               28        // линия                          +
#define LAYER_OBJ              29        // слой                           +
#define FRAGMENT_OBJ           30        // вставной фрагмент              +
#define POLYLINE_OBJ           31        // полилиния                      +
#define ELLIPSE_OBJ            32        // эллипс                         +
#define NURBS_OBJ              33        // NURBS-кривая по полюсам        +
#define ELLIPSE_ARC_OBJ        34        // дуга эллипса                   +
#define RECTANGLE_OBJ          35        // прямоугольник                  +
#define REGULARPOLYGON_OBJ     36        // многоугольник                  +
#define EQUID_OBJ              37        // эквидистанта                   +
#define LBREAKDIMENSION_OBJ    38        // линейный размер с обрывом      +
#define ABREAKDIMENSION_OBJ    39        // угловой размер с обрывом       +
#define ORDINATEDIMENSION_OBJ  40        // размер высоты                  +
#define COLORFILL_OBJ          41        // фоновая заливка цветом         +
#define CENTREMARKER_OBJ       42        // обозначение центра             +
#define ARCDIMENSION_OBJ       43        // размер длины дуги              -
#define SPC_OBJ                44        // объект спецификации            +
#define RASTER_OBJ             45        // растровый объект               +
#define CHANGE_LEADER_OBJ      46        // Обозначение изменения          +
#define REMOTE_ELEMENT_OBJ     47        // выносной элемент               +
#define AXISLINE_OBJ           48        // Осевая линия                   +
#define OLEOBJECT_OBJ          49        // вставка ole объекта            -
#define KNOTNUMBER_OBJ         50        // объект номер узла              -
#define BRACE_OBJ              51        // объект фигурная скобка         -
#define POSNUM_OBJ             52        // марка/позиционное обозначение с линией-выноской - 
#define MARKONLDR_OBJ          53        // марка/позиционное обозначение на линии          -
#define MARKWOLDR_OBJ          54        // марка/позиционное обозначение без линии-выноски -
#define WAVELINE_OBJ           55        // волнистая линия                -
#define DIRAXIS_OBJ            56        // прямая ось                     -
#define BROKENLINE_OBJ         57        // линия обрыва с изломами        -
#define CIRCLEAXIS_OBJ         58        // круговая ось                   -
#define ARCAXIS_OBJ            59        // дуговая ось                    -
#define CUTUNITMARKING         60        // Обозначение узла в сечении     -
#define UNITMARKING            61        // Обозначение узла               -
#define MULTITEXTLEADER        62        // Выносная надпись к многослойным конструкциям.      -
#define EXTERNALVIEW_OBJ       63        // Вставка внешнего вида                              -
#define ANNLINESEG_OBJ         64        // Аннотационный отрезок                 +- Для GetObjParam используется структура LineSegParam
#define ANNCIRCLE_OBJ          65        // Аннотационная окружность              +- Для GetObjParam используется структура CircleParam
#define ANNELLIPSE_OBJ         66        // Аннотационный эллипс                  +- Для GetObjParam используется структура EllipseParam
#define ANNARC_OBJ             67        // Аннотационная дуга                    +- Для GetObjParam используется структура ArcParam
#define ANNELLIPSE_ARC_OBJ     68        // Аннотационная дуга эллипса            +- Для GetObjParam используется структура EllipseArcParam
#define ANNPOLYLINE_OBJ        69        // Аннотационная полилиния               +- Для GetObjParam используется структура PolylineParam
#define ANNPOINT_OBJ           70        // Аннотационная точка                   +- Для GetObjParam используется структура PointParam
#define ANNTEXT_OBJ            71        // Текст с аннотационной точкой привязки +- Для GetObjParam используется структура TextParam
#define MULTILINE_OBJ          72        // Мультилиния                    - 
#define BUILDINGCUTLINE_OBJ    73        // Линия разреза/сечения для СПДС + используется структура CutLineParam
#define ATTACHED_LEADER_OBJ    74        // Присоединенная линия выноски ( не имеет текстов )  +
#define CONDITIONCROSSING_OBJ  75        // Условное пересечение           - 
#define REPORTTABLE_OBJ        76        // ассоциативная таблица отчета
#define EMBODIMENTSTABLE_OBJ   77        // таблица исполнений
#define SPECIALCURVE_OBJ       78        // Кривая общего вида
#define ARRAYPARAMTABLE_OBJ    79        // Таблица параметров массива
#define NURBS_BY_POINTS_OBJ    80        // NURBS-кривая по точкам                        +- Для GetObjParam используется структура NurbsParam
#define CONIC_CURVE_OBJ        81        // Коническая кривая                             +- Для GetObjParam используется структура NurbsParam
#define CONIC_CURVE_4POINT_OBJ 82        // Коническая кривая  по 4 точками и направлению +- Для GetObjParam используется структура NurbsParam
#define CONIC_CURVE_5POINT_OBJ 83        // Коническая кривая  по 5 точками               +- Для GetObjParam используется структура NurbsParam

#define MAX_VIEWTIP_SEARCH     82        // верхняя граница типов поиска для объектов вида  -

#define SPECIFICATION_OBJ      121       // спецификация на листе
#define SPECROUGH_OBJ          122       // неуказанная шероховатость      +
#define VIEW_OBJ               123       // вид                            +
#define DOCUMENT_OBJ           124       // документ  графический          +   (лист или фрагмент)
#define TECHNICALDEMAND_OBJ    125       // технические требования         +
#define STAMP_OBJ              126       // штамп                          -  //тексты
#define SELECT_GROUP_OBJ       127       // группа селектирования          -
#define NAME_GROUP_OBJ         128       // именная группа                 -
#define WORK_GROUP_OBJ         129       // рабочая группа                 -
#define SPC_DOCUMENT_OBJ       130       // документ  спецификация         +
#define D3_DOCUMENT_OBJ        131       // 3d документ  модели или сборки +
#define CHANGE_LIST_OBJ        132       // таблица изменений              -
#define TXT_DOCUMENT_OBJ       133       // текстовый документ             +
#define ALL_DOCUMENTS          134       // документы всех типов

#define MAX_TIP_SEARCH         134       // верхняя граница типов поиска   -
#define ALL_OBJ_SHOW_ORDER   -1000       // все объекты которые могут входить в вид в порядке отрисовки

// Поле тип линии имеет значения( системные стили ) см ksCurveStyleEnum:
//  1  - основная,
//  2  - тонкая,
//  3  - осевая,
//  4  - штриховая,
//  5  - волнистая линия
//  6  - вспомогательная,
//  7  - утолщенная,
//  8  - штрих-пунктир с 2 точками,
//  9  - штриховая толстая
//  10 - осевая толстая
//  11 - тонкая, включаемая в штриховку
//  12 - ISO штриховая линия
//  13 - ISO штриховая линия (дл. пробел)
//  14 - ISO штрихпунктирная линия (дл. штрих)
//  15 - ISO штрихпунктирная линия (дл. штрих 2 пунктира)
//  16 - ISO штрихпунктирная линия (дл. штрих 3 пунктира)
//  17 - ISO пунктирная линия
//  18 - ISO штрихпунктирная линия (дл. и кор. штрихи)
//  19 - ISO штрихпунктирная линия (дл. и 2 кор. штриха) 
//  20 - ISO штрихпунктирная линия
//  21 - ISO штрихпунктирная линия (2 штриха)
//  22 - ISO штрихпунктирная линия (2 пунктира)
//  23 - ISO штрихпунктирная линия (3пунктира)
//  24 - ISO штрихпунктирная линия (2 штриха 2 пунктира)
//  25 - ISO штрихпунктирная линия (2 штриха 3 пунктира)


// Поле тип точки для точки( системные стили ) :
//0 - точка
//1 - крестик
//2 - х-точка
//3  -  квадрат
//4  -  треугольник
//5  -  окружность
//6  -  звезда
//7  -  перечеркнутый квадрат
//8  -  утолщенный плюс



// Поле тип штриховки для штриховки( системные стили ) см ksHatchStyleEnum:
// 0  - металл
// 1  - неметалл 
// 2  - дерево
// 3  - камень естественный
// 4  - керамика
// 5  - бетон
// 6  - стекло
// 7  - жидкость
// 8  - естественный грунт
// 9  - насыпной грунт
// 10 - камень искусственный
// 11 - железобетон
// 12 - напряженный железобетон
// 13 - дерево в продольном сечении
// 14 - песок

// Определения флагов для текста см ksTextItemEnum
#define INVARIABLE           0    //не менять флаги текста

#define NUMERATOR          0x1    //числитель
#define DENOMINATOR        0x2    //знаменатель
#define END_FRACTION       0x3    //конец дроби
#define UPPER_DEVIAT       0x4    //верхнее отклонение
#define LOWER_DEVIAT       0x5    //нижнее отклонение
#define END_DEVIAT         0x6    //конец  отклонений
#define S_BASE             0x7    //основание выражения типа суммы
#define S_UPPER_INDEX      0x8    //верхний индекс выражения типа суммы
#define S_LOWER_INDEX      0x9    //нижний индекс выражения типа суммы
#define S_END              0x10   //конец выражения типа суммы
#define SPECIAL_SYMBOL     0x11   //спецзнак
#define SPECIAL_SYMBOL_END 0x12   //для спецзнаков с текстом
#define RETURN_BEGIN       0x13   //начало для ввода следующих строк в спецзнаке с текстом, дробях, отклонениях
#define RETURN_DOWN        0x14   //для ввода следующих строк в спецзнаке с текстом, дробях, отклонениях
#define RETURN_RIGHT       0x15   //для ввода строк справа в спецзнаке с текстом, дробях, отклонениях
#define TAB                0x16   //табуляция по текущему стилю
#define FONT_SYMBOL        0x17   //символ фонта
#define MARK_SEPARATOR     0x18   //разделитель в обозначении 
#define HYPER_TEXT_END     0x19   //завершение гипертекста
#define FONT_SYMBOL_W      0x2017 //символ фонта Unicode
#define HYPER_TEXT         0x2000 //ссылка на текст или положение объекта


#define ITALIC_ON      0x40   //включить наклон
#define ITALIC_OFF     0x80   //выключть наклон
#define BOLD_ON        0x100  //включить толщину
#define BOLD_OFF       0x200  //выключть толщину
#define UNDERLINE_ON   0x400  //включить подчеркивание
#define UNDERLINE_OFF  0x800  //выключть подчеркивание
#define NEW_LINE       0x1000 //новая строка в параграфе

#define FONT_NAME       1       //имя фонта
#define NARROWING       2       //коэффициент сужения фонта
#define HEIGHT          3       //высота фонта
#define COLOR           4       //цвет текста
#define SPECIAL         5       //спецзнак
#define FRACTION_TYPE   6       //высота дроби по отношению к тексту 1-полная высота 2-в 1.5 раза меньше 3-в 2 раза меньше
#define SUM_TYPE        7       //высота выражения типа суммы по отношению к тексту 1-полная высота 2-в 1.5 раза больше


//Определения для динамических массивов
#define CHAR_STR_ARR          1  // динамический массив указателей на строки символов
#define POINT_ARR             2  // динамический массив указателей на математические точки -структура MathPointParam
#define CURVE_PATTERN_ARR     2  // динамический массив указателей на участки штриховой линии -структура CurvePattern
#define TEXT_LINE_ARR         3  // динамический массив строк текста - структура TextLineParam
#define TEXT_ITEM_ARR         4  // динамический массив компонент строк текста структура TextItemParam
#define ATTR_COLUMN_ARR       5  // динамический массив колонок атрибутов- структура  ColumnInfo
#define USER_ARR              6  // динамический пользовательский массив
#define POLYLINE_ARR          7  // динамический массив полилиний-(указателей массивов POINT_ARR)
#define RECT_ARR              8  // динамический массив габаритных прямоугольников-(структура RectParam)
#define LIBRARY_STYLE_ARR     9  // динамический массив структур параметров для стиля в библиотеке стилей( LibraryStyleParam )
#define VARIABLE_ARR          10 // динамический массив структур параметров параметрических переменных( VariableParam )
#define CURVE_PATTERN_ARR_EX  11 // динамический массив указателей на участки штриховой линии -структура CurvePatternEx
#define LIBRARY_ATTR_TYPE_ARR 12 // динамический массив структур параметров для типа атрибута в библиотеке типов атрибутов( LibraryAttrTypeParam )
#define NURBS_POINT_ARR       13 // динамический массив структур NurbsPointParam
#define DOUBLE_ARR            14 // динамический массив double
#define CONSTRAINT_ARR        15 // динамический массив параметрических ограничений - структура ConstraintParam
#define CORNER_ARR            16 // динамический массив структур параметров углов CornerParam для прямоугольников и многоугольников
#define DOC_SPCOBJ_ARR        17 // динамический массив структур параметров прикрепленных документов к объекту спецификации DocAttachedSpcParam
#define SPCSUBSECTION_ARR     18 // динамический массив структур параметров подраздела спецификации SpcSubSectionParam
#define SPCTUNINGSEC_ARR      19 // динамический массив структур параметров настройки раздела спецификации SpcTuningSectionParam
#define SPCSTYLECOLUMN_ARR    20 // динамический массив структур параметров стиля колонки таблицы спецификации SpcStyleColumnParam
#define SPCSTYLESEC_ARR       21 // динамический массив структур параметров стиля разделa спецификации SpcStyleSectionParam
#define QUALITYITEM_ARR       22 // динамический массив структур QualityItemParam - запись об одном интервале для какого-то квалитета
#define LTVARIANT_ARR         23 // динамический массив структур LtVariant
#define TOLERANCEBRANCH_ARR   24 // динамический массив структур ToleranceBranch
#define HATCHLINE_ARR         25 // динамический массив структур HatchLineParam
#define TREENODEPARAM_ARR     26 // динамический массив структур узла дерева TreeNodeParam
#define CHAR_STR_ARR_W        27 // динамический массив указателей на строки символов типа wchar

#ifdef _UNICODE
#define CHAR_STR_ARR_T  CHAR_STR_ARR_W
#else
#define CHAR_STR_ARR_T  CHAR_STR_ARR
#endif // !UNICODE


// Поле style для текста( системные стили ) см ksTextStyleEnum :  
// 0 -умолчательный стиль для данного типа объекта
// 1  текст на чертеже
// 2  текст для технических требований
// 3  текст размерных надписей
// 4  текст шероховатости
// 5  текст для линии выноски  ( позиционной )
// 6  текст для линии выноски  ( над\под полкой )
// 7  текст для линии выноски  ( сбоку )
// 8  текст для допуска формы
// 9  текст для таблицы ( заголовок )
// 10 текст для таблицы ( ячейка )
// 11 текст для линии разреза
// 12 текст для стрелки вида
// 13 текст для для неуказанной шероховатости
// 14 текст для обозначения изменения
// 15 текст для фигурной скобки
// 16 текст для номера узла
// 17 текст для выносной надписи
// 18 текст для обозначения узла
// 19 текст для марки координационной оси
// 20 текст для МПО(марка/позиционное обозначение с линией-выноской)
// 21 текст для МПО(марка/позиционное обозначение) на линии
// 22 текст для МПО(марка/позиционное обозначение) без линии выноски
// 23 текст для заголовков спецификации
// 24 текст для линия разреза для СПДС
// 25 Текст для таблицы отчета ( заголовок ).
// 26 Текст для таблицы отчета ( ячейка ).

// Структуры для работы с табличными атрибутами  
#define CHAR_ATTR_TYPE     1
#define UCHAR_ATTR_TYPE    2
#define INT_ATTR_TYPE      3
#define UINT_ATTR_TYPE     4
#define LINT_ATTR_TYPE     5
#define FLOAT_ATTR_TYPE    6
#define DOUBLE_ATTR_TYPE   7
#define STRING_ATTR_TYPE   8   //строка фиксированной длины MAX_TEXT_LENGTH
#define RECORD_ATTR_TYPE   9

// признаки формирования размерной надписи
#define _AUTONOMINAL       0x1   // >0 автоматическое определение номинального значения размера
#define _RECTTEXT          0x2   // >0 текст в рамочке
#define _PREFIX            0x4   // >0 есть текст до номинала
#define _NOMINALOFF        0x8   // >0 нет  номинала
#define _TOLERANCE         0x10  // >0 квалитет проставлять
#define _DEVIATION         0x20  // >0 отклонения проставлять
#define _UNIT              0x40  // >0 единица измерения
#define _SUFFIX            0x80  // >0 есть текст после номинала
#define _DEVIATION_INFORM  0x100 // >0 при включенном _DEVIATION, отклонения есть в массиве текстов( даже если не ручное проставление отклонений).
                                 //    Появляется после  функции GetObjParam, чтобы пользователь мог получить отклонения.
#define _UNDER_LINE        0x200 // >0 размер с подчеркиванием
#define _BRACKETS          0x400 // >0 размер в скобках
#define _SQUARE_BRACKETS   0x800 // >0 размер в квадратных скобках, используется вместе с _BRACKETS
                                 //    _BRACKETS                    - размер в круглых скобках
                                 //    _BRACKETS | _SQUARE_BRACKETS - размер в квадратных скобках

#define   INDICATIN_TEXT_LINE_ARR        0xFFFF  //для шероховаиости, позиционной линии выноски, маркировки и клеймения
                                                 //признак, что для текста используется динамический массив TEXT_LINE_ARR

// типы стилей
#define CURVE_STYLE      1  // стиль криивых
#define HATCH_STYLE      2  // стиль штриховок
#define TEXT_STYLE       3  // стиль текста
#define STAMP_STYLE      4  // стиль штампа
#define CURVE_STYLE_EX   5  // стиль криивых расширенный

// curveType | LIKE_BASIC_LINE - параметры пера как у  основной линии
#define  LIKE_BASIC_LINE 0x10 // параметры пера как у  основной линии
#define  LIKE_THIN_LINE  0x20 // параметры пера как у  тонкой линии
#define  LIKE_HEAVY_LINE 0x30 // параметры пера как у  утолщенной линии

// Определения для функций Get/SetDocOptions и ksGet/SetSysOptions
#define DIMENTION_OPTIONS            1 // Настройки размера
#define SNAP_OPTIONS                 1 // Настройки привязок
#define ARROWFILLING_OPTIONS         2 // Зачернять стрелки ?
#define SHEET_OPTIONS                3 // Параметры листа для новых документов
#define SHEET_OPTIONS_EX             4 // Параметры листа документа
#define LENGTHUNITS_OPTIONS          5 // Настройки единиц измерений
#define SNAP_OPTIONS_EX              6 // Настройки привязок документа
#define VIEWCOLOR_OPTIONS            7 // Настройки цвета фона рабочего поля 2d - документов
#define TEXTEDIT_VIEWCOLOR_OPTIONS   8 // Настройки цвета фона редактирования текста
#define MODEL_VIEWCOLOR_OPTIONS      9 // Настройки цвета фона для моделей
#define OVERLAP_OBJECT_OPTIONS      10 // Настройки перекрывающихся объектов
#define DIMENTION_OPTIONS_EX        11 // Настройки размера

//типы колонок для спецификации
#define   SPC_CLM_FORMAT   1   // формат
#define   SPC_CLM_ZONE     2   // зона
#define   SPC_CLM_POS      3   // позиция
#define   SPC_CLM_MARK     4   // обозначение
#define   SPC_CLM_NAME     5   // наименование
#define   SPC_CLM_COUNT    6   // количество
#define   SPC_CLM_NOTE     7   // примечание
#define   SPC_CLM_MASSA    8   // масса
#define   SPC_CLM_MATERIAL 9   // материал
#define   SPC_CLM_USER     10  // пользовательская
#define   SPC_CLM_KOD      11  // код
#define   SPC_CLM_FACTORY  12  // завод изготовитель

//типы значений для колонки спецификации
// констаны не используются нужно использовать LtVariantType
//#define   SPC_INT      1   // целый
//#define   SPC_DOUBLE   2   // вещественный
//#define   SPC_STRING   3   // строка
//#define   SPC_RECORD   4   // запись

//типы блиотек стилей
#define CURVE_STYLE_LIBRARY               1 // библиотека стилей кривых (*.lcs)
#define HATCH_STYLE_LIBRARY               2 // библиотека стилей штриховок (*.lhs)
#define TEXT_STYLE_LIBRARY                3 // библиотека стилей текстов   (*.lts)
#define STAMP_LAYOUT_STYLE_LIBRARY        4 // библиотека стилей описаний штампов (*.lyt)
#define GRAPHIC_LAYOUT_STYLE_LIBRARY      5 // библиотека стилей оформлений графических документов (*.lyt)
#define TEXT_LAYOUT_STYLE_LIBRARY         6 // библиотека стилей оформлений текстовых документов (*.lyt)
#define SPC_LAYOUT_STYLE_LIBRARY          7 // библиотека стилей оформлений спецификаций (*.lyt)

//размерности и типы детали для рассчета массо-ценровочных характеристик
#define  ST_MIX_MM      0x1  // миллиметры
#define  ST_MIX_SM      0    // сантиметры
#define  ST_MIX_DM      0x2  // дециметры
#define  ST_MIX_M       0x3  // метры
#define  ST_MIX_GR      0    // граммы
#define  ST_MIX_KG      0x10 // килограммы
#define  ST_MIX_EXT     0    // выдавливание
#define  ST_MIX_RV      0x20 // вращение

// тип локальной привязки
#define  SN_NEAREST_POINT    1    // Ближайшая точка
#define  SN_NEAREST_MIDDLE   2    // Середина
#define  SN_CENTRE           3    // Центр
#define  SN_INTERSECT        4    // Пересечение
#define  SN_GRID             5    // По сетке
#define  SN_XY_ALIGN         6    // Выравнивание
#define  SN_ANGLE            7    // Угловая привязка
#define  SN_POINT_CURVE      8    // Точка на кривой

// типы общих настроек для привязок
#define  SN_DYNAMICALLY               0x1  // Привязки отслеживать динамически
#define  SN_ASSISTANT                 0x2  // Писать текст
#define  SN_BACKGROUND_LAYER          0x4  // Учитывать фоновые слои и виды
#define  SN_SUSPENDED                 0x8  // Подавить привязки
#define  SN_VISIBLE_GRID_POINTS_ONLY  0x10 // Привязка только к видимым точкам сетки


// Типы параметрических ограничений см ksConstraintTypeEnum
#define CONSTRAINT_FIXED_POINT           1  // фиксировать точку
#define CONSTRAINT_POINT_ON_CURVE        2  // точка на кривой
#define CONSTRAINT_HORIZONTAL            3  // горизонталь
#define CONSTRAINT_VERTICAL              4  // вертикаль
#define CONSTRAINT_PARALLEL              5  // параллельность двух прямых или отрезков
#define CONSTRAINT_PERPENDICULAR         6  // перпендикулярность двух прямых или отрезков
#define CONSTRAINT_EQUAL_LENGTH          7  // равенство длин двух отрезков
#define CONSTRAINT_EQUAL_RADIUS          8  // равенство радиусов двух дуг/окружностей
#define CONSTRAINT_HOR_ALIGN_POINTS      9  // выравнивать две точки по горизонтали
#define CONSTRAINT_VER_ALIGN_POINTS      10 // выравнивать две точки по вертикали
#define CONSTRAINT_MERGE_POINTS          11 // совпадение двух точек
#define CONSTRAINT_TANGENT_TWO_CURVES    15 // касание двух кривых
#define CONSTRAINT_SYMMETRY_TWO_POINTS   16 // симетрия двух точек
#define CONSTRAINT_COLLINEAR             17 // колинеарность отрезков
#define CONSTRAINT_FIXED_ANGLE           18 // фиксированный угол
#define CONSTRAINT_FIXED_LENGHT          19 // фиксированная длина
#define CONSTRAINT_POINT_ON_CURVE_MIDDLE 20 // точка на серидине кривой
#define CONSTRAINT_BISECTOR              21 // биссектриса


// типы объектов спецификации см ksSpecificationObjectTypeEnum
#define  SPC_BASE_OBJECT     1    // базовый объект (редактируется пользователем)
#define  SPC_COMMENT         2    // комментарий    (редактируется пользователем)
#define  SPC_SECTION_NAME    3    // наименование раздела ( создается по стилю СП автоматически )
#define  SPC_BLOCK_NAME      4    // наименование блока исполнений ( создается по стилю СП автоматически )
#define  SPC_RESERVE_STR     5    // резервная строка ( создается по стилю СП автоматически )
#define  SPC_EMPTY_STR       6    // пустая строка ( создается по стилю СП автоматически )

// типы сортировки см ksSortTypeEnum
#define SPC_SORT_OFF        0   // нет сортировки
#define SPC_SORT_COMPOS     1   // составная сортировка
#define SPC_SORT_ALPHABET   2   // сортировка по алфавиту (1.06.01 - больше не используется)
#define SPC_SORT_UP         3   // сортировка по возрастанию колонок
#define SPC_SORT_DOCUMENT   4   // сортировка раздела документация
#define SPC_SORT_DOWN       5   // сортировка по убыванию колонок
#define SPC_SORT_COMPOSDOWN 6   // составная сортировка по убыванию 

////////////////////////////////////////////////////////////////////////////////
//
//  типы специальных символов ( аннотационный объект )
//
////////////////////////////////////////////////////////////////////////////////
#define  ARROW_INSIDE_SYMBOL       1  // стрелка (ласточкин хвост) изнутри
#define  ARROW_OUT_SIDE_SYMBOL     2  // стрелка (ласточкин хвост) снаружи
#define  TICK_TAIL_SYMBOL          3  // засечка с продолжением кривой (с хвостиком)
#define  UP_HALF_ARROW_SYMBOL      4  // верхняя половина стрелки изнутри
#define  DOWN_HALF_ARROW_SYMBOL    5  // нижняя половина стрелки изнутри
#define  BIG_ARROW_INSIDE_SYMBOL   6  // большая стрелка изнутри (7мм)
#define  ARROW_ORDINATE_DIM_SYMBOL 7  // стрелка для размера высоты(штрихи длиной 4 мм под углом 45 гр)
#define  TRIANGLE_SYMBOL           8  // треугольник по напр-нию кривой
#define  CIRCLE_RAD2_SYMBOL        9  // окружность радиусом 2 мм тонкой линией - для шер-сти и линии-выноски
#define  CENTRE_MARKER_SYMBOL      10 // обозначение фиктивного центра в виде большого креста
#define  GLUE_SIGN_SYMBOL          11 // знак склеивания
#define  SOLDER_SIGN_SYMBOL        12 // знак пайки
#define  SEWING_SIGN_SYMBOL        13 // знак сшивания
#define  CRAMP_SIGN_SYMBOL         14 // знак соединения внахлестку металл.скобами
#define  CORNER_CRAMP_SIGN_SYMBOL  15 // знак углового соединения металл.скобами
#define  MONTAGE_JOINT_SYMBOL      16 // знак монтажного шва
#define  TICK_SYMBOL               17 // засечка без продолжения кривой (без хвостика)
#define  TRIANGLE_CURR_CS          18 // треугольник по текущей СК - для базы
#define  ARROW_CLOSED_INSIDE       19 // закрытая стрелка изнутри
#define  ARROW_CLOSED_OUTSIDE      20 // закрытая стрелка снаружи
#define  ARROW_OPEN_INSIDE         21 // открытая стрелка изнутри
#define  ARROW_OPEN_OUTSIDE        22 // открытая стрелка снаружи
#define  ARROW_RIGHTANGLE_INSIDE   23 // стрелка 90 град изнутри
#define  ARROW_RIGHTANGLE_OUTSIDE  24 // стрелка 90 град снаружи
#define  SYMBOL_DOT                25 // точка (диаметр равен длины стрелки размера)
#define  SYMBOL_SMALLDOT           26 // точка маленькая (диаметр равен 0.6 длины стрелки размера)
#define  AUXILIARY_POINT           27 // вспомогательная точка
#define  LEFT_TICK_SYMBOL          28 // засечка с наклоном влево

//------------------------------------------------------------------------------
// Битовые флаги для функции ksSetMacroParam;
// указание какой тип редактирования поддерживает макро
// ---
#define MP_DBL_CLICK_OFF   0x01 //>0 редактирование по двойному нажанию выключено
#define MP_HOTPOINTS       0x02 //>0 интерфейс hot точек включен
#define MP_EXTERN_EDIT     0x04 //>0 интерфейс внешнего управления
#define MP_PROPERTY_OBJECT 0x08 //>0 интерфейс внешних свойств объекта

//-----------------------------------------------------------------------------
//определения для конвертации в растровый формат
// ---
#define FORMAT_BMP   0
#define FORMAT_GIF   1
#define FORMAT_JPG   2
#define FORMAT_PNG   3
#define FORMAT_TIF   4
#define FORMAT_TGA   5
#define FORMAT_PCX   6
#define FORMAT_WMF   16
#define FORMAT_EMF   17

//-----------------------------------------------------------------------------
//определения для настройки цвета растрового формата
// ---
#define BLACKWHITE   0   //цвет черный
#define COLORVIEW    1   //цвет установленный для вида
#define COLORLAYER   2   //цвет установленный для слоя
#define COLOROBJECT  3   //цвет установленный для объекта

//-----------------------------------------------------------------------------
// определения бит на пиксел для конвертации в растровый формат
// ---
#define BPP_COLOR_01 1  //"Черный" 
#define BPP_COLOR_02 2  //"4 цвета" 
#define BPP_COLOR_04 4  //"16 цветов" 
#define BPP_COLOR_08 8  //"256 цветов"
#define BPP_COLOR_16 16 //"16 разрядов"
#define BPP_COLOR_24 24 //"24 разряда"
#define BPP_COLOR_32 32 //"32 разряда"

//------------------------------------------------------------------------------
// Типы стандартных видов
// ---
#define VIEW_FRONT       0x1 //  Спереди
#define VIEW_REAR        0x2 //  Сзади
#define VIEW_UP          0x4 //  Сверху
#define VIEW_DOWN        0x8 //  Снизу
#define VIEW_LEFT        0x10 //  Слева
#define VIEW_RIGHT       0x20 //  Справа
#define VIEW_ISO         0x40 //  Изометрия

//------------------------------------------------------------------------------
// типы значка объекта "Выносной элемент"
// ---
enum LtRemoteElmSignType {
  re_Circle,         // 0 - окружность
  re_Rectangle,      // 1 - прямоугольник
  re_Ballon,         // 2 - скругленный прямоугольник
};


//------------------------------------------------------------------------------
// Тип изменения порядка объектов
// ---
enum ChangeOrderType {
  co_Top = 1,                   // Выше всех  
  co_Bottom,                    // Ниже всех 
  co_BeforeObject,              // Перед объектом 
  co_AfterObject,               // За объектом
  co_UpLevel,                   // На уровень вперед
  co_DownLevel,                 // На уровень назад
  co_BeforeObjectByGroupOrder,  // Перед объектом (расположить объекты группы в порядке их добавления в группу)
  co_AfterObjectByGroupOrder,   // За объектом (расположить объекты группы в порядке их добавления в группу)
};

//------------------------------------------------------------------------------
// Стандартые курсоры Компас
// ---
#define OCR_SELECT     0xFFFE //  Курсор в виде SELECT 
#define OCR_SNAP       0xFFFD //  Курсор в виде SNAP 
#define OCR_CATCH      0xFFFC //  Курсор в виде CATCH 
#define OCR_DEFAULT    0      //  Курсор в виде креста

#define OCR_DEDAULT    OCR_DEFAULT  //  Курсор в виде креста(Для компиляции старых библиотек)

//-----------------------------------------------------------------------------
// Неопределенный цвет для TextItemFont.color
// В стиле может быть выставлен по умолчанию цвет отличный он 0
// В этом случае если TextItemFont.color будет значение 0 то создастся
// модификатор на цвет и он не будет отображаться как цвет по умолчанию
// для того чтобы модификатор цвета не создался нужно
// или прислать цвет из настроек или константу FREE_COLOR       
// ---
#define FREE_COLOR (0xff000000l)      //  Неопределенный цвет


#endif
