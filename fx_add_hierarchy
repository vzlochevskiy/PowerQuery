let
    unzip_path = "https://raw.githubusercontent.com/vzlochevskiy/buchlotnik_functions/main/UnZip.pq",
    // full_path = "\\sigma.icover.ru\share\Аналитика\Данные\Справочники\Товары\Бренд-Категория-Артикул.xlsx",
    fx_add_levels = (full_path as text) => [
        // получаем код функции UnZip (@buchlotnik)
        // функция считывает содержимое zip-архива и возвращает таблицу с полями FileName - имя файла (с путём, если он в подпапке) и Content - бинарное содержимое конкретного файла
        fx_unzip = Expression.Evaluate( Text.FromBinary( Web.Contents( unzip_path ) ), #shared ),
        // получили бинарное содержимое файла и поместили в буфер
        workbook = Binary.Buffer( File.Contents( full_path ) ),
        // получили данные с листа традиционным способом
        sheet = Excel.Workbook( workbook ){0}[Data],
        // применили функцию, получающую разметку файла
        unzip = fx_unzip( workbook ),
        // для каждой строки получаем таблицу Attributes
        sheetData = Xml.Document(unzip{[FileName = "xl/worksheets/sheet1.xml"]}[Content]){0}[Value]{[Name = "sheetData"]}[Value][Attributes],
        // получаем для каждой строки значение её outlineLevel
        lvl = List.Transform( sheetData, (x) => x{[Name="outlineLevel"]}?[Value]?  ),
        // получаем для каждой строки её номер 
        row_num = List.Transform( sheetData, (x) => x{[Name="r"]}?[Value]?), 
        // соберем из номеров и уровней словарь на записях
        dict = Record.FromList( lvl, row_num ),
        // добавляем столбец индекса, но не с 1, а с первой строки в dimension
        add = Table.AddIndexColumn( sheet, "Уровень", Number.From( row_num{0} ), 1 ),
        // подцепляем данные из словаря через Record.FieldOrDefault, поскольку не обо всех строках у нас имеется информация
        to = Table.TransformColumns( add, { "Уровень", (x)=> Record.FieldOrDefault( dict, Text.From(x) ), type text } ) ] [to],
    // добавляем "дружелюбные" названия и описания к функции
    docs = type function (
        full_path as (type text meta [Documentation.FieldCaption = "Полный путь к файлу"]))
        as table meta [ Documentation.Name = "fx_add_levels",
                        Documentation.LongDescription = "Функция добавляет столбец с уровнем иерархии строки. <br> <b>Актуально для преобразования выгрузок из 1С.</b> <br> <i> Содержит функцию UnZip.pq от buchlotnik, размещенную по адресу: https://github.com/buchlotnik/buchlotnik_functions/blob/main/UnZip.pq </i>" ],
    result = Value.ReplaceType(fx_add_levels,docs)
in
    result
