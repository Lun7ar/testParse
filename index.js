let XLSX = require("xlsx");

const parse = (filename) => {
    const excelData = XLSX.readFile(filename);
    const worksheet = excelData.Sheets[excelData.SheetNames[0]];

    return Object.keys(worksheet).filter(x => /^D\d+/.test(x)).map((x) => ({
        x,
        column: worksheet[x].v,
    }));
};

// parse("./example.xlsx").forEach((element) => {
//     console.log(element.column);
// });


const parse2 = (filename) => {
    const excelData = XLSX.readFile(filename);
    const worksheet = excelData.Sheets[excelData.SheetNames[0]];
    /* Строка `let parsedData = XLSX.utils.sheet_to_json(worksheet);` использует функцию `sheet_to_json` из
    библиотеки `XLSX` для преобразования данных рабочего листа в объект JSON. Эта функция принимает
    рабочий лист в качестве входных данных и возвращает массив объектов, где каждый объект представляет
    строку на рабочем листе. Каждая пара ключ-значение в объекте соответствует ячейке в строке, где ключ
    — это заголовок столбца, а значение — значение ячейки. Полученным объектом JSON можно затем
    манипулировать или использовать по мере необходимости. */
    let parsedData = XLSX.utils.sheet_to_json(worksheet);

    return parsedData.forEach(object => {
        if (object['Координаты'] !== undefined) {  // проверяем, что ячейка с координатами не пустая
            object['Координаты'] = object['Координаты'] //заменяем исходные значения в ячейке на обработанные
                .split(' ') //разбиваем значения в ячейке по пробелу, получается массив (ячейка) с вложенными массивами элементами которых являются координаты и пробелы
                .map(element => { // обрабатываем непосредственно значения ячейки, элементом является пробел или координата
                    return element = element //заменяем исходную координату или пробел на обработанную 
                        .split(new RegExp("\\r?\\n", "g")) //разбиваем элемент по переносу строки
                        .filter(el => el !== '') //убираем лишние элементы в массиве в виде пробелов, оставшихся после обработки
                })
                .flat(); //преобразуем в единый массив верхнего уровня

            console.log(object['Координаты'])
            return object
        }
        return object
    })
};
parse2("./example.xlsx")