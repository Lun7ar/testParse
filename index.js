let XLSX = require("xlsx");

// const parse2 = (filename) => {
//     const excelData = XLSX.readFile(filename);
//     const worksheet = excelData.Sheets[excelData.SheetNames[0]];
//     /* Строка `let parsedData = XLSX.utils.sheet_to_json(worksheet);` использует функцию `sheet_to_json` из
//     библиотеки `XLSX` для преобразования данных рабочего листа в объект JSON. Эта функция принимает
//     рабочий лист в качестве входных данных и возвращает массив объектов, где каждый объект представляет
//     строку на рабочем листе. Каждая пара ключ-значение в объекте соответствует ячейке в строке, где ключ
//     — это заголовок столбца, а значение — значение ячейки. Полученным объектом JSON можно затем
//     манипулировать или использовать по мере необходимости. */
//     let parsedData = XLSX.utils.sheet_to_json(worksheet);

//     return parsedData.forEach(object => {
//         if (object['Координаты'] !== undefined) {  // проверяем, что ячейка с координатами не пустая
//             object['Координаты'] = object['Координаты'] //заменяем исходные значения в ячейке на обработанные
//                 .split(' ') //разбиваем значения в ячейке по пробелу, получается массив (ячейка) с вложенными массивами элементами которых являются координаты и пробелы
//                 .map(element => { // обрабатываем непосредственно значения ячейки, элементом является пробел или координата
//                     return element = element //заменяем исходную координату или пробел на обработанную 
//                         .split(new RegExp("\\r?\\n", "g")) //разбиваем элемент по переносу строки
//                         .filter(el => el !== '') //убираем лишние элементы в массиве в виде пробелов, оставшихся после обработки
//                 })
//                 .flat(); //преобразуем в единый массив верхнего уровня

//             console.log(object['Координаты'])
//             return object
//         }
//         return object
//     })
// };
  
  const parse = (filename) => { 
    const excelData = XLSX.readFile(filename); 
    const worksheet = excelData.Sheets[excelData.SheetNames[0]]; 
     
    let parsedData = XLSX.utils.sheet_to_json(worksheet); 
 
    return parsedData.map((object, index) => { 
        if (object['Координаты'] !== undefined) { 
            object['Координаты'] = object['Координаты']  
                .split(new RegExp("\\r?\\n", "g")) // разделение по строкам 
                .filter(el => el !== '') // удаление пустых строк 
                .map(line => { 
                    const decimalRegex = /[0-9]+\.[0-9]+/g;
                    const coordinates = line.match(decimalRegex); // извлечение чисел из строки 
                    if (coordinates !== null && coordinates.length % 2 === 0) { 
                        const pairs = []; 
                        for (let i = 0; i < coordinates.length; i += 2) { 
                            const x = parseFloat(coordinates[i]); 
                            const y = parseFloat(coordinates[i + 1]); 
                            if (!isNaN(x) && !isNaN(y)) { 
                                pairs.push([x, y]); // создание пар координат 
                            } else { 
                                console.log(`Номер строки: ${index + 2}, Неверные координаты: ${line}`);
                            } 
                        } 
                        return pairs; 
                    } else { 
                        console.log(`Номер строки: ${index + 2}, Неверные координаты: ${line}`);
                        return null; 
                    } 
                }) 
                .flat() 
                .filter(coord => coord !== null); // удаление неверных координат 
            console.log(object['Координаты']) 
            return object 
        } 
        return object;  
    }); 
} 

parse('./example.xlsx');

