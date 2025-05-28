var initialState = {
    sheetsData: {},
    hiddenSheets: []
};

function onOpen() {
    addMenu();
    setActiveSheet(); // Встановлюємо активний лист
    updateDropdownMenu1FromQuestionnaire();
    updateDropdownMenu1_1FromQuestionnaire(); // Оновлюємо другий випадаючий список на основі першого
    updateDropdownMenu2FromQuestionnaire();
    updateDropdownMenu3FromQuestionnaire();
    updateDropdownMenu4FromQuestionnaire();
    showOpenCompleteNotification(); // Показуємо повідомлення про успішне відкриття файлу
    createTriggerOnEditForDropdownMenu1_1(); // Створюємо тригер для оновлення другого випадаючого списку при зміні першого
}

function addMenu() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('File status')
        .addItem('Save Initial State', 'saveInitialState') // Додати пункт меню для збереження початкового стану
        .addItem('Restore Initial State', 'restoreInitialState') // Додати пункт меню для відновлення початкового стану
        .addToUi();
}
function saveInitialState() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();

    var sheetsData = {}; // Об'єкт для збереження даних
    var hiddenSheets = [];

    sheets.forEach(sheet => {
        sheetsData[sheet.getName()] = sheet.getDataRange().getValues();
        if (sheet.isSheetHidden()) {
            hiddenSheets.push(sheet.getName());
        }
    });

    // Зберігаємо дані в PropertiesService
    var properties = PropertiesService.getDocumentProperties();
    properties.setProperty("sheetsData", JSON.stringify(sheetsData));
    properties.setProperty("hiddenSheets", JSON.stringify(hiddenSheets));

    Logger.log("Всі листи збережено!");
}

function restoreInitialState() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();

    var properties = PropertiesService.getDocumentProperties();
    var sheetsData = JSON.parse(properties.getProperty("sheetsData") || "{}");
    var hiddenSheets = JSON.parse(properties.getProperty("hiddenSheets") || "[]");

    sheets.forEach(sheet => {
        var sheetName = sheet.getName();
        if (sheetsData[sheetName]) {
            sheet.getDataRange().setValues(sheetsData[sheetName]); // Відновлення даних
        }

        if (hiddenSheets.includes(sheetName)) {
            sheet.hideSheet();
        } else {
            sheet.showSheet();
        }
    });

    Logger.log("Відновлено початковий стан!");
}

function setActiveSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Template room"); // Назва листа, який потрібно активувати
    if (sheet) {
        sheet.activate();
    }
}

// Функція для відкриття файлу та показу повідомлення про успішне відкриття
function showOpenCompleteNotification() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("Файл успішно відкрито! Ви можете приступати до роботи.");
}



// Функція для оновлення 1 випадаючого списку з Questionaire на Template room
function updateDropdownMenu1FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName("Template room");
    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    // Отримуємо дані з комірок B4:B16
    var dataRange = sourceSheet.getRange("B4:B16");
    var values = dataRange.getValues().flat(); // Перетворюємо 2D масив у 1D список

    // Очищаємо пусті значення
    var filteredValues = values.filter(value => value.toString().trim() !== "");

    // Додаємо "ALL" як перший елемент
    filteredValues.unshift("ALL");

    // Заповнюємо випадаючий список у комірці A2
    var dropdownCell = targetSheet.getRange("A2"); // Комірка, де буде випадаючий список
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValues).build();
    dropdownCell.setDataValidation(rule);

    // Встановлюємо "ALL" як початкове значення
    dropdownCell.setValue("ALL");

    Logger.log("Випадаючий список оновлено, перший пункт - 'ALL', і він встановлений як початкове значення!");
}

function updateDropdownMenu1_1FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName("Template room");

    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    var firstDropdownCell = targetSheet.getRange("A2"); // Перше меню
    var secondDropdownCell = targetSheet.getRange("B2"); // Друге меню
    var selectedValue = firstDropdownCell.getValue().toString().trim();

    if (selectedValue === "ALL") {
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(["ALL"]).build();
        secondDropdownCell.setDataValidation(rule);

        Logger.log("Другий список оновлено для 'ALL'.");
        return;
    }

    var dataRange = sourceSheet.getDataRange().getValues(); // Отримання всіх даних
    var matchingRow = dataRange.find(row => row.includes(selectedValue)); // Пошук рядка з відповідним значенням

    if (!matchingRow) {
        Logger.log("Помилка: відповідний рядок не знайдено.");
        return;
    }

    var index = matchingRow.indexOf(selectedValue);
    var filteredValues = matchingRow.slice(index + 1).filter(value => value.toString().trim() !== "");

    if (filteredValues.length === 0) {
        Logger.log("Помилка: немає доступних значень для другого списку.");
        return;
    }

    // Додаємо "ALL" у початок списку
    filteredValues.unshift("ALL");

    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValues).build();
    secondDropdownCell.setDataValidation(rule);

    Logger.log("Другий випадаючий список оновлено, 'ALL' додано першим пунктом.");
}

function createTriggerOnEditForDropdownMenu1_1() {
    var triggers = ScriptApp.getProjectTriggers();

    // Перевіряємо, чи тригер вже існує, щоб не створювати дублікати
    var triggerExists = triggers.some(trigger => trigger.getHandlerFunction() === "updateDropdownMenu1_1FromQuestionnaire");

    if (!triggerExists) {
        ScriptApp.newTrigger("updateDropdownMenu1_1FromQuestionnaire")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
        Logger.log("Тригер на зміну першого меню створено!");
    } else {
        Logger.log("Тригер вже існує, повторне створення не потрібно.");
    }
}

// Функція для оновлення 2 випадаючого списку з Questionaire на Template MFG
function updateDropdownMenu2FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName("Template room");

    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    // Отримуємо дані з комірок B22:B27
    var dataRange = sourceSheet.getRange("B22:B27");
    var values = dataRange.getValues().flat();

    // Очищаємо пусті значення
    var filteredValues = values.filter(value => value.toString().trim() !== "");

    // Додаємо "ALL" як перший елемент
    filteredValues.unshift("ALL");

    // Заповнюємо випадаючий список у комірці A4
    var dropdownCell = targetSheet.getRange("A4");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValues).build();
    dropdownCell.setDataValidation(rule);

    // Примусово застосовуємо зміни, щоб уникнути асинхронних проблем
    SpreadsheetApp.flush();

    // Встановлюємо "ALL" як початкове значення після оновлення валідації
    dropdownCell.setValue("ALL");

    Logger.log("Випадаючий список оновлено! 'ALL' додано першим пунктом і встановлено як початкове значення.");
}

// Функція для оновлення 3 випадаючого списку в комірці A6
function updateDropdownMenu3FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName("Template room");

    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    // Отримуємо дані з комірок B32:B37
    var dataRange = sourceSheet.getRange("B32:B37");
    var values = dataRange.getValues().flat(); // Перетворюємо 2D масив у 1D список

    // Очищаємо пусті значення
    var filteredValues = values.filter(value => value.toString().trim() !== "");

    // Додаємо "ALL" як перший елемент списку
    filteredValues.unshift("ALL");

    // Заповнюємо випадаючий список у комірці A6
    var dropdownCell = targetSheet.getRange("A6");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValues).build();
    dropdownCell.setDataValidation(rule);

    // Примусово застосовуємо зміни, щоб уникнути асинхронних проблем
    SpreadsheetApp.flush();

    // Встановлюємо "ALL" як початкове значення після оновлення списку
    dropdownCell.setValue("ALL");

    Logger.log("✅ Випадаючий список оновлено! 'ALL' додано першим пунктом і встановлено як початкове значення.");
}

// Функція для оновлення 4 випадаючого списку з Questionaire на Template MFG
function updateDropdownMenu4FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName("Template room");

    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    // Отримуємо дані з комірок B40:C41
    var dataRange = sourceSheet.getRange("B40:C41");
    var values = dataRange.getValues().flat(); // Перетворюємо 2D масив у 1D список

    // Очищаємо пусті значення
    var filteredValues = values.filter(value => value.toString().trim() !== "");

    // Додаємо "ALL" як перший елемент списку
    filteredValues.unshift("ALL");

    // Заповнюємо випадаючий список у комірці A8
    var dropdownCell = targetSheet.getRange("A8");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValues).build();
    dropdownCell.setDataValidation(rule);

    // Примусово застосовуємо зміни, щоб уникнути асинхронних проблем
    SpreadsheetApp.flush();

    // Встановлюємо "ALL" як початкове значення після оновлення списку
    dropdownCell.setValue("ALL");

    Logger.log("✅ Випадаючий список оновлено! 'ALL' додано першим пунктом і встановлено як початкове значення.");
}



// Функція для додавання нового аркуша "Customer Order" до таблиці
// якщо він ще не існує
function ensureCustomerOrderSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Customer Order";

    // Перевіряємо, чи лист вже існує
    var existingSheet = spreadsheet.getSheetByName(sheetName);

    if (!existingSheet) {
        // Якщо листа немає, створюємо його
        var newSheet = spreadsheet.insertSheet(sheetName);
        Logger.log("Лист 'Customer Order' створено.");
    } else {
        Logger.log("Лист 'Customer Order' вже існує.");
    }
}

// Функція для очищення листа "Customer Order"
// Вона видаляє всі дані з листа, але не сам лист
function clearCustomerOrderSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Customer Order");
    if (sheet) {
        sheet.clear(); // Очищаємо весь лист
        Logger.log("Лист 'Customer Order' очищено!");
    } else {
        Logger.log("Лист 'Customer Order' не знайдено.");
    }
}



function addRoomToСustomerOrderSheet() {
    // ensureCustomerOrderSheet();// Переконуємося, що лист існує
    filterCustomerOrderByDropMenu1(); // Викликаємо функцію для фільтрації та копіювання рядків в залежності від значення з випадаючого меню 1
    filterCustomerOrderByDropMenu2(); // Викликаємо функцію для фільтрації рядків в залежності від значення з випадаючого меню 2
    filterCustomerOrderByDropMenu3(); // Викликаємо функцію для фільтрації рядків в залежності від значення з випадаючого меню 3
    filterCustomerOrderByDropMenu4(); // Викликаємо функцію для фільтрації рядків в залежності від значення з випадаючого меню 4
}

function valueOfTheFirstDropMenuFromTheQuestionaireSheet() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Template room");
    var questionnaireSheet = ss.getSheetByName("Questionaire");

    if (!templateSheet || !questionnaireSheet) {
        Logger.log("Помилка: один із листів не знайдено.");
        return;
    }

    // 1️⃣ Отримуємо значення з випадаючого списку A2
    var selectedValue = templateSheet.getRange("A2").getValue();
    if (!selectedValue) {
        Logger.log("Помилка: значення випадаючого списку порожнє.");
        return;
    }

    // 2️⃣ Шукаємо це значення у B4:I16
    var dataRange = questionnaireSheet.getRange("B4:I16");
    var values = dataRange.getValues();
    var foundRow = -1;
    var foundColumn = -1;

    for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
            if (values[row][col] === selectedValue) {
                foundRow = row + 4; // Додаємо зсув, бо починаємо з B4
                foundColumn = col + 2; // Додаємо зсув, бо починаємо з B4 (B = 2)
                break;
            }
        }
        if (foundRow !== -1) break;
    }

    if (foundRow === -1 || foundColumn === -1) {
        Logger.log("Помилка: значення '" + selectedValue + "' не знайдено.");
        return;
    }

    // 3️⃣ Отримуємо значення з рядка 3 і відповідного стовпця
    var headerValue = questionnaireSheet.getRange(3, foundColumn).getValue();

    // 4️⃣ Отримуємо значення з стовпця A знайденого рядка
    var rowValue = questionnaireSheet.getRange(foundRow, 1).getValue();

    // 5️⃣ Виводимо результат у консоль
    //Logger.log("🔹 Знайдене значення: " + selectedValue);
    //Logger.log("📍 Адреса: " + foundRow + foundColumn);
    //Logger.log("🛠 Значення з рядка 3, стовпця " + foundColumn + ": " + headerValue);
    //Logger.log("💡 Значення з стовпця A, рядка " + foundRow + ": " + rowValue);

    var finalResultMenu1 = headerValue + rowValue;
    //Logger.log("✅ Остаточний результат: " + finalResult);
    var allResultMenu1 = headerValue + "ALL";

    // ✅ Повертаємо об'єкт з двома значеннями
    return { finalResultMenu1: finalResultMenu1, allResultMenu1: allResultMenu1 };

}

function filterCustomerOrderByDropMenu1() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Room components database");
    var resultSheet = ss.getSheetByName("Customer Order") || ss.insertSheet("Customer Order");

    if (!templateSheet) {
        Logger.log("❌ Ошибка: Лист 'Room components database' не найден.");
        return;
    }

    // Получаем значения из выпадающего меню
    var resultValues = valueOfTheFirstDropMenuFromTheQuestionaireSheet();
    if (!resultValues) {
        Logger.log("❌ Ошибка: Не удалось получить значение для фильтрации.");
        return;
    }

    var finalResultMenu1 = resultValues.finalResultMenu1;
    var allResultMenu1 = resultValues.allResultMenu1;

    // Определяем диапазон данных (начиная с первой строки)
    var lastRow = templateSheet.getLastRow();
    var lastColumn = templateSheet.getLastColumn();
    var range = templateSheet.getRange(1, 1, lastRow, lastColumn);
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();

    var filteredRows = [];

    for (var row = 0; row < values.length; row++) {
        var cellValue = values[row][0]; // Значение в столбце A

        if (cellValue !== finalResultMenu1 && cellValue !== allResultMenu1 || backgrounds[row][0] === "#00AEEF") {
            filteredRows.push(values[row]);
        }
    }

    // Копируем отфильтрованные данные на целевой лист
    if (filteredRows.length > 0) {
        var targetRange = resultSheet.getRange(1, 1, filteredRows.length, lastColumn);
        targetRange.setValues(filteredRows);
        targetRange.setBackgrounds(backgrounds);
        Logger.log("✅ Фильтрованные строки успешно скопированы!");
    } else {
        Logger.log("⚠️ Нет строк, удовлетворяющих критериям.");
    }
}

function filterCustomerOrderByDropMenu2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Template room");
    var resultSheet = ss.getSheetByName("Customer Order");

    if (!templateSheet || !resultSheet) {
        Logger.log("❌ Помилка: Один із листів не знайдено.");
        return;
    }

    var filterValue = templateSheet.getRange("A4").getValue(); // Значення для перевірки
    var values = resultSheet.getDataRange().getValues(); // Отримуємо всі дані з аркуша

    var rangesToCheck = [
        ["1. Cabinet Construction", "2. Finish Panel and Door Material"],
        ["4. Hardware", "5. Extras + Other"]
    ];

    var rowsToDelete = [];

    // 🔍 Проходимо кожну пару меж
    rangesToCheck.forEach(function (bounds) {
        var startRow = null;
        var endRow = null;

        // Знаходимо межі для поточного блоку
        for (var row = 0; row < values.length; row++) {
            if (values[row][3] === bounds[0]) {
                startRow = row;
            } else if (values[row][3] === bounds[1]) {
                endRow = row;
                break;
            }
        }

        // Якщо знайдено межі, перевіряємо рядки між ними
        if (startRow !== null && endRow !== null && startRow < endRow) {
            for (var row = startRow + 1; row < endRow; row++) {
                var cellValueC = values[row][2]; // Колонка C

                // Якщо значення C НЕ дорівнює `A4` або `"ALL"`, позначаємо рядок для видалення
                if (cellValueC !== filterValue && cellValueC !== "ALL") {
                    rowsToDelete.push(row + 1); // Зберігаємо номер рядка для видалення
                }
            }
        }
    });

    // 🔥 Видаляємо рядки у зворотному порядку, щоб не зміщувати індекси
    if (rowsToDelete.length > 0) {
        rowsToDelete.reverse().forEach(rowNum => resultSheet.deleteRow(rowNum));
        Logger.log(`✅ Видалено ${rowsToDelete.length} рядків.`);
    } else {
        Logger.log("⚠️ Усі рядки відповідали критеріям, нічого не видалено.");
    }
}

function filterCustomerOrderByDropMenu3() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Template room");
    var resultSheet = ss.getSheetByName("Customer Order");

    if (!templateSheet || !resultSheet) {
        Logger.log("❌ Помилка: Один із листів не знайдено.");
        return;
    }

    var filterValue = templateSheet.getRange("A6").getValue(); // Значення для перевірки
    var values = resultSheet.getDataRange().getValues(); // Отримуємо всі дані з аркуша

    var rangesToCheck = [
        ["2. Finish Panel and Door Material", "3. Finishing Type"],
        ["3. Finishing Type", "4. Hardware"]
    ];

    var rowsToDelete = [];

    // 🔍 Проходимо кожну пару меж
    rangesToCheck.forEach(function (bounds) {
        var startRow = null;
        var endRow = null;

        // Знаходимо межі для поточного блоку
        for (var row = 0; row < values.length; row++) {
            if (values[row][3] === bounds[0]) {
                startRow = row;
            } else if (values[row][3] === bounds[1]) {
                endRow = row;
                break;
            }
        }

        // Якщо знайдено межі, перевіряємо рядки між ними
        if (startRow !== null && endRow !== null && startRow < endRow) {
            for (var row = startRow + 1; row < endRow; row++) {
                var cellValueC = values[row][2]; // Колонка C

                // Якщо значення C НЕ дорівнює `A4` або `"ALL"`, позначаємо рядок для видалення
                if (cellValueC !== filterValue && cellValueC !== "ALL") {
                    rowsToDelete.push(row + 1); // Зберігаємо номер рядка для видалення
                }
            }
        }
    });

    // 🔥 Видаляємо рядки у зворотному порядку, щоб не зміщувати індекси
    if (rowsToDelete.length > 0) {
        rowsToDelete.reverse().forEach(rowNum => resultSheet.deleteRow(rowNum));
        Logger.log(`✅ Видалено ${rowsToDelete.length} рядків.`);
    } else {
        Logger.log("⚠️ Усі рядки відповідали критеріям, нічого не видалено.");
    }
}

function filterCustomerOrderByDropMenu4() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Template room");
    var resultSheet = ss.getSheetByName("Customer Order");

    if (!templateSheet || !resultSheet) {
        Logger.log("❌ Помилка: Один із листів не знайдено.");
        return;
    }

    var filterValue = templateSheet.getRange("A8").getValue(); // Значення для перевірки
    var values = resultSheet.getDataRange().getValues(); // Отримуємо всі дані з аркуша

    var startRow = null;
    var endRow = null;

    // 🔍 Знаходимо межі пошуку у колонці D
    for (var row = 0; row < values.length; row++) {
        if (values[row][3] === "5. Extras + Other") {
            startRow = row;
        } else if (values[row][3] === "6. Overhead + Assembly") {
            endRow = row;
            break; // При знаходженні обох меж — зупиняємо цикл
        }
    }

    if (startRow === null || endRow === null || startRow >= endRow) {
        Logger.log("⚠️ Не вдалося знайти потрібні межі в колонці D.");
        return;
    }

    var rowsToDelete = [];

    // 🔍 Перевіряємо рядки між `startRow` та `endRow`
    for (var row = startRow + 1; row < endRow; row++) {
        var cellValueC = values[row][1]; // Колонка B

        // Якщо значення C НЕ дорівнює `A8` або `"ALL"`, позначаємо рядок для видалення
        if (cellValueC !== filterValue && cellValueC !== "ALL") {
            rowsToDelete.push(row + 1); // Додаємо номер рядка для видалення (1-based index)
        }
    }

    // 🔥 Видаляємо рядки у зворотному порядку (щоб не зміщувати індекси)
    if (rowsToDelete.length > 0) {
        rowsToDelete.reverse().forEach(rowNum => resultSheet.deleteRow(rowNum));
        Logger.log(`✅ Видалено ${rowsToDelete.length} рядків.`);
    } else {
        Logger.log("⚠️ Усі рядки відповідали критеріям, нічого не видалено.");
    }
}