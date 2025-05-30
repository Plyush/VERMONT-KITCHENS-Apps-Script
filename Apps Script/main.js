var initialState = { // Эта переменная все еще не используется. Рассмотрите возможность ее удаления.
    sheetsData: {},
    hiddenSheets: []
};

function onOpen() {
    addMenu();
    setActiveSheet(); // Встановлюємо активний лист (сейчас "Template room")
    updateDropdownMenu1FromQuestionnaire();
    updateDropdownMenu1_1FromQuestionnaire();
    updateDropdownMenu2FromQuestionnaire();
    updateDropdownMenu3FromQuestionnaire();
    updateDropdownMenu4FromQuestionnaire();
    showOpenCompleteNotification();
    createTriggerOnEditForDropdownMenu1_1();
}

function addMenu() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('File status')
        .addItem('Save Initial State', 'saveInitialState')
        .addItem('Restore Initial State', 'restoreInitialState')
        .addToUi();
}

function saveInitialState() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    var sheetsData = {};
    var hiddenSheets = [];

    sheets.forEach(sheet => {
        sheetsData[sheet.getName()] = sheet.getDataRange().getValues();
        if (sheet.isSheetHidden()) {
            hiddenSheets.push(sheet.getName());
        }
    });

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
        var dataToRestore = sheetsData[sheetName];
        if (dataToRestore) {
            sheet.clearContents();
            if (dataToRestore.length > 0 && dataToRestore[0].length > 0) {
                sheet.getRange(1, 1, dataToRestore.length, dataToRestore[0].length).setValues(dataToRestore);
            }
        }

        if (hiddenSheets.includes(sheetName)) {
            sheet.hideSheet();
        } else {
            sheet.showSheet();
        }
    });
    Logger.log("Відновлено початковий стан!");
}

function getActiveSheetName() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function setActiveSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Template room";
    // TODO: Я незнаю який лист потрібно активовути (Убедитесь, что "Template room" - это правильный лист)
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        sheet.activate();
    } else {
        Logger.log("Помилка в setActiveSheet: Лист '" + sheetName + "' не знайдено.");
    }
}

function showOpenCompleteNotification() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("Файл успішно відкрито! Ви можете приступати до роботи.");
}

function updateDropdownMenu1FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName(getActiveSheetName());
    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка updateDropdownMenu1: один із листів не знайдено. Source: " + (sourceSheet ? sourceSheet.getName() : "null") + ", Target: " + (targetSheet ? targetSheet.getName() : "null"));
        return;
    }
    var dataRange = sourceSheet.getRange("B4:B16");
    var values = dataRange.getValues().flat().filter(value => value.toString().trim() !== "");
    values.unshift("ALL");
    var dropdownCell = targetSheet.getRange("A2");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    dropdownCell.setDataValidation(rule);
    dropdownCell.setValue("ALL");
    Logger.log("Випадаючий список 1 (A2) оновлено на листі '" + targetSheet.getName() + "'.");
}

function updateDropdownMenu1_1FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName(getActiveSheetName());

    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка updateDropdownMenu1_1: один із листів не знайдено.");
        return;
    }

    var firstDropdownCell = targetSheet.getRange("A2");
    var secondDropdownCell = targetSheet.getRange("B2");
    var selectedValue = firstDropdownCell.getValue().toString().trim();
    var filteredValuesForB2 = [];

    if (selectedValue === "ALL") {
        filteredValuesForB2 = ["ALL"];
    } else {
        var allDataQuestionaire = sourceSheet.getDataRange().getValues();
        var matchingRowFound = false;
        for (var i = 0; i < allDataQuestionaire.length; i++) {
            var row = allDataQuestionaire[i];
            var indexInRow = row.indexOf(selectedValue);
            if (indexInRow !== -1) {
                filteredValuesForB2 = row.slice(indexInRow + 1).filter(value => value.toString().trim() !== "");
                matchingRowFound = true;
                break;
            }
        }

        if (!matchingRowFound) {
            Logger.log("Помилка updateDropdownMenu1_1: значення '" + selectedValue + "' не знайдено на листі 'Questionaire' для залежного списку.");
            filteredValuesForB2 = ["ALL"];
        } else if (filteredValuesForB2.length === 0) {
            filteredValuesForB2 = ["ALL"];
        }
        filteredValuesForB2.unshift("ALL");
    }

    var rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredValuesForB2).setAllowInvalid(false).build();
    secondDropdownCell.setDataValidation(rule);
    if (filteredValuesForB2.includes("ALL")) {
        secondDropdownCell.setValue("ALL");
    } else if (filteredValuesForB2.length > 0) {
        secondDropdownCell.setValue(filteredValuesForB2[0]);
    } else {
        secondDropdownCell.clearDataValidations();
        secondDropdownCell.setValue("");
    }
    Logger.log("Другий випадаючий список (B2) оновлено на листі '" + targetSheet.getName() + "'.");
}

function createTriggerOnEditForDropdownMenu1_1() {
    var triggers = ScriptApp.getProjectTriggers();
    var triggerExists = triggers.some(trigger => trigger.getHandlerFunction() === "updateDropdownMenu1_1FromQuestionnaire");

    if (!triggerExists) {
        ScriptApp.newTrigger("updateDropdownMenu1_1FromQuestionnaire")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
        Logger.log("Тригер 'updateDropdownMenu1_1FromQuestionnaire' на onEdit створено!");
    } else {
        Logger.log("Тригер 'updateDropdownMenu1_1FromQuestionnaire' вже існує.");
    }
}

function updateDropdownMenu2FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName(getActiveSheetName());
    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка updateDropdownMenu2: один із листів не знайдено.");
        return;
    }
    var dataRange = sourceSheet.getRange("B22:B27");
    var values = dataRange.getValues().flat().filter(value => value.toString().trim() !== "");
    values.unshift("ALL");
    var dropdownCell = targetSheet.getRange("A4");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    dropdownCell.setDataValidation(rule);
    SpreadsheetApp.flush();
    dropdownCell.setValue("ALL");
    Logger.log("Випадаючий список 2 (A4) оновлено на листі '" + targetSheet.getName() + "'.");
}

function updateDropdownMenu3FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName(getActiveSheetName());
    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка updateDropdownMenu3: один із листів не знайдено.");
        return;
    }
    var dataRange = sourceSheet.getRange("B32:B37");
    var values = dataRange.getValues().flat().filter(value => value.toString().trim() !== "");
    values.unshift("ALL");
    var dropdownCell = targetSheet.getRange("A6");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    dropdownCell.setDataValidation(rule);
    SpreadsheetApp.flush();
    dropdownCell.setValue("ALL");
    Logger.log("Випадаючий список 3 (A6) оновлено на листі '" + targetSheet.getName() + "'.");
}

function updateDropdownMenu4FromQuestionnaire() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Questionaire");
    var targetSheet = ss.getSheetByName(getActiveSheetName());
    if (!sourceSheet || !targetSheet) {
        Logger.log("Помилка updateDropdownMenu4: один із листів не знайдено.");
        return;
    }
    var dataRange = sourceSheet.getRange("B40:C41");
    var values = dataRange.getValues().flat().filter(value => value.toString().trim() !== "");
    values.unshift("ALL");
    var dropdownCell = targetSheet.getRange("A8");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    dropdownCell.setDataValidation(rule);
    SpreadsheetApp.flush();
    dropdownCell.setValue("ALL");
    Logger.log("Випадаючий список 4 (A8) оновлено на листі '" + targetSheet.getName() + "'.");
}

function ensureCustomerOrderSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Customer Order";
    var existingSheet = spreadsheet.getSheetByName(sheetName);
    if (!existingSheet) {
        spreadsheet.insertSheet(sheetName);
        Logger.log("Лист '" + sheetName + "' створено.");
    } else {
        Logger.log("Лист '" + sheetName + "' вже існує.");
    }
}

function clearCustomerOrderSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Customer Order";
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
        sheet.clearContents();
        Logger.log("Лист '" + sheetName + "' очищено (тільки вміст)!");
    } else {
        Logger.log("Лист '" + sheetName + "' не знайдено для очищення.");
    }
}

function addRoomToСustomerOrderSheet() {
    Logger.log("Розпочато addRoomToСustomerOrderSheet на активному листі: " + SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName());
    filterCustomerOrderByDropMenu1();
    filterCustomerOrderByDropMenu2();
    filterCustomerOrderByDropMenu3();
    filterCustomerOrderByDropMenu4();
    Logger.log("Завершено addRoomToСustomerOrderSheet.");
}

function valueOfTheFirstDropMenuFromTheQuestionaireSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ИСПОЛЬЗУЕМ АКТИВНЫЙ ЛИСТ для чтения A2/B2 и для переименования
    var activeSheet = ss.getActiveSheet();
    var questionnaireSheet = ss.getSheetByName("Questionaire");
    // var ui = SpreadsheetApp.getUi(); // Раскомментируйте, если нужны ui.alert для ошибок переименования

    // activeSheet почти всегда будет существовать, если скрипт запущен из таблицы.
    // Поэтому явная проверка if (!activeSheet) здесь менее критична, чем для именованных листов.
    if (!questionnaireSheet) {
        Logger.log("Помилка valueOfTheFirstDropMenu: Лист 'Questionaire' не знайдено.");
        return { finalResultMenu1: "ERROR_QUESTIONNAIRE_NOT_FOUND", allResultMenu1: "ERROR_QUESTIONNAIRE_NOT_FOUND" };
    }

    // Читаем значения A2 и B2 с АКТИВНОГО листа
    var valueA2 = activeSheet.getRange("A2").getValue();
    var valueB2 = activeSheet.getRange("B2").getValue();

    // Это значение используется для поиска в 'Questionaire' и для формирования имени листа (если не "ALL")
    var derivedSelectedValue = (valueB2 === "ALL" || valueB2 === "") ? valueA2 : valueB2;
    var newSheetName = ""; // Будущее имя для активного листа

    if (valueA2 === "ALL" && (valueB2 === "ALL" || valueB2 === "")) {
        Logger.log("✅ valueOfTheFirstDropMenu: A2 и B2 на активном листе ('" + activeSheet.getName() + "') = ALL.");
        newSheetName = "ALL rooms"; // Формируем новое имя листа

        try {
            var oldName = activeSheet.getName();
            if (oldName !== newSheetName) {
                activeSheet.setName(newSheetName);
                SpreadsheetApp.flush(); // Применяем изменения немедленно
                Logger.log("✅ Активний лист '" + oldName + "' перейменовано на: " + newSheetName);
            } else {
                Logger.log("✅ Активний лист вже має назву: '" + newSheetName + "'. Перейменування не потрібне.");
            }
        } catch (e) {
            Logger.log('Помилка перейменування активного листа (' + activeSheet.getName() + ') на "' + newSheetName + '": ' + e.message);
            // ui.alert('Помилка перейменування', 'Не вдалося перейменувати активний лист на "' + newSheetName + '".\nПомилка: ' + e.message, ui.ButtonSet.OK);
        }
        return { finalResultMenu1: "ALL", allResultMenu1: "ALL" };
    }

    if (!derivedSelectedValue || derivedSelectedValue.toString().trim() === "") {
        Logger.log("Помилка valueOfTheFirstDropMenu: значення випадаючого списку (A2/B2) на активному листі ('" + activeSheet.getName() + "') порожнє. Активний лист не перейменовано.");
        return { finalResultMenu1: "ERROR_NO_SELECTION", allResultMenu1: "ERROR_NO_SELECTION" };
    }

    // Логика поиска для finalResultMenu1 и allResultMenu1
    var dataRangeValues = questionnaireSheet.getRange("B4:I16").getValues();
    var foundRowInSpreadsheet = -1;
    var foundColInSpreadsheet = -1;

    for (var r = 0; r < dataRangeValues.length; r++) {
        for (var c = 0; c < dataRangeValues[r].length; c++) {
            if (dataRangeValues[r][c].toString().trim() === derivedSelectedValue.toString().trim()) {
                foundRowInSpreadsheet = r + 4;
                foundColInSpreadsheet = c + 2;
                break;
            }
        }
        if (foundRowInSpreadsheet !== -1) break;
    }

    if (foundRowInSpreadsheet === -1 || foundColInSpreadsheet === -1) {
        Logger.log("Помилка valueOfTheFirstDropMenu: значення '" + derivedSelectedValue + "' (з активного листа '" + activeSheet.getName() + "') не знайдено в діапазоні B4:I16 на 'Questionaire'. Активний лист не перейменовано.");
        return { finalResultMenu1: "ERROR_NOT_FOUND", allResultMenu1: "ERROR_NOT_FOUND" };
    }

    var headerValue = questionnaireSheet.getRange(3, foundColInSpreadsheet).getValue();
    var rowValue = questionnaireSheet.getRange(foundRowInSpreadsheet, 1).getValue();

    var finalResultMenu1 = (valueA2 === "ALL" || valueA2 === "") ? "ALL" : headerValue.toString() + rowValue.toString();
    var allResultMenu1 = (valueA2 === "ALL" || valueA2 === "") ? "ALL" : rowValue.toString() + "ALL";

    // Логика переименования для случая, когда выбрано не "ALL"
    newSheetName = derivedSelectedValue.toString().trim() + " room"; // Формируем новое имя листа
    try {
        var oldName = activeSheet.getName();
        if (oldName !== newSheetName) {
            activeSheet.setName(newSheetName);
            SpreadsheetApp.flush();
            Logger.log("✅ Активний лист '" + oldName + "' перейменовано на: " + newSheetName);
        } else {
            Logger.log("✅ Активний лист вже має назву: '" + newSheetName + "'. Перейменування не потрібне.");
        }
    } catch (e) {
        Logger.log('Помилка перейменування активного листа (' + activeSheet.getName() + ') на "' + newSheetName + '": ' + e.message);
        // ui.alert('Помилка перейменування', 'Не вдалося перейменувати активний лист на "' + newSheetName + '".\nПомилка: ' + e.message, ui.ButtonSet.OK);
    }

    Logger.log("✅ valueOfTheFirstDropMenu: finalResultMenu1: " + finalResultMenu1 + ", allResultMenu1: " + allResultMenu1);
    return { finalResultMenu1: finalResultMenu1, allResultMenu1: allResultMenu1 };
}

function filterCustomerOrderByDropMenu1() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceDbSheet = ss.getSheetByName("Room components database");
    var targetActiveSheet = ss.getActiveSheet();

    if (!sourceDbSheet) {
        Logger.log("❌ filterCustomerOrderByDropMenu1: Лист 'Room components database' не найден.");
        return;
    }

    var activeSheetData = targetActiveSheet.getDataRange().getValues();
    var markerRowIndex = -1;
    var markerText = "Customer Order";

    for (var i = 0; i < activeSheetData.length; i++) {
        if (activeSheetData[i][0] === markerText) {
            markerRowIndex = i;
            break;
        }
    }

    if (markerRowIndex === -1) {
        Logger.log("❌ filterCustomerOrderByDropMenu1: Маркер '" + markerText + "' не найден в первой колонке активного листа '" + targetActiveSheet.getName() + "'. Данные не будут скопированы.");
        return;
    }
    var outputStartRow = markerRowIndex + 2;

    var resultValues = valueOfTheFirstDropMenuFromTheQuestionaireSheet();
    if (!resultValues || (resultValues.finalResultMenu1 && resultValues.finalResultMenu1.toString().startsWith("ERROR_"))) {
        Logger.log("❌ filterCustomerOrderByDropMenu1: Не удалось получить корректные значения для фильтрации из valueOfTheFirstDropMenu. " + (resultValues ? resultValues.finalResultMenu1 : ""));
        return;
    }

    var finalResultMenu1 = resultValues.finalResultMenu1;
    var allResultMenu1 = resultValues.allResultMenu1;

    var lastRowSource = sourceDbSheet.getLastRow();
    if (lastRowSource === 0) {
        Logger.log("⚠️ filterCustomerOrderByDropMenu1: Исходный лист 'Room components database' пуст.");
        return;
    }
    var lastColumnSource = sourceDbSheet.getLastColumn();
    var rangeSource = sourceDbSheet.getRange(1, 1, lastRowSource, lastColumnSource);
    var valuesSource = rangeSource.getValues();
    var backgroundsSource = rangeSource.getBackgrounds();

    if (targetActiveSheet.getMaxRows() >= outputStartRow) {
        targetActiveSheet.getRange(outputStartRow, 1, targetActiveSheet.getMaxRows() - outputStartRow + 1, targetActiveSheet.getMaxColumns()).clearContent();
    }

    if (finalResultMenu1 === "ALL" && allResultMenu1 === "ALL") {
        if (valuesSource.length > 0) {
            var targetRangeAll = targetActiveSheet.getRange(outputStartRow, 1, valuesSource.length, valuesSource[0].length);
            targetRangeAll.setValues(valuesSource);
            targetRangeAll.setBackgrounds(backgroundsSource);
            Logger.log("✅ filterCustomerOrderByDropMenu1: Все строки скопированы на '" + targetActiveSheet.getName() + "' после маркера.");
        } else {
            Logger.log("⚠️ filterCustomerOrderByDropMenu1: Нет данных для копирования из 'Room components database' (при ALL).");
        }
        return;
    }

    var filteredRows = [];
    var filteredBackgrounds = [];
    for (var row = 0; row < valuesSource.length; row++) {
        var cellValue = valuesSource[row][0];
        if (cellValue !== finalResultMenu1 && cellValue !== allResultMenu1 || backgroundsSource[row][0] === "#00AEEF") {
            filteredRows.push(valuesSource[row]);
            filteredBackgrounds.push(backgroundsSource[row]);
        }
    }

    if (filteredRows.length > 0) {
        var targetRangeFiltered = targetActiveSheet.getRange(outputStartRow, 1, filteredRows.length, filteredRows[0].length);
        targetRangeFiltered.setValues(filteredRows);
        targetRangeFiltered.setBackgrounds(filteredBackgrounds);
        Logger.log("✅ filterCustomerOrderByDropMenu1: Фильтрованные строки скопированы на '" + targetActiveSheet.getName() + "' после маркера!");
    } else {
        Logger.log("⚠️ filterCustomerOrderByDropMenu1: Нет строк для копирования на '" + targetActiveSheet.getName() + "' после фильтрации.");
    }
}

function filterCustomerOrderByDropMenu2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ИЗМЕНЕНО: Обе операции (чтение критерия и фильтрация данных) выполняются на активном листе
    var activeSheet = ss.getActiveSheet();

    // Получаем значение из выпадающего меню A4 на активном листе
    var filterValue = activeSheet.getRange("A4").getValue();

    if (filterValue === "ALL" || filterValue === "") {
        Logger.log("✅ filterCustomerOrderByDropMenu2: A4 на активном листе = ALL или пусто. Фильтрация не нужна.");
        return;
    }

    // Отримуємо всі дані з активного аркуша
    var values = activeSheet.getDataRange().getValues();
    var rangesToCheck = [
        ["1. Cabinet Construction", "2. Finish Panel and Door Material"],
        ["4. Hardware", "5. Extras + Other"]
    ];
    var rowsToDelete = [];

    rangesToCheck.forEach(function (bounds) {
        var startRowDataIdx = -1;
        var endRowDataIdx = -1;

        for (var r = 0; r < values.length; r++) {
            if (values[r][3] === bounds[0]) {
                startRowDataIdx = r;
            } else if (values[r][3] === bounds[1]) {
                endRowDataIdx = r;
                break;
            }
        }

        if (startRowDataIdx !== -1 && endRowDataIdx !== -1 && startRowDataIdx < endRowDataIdx) {
            for (var r = startRowDataIdx + 1; r < endRowDataIdx; r++) {
                var cellValueC = values[r][2];
                if (cellValueC !== filterValue && cellValueC !== "ALL") {
                    rowsToDelete.push(r + 1);
                }
            }
        }
    });

    if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a).forEach(rowNum => activeSheet.deleteRow(rowNum)); // Удаляем с активного листа
        Logger.log(`✅ filterCustomerOrderByDropMenu2: Удалено ${rowsToDelete.length} рядков с активного листа '${activeSheet.getName()}'.`);
    } else {
        Logger.log("⚠️ filterCustomerOrderByDropMenu2: Нет строк для удаления на активном листе '" + activeSheet.getName() + "'.");
    }
}

function filterCustomerOrderByDropMenu3() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ИЗМЕНЕНО: Обе операции на активном листе
    var activeSheet = ss.getActiveSheet();

    // Получаем значение из выпадающего меню A6 на активном листе
    var filterValue = activeSheet.getRange("A6").getValue();

    if (filterValue === "ALL" || filterValue === "") {
        Logger.log("✅ filterCustomerOrderByDropMenu3: A6 на активном листе = ALL или пусто. Фильтрация не нужна.");
        return;
    }

    var values = activeSheet.getDataRange().getValues();
    var rangesToCheck = [
        ["2. Finish Panel and Door Material", "3. Finishing Type"],
        ["3. Finishing Type", "4. Hardware"]
    ];
    var rowsToDelete = [];

    rangesToCheck.forEach(function (bounds) {
        var startRowDataIdx = -1;
        var endRowDataIdx = -1;
        for (var r = 0; r < values.length; r++) {
            if (values[r][3] === bounds[0]) { startRowDataIdx = r; }
            else if (values[r][3] === bounds[1]) { endRowDataIdx = r; break; }
        }
        if (startRowDataIdx !== -1 && endRowDataIdx !== -1 && startRowDataIdx < endRowDataIdx) {
            for (var r = startRowDataIdx + 1; r < endRowDataIdx; r++) {
                var cellValueC = values[r][2];
                if (cellValueC !== filterValue && cellValueC !== "ALL") {
                    rowsToDelete.push(r + 1);
                }
            }
        }
    });

    if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a).forEach(rowNum => activeSheet.deleteRow(rowNum)); // Удаляем с активного листа
        Logger.log(`✅ filterCustomerOrderByDropMenu3: Удалено ${rowsToDelete.length} рядков с активного листа '${activeSheet.getName()}'.`);
    } else {
        Logger.log("⚠️ filterCustomerOrderByDropMenu3: Нет строк для удаления на активном листе '" + activeSheet.getName() + "'.");
    }
}

function filterCustomerOrderByDropMenu4() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ИЗМЕНЕНО: Обе операции на активном листе
    var activeSheet = ss.getActiveSheet();

    // Получаем значение из выпадающего меню A8 на активном листе
    var filterValue = activeSheet.getRange("A8").getValue();

    if (filterValue === "ALL" || filterValue === "") {
        Logger.log("✅ filterCustomerOrderByDropMenu4: A8 на активном листе = ALL или пусто. Фильтрация не нужна.");
        return;
    }

    var values = activeSheet.getDataRange().getValues();
    var startRowDataIdx = -1;
    var endRowDataIdx = -1;

    for (var r = 0; r < values.length; r++) {
        if (values[r][3] === "5. Extras + Other") { startRowDataIdx = r; }
        else if (values[r][3] === "6. Overhead + Assembly") { endRowDataIdx = r; break; }
    }

    if (startRowDataIdx === -1 || endRowDataIdx === -1 || startRowDataIdx >= endRowDataIdx) {
        Logger.log("⚠️ filterCustomerOrderByDropMenu4: Не удалось найти межі ('5. Extras + Other', '6. Overhead + Assembly') в колонці D на активном листе '" + activeSheet.getName() + "'.");
        return;
    }

    var rowsToDelete = [];
    for (var r = startRowDataIdx + 1; r < endRowDataIdx; r++) {
        var cellValueB = values[r][1];
        if (cellValueB !== filterValue && cellValueB !== "ALL") {
            rowsToDelete.push(r + 1);
        }
    }

    if (rowsToDelete.length > 0) {
        rowsToDelete.sort((a, b) => b - a).forEach(rowNum => activeSheet.deleteRow(rowNum)); // Удаляем с активного листа
        Logger.log(`✅ filterCustomerOrderByDropMenu4: Удалено ${rowsToDelete.length} рядков с активного листа '${activeSheet.getName()}'.`);
    } else {
        Logger.log("⚠️ filterCustomerOrderByDropMenu4: Нет строк для удаления на активном листе '" + activeSheet.getName() + "'.");
    }
}