function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
}

function generateTable(eventData) {
    const dateFrom = new Date(document.getElementById('dateFrom').value);
    const dateTo = new Date(document.getElementById('dateTo').value);
    const table = document.getElementById('dataTable');

    if (!dateFrom || !dateTo) {
        alert('Пожалуйста, выберите даты');
        return;
    }

    table.innerHTML = '';

    const diffTime = Math.abs(dateTo - dateFrom);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

    // Заголовок таблицы
    let headerRow = '<tr><th class="fixed-header fixed-column-1" rowspan="2">№</th><th class="fixed-header fixed-column-2" rowspan="2">ФИО</th>';
    let currentDate = new Date(dateFrom);
    for (let i = 0; i < diffDays; i++) {
        headerRow += `<th class="fixed-header date-header" colspan="2">${formatDate(currentDate)}</th>`;
        currentDate.setDate(currentDate.getDate() + 1);
    }
    headerRow += '</tr>';

    let subHeaderRow = '<tr>';
    for (let i = 0; i < diffDays; i++) {
        subHeaderRow += '<th class="fixed-header subheader">Вход</th><th class="fixed-header subheader">Выход</th>';
    }
    subHeaderRow += '</tr>';

    // Данные только из вашего массива
    const person = {
        id: eventData[0].ID,
        fio: eventData[0].PersonName
    };

    // Подсчитываем максимальное количество событий в день
    const eventsByDate = {};
    currentDate = new Date(dateFrom);
    for (let i = 0; i < diffDays; i++) {
        const currentDateStr = currentDate.toISOString().split('T')[0];
        eventsByDate[currentDateStr] = eventData[0].events.filter(event => event.IN.Date === currentDateStr);
        currentDate.setDate(currentDate.getDate() + 1);
    }
    const maxEvents = Math.max(...Object.values(eventsByDate).map(events => events.length), 1); // Минимум 1 строка

    // Генерация таблицы
    let tableContent = headerRow + subHeaderRow;
    for (let rowIndex = 0; rowIndex < maxEvents; rowIndex++) {
        let row = '<tr>';

        // Добавляем ID и ФИО только в первой строке с rowspan
        if (rowIndex === 0) {
            row += `<td class="fixed-column-1" rowspan="${maxEvents}">${person.id}</td>`;
            row += `<td class="fixed-column-2 fio" rowspan="${maxEvents}">${person.fio}</td>`;
        }

        // Заполняем данные по дням
        currentDate = new Date(dateFrom);
        for (let i = 0; i < diffDays; i++) {
            const currentDateStr = currentDate.toISOString().split('T')[0];
            const dayEvents = eventsByDate[currentDateStr];

            let inTime = '';
            let outTime = '';

            if (dayEvents && rowIndex < dayEvents.length) {
                inTime = dayEvents[rowIndex].IN.Time.slice(0, 5);
                outTime = dayEvents[rowIndex].OUT ? dayEvents[rowIndex].OUT.Time.slice(0, 5) : '';
            }

            row += `<td>${inTime}</td><td>${outTime}</td>`;
            currentDate.setDate(currentDate.getDate() + 1);
        }
        row += '</tr>';
        tableContent += row;
    }

    table.innerHTML = tableContent;
}

function exportToExcel(eventData) {
    const dateFrom = new Date(document.getElementById('dateFrom').value);
    const dateTo = new Date(document.getElementById('dateTo').value);
    const table = document.getElementById('dataTable');

    const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
    const ws = wb.Sheets["Sheet1"];

    // Имя файла с датами
    const fromStr = formatDate(dateFrom).replace(/\./g, '');
    const toStr = formatDate(dateTo).replace(/\./g, '');
    const fileName = `Attendance_${fromStr}-${toStr}.xlsx`;

    // Установка ширины колонок
    const diffTime = Math.abs(dateTo - dateFrom);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    const colWidths = [
        { wch: 5 },  // №
        { wch: 30 }, // ФИО
    ];
    for (let i = 0; i < diffDays * 2; i++) {
        colWidths.push({ wch: 10 });
    }
    ws['!cols'] = colWidths;

    // Центрирование данных
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 0; R <= range.e.r; R++) {
        for (let C = 2; C <= range.e.c; C++) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
            if (cell) {
                cell.s = {
                    alignment: {
                        horizontal: 'center',
                        vertical: 'center'
                    }
                };
            }
        }
    }

    XLSX.writeFile(wb, fileName);
}


// Пример использования с вашими данными
const eventData = [
    {
        "DeviceName": "192.168.1.56",
        "ID": "41",
        "PersonName": "Kabilov Ilxom",
        "events": [
            {
                "IN": {
                    "DateTime": "2025-03-04T07:10:26",
                    "Date": "2025-03-04",
                    "Time": "07:10:26"
                },
                "OUT": {
                    "DateTime": "2025-03-04T07:40:28",
                    "Date": "2025-03-04",
                    "Time": "07:40:28"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-04T07:40:56",
                    "Date": "2025-03-04",
                    "Time": "07:40:56"
                },
                "OUT": {
                    "DateTime": "2025-03-04T07:54:35",
                    "Date": "2025-03-04",
                    "Time": "07:54:35"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-05T07:04:01",
                    "Date": "2025-03-05",
                    "Time": "07:04:01"
                },
                "OUT": {
                    "DateTime": "2025-03-05T07:33:24",
                    "Date": "2025-03-05",
                    "Time": "07:33:24"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-06T07:04:38",
                    "Date": "2025-03-06",
                    "Time": "07:04:38"
                },
                "OUT": {
                    "DateTime": "2025-03-06T07:23:11",
                    "Date": "2025-03-06",
                    "Time": "07:23:11"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-06T12:49:51",
                    "Date": "2025-03-06",
                    "Time": "12:49:51"
                },
                "OUT": {
                    "DateTime": "2025-03-06T18:08:20",
                    "Date": "2025-03-06",
                    "Time": "18:08:20"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-07T09:23:17",
                    "Date": "2025-03-07",
                    "Time": "09:23:17"
                },
                "OUT": {
                    "DateTime": "2025-03-07T15:28:16",
                    "Date": "2025-03-07",
                    "Time": "15:28:16"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-08T07:28:07",
                    "Date": "2025-03-08",
                    "Time": "07:28:07"
                },
                "OUT": {
                    "DateTime": "2025-03-08T07:51:40",
                    "Date": "2025-03-08",
                    "Time": "07:51:40"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-08T14:13:19",
                    "Date": "2025-03-08",
                    "Time": "14:13:19"
                },
                "OUT": {
                    "DateTime": "2025-03-08T18:06:36",
                    "Date": "2025-03-08",
                    "Time": "18:06:36"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-09T07:09:25",
                    "Date": "2025-03-09",
                    "Time": "07:09:25"
                },
                "OUT": {
                    "DateTime": "2025-03-09T07:11:52",
                    "Date": "2025-03-09",
                    "Time": "07:11:52"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-09T07:18:07",
                    "Date": "2025-03-09",
                    "Time": "07:18:07"
                },
                "OUT": {
                    "DateTime": "2025-03-09T07:27:39",
                    "Date": "2025-03-09",
                    "Time": "07:27:39"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-09T07:28:16",
                    "Date": "2025-03-09",
                    "Time": "07:28:16"
                },
                "OUT": {
                    "DateTime": "2025-03-09T08:16:28",
                    "Date": "2025-03-09",
                    "Time": "08:16:28"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-09T10:02:13",
                    "Date": "2025-03-09",
                    "Time": "10:02:13"
                },
                "OUT": {
                    "DateTime": "2025-03-09T11:02:51",
                    "Date": "2025-03-09",
                    "Time": "11:02:51"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-09T11:43:18",
                    "Date": "2025-03-09",
                    "Time": "11:43:18"
                },
                "OUT": {
                    "DateTime": "2025-03-09T11:51:30",
                    "Date": "2025-03-09",
                    "Time": "11:51:30"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-10T07:33:05",
                    "Date": "2025-03-10",
                    "Time": "07:33:05"
                },
                "OUT": {
                    "DateTime": "2025-03-10T12:00:36",
                    "Date": "2025-03-10",
                    "Time": "12:00:36"
                }
            },
            {
                "IN": {

                    "DateTime": "2025-03-12T07:22:16",
                    "Date": "2025-03-12",
                    "Time": "07:22:16"
                },
                "OUT": {
                    "DateTime": "2025-03-12T07:31:50",
                    "Date": "2025-03-12",
                    "Time": "07:31:50"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-12T08:23:10",
                    "Date": "2025-03-12",
                    "Time": "08:23:10"
                },
                "OUT": {
                    "DateTime": "2025-03-12T12:48:05",
                    "Date": "2025-03-12",
                    "Time": "12:48:05"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-12T12:49:32",
                    "Date": "2025-03-12",
                    "Time": "12:49:32"
                },
                "OUT": {
                    "DateTime": "2025-03-12T18:10:47",
                    "Date": "2025-03-12",
                    "Time": "18:10:47"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-13T14:41:03",
                    "Date": "2025-03-13",
                    "Time": "14:41:03"
                },
                "OUT": {
                    "DateTime": "2025-03-13T18:06:44",
                    "Date": "2025-03-13",
                    "Time": "18:06:44"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-14T07:35:14",
                    "Date": "2025-03-14",
                    "Time": "07:35:14"
                },
                "OUT": {
                    "DateTime": "2025-03-14T07:58:03",
                    "Date": "2025-03-14",
                    "Time": "07:58:03"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-15T13:20:37",
                    "Date": "2025-03-15",
                    "Time": "13:20:37"
                },
                "OUT": {
                    "DateTime": "2025-03-15T18:31:32",
                    "Date": "2025-03-15",
                    "Time": "18:31:32"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-16T07:34:01",
                    "Date": "2025-03-16",
                    "Time": "07:34:01"
                },
                "OUT": {
                    "DateTime": "2025-03-16T07:48:45",
                    "Date": "2025-03-16",
                    "Time": "07:48:45"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-17T07:12:46",
                    "Date": "2025-03-17",
                    "Time": "07:12:46"
                },
                "OUT": {
                    "DateTime": "2025-03-17T10:44:05",
                    "Date": "2025-03-17",
                    "Time": "10:44:05"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-17T16:15:18",
                    "Date": "2025-03-17",
                    "Time": "16:15:18"
                },
                "OUT": {
                    "DateTime": "2025-03-17T18:03:45",
                    "Date": "2025-03-17",
                    "Time": "18:03:45"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-18T09:27:56",
                    "Date": "2025-03-18",
                    "Time": "09:27:56"
                },
                "OUT": {
                    "DateTime": "2025-03-18T10:29:00",
                    "Date": "2025-03-18",
                    "Time": "10:29:00"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-18T14:11:16",
                    "Date": "2025-03-18",
                    "Time": "14:11:16"
                },
                "OUT": {
                    "DateTime": "2025-03-18T15:09:54",
                    "Date": "2025-03-18",
                    "Time": "15:09:54"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-18T15:14:29",
                    "Date": "2025-03-18",
                    "Time": "15:14:29"
                },
                "OUT": {
                    "DateTime": "2025-03-18T18:08:48",
                    "Date": "2025-03-18",
                    "Time": "18:08:48"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T07:03:20",
                    "Date": "2025-03-19",
                    "Time": "07:03:20"
                },
                "OUT": {
                    "DateTime": "2025-03-19T07:13:11",
                    "Date": "2025-03-19",
                    "Time": "07:13:11"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T07:13:41",
                    "Date": "2025-03-19",
                    "Time": "07:13:41"
                },
                "OUT": {
                    "DateTime": "2025-03-19T07:24:16",
                    "Date": "2025-03-19",
                    "Time": "07:24:16"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T07:25:45",
                    "Date": "2025-03-19",
                    "Time": "07:25:45"
                },
                "OUT": {

                    "DateTime": "2025-03-19T07:31:17",
                    "Date": "2025-03-19",
                    "Time": "07:31:17"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T07:33:41",
                    "Date": "2025-03-19",
                    "Time": "07:33:41"
                },
                "OUT": {
                    "DateTime": "2025-03-19T08:32:06",
                    "Date": "2025-03-19",
                    "Time": "08:32:06"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T09:13:34",
                    "Date": "2025-03-19",
                    "Time": "09:13:34"
                },
                "OUT": {
                    "DateTime": "2025-03-19T09:48:47",
                    "Date": "2025-03-19",
                    "Time": "09:48:47"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T09:53:22",
                    "Date": "2025-03-19",
                    "Time": "09:53:22"
                },
                "OUT": {
                    "DateTime": "2025-03-19T10:21:44",
                    "Date": "2025-03-19",
                    "Time": "10:21:44"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T10:22:06",
                    "Date": "2025-03-19",
                    "Time": "10:22:06"
                },
                "OUT": {
                    "DateTime": "2025-03-19T10:32:41",
                    "Date": "2025-03-19",
                    "Time": "10:32:41"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T11:17:41",
                    "Date": "2025-03-19",
                    "Time": "11:17:41"
                },
                "OUT": {
                    "DateTime": "2025-03-19T13:51:24",
                    "Date": "2025-03-19",
                    "Time": "13:51:24"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T14:03:56",
                    "Date": "2025-03-19",
                    "Time": "14:03:56"
                },
                "OUT": {
                    "DateTime": "2025-03-19T15:28:04",
                    "Date": "2025-03-19",
                    "Time": "15:28:04"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-19T15:34:24",
                    "Date": "2025-03-19",
                    "Time": "15:34:24"
                },
                "OUT": {
                    "DateTime": "2025-03-19T16:05:06",
                    "Date": "2025-03-19",
                    "Time": "16:05:06"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-20T07:04:06",
                    "Date": "2025-03-20",
                    "Time": "07:04:06"
                },
                "OUT": {
                    "DateTime": "2025-03-20T07:12:48",
                    "Date": "2025-03-20",
                    "Time": "07:12:48"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-20T08:03:48",
                    "Date": "2025-03-20",
                    "Time": "08:03:48"
                },
                "OUT": {
                    "DateTime": "2025-03-20T11:28:29",
                    "Date": "2025-03-20",
                    "Time": "11:28:29"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-20T13:58:25",
                    "Date": "2025-03-20",
                    "Time": "13:58:25"
                },
                "OUT": {
                    "DateTime": "2025-03-20T17:59:16",
                    "Date": "2025-03-20",
                    "Time": "17:59:16"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-21T07:13:31",
                    "Date": "2025-03-21",
                    "Time": "07:13:31"
                },
                "OUT": {
                    "DateTime": "2025-03-21T07:36:42",
                    "Date": "2025-03-21",
                    "Time": "07:36:42"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-22T07:09:07",
                    "Date": "2025-03-22",
                    "Time": "07:09:07"
                },
                "OUT": {
                    "DateTime": "2025-03-22T07:14:50",
                    "Date": "2025-03-22",
                    "Time": "07:14:50"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-22T07:15:19",
                    "Date": "2025-03-22",
                    "Time": "07:15:19"
                },
                "OUT": {
                    "DateTime": "2025-03-22T07:29:22",
                    "Date": "2025-03-22",
                    "Time": "07:29:22"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-22T09:31:28",
                    "Date": "2025-03-22",
                    "Time": "09:31:28"
                },
                "OUT": {
                    "DateTime": "2025-03-22T11:39:13",
                    "Date": "2025-03-22",
                    "Time": "11:39:13"
                }
            },
            {

                "IN": {
                    "DateTime": "2025-03-22T11:44:42",
                    "Date": "2025-03-22",
                    "Time": "11:44:42"
                },
                "OUT": {
                    "DateTime": "2025-03-22T14:42:25",
                    "Date": "2025-03-22",
                    "Time": "14:42:25"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-22T14:43:30",
                    "Date": "2025-03-22",
                    "Time": "14:43:30"
                },
                "OUT": {
                    "DateTime": "2025-03-22T18:08:19",
                    "Date": "2025-03-22",
                    "Time": "18:08:19"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-23T07:11:27",
                    "Date": "2025-03-23",
                    "Time": "07:11:27"
                },
                "OUT": {
                    "DateTime": "2025-03-23T07:13:35",
                    "Date": "2025-03-23",
                    "Time": "07:13:35"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-23T11:05:14",
                    "Date": "2025-03-23",
                    "Time": "11:05:14"
                },
                "OUT": {
                    "DateTime": "2025-03-23T15:00:56",
                    "Date": "2025-03-23",
                    "Time": "15:00:56"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-23T15:01:18",
                    "Date": "2025-03-23",
                    "Time": "15:01:18"
                },
                "OUT": {
                    "DateTime": "2025-03-24T07:39:54",
                    "Date": "2025-03-24",
                    "Time": "07:39:54"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-23T15:01:19",
                    "Date": "2025-03-23",
                    "Time": "15:01:19"
                },
                "OUT": {
                    "DateTime": "2025-03-23T17:43:09",
                    "Date": "2025-03-23",
                    "Time": "17:43:09"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-24T07:09:59",
                    "Date": "2025-03-24",
                    "Time": "07:09:59"
                },
                "OUT": {
                    "DateTime": "2025-03-24T07:19:56",
                    "Date": "2025-03-24",
                    "Time": "07:19:56"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-24T07:21:12",
                    "Date": "2025-03-24",
                    "Time": "07:21:12"
                },
                "OUT": {
                    "DateTime": "2025-03-24T07:29:28",
                    "Date": "2025-03-24",
                    "Time": "07:29:28"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-24T11:10:31",
                    "Date": "2025-03-24",
                    "Time": "11:10:31"
                },
                "OUT": {
                    "DateTime": "2025-03-24T13:23:56",
                    "Date": "2025-03-24",
                    "Time": "13:23:56"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-24T16:42:16",
                    "Date": "2025-03-24",
                    "Time": "16:42:16"
                },
                "OUT": {
                    "DateTime": "2025-03-24T18:09:13",
                    "Date": "2025-03-24",
                    "Time": "18:09:13"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-25T07:43:30",
                    "Date": "2025-03-25",
                    "Time": "07:43:30"
                },
                "OUT": {
                    "DateTime": "2025-03-25T13:21:40",
                    "Date": "2025-03-25",
                    "Time": "13:21:40"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-25T13:22:34",
                    "Date": "2025-03-25",
                    "Time": "13:22:34"
                },
                "OUT": {
                    "DateTime": "2025-03-25T15:52:28",
                    "Date": "2025-03-25",
                    "Time": "15:52:28"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-26T07:10:33",
                    "Date": "2025-03-26",
                    "Time": "07:10:33"
                },
                "OUT": {
                    "DateTime": "2025-03-26T07:20:05",
                    "Date": "2025-03-26",
                    "Time": "07:20:05"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-26T08:23:31",
                    "Date": "2025-03-26",
                    "Time": "08:23:31"
                },
                "OUT": {
                    "DateTime": "2025-03-26T09:22:36",
                    "Date": "2025-03-26",
                    "Time": "09:22:36"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-26T09:23:34",
                    "Date": "2025-03-26",
                    "Time": "09:23:34"
                },

                "OUT": {
                    "DateTime": "2025-03-26T14:36:01",
                    "Date": "2025-03-26",
                    "Time": "14:36:01"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-27T07:13:45",
                    "Date": "2025-03-27",
                    "Time": "07:13:45"
                },
                "OUT": {
                    "DateTime": "2025-03-27T07:26:56",
                    "Date": "2025-03-27",
                    "Time": "07:26:56"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-27T09:07:31",
                    "Date": "2025-03-27",
                    "Time": "09:07:31"
                },
                "OUT": {
                    "DateTime": "2025-03-27T11:57:11",
                    "Date": "2025-03-27",
                    "Time": "11:57:11"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-27T12:01:40",
                    "Date": "2025-03-27",
                    "Time": "12:01:40"
                },
                "OUT": {
                    "DateTime": "2025-03-27T12:39:03",
                    "Date": "2025-03-27",
                    "Time": "12:39:03"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-27T12:40:47",
                    "Date": "2025-03-27",
                    "Time": "12:40:47"
                },
                "OUT": {
                    "DateTime": "2025-03-27T18:07:13",
                    "Date": "2025-03-27",
                    "Time": "18:07:13"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-28T07:58:00",
                    "Date": "2025-03-28",
                    "Time": "07:58:00"
                },
                "OUT": {
                    "DateTime": "2025-03-28T14:09:40",
                    "Date": "2025-03-28",
                    "Time": "14:09:40"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-28T14:12:02",
                    "Date": "2025-03-28",
                    "Time": "14:12:02"
                },
                "OUT": {
                    "DateTime": "2025-03-28T14:31:46",
                    "Date": "2025-03-28",
                    "Time": "14:31:46"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-28T15:47:41",
                    "Date": "2025-03-28",
                    "Time": "15:47:41"
                },
                "OUT": {
                    "DateTime": "2025-03-28T17:11:17",
                    "Date": "2025-03-28",
                    "Time": "17:11:17"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-28T17:13:45",
                    "Date": "2025-03-28",
                    "Time": "17:13:45"
                },
                "OUT": {
                    "DateTime": "2025-03-28T18:19:05",
                    "Date": "2025-03-28",
                    "Time": "18:19:05"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-29T07:01:59",
                    "Date": "2025-03-29",
                    "Time": "07:01:59"
                },
                "OUT": {
                    "DateTime": "2025-03-29T07:15:58",
                    "Date": "2025-03-29",
                    "Time": "07:15:58"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-29T10:08:07",
                    "Date": "2025-03-29",
                    "Time": "10:08:07"
                },
                "OUT": {
                    "DateTime": "2025-03-29T11:06:40",
                    "Date": "2025-03-29",
                    "Time": "11:06:40"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-29T11:08:15",
                    "Date": "2025-03-29",
                    "Time": "11:08:15"
                },
                "OUT": {
                    "DateTime": "2025-03-29T11:27:11",
                    "Date": "2025-03-29",
                    "Time": "11:27:11"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-29T12:09:01",
                    "Date": "2025-03-29",
                    "Time": "12:09:01"
                },
                "OUT": {
                    "DateTime": "2025-03-29T13:41:02",
                    "Date": "2025-03-29",
                    "Time": "13:41:02"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-29T14:47:41",
                    "Date": "2025-03-29",
                    "Time": "14:47:41"
                },
                "OUT": {
                    "DateTime": "2025-03-29T18:13:53",
                    "Date": "2025-03-29",
                    "Time": "18:13:53"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-03-31T07:32:46",
                    "Date": "2025-03-31",
                    "Time": "07:32:46"
                },
                "OUT": {
                    "DateTime": "2025-03-31T08:21:46",
                    "Date": "2025-03-31",
                    "Time": "08:21:46"
                }
            },

            {
                "IN": {
                    "DateTime": "2025-03-31T08:23:32",
                    "Date": "2025-03-31",
                    "Time": "08:23:32"
                },
                "OUT": {
                    "DateTime": "2025-03-31T10:11:42",
                    "Date": "2025-03-31",
                    "Time": "10:11:42"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-01T07:21:59",
                    "Date": "2025-04-01",
                    "Time": "07:21:59"
                },
                "OUT": {
                    "DateTime": "2025-04-01T09:34:37",
                    "Date": "2025-04-01",
                    "Time": "09:34:37"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-01T13:36:09",
                    "Date": "2025-04-01",
                    "Time": "13:36:09"
                },
                "OUT": {
                    "DateTime": "2025-04-01T15:41:01",
                    "Date": "2025-04-01",
                    "Time": "15:41:01"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-01T15:41:50",
                    "Date": "2025-04-01",
                    "Time": "15:41:50"
                },
                "OUT": {
                    "DateTime": "2025-04-01T18:12:43",
                    "Date": "2025-04-01",
                    "Time": "18:12:43"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T07:37:55",
                    "Date": "2025-04-02",
                    "Time": "07:37:55"
                },
                "OUT": {
                    "DateTime": "2025-04-02T08:16:10",
                    "Date": "2025-04-02",
                    "Time": "08:16:10"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T08:16:45",
                    "Date": "2025-04-02",
                    "Time": "08:16:45"
                },
                "OUT": {
                    "DateTime": "2025-04-02T11:32:51",
                    "Date": "2025-04-02",
                    "Time": "11:32:51"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T11:36:27",
                    "Date": "2025-04-02",
                    "Time": "11:36:27"
                },
                "OUT": {
                    "DateTime": "2025-04-02T12:46:22",
                    "Date": "2025-04-02",
                    "Time": "12:46:22"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T12:50:09",
                    "Date": "2025-04-02",
                    "Time": "12:50:09"
                },
                "OUT": null
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T13:30:10",
                    "Date": "2025-04-02",
                    "Time": "13:30:10"
                },
                "OUT": {
                    "DateTime": "2025-04-02T15:46:55",
                    "Date": "2025-04-02",
                    "Time": "15:46:55"
                }
            },
            {
                "IN": {
                    "DateTime": "2025-04-02T16:01:43",
                    "Date": "2025-04-02",
                    "Time": "16:01:43"
                },
                "OUT": null
            },
            {
                "IN": {
                    "DateTime": "2025-04-03T07:10:00",
                    "Date": "2025-04-03",
                    "Time": "07:10:00"
                },
                "OUT": null
            },
            {
                "IN": {
                    "DateTime": "2025-04-03T12:16:10",
                    "Date": "2025-04-03",
                    "Time": "12:16:10"
                },
                "OUT": null
            },
            {
                "IN": {
                    "DateTime": "2025-04-03T12:16:11",
                    "Date": "2025-04-03",
                    "Time": "12:16:11"
                },
                "OUT": null
            },
            {
                "IN": {
                    "DateTime": "2025-04-03T12:38:30",
                    "Date": "2025-04-03",
                    "Time": "12:38:30"
                },
                "OUT": null
            }
        ]
    },];
// Привязка к кнопкам
document.getElementById('generateButton')?.addEventListener('click', () => generateTable(eventData));
document.getElementById('exportButton')?.addEventListener('click', () => exportToExcel(eventData));