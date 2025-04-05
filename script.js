function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
}

let hasShownError = false;

async function fetchEventData() {
    try {
        const response = await fetch('http://localhost:8000/users');

        if (!response.ok) {
            throw new Error(`Ошибка HTTP: ${ response.status }`);
        }
        const data = await response.json();
        hasShownError = false;
        return data;
    } catch (error) {
        console.error('Ошибка при загрузке данных:', error);
        if (!hasShownError) {
            alert('Не удалось загрузить данные с сервера. Попробуйте позже.');
            hasShownError = true;
        }
        return null;
    }
}

async function generateTable() {
    const dateFrom = new Date(document.getElementById('dateFrom').value);
    const dateTo = new Date(document.getElementById('dateTo').value);
    const table = document.getElementById('dataTable');

    if (isNaN(dateFrom.getTime()) || isNaN(dateTo.getTime())) {
        alert('Пожалуйста, выберите даты');
        return;
    }

    const eventData = await fetchEventData();
    if (!eventData || !eventData.length) {
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

    let tableContent = headerRow + subHeaderRow;

    eventData.forEach((personData, personIndex) => {
        const person = {
            id: personData.ID,
            fio: personData.PersonName
        };

        const eventsByDate = {};
        currentDate = new Date(dateFrom);
        for (let i = 0; i < diffDays; i++) {
            const currentDateStr = currentDate.toISOString().split('T')[0];
            eventsByDate[currentDateStr] = personData.events.filter(event => event.IN.Date === currentDateStr);
            currentDate.setDate(currentDate.getDate() + 1);
        }
        const maxEvents = Math.max(...Object.values(eventsByDate).map(events => events.length), 1);

        for (let rowIndex = 0; rowIndex < maxEvents; rowIndex++) {
            let row = '<tr>';

            if (rowIndex === 0) {
                row += `<td class="fixed-column-1" rowspan="${maxEvents}">${person.id}</td>`;
                row += `<td class="fixed-column-2 fio" rowspan="${maxEvents}">${person.fio}</td>`;
            }

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
    });

    table.innerHTML = tableContent;
}

async function exportToExcel() {
    const dateFrom = new Date(document.getElementById('dateFrom').value);
    const dateTo = new Date(document.getElementById('dateTo').value);
    const table = document.getElementById('dataTable');

    if (isNaN(dateFrom.getTime()) || isNaN(dateTo.getTime())) {
        alert('Пожалуйста, выберите даты');
        return;
    }

    if (!table.innerHTML) {
        await generateTable();
        if (!table.innerHTML) {
            // Если таблица не сгенерировалась из-за ошибки сервера, выходим
            return;
        }
    }

    const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
    const ws = wb.Sheets["Sheet1"];

    const fromStr = formatDate(dateFrom).replace(/\./g, '');
    const toStr = formatDate(dateTo).replace(/\./g, '');
    const fileName = `Attendance_${fromStr}-${toStr}.xlsx`;

    const diffTime = Math.abs(dateTo - dateFrom);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    const colWidths = [
        { wch: 5 },
        { wch: 30 },
    ];
    for (let i = 0; i < diffDays * 2; i++) {
        colWidths.push({ wch: 10 });
    }
    ws['!cols'] = colWidths;

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

    XLSX.writeFile(wb, gesturedName);
}

// Привязка к кнопкам
document.getElementById('generateButton')?.addEventListener('click', generateTable);
document.getElementById('exportButton')?.addEventListener('click', exportToExcel);