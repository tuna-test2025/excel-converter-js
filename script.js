document.addEventListener('DOMContentLoaded', function() {
const tableInput = document.getElementById('tableInput');
const convertBtn = document.getElementById('convertBtn');

// Пытаемся получить данные от ДБП через postMessage
let parentData = window.parent?.message?.data;

// Если не пришли через parent — читаем из URL-параметров
if (!parentData) {
    const urlParams = new URLSearchParams(window.location.search);
    parentData = {};
    if (urlParams.has('tableHtml')) parentData.tableHtml = urlParams.get('tableHtml');
    if (urlParams.has('tableBB')) parentData.tableBB = urlParams.get('tableBB');
}

let tableData = '';
if (parentData && parentData.tableHtml) tableData = parentData.tableHtml;
else if (parentData && parentData.tableBB) tableData = parentData.tableBB;

tableInput.value = tableData;

if (!tableData) {
    tableInput.placeholder = 'Таблица не передана из бизнес-процесса. Вставьте HTML или BB-код.';
    tableInput.style.display = 'block';
    convertBtn.textContent = 'Преобразовать в XLSX';
} else {
    tableInput.style.display = 'block';
    convertBtn.textContent = 'Преобразовать и завершить';
}

convertBtn.oncl ick = function() {
    const input = tableInput.value.trim();
    if (!input) {
     alert('Пожалуйста, введите таблицу.');
     return;
    }

    // Преобразуем BB-код в HTML
    let htmlTable = input;
    if (input.includes('[table]')) {
     htmlTable = input
        .replace(/\[table\]/g, '<table>')
        .replace(/\[\/table\]/g, '</table>')
        .replace(/\[tr\]/g, '<tr>')
        .replace(/\[\/tr\]/g, '</tr>')
        .replace(/\[td\]/g, '<td>')
        .replace(/\[\/td\]/g, '</td>');
    }

    // Парсим таблицу
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = '<table>' + htmlTable + '</table>';
    const table = tempDiv.querySelector('table');

    if (!table) {
     alert('Не удалось распознать таблицу. Проверьте формат.');
     return;
    }

    // Собираем данные в массив
    const worksheetData = [];
    const rows = table.querySelectorAll('tr');
    rows.forEach(row => {
     const rowData = [];
     const cells = row.querySelectorAll('td, th');
     cells.forEach(cell => {
        rowData.push(cell.textContent.trim());
     });
     worksheetData.push(rowData);
    });

    // Создаём XLSX
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Таблица');

    // Генерируем файл
    const excelBuffer = XLSX.write(workbook, { type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const fileUrl = URL.createObjectURL(blob);
    const fileName = 'таблица.xlsx';

    // Отправляем результат обратно в ДБП
    if (window.parent && window.parent.postMessage) {
     window.parent.postMessage({
        type: 'xlsx-generated',
        fileUrl: fileUrl,
        fileName: fileName
     }, '*');
    }

    // Закрываем окно через 1 секунду
    setTimeout(() => {
     window.close();
    }, 1000);
};
});