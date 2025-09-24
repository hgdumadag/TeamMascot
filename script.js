const fileInput = document.getElementById('fileInput');
const winnerCountInput = document.getElementById('winnerCount');
const prizeInput = document.getElementById('prizeInput');
const durationSlider = document.getElementById('durationSlider');
const durationDisplay = document.getElementById('durationDisplay');
const fieldSelector = document.getElementById('fieldSelector');
const pickButton = document.getElementById('pickButton');
const winnersDisplay = document.getElementById('winnersDisplay');
const loadingIndicator = document.getElementById('loadingIndicator');
const remainingCount = document.getElementById('remainingCount');
const winnerHistoryTable = document.getElementById('winnerHistoryTable');
const downloadButton = document.getElementById('downloadButton');

let participants = [];
let winnersHistory = [];
let availableFields = [];
let originalFields = [];

fileInput.addEventListener('change', handleFileUpload);
durationSlider.addEventListener('input', () => {
    durationDisplay.textContent = `${durationSlider.value}s`;
});
pickButton.addEventListener('click', startWinnerSelection);
downloadButton.addEventListener('click', downloadWinnersCSV);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) {
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (!json.length) {
                resetState();
                showFieldSelectorMessage('The uploaded file does not contain any rows.');
                alert('No data found in the uploaded file.');
                return;
            }

            participants = json.slice();
            availableFields = collectFields(json);
            originalFields = availableFields.slice();
            winnersHistory = [];

            enableControls();
            buildFieldSelector(availableFields);
            updateRemainingCount();
            updateWinnerCountLimits();
            clearWinnersDisplay();
            renderHistoryTable();
            downloadButton.disabled = true;
        } catch (error) {
            console.error(error);
            resetState();
            alert('There was a problem reading the file. Please make sure it is a valid CSV or Excel file.');
        }
    };

    reader.readAsArrayBuffer(file);
}

function collectFields(rows) {
    const fieldSet = new Set();
    rows.forEach((row) => {
        Object.keys(row).forEach((key) => {
            if (key) {
                fieldSet.add(key);
            }
        });
    });
    return Array.from(fieldSet);
}

function enableControls() {
    winnerCountInput.disabled = false;
    prizeInput.disabled = false;
    durationSlider.disabled = false;
    pickButton.disabled = participants.length === 0;
}

function resetState() {
    participants = [];
    winnersHistory = [];
    availableFields = [];
    originalFields = [];
    winnerCountInput.disabled = true;
    prizeInput.disabled = true;
    durationSlider.disabled = true;
    pickButton.disabled = true;
    downloadButton.disabled = true;
    winnerCountInput.value = 1;
    durationSlider.value = 3;
    durationDisplay.textContent = '3s';
    fieldSelector.innerHTML = '<p class="hint">Upload a file to select fields to display.</p>';
    remainingCount.textContent = 'Participants remaining: 0';
    clearWinnersDisplay();
    renderHistoryTable();
}

function buildFieldSelector(fields) {
    if (!fields.length) {
        showFieldSelectorMessage('No fields detected in the uploaded file.');
        return;
    }

    fieldSelector.innerHTML = '';
    const instructions = document.createElement('p');
    instructions.className = 'hint';
    instructions.textContent = 'Select the fields to show in the winners list:';
    fieldSelector.appendChild(instructions);
    fields.forEach((field) => {
        const id = `field-${field.replace(/\s+/g, '-').toLowerCase()}`;
        const label = document.createElement('label');
        label.className = 'checkbox-chip';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = id;
        checkbox.value = field;
        checkbox.checked = true;

        const span = document.createElement('span');
        span.textContent = field;

        label.appendChild(checkbox);
        label.appendChild(span);
        fieldSelector.appendChild(label);
    });
}

function showFieldSelectorMessage(message) {
    fieldSelector.innerHTML = `<p class="hint">${escapeHtml(message)}</p>`;
}

function updateWinnerCountLimits() {
    winnerCountInput.max = participants.length;
    if (Number(winnerCountInput.value) > participants.length) {
        winnerCountInput.value = participants.length || 1;
    }
}

function updateRemainingCount() {
    remainingCount.textContent = `Participants remaining: ${participants.length}`;
    if (participants.length === 0) {
        pickButton.disabled = true;
        winnerCountInput.disabled = true;
        prizeInput.disabled = true;
        durationSlider.disabled = true;
    } else {
        winnerCountInput.disabled = false;
        prizeInput.disabled = false;
        durationSlider.disabled = false;
    }
}

function clearWinnersDisplay() {
    winnersDisplay.classList.add('empty');
    winnersDisplay.innerHTML = '<p>No winners yet. Configure the draw and click &ldquo;Pick Winners&rdquo;.</p>';
}

function startWinnerSelection() {
    const winnerCount = Number(winnerCountInput.value);
    const prize = prizeInput.value.trim();
    const selectedFields = getSelectedFields();

    if (!participants.length) {
        alert('Please upload a file with participants before picking winners.');
        return;
    }

    if (!winnerCount || winnerCount < 1) {
        alert('Please specify how many winners to pick.');
        return;
    }

    if (winnerCount > participants.length) {
        alert('There are not enough participants remaining for that many winners.');
        return;
    }

    if (!prize) {
        alert('Please enter a prize for the winners.');
        return;
    }

    if (!selectedFields.length) {
        alert('Select at least one field to display for the winners.');
        return;
    }

    toggleLoading(true);
    disableControlsDuringDraw(true);

    const duration = Number(durationSlider.value) * 1000;

    setTimeout(() => {
        const winners = drawRandomWinners(winnerCount);
        const timestamp = new Date().toLocaleString();

        updateWinnersDisplay(winners, selectedFields, prize);
        recordWinnersHistory(winners, prize, timestamp);
        renderHistoryTable();

        updateWinnerCountLimits();
        updateRemainingCount();
        downloadButton.disabled = winnersHistory.length === 0;

        toggleLoading(false);
        disableControlsDuringDraw(false);
    }, duration);
}

function disableControlsDuringDraw(disabled) {
    pickButton.disabled = disabled || participants.length === 0;
    fileInput.disabled = disabled;
    winnerCountInput.disabled = disabled;
    prizeInput.disabled = disabled;
    durationSlider.disabled = disabled;
    fieldSelector.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
        checkbox.disabled = disabled;
    });
}

function toggleLoading(isLoading) {
    if (isLoading) {
        loadingIndicator.classList.remove('hidden');
        winnersDisplay.classList.add('hidden');
    } else {
        loadingIndicator.classList.add('hidden');
        winnersDisplay.classList.remove('hidden');
    }
}

function drawRandomWinners(count) {
    const selected = [];
    for (let i = 0; i < count; i += 1) {
        const index = Math.floor(Math.random() * participants.length);
        const [winner] = participants.splice(index, 1);
        selected.push(winner);
    }
    return selected;
}

function getSelectedFields() {
    const checkboxes = fieldSelector.querySelectorAll('input[type="checkbox"]');
    return Array.from(checkboxes)
        .filter((checkbox) => checkbox.checked)
        .map((checkbox) => checkbox.value);
}

function updateWinnersDisplay(winners, fields, prize) {
    if (!winners.length) {
        clearWinnersDisplay();
        return;
    }

    winnersDisplay.classList.remove('empty');
    winnersDisplay.innerHTML = '';

    winners.forEach((winner, index) => {
        const card = document.createElement('article');
        card.className = 'winner-card';

        const title = document.createElement('h3');
        const titleField = fields[0] || originalFields[0];
        const titleValue = titleField ? winner[titleField] : '';
        title.textContent = titleValue ? `${titleValue}` : `Winner ${index + 1}`;

        const fieldsList = document.createElement('div');
        fieldsList.className = 'winner-fields';

        fields.forEach((field) => {
            const value = winner[field];
            const fieldRow = document.createElement('div');
            fieldRow.innerHTML = `<strong>${escapeHtml(field)}:</strong> ${escapeHtml(value ?? '')}`;
            fieldsList.appendChild(fieldRow);
        });

        const prizeRow = document.createElement('div');
        prizeRow.innerHTML = `<strong>Prize:</strong> ${escapeHtml(prize)}`;
        fieldsList.appendChild(prizeRow);

        card.appendChild(title);
        card.appendChild(fieldsList);
        winnersDisplay.appendChild(card);
    });
}

function recordWinnersHistory(winners, prize, timestamp) {
    winners.forEach((winner) => {
        const entry = { ...winner, Prize: prize, 'Picked At': timestamp };
        winnersHistory.push(entry);
    });
}

function renderHistoryTable() {
    const thead = winnerHistoryTable.querySelector('thead');
    const tbody = winnerHistoryTable.querySelector('tbody');

    if (!winnersHistory.length) {
        thead.innerHTML = '';
        const colspan = Math.max(originalFields.length + 2, 1);
        tbody.innerHTML = `<tr><td class="empty" colspan="${colspan}">No winners selected yet.</td></tr>`;
        return;
    }

    const columns = [...originalFields, 'Prize', 'Picked At'];
    thead.innerHTML = `<tr>${columns.map((col) => `<th>${escapeHtml(col)}</th>`).join('')}</tr>`;
    tbody.innerHTML = winnersHistory
        .map((row) => {
            const cells = columns
                .map((col) => `<td>${escapeHtml(row[col] ?? '')}</td>`)
                .join('');
            return `<tr>${cells}</tr>`;
        })
        .join('');
}

function downloadWinnersCSV() {
    if (!winnersHistory.length) {
        return;
    }

    const columns = [...originalFields, 'Prize', 'Picked At'];
    const header = columns.join(',');
    const rows = winnersHistory.map((row) =>
        columns
            .map((column) => formatCSVValue(row[column]))
            .join(',')
    );

    const csvContent = [header, ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `winner-list-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

function formatCSVValue(value) {
    if (value == null) {
        return '';
    }
    const stringValue = String(value);
    if (/([",\n])/.test(stringValue)) {
        return `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
}

function escapeHtml(value) {
    if (value == null) {
        return '';
    }
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}
