/**
 * @typedef {Object} DOMelements
 * @property {HTMLElement} jsonFileInput - The file input element for JSON.
 * @property {HTMLElement} convertButton - The button to trigger conversion.
 * @property {HTMLElement} messageDiv - The div for displaying messages.
 * @property {HTMLElement} selectedFileNameDiv - The div to display the selected file name.
 * @property {HTMLElement} previewTableContainer - The container for the preview table.
 * @property {HTMLElement} tableHeadersRow - The table row for headers.
 * @property {HTMLElement} tableBody - The table body.
 * @property {HTMLElement} fieldSelectionContainer - The container for field selection checkboxes.
 * @property {HTMLElement} fieldCheckboxesDiv - The div containing the field checkboxes.
 */

class JsonToXlsxConverter {
    /** @type {DOMelements} */
    #elements;
    /** @type {Array<object>} */
    #originalJsonData = [];
    /** @type {Array<object>} */
    #flattenedAndCleanedData = [];

    constructor() {
        this.#elements = this.#getDomElements();
        this.#addEventListeners();
    }

    /**
     * Retrieves all necessary DOM elements.
     * @returns {DOMelements} An object containing references to the DOM elements.
     */
    #getDomElements() {
        return {
            jsonFileInput: document.getElementById('jsonFile'),
            convertButton: document.getElementById('convertBtn'),
            messageDiv: document.getElementById('message'),
            selectedFileNameDiv: document.getElementById('selectedFileName'),
            previewTableContainer: document.getElementById('previewTableContainer'),
            tableHeadersRow: document.getElementById('tableHeaders'),
            tableBody: document.getElementById('tableBody'),
            fieldSelectionContainer: document.getElementById('fieldSelectionContainer'),
            fieldCheckboxesDiv: document.getElementById('fieldCheckboxes')
        };
    }

    /**
     * Adds event listeners to the relevant DOM elements.
     */
    #addEventListeners() {
        this.#elements.jsonFileInput.addEventListener('change', this.#handleJsonFileChange.bind(this));
        this.#elements.convertButton.addEventListener('click', this.#handleConvertButtonClick.bind(this));
    }

    /**
     * Handles the change event of the JSON file input.
     * @param {Event} event - The change event.
     */
    #handleJsonFileChange(event) {
        const file = event.target.files[0];
        if (file) {
            this.#elements.selectedFileNameDiv.textContent = file.name;
            this.#elements.messageDiv.textContent = ''; // Clear any previous messages
            this.#elements.previewTableContainer.classList.add('hidden'); // Hide table until conversion
            this.#elements.fieldSelectionContainer.classList.add('hidden'); // Hide field selection initially
            this.#elements.tableHeadersRow.innerHTML = '';
            this.#elements.tableBody.innerHTML = '';
            this.#elements.fieldCheckboxesDiv.innerHTML = '';

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    this.#originalJsonData = JSON.parse(e.target.result);
                    if (!Array.isArray(this.#originalJsonData)) {
                        this.#elements.messageDiv.textContent = 'JSON data must be an array of objects.';
                        return;
                    }

                    this.#flattenedAndCleanedData = this.#originalJsonData.map(item => this.#flattenAndCleanObject(item));
                    
                    this.#populateFieldSelection(this.#flattenedAndCleanedData);
                    this.#renderTable(this.#flattenedAndCleanedData); // Render table with all fields initially

                } catch (error) {
                    this.#elements.messageDiv.textContent = 'Error reading JSON file: ' + error.message;
                    console.error('Error reading JSON:', error);
                }
            };
            reader.readAsText(file);
        } else {
            this.#elements.selectedFileNameDiv.textContent = 'No file chosen';
            this.#elements.messageDiv.textContent = '';
            this.#elements.previewTableContainer.classList.add('hidden');
            this.#elements.fieldSelectionContainer.classList.add('hidden');
            this.#elements.tableHeadersRow.innerHTML = '';
            this.#elements.tableBody.innerHTML = '';
            this.#elements.fieldCheckboxesDiv.innerHTML = '';
            this.#originalJsonData = [];
            this.#flattenedAndCleanedData = [];
        }
    }

    /**
     * Handles the click event of the convert button.
     */
    #handleConvertButtonClick() {
        if (!this.#flattenedAndCleanedData || this.#flattenedAndCleanedData.length === 0) {
            this.#elements.messageDiv.textContent = 'No data to convert. Please upload a JSON file first.';
            return;
        }

        this.#elements.messageDiv.textContent = ''; // Clear previous messages

        const selectedFields = Array.from(this.#elements.fieldCheckboxesDiv.querySelectorAll('input[type="checkbox"]:checked'))
                                .map(checkbox => checkbox.value);

        if (selectedFields.length === 0) {
            this.#elements.messageDiv.textContent = 'Please select at least one field to export.';
            return;
        }

        // Filter data based on selected fields
        const dataToExport = this.#flattenedAndCleanedData.map(row => {
            const newRow = {};
            selectedFields.forEach(field => {
                newRow[field] = row[field];
            });
            return newRow;
        });

        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');

        const fileName = this.#elements.jsonFileInput.files[0].name.replace('.json', '.xlsx');
        XLSX.writeFile(workbook, fileName);
        this.#elements.messageDiv.textContent = 'Conversion successful! Downloading ' + fileName;
    }

    /**
     * Formats a given timestamp into a readable date string.
     * @param {number} timestamp - The timestamp to format (assumed to be in milliseconds).
     * @returns {string} The formatted date string.
     */
    #formatTimestampField(timestamp) {
        return moment(timestamp).format('YYYY-MM-DD HH:mm:ss');
    }

    /**
     * Recursively flattens nested objects and cleans data.
     * @param {object} obj - The object to flatten.
     * @param {string} [parentKey=''] - The key of the parent object.
     * @param {object} [result={}] - The object to store the flattened data.
     * @returns {object} The flattened and cleaned object.
     */
    #flattenAndCleanObject(obj, parentKey = '', result = {}) {
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                const newKey = parentKey ? `${parentKey}.${key}` : key;
                if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
                    this.#flattenAndCleanObject(obj[key], newKey, result);
                } else if (Array.isArray(obj[key])) {
                    // Handle arrays:
                    // If all elements are simple types, join them with ', '
                    // Otherwise, stringify the array to preserve structure for complex cases
                    const allSimple = obj[key].every(item => typeof item !== 'object' && !Array.isArray(item));
                    if (allSimple) {
                        result[newKey] = obj[key].join(', ');
                    } else {
                        result[newKey] = JSON.stringify(obj[key]);
                    }
                } else if (['createdOn', 'modifiedOn', 'dueDate', 'reportedTime', 'remainingTime'].includes(key) && typeof obj[key] === 'number' && obj[key] > 0) {
                    result[newKey] = this.#formatTimestampField(obj[key]);
                }
                else {
                    result[newKey] = obj[key];
                }
            }
        }
        return result;
    }

    /**
     * Creates a checkbox element for a given header.
     * @param {string} header - The header for which to create the checkbox.
     * @param {function} changeHandler - The function to call when the checkbox state changes.
     * @returns {HTMLLabelElement} The label element containing the checkbox and text.
     */
    #createFieldCheckbox(header, changeHandler) {
        const checkboxId = `field-${header}`;
        const label = document.createElement('label');
        label.setAttribute('for', checkboxId);

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = checkboxId;
        checkbox.name = 'field';
        checkbox.value = header;
        checkbox.checked = true; // All fields selected by default
        
        checkbox.addEventListener('change', changeHandler);

        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(header));
        return label;
    }

    /**
     * Populates the field selection checkboxes based on the data.
     * @param {Array<object>} data - The flattened and cleaned JSON data.
     */
    #populateFieldSelection(data) {
        if (data.length === 0) {
            this.#elements.fieldSelectionContainer.classList.add('hidden');
            return;
        }

        const allHeaders = Array.from(new Set(data.flatMap(Object.keys)));
        this.#elements.fieldCheckboxesDiv.innerHTML = ''; // Clear previous checkboxes

        allHeaders.forEach(header => {
            const changeHandler = () => {
                // Re-render table when selection changes
                const currentSelectedFields = Array.from(this.#elements.fieldCheckboxesDiv.querySelectorAll('input[type="checkbox"]:checked'))
                                                .map(cb => cb.value);
                const filteredData = this.#flattenedAndCleanedData.map(row => {
                    const newRow = {};
                    currentSelectedFields.forEach(field => {
                        newRow[field] = row[field];
                    });
                    return newRow;
                });
                this.#renderTable(filteredData);
            };
            this.#elements.fieldCheckboxesDiv.appendChild(this.#createFieldCheckbox(header, changeHandler));
        });

        this.#elements.fieldSelectionContainer.classList.remove('hidden'); // Show the field selection area
    }

    /**
     * Renders the table headers.
     * @param {string[]} headers - An array of header strings.
     */
    /**
     * Renders the table headers.
     * @param {string[]} headers - An array of header strings.
     */
    #renderTableHeaders(headers) {
        this.#elements.tableHeadersRow.innerHTML = ''; // Clear previous content
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            th.className = 'px-4 py-2 font-semibold text-gray-700 uppercase tracking-wider border-b border-gray-300'; // Tailwind classes for header
            this.#elements.tableHeadersRow.appendChild(th);
        });
    }

    /**
     * Renders the table body rows.
     * @param {Array<object>} data - The data to render.
     * @param {string[]} headers - An array of header strings to determine column order.
     */
    #renderTableRows(data, headers) {
        this.#elements.tableBody.innerHTML = ''; // Clear previous content
        data.forEach((rowData, index) => {
            const tr = document.createElement('tr');
            tr.className = 'hover:bg-gray-100 transition-colors duration-150 ease-in-out'; // Tailwind classes for row hover effect
            if (index % 2 === 1) { // Apply different background for odd rows for striping
                tr.classList.add('bg-gray-50');
            }
            headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = rowData[header] !== undefined ? rowData[header] : '';
                td.className = 'px-4 py-2 border-b border-gray-200 text-gray-800 whitespace-nowrap overflow-hidden text-ellipsis max-w-xs'; // Tailwind classes for cell
                tr.appendChild(td);
            });
            this.#elements.tableBody.appendChild(tr);
        });
    }

    /**
     * Renders data into the HTML table.
     * @param {Array<object>} data - The data to render.
     */
    #renderTable(data) {
        if (data.length === 0) {
            this.#elements.previewTableContainer.classList.add('hidden');
            return;
        }

        // Get all unique headers from all objects in the current data set
        // This ensures the table headers reflect only the selected/available fields
        const headers = Array.from(new Set(data.flatMap(Object.keys)));
        
        this.#renderTableHeaders(headers);
        this.#renderTableRows(data, headers);

        this.#elements.previewTableContainer.classList.remove('hidden');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new JsonToXlsxConverter();
});