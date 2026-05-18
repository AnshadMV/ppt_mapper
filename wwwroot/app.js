const pointSections = [
    { textId: "page1-points", fieldPrefix: "Page1", bulletId: "page1-bullets", boldId: "page1-bold" },
    { textId: "page2-points", fieldPrefix: "Page2", bulletId: "page2-bullets", boldId: "page2-bold" },
    { textId: "page3-points", fieldPrefix: "Page3", bulletId: "page3-bullets", boldId: "page3-bold" },
    { textId: "page6-points", fieldPrefix: "Page6", bulletId: "page6-bullets", boldId: "page6-bold" },
    { textId: "page7-points", fieldPrefix: "Page7", bulletId: "page7-bullets", boldId: "page7-bold" }
];

const tableHeaders = [
    "Risk/Challenge Description",
    "Impact (H=3, M=2, L=1)",
    "Likelihood (H=3, M=2, L=1)",
    "Detection Rating (1-5)",
    "RPN",
    "Owner",
    "Mitigation Plan",
    "Contingency Plan",
    "Status"
];

const numericMetricIds = [
    "num-1", "num-2", "num-3", "num-4",
    "num-5", "num-6", "num-7", "num-8",
    "num-9", "num-10", "num-11", "num-12",
    "num-13", "num-14", "num-15", "num-16",
    "num-17", "num-19", "num-20", "num-21"
];

const percentageMetricIds = [
    "perc-1", "perc-2", "perc-3", "perc-4",
    "perc-5", "perc-6", "perc-7", "perc-8"
];

const nextSprintFields = [
    {
        valueId: "total-stories-committed",
        formField: "TotalStoriesCommitted",
        bulletId: "total-stories-committed-bullet",
        bulletField: "TotalStoriesCommittedIsBulletPoint",
        boldId: "total-stories-committed-bold",
        boldField: "TotalStoriesCommittedIsBold"
    },
    {
        valueId: "total-story-points",
        formField: "TotalStoryPoints",
        bulletId: "total-story-points-bullet",
        bulletField: "TotalStoryPointsIsBulletPoint",
        boldId: "total-story-points-bold",
        boldField: "TotalStoryPointsIsBold"
    },
    {
        valueId: "new-user-stories",
        formField: "NewUserStories",
        bulletId: "new-user-stories-bullet",
        bulletField: "NewUserStoriesIsBulletPoint",
        boldId: "new-user-stories-bold",
        boldField: "NewUserStoriesIsBold"
    },
    {
        valueId: "spillover-stories",
        formField: "SpilloverStories",
        bulletId: "spillover-stories-bullet",
        bulletField: "SpilloverStoriesIsBulletPoint",
        boldId: "spillover-stories-bold",
        boldField: "SpilloverStoriesIsBold"
    },
    {
        valueId: "agile-ceremony-process-item",
        formField: "AgileCeremonyProcessItem",
        bulletId: "agile-ceremony-process-item-bullet",
        bulletField: "AgileCeremonyProcessItemIsBulletPoint",
        boldId: "agile-ceremony-process-item-bold",
        boldField: "AgileCeremonyProcessItemIsBold"
    }
];

document.addEventListener("DOMContentLoaded", () => {
    renderRiskTable("page4-table", 1, 4, "page4");
    renderRiskTable("page5-table", 5, 4, "page5");
    renderImageInputRow();

    document.getElementById("ppt-form").addEventListener("submit", handleSubmit);
    document.getElementById("sample-fill").addEventListener("click", fillSampleData);
    document.getElementById("add-image-input").addEventListener("click", () => {
        renderImageInputRow();
        updateImageSelectionUi();
    });

    updateImageSelectionUi();
});

function renderRiskTable(containerId, startRowNumber, rowCount, tableKey) {
    const container = document.getElementById(containerId);
    const table = document.createElement("table");
    table.className = "risk-table";

    table.innerHTML = `
        <thead>
            <tr>
                <th>#</th>
                <th class="risk-table__desc">${tableHeaders[0]}</th>
                <th class="risk-table__small">${tableHeaders[1]}</th>
                <th class="risk-table__small">${tableHeaders[2]}</th>
                <th class="risk-table__small">${tableHeaders[3]}</th>
                <th class="risk-table__small">${tableHeaders[4]}</th>
                <th class="risk-table__owner">${tableHeaders[5]}</th>
                <th class="risk-table__plan">${tableHeaders[6]}</th>
                <th class="risk-table__plan">${tableHeaders[7]}</th>
                <th class="risk-table__small">${tableHeaders[8]}</th>
            </tr>
        </thead>
        <tbody></tbody>
    `;

    const tbody = table.querySelector("tbody");

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        const rowNumber = startRowNumber + rowIndex;
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${rowNumber}</td>
            ${tableHeaders.map((header, cellIndex) => `
                <td>
                    <textarea
                        rows="${cellIndex === 0 || cellIndex >= 6 ? 5 : 3}"
                        data-table="${tableKey}"
                        data-row="${rowIndex}"
                        data-cell="${cellIndex}"
                        placeholder="${header}"></textarea>
                </td>
            `).join("")}
        `;
        tbody.appendChild(row);
    }

    container.innerHTML = "";
    container.appendChild(table);
}

async function handleSubmit(event) {
    event.preventDefault();

    const submitButton = document.getElementById("submit-button");
    submitButton.disabled = true;
    setStatus("Generating PPTX... Please wait.", "working");

    try {
        const response = await fetch("/api/ppt/generate", {
            method: "POST",
            body: buildFormData()
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(errorText || `Request failed with status ${response.status}.`);
        }

        const blob = await response.blob();
        downloadBlob(blob, getDownloadFileName(response));
        setStatus("PPTX generated successfully. Download started.", "success");
    } catch (error) {
        setStatus(error.message || "Could not generate the PPTX.", "error");
    } finally {
        submitButton.disabled = false;
    }
}

function buildFormData() {
    const formData = new FormData();

    appendText(formData, "Title", getValue("title"));
    appendBoolean(formData, "TitleIsBold", isChecked("title-bold"));
    appendText(formData, "Date", getValue("date"));
    appendBoolean(formData, "DateIsBold", isChecked("date-bold"));
    appendText(formData, "SprintNumber", getValue("sprint-number"));
    appendBoolean(formData, "SprintNumberIsBold", isChecked("sprint-bold"));

    pointSections.forEach(section => {
        const values = parseLines(getValue(section.textId));
        const showAsBullet = isChecked(section.bulletId);
        const makeBold = isChecked(section.boldId);

        values.forEach(value => {
            appendText(formData, `${section.fieldPrefix}Points`, value);
            appendBoolean(formData, `${section.fieldPrefix}PointIsBulletPoint`, showAsBullet);
            appendBoolean(formData, `${section.fieldPrefix}PointIsBold`, makeBold);
        });
    });

    const page4Values = [...tableHeaders, ...collectTableValues("page4", 4)];
    page4Values.forEach(value => appendText(formData, "Page4Points", value));
    page4Values.forEach(() => appendBoolean(formData, "Page4PointIsBold", isChecked("page4-bold")));

    const page5Values = collectTableValues("page5", 4).slice(0, 35);
    page5Values.forEach(value => appendText(formData, "Page5Points", value));
    page5Values.forEach(() => appendBoolean(formData, "Page5PointIsBold", isChecked("page5-bold")));

    numericMetricIds.forEach(id => {
        appendText(formData, "NumericValues", getValue(id));
        appendBoolean(formData, "NumericValueIsBold", isChecked("numeric-bold"));
    });

    percentageMetricIds.forEach(id => {
        appendText(formData, "PercentageValues", getValue(id));
        appendBoolean(formData, "PercentageValueIsBold", isChecked("percentage-bold"));
    });

    nextSprintFields.forEach(field => {
        appendText(formData, field.formField, getValue(field.valueId));
        appendBoolean(formData, field.bulletField, isChecked(field.bulletId));
        appendBoolean(formData, field.boldField, isChecked(field.boldId));
    });

    const files = getSelectedImageFiles();
    for (const file of files) {
        formData.append("Images", file);
    }

    return formData;
}

function collectTableValues(tableKey, rowCount) {
    const values = [];

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        for (let cellIndex = 0; cellIndex < tableHeaders.length; cellIndex++) {
            const input = document.querySelector(`[data-table="${tableKey}"][data-row="${rowIndex}"][data-cell="${cellIndex}"]`);
            values.push(input ? input.value.trim() : "");
        }
    }

    return values;
}

function fillSampleData() {
    setValue("title", "Retrospective Report - Team Pegasus");
    setValue("date", "May 2026");
    setValue("sprint-number", "18");

    setValue("page1-points", [
        "Delivered high-priority items on schedule",
        "Improved collaboration with QA",
        "Faster review turnaround this sprint"
    ].join("\n"));

    setValue("page2-points", [
        "Team communication stayed clear",
        "Release planning was smoother",
        "Defect triage became more predictable"
    ].join("\n"));

    setValue("page3-points", [
        "A few dependencies were blocked externally",
        "Story splitting still needs improvement",
        "Some reporting tasks took longer than expected"
    ].join("\n"));

    fillTableRow("page4", 0, [
        "Delay in Nphies External Api",
        "3 (High)",
        "2 (Medium)",
        "4",
        "24",
        "Arunnath",
        "Currently available in 8443 env",
        "",
        "Open"
    ]);

    fillTableRow("page4", 1, [
        "PR delays due to ongoing OP panel and copay development, avoiding partial PRs, pending review comments, and ensuring NPHIES changes do not impact existing flows",
        "3 (High)",
        "3 (High)",
        "4",
        "36",
        "Viveka",
        "Plan PRs after dependent modules are completed; allocate time for resolving review comments; perform impact analysis before merging NPHIES changes",
        "",
        "Open"
    ]);

    setValue("page6-points", [
        "Refine story grooming before sprint start",
        "Reduce deployment wait time",
        "Create a shared checklist for release readiness"
    ].join("\n"));

    setValue("page7-points", [
        "Track blockers earlier in daily stand-up",
        "Review estimation variance weekly"
    ].join("\n"));

    setValue("num-13", "0");
    setValue("num-14", "305");

    setValue("num-1", "305");
    setValue("num-5", "281");
    setValue("num-9", "24");

    setValue("num-2", "87");
    setValue("num-6", "87");
    setValue("num-10", "0");

    setValue("num-3", "0");
    setValue("num-7", "0");
    setValue("num-11", "0");

    setValue("num-4", "4");
    setValue("num-8", "4");
    setValue("num-12", "0");

    setValue("num-15", "20");
    setValue("num-16", "20");
    setValue("num-17", "21");
    setValue("num-19", "3");
    setValue("num-20", "0");
    setValue("num-21", "10");

    setValue("perc-1", "66%");
    setValue("perc-2", "19%");
    setValue("perc-3", "0%");
    setValue("perc-4", "3%");
    setValue("perc-5", "0%");
    setValue("perc-6", "12%");
    setValue("perc-7", "464 hours(100%)");
    setValue("perc-8", "100");

    setValue("total-stories-committed", "24");
    setValue("total-story-points", "88");
    setValue("new-user-stories", "15");
    setValue("spillover-stories", "6");
    setValue("agile-ceremony-process-item", "3");

    setStatus("Sample data loaded. You can edit any field before generating.", "idle");
}

function updateImageSelectionUi() {
    const status = document.getElementById("image-upload-status");
    const list = document.getElementById("image-upload-list");
    const files = getSelectedImageFiles();

    list.innerHTML = "";

    if (files.length === 0) {
        status.textContent = "No images selected.";
        return;
    }

    status.textContent = `${files.length} image${files.length === 1 ? "" : "s"} selected for upload.`;

    files.forEach((file, index) => {
        const item = document.createElement("div");
        item.className = "upload-item";
        item.innerHTML = `
            <span class="upload-item__name">${index + 1}. ${escapeHtml(file.name)}</span>
            <span class="upload-item__meta">${formatFileSize(file.size)}</span>
        `;
        list.appendChild(item);
    });
}

function renderImageInputRow() {
    const container = document.getElementById("image-input-list");
    const index = container.children.length;
    const row = document.createElement("div");
    row.className = "image-input-row";
    row.innerHTML = `
        <input type="file" accept="image/*" data-image-input="true" aria-label="Image ${index + 1}" />
        <button type="button" class="image-remove" aria-label="Remove image field">-</button>
    `;

    const input = row.querySelector('[data-image-input="true"]');
    const removeButton = row.querySelector(".image-remove");

    input.addEventListener("change", updateImageSelectionUi);
    removeButton.addEventListener("click", () => {
        if (container.children.length === 1) {
            input.value = "";
        } else {
            row.remove();
        }

        updateImageSelectionUi();
    });

    container.appendChild(row);
}

function getSelectedImageFiles() {
    return Array.from(document.querySelectorAll('[data-image-input="true"]'))
        .map(input => input.files && input.files[0] ? input.files[0] : null)
        .filter(file => file !== null);
}

function fillTableRow(tableKey, rowIndex, values) {
    values.forEach((value, cellIndex) => {
        const input = document.querySelector(`[data-table="${tableKey}"][data-row="${rowIndex}"][data-cell="${cellIndex}"]`);
        if (input) {
            input.value = value;
        }
    });
}

function parseLines(value) {
    return value
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line.length > 0);
}

function appendText(formData, fieldName, value) {
    formData.append(fieldName, value ?? "");
}

function appendBoolean(formData, fieldName, value) {
    formData.append(fieldName, value ? "true" : "false");
}

function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
}

function getDownloadFileName(response) {
    const contentDisposition = response.headers.get("Content-Disposition") || response.headers.get("content-disposition");
    if (!contentDisposition) {
        return "Pegasus_sprint_retro.pptx";
    }

    const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match?.[1]) {
        return decodeURIComponent(utf8Match[1]);
    }

    const basicMatch = contentDisposition.match(/filename="?([^"]+)"?/i);
    if (basicMatch?.[1]) {
        return basicMatch[1];
    }

    return "Pegasus_sprint_retro.pptx";
}

function setStatus(message, state) {
    const banner = document.getElementById("status-banner");
    banner.textContent = message;
    banner.className = `status status-${state}`;
}

function getValue(id) {
    const element = document.getElementById(id);
    if (!element) {
        return "";
    }

    if ("value" in element) {
        return element.value;
    }

    return element.innerText ?? "";
}

function setValue(id, value) {
    const element = document.getElementById(id);
    if (!element) {
        return;
    }

    if ("value" in element) {
        element.value = value;
        return;
    }

    element.innerText = value;
}

function isChecked(id) {
    const element = document.getElementById(id);
    return !!element?.checked;
}

function formatFileSize(bytes) {
    if (!Number.isFinite(bytes) || bytes < 1024) {
        return `${bytes || 0} B`;
    }

    const kb = bytes / 1024;
    if (kb < 1024) {
        return `${kb.toFixed(1)} KB`;
    }

    return `${(kb / 1024).toFixed(1)} MB`;
}

function escapeHtml(value) {
    return (value || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}
