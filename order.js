(function () {
    const form = document.getElementById("orderForm");
    const status = document.getElementById("formStatus");

    if (!form || !status) {
        return;
    }

    const STORAGE_KEY = "avan-candies-orders";
    const HEADERS = [
        "Name",
        "Email",
        "Class",
        "School",
        "Phone Number",
        "Candy Item",
        "Quantity",
        "Notes",
    ];

    const COLUMN_WIDTHS = [22, 28, 18, 28, 20, 22, 14, 28];

    form.addEventListener("submit", function (event) {
        event.preventDefault();

        const formData = new FormData(form);
        const order = {
            name: String(formData.get("name") || "").trim(),
            email: String(formData.get("email") || "").trim(),
            class: String(formData.get("class") || "").trim(),
            school: String(formData.get("school") || "").trim(),
            phone: String(formData.get("phone") || "").trim(),
            item: String(formData.get("item") || "").trim(),
            quantity: String(formData.get("quantity") || "").trim(),
            notes: String(formData.get("notes") || "").trim(),
        };

        if (!order.name || !order.email || !order.class || !order.school || !order.phone || !order.item || !order.quantity) {
            status.textContent = "Please complete all required fields before submitting.";
            status.dataset.state = "error";
            return;
        }

        const savedOrders = loadOrders();
        savedOrders.push(order);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(savedOrders));

        const workbookBlob = buildWorkbook(savedOrders);
        downloadBlob(workbookBlob, "A_VAN_CANDIES_Orders.xlsx");

        status.textContent = "Order saved and Excel file downloaded.";
        status.dataset.state = "success";
        form.reset();
        form.elements.quantity.value = "1";
    });

    function loadOrders() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            return raw ? JSON.parse(raw) : [];
        } catch (error) {
            return [];
        }
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 1000);
    }

    function buildWorkbook(orders) {
        const files = [
            {
                name: "[Content_Types].xml",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
                    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
                    '</Types>'),
            },
            {
                name: "_rels/.rels",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
                    '</Relationships>'),
            },
            {
                name: "docProps/core.xml",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                    '<dc:title>A VAN CANDIES Orders</dc:title>' +
                    '<dc:creator>A VAN CANDIES</dc:creator>' +
                    '<cp:lastModifiedBy>A VAN CANDIES</cp:lastModifiedBy>' +
                    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-04-25T00:00:00Z</dcterms:created>' +
                    '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-25T00:00:00Z</dcterms:modified>' +
                    '</cp:coreProperties>'),
            },
            {
                name: "docProps/app.xml",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
                    '<Application>Microsoft Excel</Application>' +
                    '</Properties>'),
            },
            {
                name: "xl/workbook.xml",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    '<sheets><sheet name="Orders" sheetId="1" r:id="rId1"/></sheets>' +
                    '</workbook>'),
            },
            {
                name: "xl/_rels/workbook.xml.rels",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                    '</Relationships>'),
            },
            {
                name: "xl/styles.xml",
                content: xml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
                    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
                    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
                    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
                    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
                    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
                    '</styleSheet>'),
            },
            {
                name: "xl/worksheets/sheet1.xml",
                content: xml(buildWorksheetXml(orders)),
            },
        ];

        return createZip(files);
    }

    function buildWorksheetXml(orders) {
        const rows = [];
        rows.push(buildRow(1, HEADERS));

        orders.forEach(function (order, index) {
            rows.push(buildRow(index + 2, [
                order.name,
                order.email,
                order.class,
                order.school,
                order.phone,
                order.item,
                order.quantity,
                order.notes,
            ]));
        });

        const cols = COLUMN_WIDTHS.map(function (width, index) {
            const column = index + 1;
            return '<col min="' + column + '" max="' + column + '" width="' + width + '" customWidth="1"/>';
        }).join("");

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="15"/>' +
            '<cols>' + cols + '</cols>' +
            '<sheetData>' + rows.join("") + '</sheetData>' +
            '</worksheet>';
    }

    function buildRow(rowNumber, values) {
        const cells = values.map(function (value, index) {
            const ref = columnName(index + 1) + rowNumber;
            return '<c r="' + ref + '" t="inlineStr"><is><t>' + escapeXml(String(value || "")) + '</t></is></c>';
        }).join("");

        return '<row r="' + rowNumber + '">' + cells + '</row>';
    }

    function columnName(number) {
        let name = "";
        let current = number;

        while (current > 0) {
            const remainder = (current - 1) % 26;
            name = String.fromCharCode(65 + remainder) + name;
            current = Math.floor((current - 1) / 26);
        }

        return name;
    }

    function escapeXml(value) {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    function xml(text) {
        return new TextEncoder().encode(text);
    }

    function createZip(files) {
        const localParts = [];
        const centralParts = [];
        let offset = 0;

        files.forEach(function (file) {
            const nameBytes = new TextEncoder().encode(file.name);
            const data = file.content;
            const crc = crc32(data);
            const localHeader = new Uint8Array(30 + nameBytes.length);
            const localView = new DataView(localHeader.buffer);

            localView.setUint32(0, 0x04034b50, true);
            localView.setUint16(4, 20, true);
            localView.setUint16(6, 0, true);
            localView.setUint16(8, 0, true);
            localView.setUint16(10, 0, true);
            localView.setUint16(12, 0, true);
            localView.setUint32(14, crc, true);
            localView.setUint32(18, data.length, true);
            localView.setUint32(22, data.length, true);
            localView.setUint16(26, nameBytes.length, true);
            localView.setUint16(28, 0, true);
            localHeader.set(nameBytes, 30);

            localParts.push(localHeader, data);

            const centralHeader = new Uint8Array(46 + nameBytes.length);
            const centralView = new DataView(centralHeader.buffer);
            centralView.setUint32(0, 0x02014b50, true);
            centralView.setUint16(4, 20, true);
            centralView.setUint16(6, 20, true);
            centralView.setUint16(8, 0, true);
            centralView.setUint16(10, 0, true);
            centralView.setUint16(12, 0, true);
            centralView.setUint16(14, 0, true);
            centralView.setUint32(16, crc, true);
            centralView.setUint32(20, data.length, true);
            centralView.setUint32(24, data.length, true);
            centralView.setUint16(28, nameBytes.length, true);
            centralView.setUint16(30, 0, true);
            centralView.setUint16(32, 0, true);
            centralView.setUint16(34, 0, true);
            centralView.setUint16(36, 0, true);
            centralView.setUint32(38, 0, true);
            centralView.setUint32(42, offset, true);
            centralHeader.set(nameBytes, 46);

            centralParts.push(centralHeader);
            offset += localHeader.length + data.length;
        });

        const centralSize = centralParts.reduce(function (total, part) {
            return total + part.length;
        }, 0);

        const endHeader = new Uint8Array(22);
        const endView = new DataView(endHeader.buffer);
        endView.setUint32(0, 0x06054b50, true);
        endView.setUint16(4, 0, true);
        endView.setUint16(6, 0, true);
        endView.setUint16(8, files.length, true);
        endView.setUint16(10, files.length, true);
        endView.setUint32(12, centralSize, true);
        endView.setUint32(16, offset, true);
        endView.setUint16(20, 0, true);

        return new Blob(localParts.concat(centralParts, [endHeader]), {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
    }

    function crc32(data) {
        let crc = 0 ^ -1;

        for (let i = 0; i < data.length; i += 1) {
            crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ data[i]) & 0xff];
        }

        return (crc ^ -1) >>> 0;
    }

    const CRC_TABLE = (function () {
        const table = [];

        for (let i = 0; i < 256; i += 1) {
            let c = i;
            for (let j = 0; j < 8; j += 1) {
                c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
            }
            table[i] = c >>> 0;
        }

        return table;
    }());
}());
