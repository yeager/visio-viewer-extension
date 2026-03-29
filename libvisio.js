/**
 * Complete VSDX to SVG converter - JavaScript port of libvisio-ng Python converter
 * Author: Daniel Nylander <daniel@danielnylander.se>
 * Port maintains exact same rendering quality and API as Python version
 * 
 * Requirements: JSZip for ZIP archive handling
 */

class VisioConverter {
    constructor() {
        this.pages = [];
        this.pageNames = [];
        this.masters = {};
        this.media = {};
        this.themeColors = {};
        this.zipFile = null;
        this.isNodeJs = typeof window === 'undefined' || typeof window.document === 'undefined';
        
        // Constants from Python version
        this.VISIO_COLORS = {
            0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00", 4: "#0000FF",
            5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF", 8: "#800000", 9: "#008000",
            10: "#000080", 11: "#808000", 12: "#800080", 13: "#008080",
            14: "#C0C0C0", 15: "#808080", 16: "#993366", 17: "#333399",
            18: "#333333", 19: "#003300", 20: "#003366", 21: "#993300",
            22: "#993366", 23: "#333399", 24: "#E6E6E6"
        };
        
        this.LINE_PATTERNS = {
            0: "none", 1: "", 2: "4,3", 3: "1,3", 4: "4,3,1,3",
            5: "4,3,1,3,1,3", 6: "8,3", 7: "1,1", 8: "8,3,1,3",
            9: "8,3,1,3,1,3", 10: "12,6", 16: "6,3,6,3"
        };
        
        this.ARROW_SIZES = {0: 0.6, 1: 0.8, 2: 1.0, 3: 1.2, 4: 1.6, 5: 2.0, 6: 2.5};
        
        this.INCH_TO_PX = 72.0;
        
        // XML namespaces
        this.NS = {
            v: "http://schemas.microsoft.com/office/visio/2012/main",
            r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        };
        this.VNS = this.NS.v;
        this.VTAG = "{" + this.VNS + "}";
        this.RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
        
        // Image MIME types
        this.IMAGE_MIMETYPES = {
            ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".gif": "image/gif", ".bmp": "image/bmp", ".emf": "image/x-emf",
            ".wmf": "image/x-wmf", ".tiff": "image/tiff", ".tif": "image/tiff",
            ".svg": "image/svg+xml"
        };
        
        // Font mapping
        this.FONT_MAP = {
            "angsana new": "Noto Sans Thai, Noto Serif Thai, sans-serif",
            "browallia new": "Noto Sans Thai, sans-serif",
            "cordia new": "Noto Sans Thai, sans-serif",
            "freesia upc": "Noto Sans Thai, sans-serif",
            "tahoma": "Tahoma, Noto Sans, sans-serif",
            "arial": "Arial, Noto Sans, sans-serif",
            "calibri": "Calibri, Noto Sans, sans-serif",
            "segoe ui": "Segoe UI, Noto Sans, sans-serif",
            "times new roman": "Times New Roman, Noto Serif, serif",
            "ms gothic": "Noto Sans JP, sans-serif",
            "ms mincho": "Noto Serif JP, serif",
            "simsun": "Noto Sans SC, sans-serif",
            "simhei": "Noto Sans SC, sans-serif",
            "microsoft yahei": "Noto Sans SC, sans-serif",
            "malgun gothic": "Noto Sans KR, sans-serif",
            "gulim": "Noto Sans KR, sans-serif"
        };
    }

    /**
     * Parse XML with correct DOM methods for Node.js vs Browser
     */
    parseXml(xmlText) {
        const parser = new DOMParser();
        return parser.parseFromString(xmlText, "text/xml");
    }

    /**
     * Query all elements - works with both browser and xmldom
     */
    querySelectorAll(element, selector) {
        if (this.isNodeJs || !element.querySelectorAll) {
            // For xmldom, use getElementsByTagNameNS or getElementsByTagName
            if (selector === "Shape") {
                return this.getElementsByTagNameRecursive(element, "Shape");
            }
            if (selector === "Cell") {
                return this.getElementsByTagNameRecursive(element, "Cell");
            }
            if (selector === "Section") {
                return this.getElementsByTagNameRecursive(element, "Section");
            }
            if (selector === "Row") {
                return this.getElementsByTagNameRecursive(element, "Row");
            }
            if (selector === "Master") {
                return this.getElementsByTagNameRecursive(element, "Master");
            }
            if (selector === "Relationship") {
                return this.getElementsByTagNameRecursive(element, "Relationship");
            }
            if (selector === "Text") {
                return this.getElementsByTagNameRecursive(element, "Text");
            }
            if (selector === "Shapes") {
                return this.getElementsByTagNameRecursive(element, "Shapes");
            }
            if (selector === "ForeignData") {
                return this.getElementsByTagNameRecursive(element, "ForeignData");
            }
            if (selector === "Rel") {
                return this.getElementsByTagNameRecursive(element, "Rel");
            }
            if (selector === "Connect") {
                return this.getElementsByTagNameRecursive(element, "Connect");
            }
            if (selector === "PageSheet") {
                return this.getElementsByTagNameRecursive(element, "PageSheet");
            }
            if (selector === "Connects") {
                return this.getElementsByTagNameRecursive(element, "Connects");
            }
            // Add more as needed
            return [];
        } else {
            return element.querySelectorAll(selector);
        }
    }

    /**
     * Query single element - works with both browser and xmldom
     */
    querySelector(element, selector) {
        if (this.isNodeJs || !element.querySelector) {
            // For xmldom, use getElementsByTagNameNS or getElementsByTagName
            const results = this.querySelectorAll(element, selector);
            return results.length > 0 ? results[0] : null;
        } else {
            return element.querySelector(selector);
        }
    }

    /**
     * Get elements by tag name recursively (for xmldom)
     */
    getElementsByTagNameRecursive(element, tagName) {
        const results = [];
        
        if (element.getElementsByTagName) {
            const nodeList = element.getElementsByTagName(tagName);
            for (let i = 0; i < nodeList.length; i++) {
                results.push(nodeList[i]);
            }
        } else if (element.childNodes) {
            // Manual traversal for xmldom
            this.traverseForTag(element, tagName, results);
        }
        
        return results;
    }

    /**
     * Traverse DOM tree manually to find elements
     */
    traverseForTag(node, tagName, results) {
        if (node.nodeType === 1 && node.tagName === tagName) {
            results.push(node);
        }
        
        if (node.childNodes) {
            for (let i = 0; i < node.childNodes.length; i++) {
                this.traverseForTag(node.childNodes[i], tagName, results);
            }
        }
    }

    /**
     * Get direct children elements with specific tag name
     */
    getDirectChildren(element, tagName) {
        const results = [];
        if (element.childNodes) {
            for (let i = 0; i < element.childNodes.length; i++) {
                const child = element.childNodes[i];
                if (child.nodeType === 1 && child.tagName === tagName) {
                    results.push(child);
                }
            }
        }
        return results;
    }

    /**
     * Load VSDX file from ArrayBuffer
     */
    async loadFromArrayBuffer(buffer) {
        this.zipFile = await JSZip.loadAsync(buffer);
        
        // Parse file structure in correct order
        this.themeColors = await this.parseTheme();
        this.masters = await this.parseMasterShapes();
        this.media = await this.extractMedia();
        
        // Parse pages
        const pageFiles = this.getPageFiles();
        this.pages = [];
        this.pageNames = [];
        
        for (let i = 0; i < pageFiles.length; i++) {
            const pageFile = pageFiles[i];
            try {
                const pageXml = await this.zipFile.file(pageFile).async("text");
                const parser = new DOMParser();
                const doc = parser.parseFromString(pageXml, "text/xml");
                
                const shapes = this.parseVsdxShapes(pageXml);
                const connects = this.parseConnects(doc);
                const layers = this.parseLayers(doc);
                const pageRels = await this.parseRels(pageFile);
                const { width, height } = this.parsePageDimensions(pageXml);
                
                this.pages.push({
                    shapes,
                    connects,
                    layers,
                    pageRels,
                    width,
                    height,
                    file: pageFile
                });
                
                // Extract page name
                this.pageNames.push(this.extractPageName(pageXml) || `Page ${i + 1}`);
            } catch (error) {
                console.warn(`Failed to parse page ${i}:`, error);
            }
        }
    }

    /**
     * Get list of pages
     */
    getPages() {
        return this.pageNames.map((name, index) => ({
            index,
            name,
            width: this.pages[index]?.width || 8.5,
            height: this.pages[index]?.height || 11.0
        }));
    }

    /**
     * Render specific page as SVG string
     */
    renderPage(pageIndex) {
        if (pageIndex < 0 || pageIndex >= this.pages.length) {
            throw new Error(`Page index ${pageIndex} out of range`);
        }
        
        const page = this.pages[pageIndex];
        return this.shapesToSvg(
            page.shapes, 
            page.width, 
            page.height,
            page.connects,
            page.layers,
            page.pageRels
        );
    }

    /**
     * Parse page files from ZIP structure
     */
    getPageFiles() {
        const files = [];
        Object.keys(this.zipFile.files).forEach(name => {
            if (name.match(/^visio\/pages\/page\d+\.xml$/)) {
                files.push(name);
            }
        });
        return files.sort();
    }

    /**
     * Parse master shapes (complete implementation)
     */
    async parseMasterShapes() {
        const masters = {};
        
        // First read masters.xml to map Master ID -> rel ID
        let masterIdToFile = {};
        try {
            const mastersXml = await this.zipFile.file("visio/masters/masters.xml").async("text");
            const root = this.parseXml(mastersXml);
            
            // Parse rels to map rId -> filename
            const ridToFile = {};
            try {
                const relsXml = await this.zipFile.file("visio/masters/_rels/masters.xml.rels").async("text");
                const relsRoot = this.parseXml(relsXml);
                const rels = this.querySelectorAll(relsRoot, "Relationship");
                for (const rel of rels) {
                    const rid = rel.getAttribute("Id");
                    const target = rel.getAttribute("Target");
                    if (rid && target) {
                        const fname = target.replace("master", "").replace(".xml", "");
                        ridToFile[rid] = fname;
                    }
                }
            } catch (e) {}
            
            const masterEls = this.querySelectorAll(root, "Master");
            for (const masterEl of masterEls) {
                const mid = masterEl.getAttribute("ID");
                const relEl = this.querySelector(masterEl, "Rel");
                if (relEl) {
                    const rid = relEl.getAttribute("r:id") || relEl.getAttributeNS(this.NS.r, "id");
                    if (rid && ridToFile[rid]) {
                        masterIdToFile[mid] = ridToFile[rid];
                        continue;
                    }
                }
                // Fallback: assume master ID matches file number
                masterIdToFile[mid] = mid;
            }
        } catch (e) {}
        
        // Parse all master files
        const fileToShapes = {};
        for (const name of Object.keys(this.zipFile.files)) {
            if (!name.match(/^visio\/masters\/master\d+\.xml$/)) continue;
            if (name.includes("masters.xml")) continue;
            
            const masterNum = name.match(/master(\d+)/)?.[1];
            if (!masterNum) continue;
            
            try {
                const xml = await this.zipFile.file(name).async("text");
                const root = this.parseXml(xml);
                
                const shapesData = {};
                const shapes = this.querySelectorAll(root, "Shape");
                for (const shape of shapes) {
                    const sd = this.parseSingleShape(shape);
                    shapesData[sd.id] = sd;
                }
                
                if (Object.keys(shapesData).length > 0) {
                    fileToShapes[masterNum] = shapesData;
                }
            } catch (e) {
                console.warn(`Failed to parse master ${name}:`, e);
            }
        }
        
        // Re-key by Master ID
        for (const [mid, fnum] of Object.entries(masterIdToFile)) {
            if (fileToShapes[fnum]) {
                masters[mid] = fileToShapes[fnum];
            }
        }
        
        // Add unmapped files
        const mappedFiles = new Set(Object.values(masterIdToFile));
        for (const [fnum, shapesData] of Object.entries(fileToShapes)) {
            if (!mappedFiles.has(fnum)) {
                masters[fnum] = shapesData;
            }
        }
        
        return masters;
    }

    /**
     * Parse single shape element into complete shape dict
     */
    parseSingleShape(shapeElem) {
        const sd = {
            id: shapeElem.getAttribute("ID") || "",
            name: shapeElem.getAttribute("Name") || "",
            name_u: shapeElem.getAttribute("NameU") || "",
            type: shapeElem.getAttribute("Type") || "Shape",
            master: shapeElem.getAttribute("Master") || "",
            master_shape: shapeElem.getAttribute("MasterShape") || "",
            line_style: shapeElem.getAttribute("LineStyle") || "",
            fill_style: shapeElem.getAttribute("FillStyle") || "",
            text_style: shapeElem.getAttribute("TextStyle") || "",
            cells: {},
            geometry: [],
            text: "",
            text_parts: [],
            char_formats: {},
            para_formats: {},
            sub_shapes: [],
            controls: {},
            connections: {},
            user: {},
            foreign_data: null,
            hyperlinks: []
        };
        
        // Parse top-level cells
        const cells = this.getDirectChildren(shapeElem, "Cell");
        for (const cell of cells) {
            const n = cell.getAttribute("N");
            const v = cell.getAttribute("V") || "";
            const f = cell.getAttribute("F") || "";
            if (n) {
                sd.cells[n] = { V: v, F: f };
            }
        }
        
        // Parse sections
        const sections = this.getDirectChildren(shapeElem, "Section");
        for (const section of sections) {
            const secName = section.getAttribute("N");
            
            if (secName === "Geometry") {
                const geo = this.parseGeometrySection(section);
                if (geo) sd.geometry.push(geo);
            } else if (secName === "Controls") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const rowIx = row.getAttribute("IX") || "0";
                    const ctrl = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        ctrl[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    sd.controls[`Row_${rowIx}`] = ctrl;
                }
            } else if (secName === "User") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const rowName = row.getAttribute("N");
                    const userVals = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        userVals[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    sd.user[rowName] = userVals;
                }
            } else if (secName === "Connection") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const rowIx = row.getAttribute("IX") || "0";
                    const conn = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        conn[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    sd.connections[rowIx] = conn;
                }
            } else if (secName === "Character") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const rowIx = row.getAttribute("IX") || "0";
                    const fmt = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        fmt[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    sd.char_formats[rowIx] = fmt;
                }
            } else if (secName === "Paragraph") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const rowIx = row.getAttribute("IX") || "0";
                    const fmt = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        fmt[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    sd.para_formats[rowIx] = fmt;
                }
            } else if (secName === "FillGradientDef") {
                const gradStops = [];
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const stopCells = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        stopCells[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    const pos = this.safeFloat(stopCells.GradientStopPosition || "0");
                    const color = stopCells.GradientStopColor || "";
                    if (color) {
                        gradStops.push([pos * 100, color]);
                    }
                }
                if (gradStops.length > 0) {
                    sd._gradient_stops = [gradStops];
                }
            }
        }
        
        // Parse text
        const textElem = this.querySelector(shapeElem, "Text");
        if (textElem) {
            sd.text = (textElem.textContent || textElem.innerText || "").trim();
            sd.text_parts = this.parseTextElement(textElem);
            sd._has_text_elem = true;
        }
        
        // Parse sub-shapes (groups)
        const shapesContainer = this.querySelector(shapeElem, "Shapes");
        if (shapesContainer) {
            const subShapes = this.querySelectorAll(shapesContainer, "Shape");
            for (const subShape of subShapes) {
                sd.sub_shapes.push(this.parseSingleShape(subShape));
            }
        }
        
        // Parse ForeignData (embedded images)
        const fdInfo = this.parseForeignData(shapeElem);
        if (fdInfo) sd.foreign_data = fdInfo;
        
        // Parse hyperlinks
        const allSections = this.querySelectorAll(shapeElem, "Section");
        for (const section of allSections) {
            if (section.getAttribute("N") === "Hyperlink") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const link = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        const n = cell.getAttribute("N");
                        const v = cell.getAttribute("V") || "";
                        if (n === "Description") link.description = v;
                        else if (n === "Address") link.address = v;
                        else if (n === "SubAddress") link.sub_address = v;
                        else if (n === "Frame") link.frame = v;
                    }
                    if (Object.keys(link).length > 0) {
                        sd.hyperlinks.push(link);
                    }
                }
            }
        }
        
        return sd;
    }

    /**
     * Parse geometry section
     */
    parseGeometrySection(section) {
        const geo = { rows: [], no_fill: false, no_line: false, no_show: false };
        
        // Check section-level cells
        const sectionCells = this.getDirectChildren(section, "Cell");
        for (const cell of sectionCells) {
            const n = cell.getAttribute("N");
            const v = cell.getAttribute("V") || "0";
            if (n === "NoFill" && v === "1") geo.no_fill = true;
            else if (n === "NoLine" && v === "1") geo.no_line = true;
            else if (n === "NoShow" && v === "1") geo.no_show = true;
        }
        
        const rows = this.querySelectorAll(section, "Row");
        for (const row of rows) {
            const rowType = row.getAttribute("T") || "";
            const rowIx = row.getAttribute("IX") || "";
            const rowData = { type: rowType, cells: {}, ix: rowIx };
            
            const cells = this.querySelectorAll(row, "Cell");
            for (const cell of cells) {
                const n = cell.getAttribute("N");
                const v = cell.getAttribute("V") || "";
                const f = cell.getAttribute("F") || "";
                if (n) {
                    rowData.cells[n] = { V: v, F: f };
                }
            }
            
            geo.rows.push(rowData);
        }
        
        geo.ix = section.getAttribute("IX") || "0";
        return geo;
    }

    /**
     * Parse text element with formatting
     */
    parseTextElement(textElem) {
        const parts = [];
        let currentCp = "0";
        let currentPp = "0";
        
        if (textElem.textContent) {
            parts.push({ text: textElem.textContent, cp: currentCp, pp: currentPp });
        }
        
        const children = Array.from(textElem.childNodes);
        for (const child of children) {
            if (child.nodeType === 1) { // Element node
                const tag = child.tagName;
                if (tag === "cp") {
                    currentCp = child.getAttribute("IX") || "0";
                } else if (tag === "pp") {
                    currentPp = child.getAttribute("IX") || "0";
                } else if (tag === "fld") {
                    const fieldText = child.textContent?.trim() || "";
                    if (fieldText) {
                        parts.push({ text: fieldText, cp: currentCp, pp: currentPp });
                    }
                }
            }
            if (child.textContent) {
                parts.push({ text: child.textContent, cp: currentCp, pp: currentPp });
            }
        }
        
        return parts;
    }

    /**
     * Parse ForeignData element
     */
    parseForeignData(shapeElem) {
        const fd = this.querySelector(shapeElem, "ForeignData");
        if (!fd) return null;
        
        const info = {
            foreign_type: fd.getAttribute("ForeignType") || "",
            compression: fd.getAttribute("CompressionType") || "",
            data: null,
            rel_id: null
        };
        
        // Check for Rel element
        let relElem = this.querySelector(fd, "Rel");
        if (relElem) {
            info.rel_id = relElem.getAttribute("r:id") || 
                         relElem.getAttributeNS(this.NS.r, "id") || "";
        } else {
            // Inline data
            const text = (fd.textContent || fd.innerText || "");
            if (text.trim()) {
                info.data = text.trim();
            }
        }
        
        return info;
    }

    /**
     * Parse theme colors from theme1.xml
     */
    async parseTheme() {
        const themeColors = {};
        const DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        
        for (const themeFile of ["visio/theme/theme1.xml", "visio/theme/theme2.xml"]) {
            try {
                const themeXml = await this.zipFile.file(themeFile).async("text");
                const parser = new DOMParser();
                const doc = parser.parseFromString(themeXml, "text/xml");
                
                const colorSchemes = doc.getElementsByTagNameNS(DML_NS, "clrScheme");
                for (const clrScheme of colorSchemes) {
                    const colorNames = [
                        "dk1", "lt1", "dk2", "lt2",
                        "accent1", "accent2", "accent3", "accent4",
                        "accent5", "accent6", "hlink", "folHlink"
                    ];
                    
                    for (const cname of colorNames) {
                        const elem = clrScheme.querySelector(`*[localName="${cname}"]`);
                        if (!elem) continue;
                        
                        let baseColor = null;
                        const srgb = elem.querySelector(`*[localName="srgbClr"]`);
                        if (srgb) {
                            const val = srgb.getAttribute("val");
                            if (val) {
                                baseColor = "#" + val;
                                baseColor = this.applyColorTransforms(srgb, baseColor, DML_NS);
                            }
                        } else {
                            const sysClr = elem.querySelector(`*[localName="sysClr"]`);
                            if (sysClr) {
                                const val = sysClr.getAttribute("lastClr") || sysClr.getAttribute("val");
                                if (val && val.length === 6) {
                                    baseColor = "#" + val;
                                    baseColor = this.applyColorTransforms(sysClr, baseColor, DML_NS);
                                }
                            }
                        }
                        
                        if (baseColor) {
                            themeColors[cname] = baseColor;
                        }
                    }
                    break; // Only use first clrScheme
                }
                
                if (Object.keys(themeColors).length > 0) break;
            } catch (e) {}
        }
        
        // Build numeric index mapping
        const idxMap = {
            0: "dk1", 1: "lt1", 2: "dk2", 3: "lt2",
            4: "accent1", 5: "accent2", 6: "accent3", 7: "accent4",
            8: "accent5", 9: "accent6", 10: "hlink", 11: "folHlink"
        };
        
        for (const [idx, name] of Object.entries(idxMap)) {
            if (themeColors[name]) {
                themeColors[idx] = themeColors[name];
            }
        }
        
        return themeColors;
    }

    /**
     * Apply DML color transforms (tint, shade, etc.)
     */
    applyColorTransforms(elem, baseColor, dmlNs) {
        if (!baseColor?.startsWith("#") || baseColor.length !== 7) {
            return baseColor;
        }
        
        let r = parseInt(baseColor.substring(1, 3), 16);
        let g = parseInt(baseColor.substring(3, 5), 16);
        let b = parseInt(baseColor.substring(5, 7), 16);
        
        const children = elem.children;
        for (const child of children) {
            const tag = child.localName;
            const val = parseInt(child.getAttribute("val") || "0");
            const pct = val / 100000.0;
            
            if (tag === "tint") {
                r = Math.round(r + (255 - r) * pct);
                g = Math.round(g + (255 - g) * pct);
                b = Math.round(b + (255 - b) * pct);
            } else if (tag === "shade") {
                r = Math.round(r * pct);
                g = Math.round(g * pct);
                b = Math.round(b * pct);
            } else if (tag === "lumMod") {
                r = Math.min(255, Math.round(r * pct));
                g = Math.min(255, Math.round(g * pct));
                b = Math.min(255, Math.round(b * pct));
            } else if (tag === "lumOff") {
                const off = Math.round(255 * pct);
                r = Math.min(255, Math.max(0, r + off));
                g = Math.min(255, Math.max(0, g + off));
                b = Math.min(255, Math.max(0, b + off));
            }
        }
        
        r = Math.max(0, Math.min(255, r));
        g = Math.max(0, Math.min(255, g));
        b = Math.max(0, Math.min(255, b));
        
        return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
    }

    /**
     * Extract media files
     */
    async extractMedia() {
        const media = {};
        
        for (const name of Object.keys(this.zipFile.files)) {
            if (name.startsWith("visio/media/")) {
                const fname = name.split("/").pop();
                if (fname) {
                    try {
                        const data = await this.zipFile.file(name).async("arraybuffer");
                        media[fname] = new Uint8Array(data);
                    } catch (e) {}
                }
            }
        }
        
        return media;
    }

    /**
     * Parse page dimensions
     */
    parsePageDimensions(pageXml) {
        const doc = this.parseXml(pageXml);
        
        let width = 8.5, height = 11.0;
        
        const pageSheet = this.querySelector(doc, "PageSheet");
        if (pageSheet) {
            const cells = this.querySelectorAll(pageSheet, "Cell");
            for (const cell of cells) {
                const name = cell.getAttribute("N");
                const value = this.safeFloat(cell.getAttribute("V"));
                if (name === "PageWidth") width = value || 8.5;
                if (name === "PageHeight") height = value || 11.0;
            }
        }
        
        return { width, height };
    }

    /**
     * Extract page name
     */
    extractPageName(pageXml) {
        const doc = this.parseXml(pageXml);
        
        const pageSheet = this.querySelector(doc, "PageSheet");
        if (pageSheet) {
            const cells = this.querySelectorAll(pageSheet, "Cell");
            for (const cell of cells) {
                if (cell.getAttribute("N") === "PageName") {
                    return cell.getAttribute("V") || "";
                }
            }
        }
        
        return null;
    }

    /**
     * Parse shapes from page XML (complete implementation)
     */
    parseVsdxShapes(pageXml) {
        const shapes = [];
        const doc = this.parseXml(pageXml);
        
        const shapesContainer = this.querySelector(doc, "Shapes");
        if (!shapesContainer) return shapes;
        
        const shapeNodes = this.querySelectorAll(shapesContainer, "Shape");
        for (const node of shapeNodes) {
            // Only get direct children of Shapes container
            if (node.parentNode === shapesContainer) {
                const shape = this.parseSingleShape(node);
                if (shape) shapes.push(shape);
            }
        }
        
        return shapes;
    }

    /**
     * Parse connections
     */
    parseConnects(pageXmlRoot) {
        const connects = [];
        const connectsEl = this.querySelector(pageXmlRoot, "Connects");
        if (!connectsEl) return connects;
        
        const connectNodes = this.querySelectorAll(connectsEl, "Connect");
        for (const c of connectNodes) {
            connects.push({
                from_sheet: c.getAttribute("FromSheet") || "",
                from_cell: c.getAttribute("FromCell") || "",
                to_sheet: c.getAttribute("ToSheet") || "",
                to_cell: c.getAttribute("ToCell") || ""
            });
        }
        
        return connects;
    }

    /**
     * Parse layers
     */
    parseLayers(pageXmlRoot) {
        const layers = {};
        const pageSheet = this.querySelector(pageXmlRoot, "PageSheet");
        if (!pageSheet) return layers;
        
        const sections = this.querySelectorAll(pageSheet, "Section");
        for (const section of sections) {
            if (section.getAttribute("N") === "Layer") {
                const rows = this.querySelectorAll(section, "Row");
                for (const row of rows) {
                    const ix = row.getAttribute("IX") || "";
                    const cells = {};
                    const rowCells = this.querySelectorAll(row, "Cell");
                    for (const cell of rowCells) {
                        cells[cell.getAttribute("N")] = cell.getAttribute("V") || "";
                    }
                    const visible = cells.Visible !== "0";
                    const name = cells.Name || `Layer ${ix}`;
                    layers[ix] = { name, visible };
                }
            }
        }
        
        return layers;
    }

    /**
     * Parse relationships
     */
    async parseRels(pageFile) {
        const pageDir = pageFile.substring(0, pageFile.lastIndexOf("/"));
        const pageBasename = pageFile.substring(pageFile.lastIndexOf("/") + 1);
        const relsPath = `${pageDir}/_rels/${pageBasename}.rels`;
        
        const rels = {};
        try {
            const relsXml = await this.zipFile.file(relsPath).async("text");
            const root = this.parseXml(relsXml);
            const relNodes = this.querySelectorAll(root, "Relationship");
            for (const rel of relNodes) {
                const rid = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                if (rid && target) {
                    rels[rid] = target;
                }
            }
        } catch (e) {}
        
        return rels;
    }

    /**
     * Convert shapes to SVG
     */
    shapesToSvg(shapes, pageW, pageH, connects = null, layers = null, pageRels = null) {
        const pageWPx = pageW * this.INCH_TO_PX;
        const pageHPx = pageH * this.INCH_TO_PX;
        
        // Calculate content bounds for optimal viewBox
        let vbX = 0, vbY = 0, vbW = pageWPx, vbH = pageHPx;
        const maxSvgPx = 4000.0;
        
        if (shapes.length > 0) {
            let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
            
            for (const s of shapes) {
                const px = this.safeFloat(s.cells.PinX?.V) * this.INCH_TO_PX;
                const py = (pageH - this.safeFloat(s.cells.PinY?.V)) * this.INCH_TO_PX;
                const sw = Math.abs(this.safeFloat(s.cells.Width?.V)) * this.INCH_TO_PX;
                const sh = Math.abs(this.safeFloat(s.cells.Height?.V)) * this.INCH_TO_PX;
                
                if (px > 0 || py > 0) {
                    minX = Math.min(minX, px - sw / 2);
                    minY = Math.min(minY, py - sh / 2);
                    maxX = Math.max(maxX, px + sw / 2);
                    maxY = Math.max(maxY, py + sh / 2);
                }
            }
            
            if (minX < Infinity) {
                const contentW = maxX - minX;
                const contentH = maxY - minY;
                const padX = Math.max(50, contentW * 0.08);
                const padY = Math.max(50, contentH * 0.08);
                
                vbX = Math.min(0, minX - padX);
                vbY = Math.min(0, minY - padY);
                vbW = Math.max(contentW + 2 * padX, pageWPx);
                vbH = Math.max(contentH + 2 * padY, pageHPx);
            }
        }
        
        let displayW = vbW;
        let displayH = vbH;
        if (Math.max(vbW, vbH) > maxSvgPx) {
            const scale = maxSvgPx / Math.max(vbW, vbH);
            displayW = vbW * scale;
            displayH = vbH * scale;
        }
        
        const usedMarkers = new Set();
        const gradients = {};
        const hasShadow = new Set();
        const textLayer = [];
        
        const svgLines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<svg xmlns="http://www.w3.org/2000/svg" ',
            'xmlns:xlink="http://www.w3.org/1999/xlink" ',
            `width="${Math.round(displayW)}" height="${Math.round(displayH)}" `,
            `viewBox="${Math.round(vbX)} ${Math.round(vbY)} ${Math.round(vbW)} ${Math.round(vbH)}">`,
            `<rect x="${Math.round(vbX)}" y="${Math.round(vbY)}" width="${Math.round(vbW)}" height="${Math.round(vbH)}" fill="white"/>`
        ];
        
        // Sort shapes by z-order (containers first)
        const sortedShapes = [...shapes].sort((a, b) => this.shapeZOrder(a) - this.shapeZOrder(b));
        
        // Render shapes
        for (const s of sortedShapes) {
            const svgElements = this.renderShapeToSvg(
                s, pageH, usedMarkers, gradients, hasShadow, textLayer, pageRels || {}
            );
            svgLines.push(...svgElements);
        }
        
        // Render connections
        if (connects && connects.length > 0) {
            const shapeIndex = this.buildShapeIndex(shapes);
            const connLines = this.renderConnectionsSvg(connects, shapeIndex, pageH);
            svgLines.push(...connLines);
        }
        
        // Text layer
        if (textLayer.length > 0) {
            svgLines.push('<!-- Text layer -->');
            const processedText = this.avoidTextCollisions(textLayer);
            svgLines.push(...processedText);
        }
        
        svgLines.push('</svg>');
        
        // Build defs
        const defsContent = [];
        
        // QUALITY FIX 1: Add default shadow filter for professional look
        defsContent.push(
            `<filter id="shadow" x="-5%" y="-5%" width="115%" height="115%">`,
            `<feDropShadow dx="2" dy="2" stdDeviation="2" flood-color="rgba(0,0,0,0.2)"/>`,
            `</filter>`
        );
        
        // Arrow markers
        if (usedMarkers.size > 0) {
            const markerLines = this.arrowMarkerDefs(usedMarkers);
            defsContent.push(...markerLines.slice(1, -1)); // Remove <defs> wrapper
        }
        
        // Patterns and gradients
        const fillPatterns = {};
        const realGradients = {};
        for (const [gid, g] of Object.entries(gradients)) {
            if (g._is_pattern) {
                const realPid = gid.replace("__pat__", "");
                fillPatterns[realPid] = g;
            } else {
                realGradients[gid] = g;
            }
        }
        
        if (Object.keys(fillPatterns).length > 0) {
            defsContent.push(...this.fillPatternDefs(fillPatterns));
        }
        if (Object.keys(realGradients).length > 0) {
            defsContent.push(...this.gradientDefs(realGradients));
        }
        
        // Shadows
        for (const shdwSpec of Array.from(hasShadow).sort()) {
            const parts = shdwSpec.split("|");
            if (parts.length === 6) {
                const [sid, sdx, sdy, scolor, sopacity, sblur] = parts;
                defsContent.push(this.shadowFilterDef(
                    sid, parseFloat(sdx), parseFloat(sdy), scolor, 
                    parseFloat(sopacity), parseFloat(sblur)
                ));
            }
        }
        
        if (defsContent.length > 0) {
            const defsLines = ["<defs>", ...defsContent, "</defs>"];
            svgLines.splice(6, 0, ...defsLines);
        }
        
        return svgLines.join('\n');
    }

    /**
     * Render single shape to SVG elements
     */
    renderShapeToSvg(shape, pageH, usedMarkers, gradients, hasShadow, textLayer, pageRels) {
        // Merge with master first
        shape = this.mergeShapeWithMaster(shape, this.masters);
        
        const elements = [];
        
        // Skip invisible shapes
        const visVal = this.getCellVal(shape, "Visible");
        if (visVal === "0") return elements;
        
        // Get shape properties
        const wInch = this.getCellFloat(shape, "Width");
        const hInch = this.getCellFloat(shape, "Height");
        const wPx = Math.abs(wInch) * this.INCH_TO_PX;
        const hPx = Math.abs(hInch) * this.INCH_TO_PX;
        
        // Check for 1D connector
        const beginX = this.getCellVal(shape, "BeginX");
        const endX = this.getCellVal(shape, "EndX");
        const is1D = !!(beginX && endX);
        const objType = this.getCellVal(shape, "ObjType");
        const isConnector = is1D || objType === "2";
        
        // Handle groups
        const shapeType = shape.type || "Shape";
        if ((shapeType === "Group" || shape.sub_shapes?.length > 0) && !is1D) {
            return this.renderGroupShape(shape, pageH, usedMarkers, gradients, hasShadow, textLayer, pageRels);
        }
        
        // Handle 1D connectors
        if (isConnector && is1D) {
            return this.render1DConnector(shape, pageH, usedMarkers);
        }
        
        // 2D shapes with geometry
        if (shape.geometry && shape.geometry.length > 0) {
            return this.render2DShape(shape, pageH, wInch, hInch, wPx, hPx, usedMarkers, gradients, hasShadow, textLayer, pageRels);
        }
        
        // Fallback rectangle
        if (wPx > 0 && hPx > 0) {
            return this.renderFallbackShape(shape, pageH, wPx, hPx, usedMarkers, gradients, hasShadow, textLayer, pageRels);
        }
        
        return elements;
    }

    /**
     * Render 1D connector
     */
    render1DConnector(shape, pageH, usedMarkers) {
        const elements = [];
        
        const beginX = this.safeFloat(this.getCellVal(shape, "BeginX")) * this.INCH_TO_PX;
        const beginY = (pageH - this.safeFloat(this.getCellVal(shape, "BeginY"))) * this.INCH_TO_PX;
        const endX = this.safeFloat(this.getCellVal(shape, "EndX")) * this.INCH_TO_PX;
        const endY = (pageH - this.safeFloat(this.getCellVal(shape, "EndY"))) * this.INCH_TO_PX;
        
        const lineColor = this.resolveColor(this.getCellVal(shape, "LineColor")) || "#333333";
        let strokeWidth = this.getCellFloat(shape, "LineWeight") * this.INCH_TO_PX;
        if (strokeWidth < 1.0) strokeWidth = 1.5;
        
        const linePattern = parseInt(this.safeFloat(this.getCellVal(shape, "LinePattern", "1")));
        const dashArray = this.getDashArray(linePattern, strokeWidth);
        
        // Arrow markers
        const beginArrow = parseInt(this.safeFloat(this.getCellVal(shape, "BeginArrow", "0")));
        const endArrow = parseInt(this.safeFloat(this.getCellVal(shape, "EndArrow", "0")));
        const beginArrowSize = parseInt(this.safeFloat(this.getCellVal(shape, "BeginArrowSize", "2")));
        const endArrowSize = parseInt(this.safeFloat(this.getCellVal(shape, "EndArrowSize", "2")));
        
        let markerAttrs = "";
        const markerColor = lineColor.replace("#", "");
        
        if (beginArrow > 0) {
            const mid = `arrow_start_${beginArrowSize}_${markerColor}`;
            usedMarkers.add(mid);
            markerAttrs += ` marker-start="url(#${mid})"`;
        }
        if (endArrow > 0) {
            const mid = `arrow_end_${endArrowSize}_${markerColor}`;
            usedMarkers.add(mid);
            markerAttrs += ` marker-end="url(#${mid})"`;
        }
        
        // Default arrow for connectors without explicit EndArrow
        const objType = this.getCellVal(shape, "ObjType");
        const shapeName = (shape.name_u || shape.name || "").toLowerCase();
        if (endArrow === 0 && (objType === "2" || shapeName.includes("connector"))) {
            const endArrowCell = shape.cells?.EndArrow;
            if (!endArrowCell || (!endArrowCell.V && !endArrowCell.F)) {
                const mid = `arrow_end_2_${markerColor}`;
                usedMarkers.add(mid);
                markerAttrs += ` marker-end="url(#${mid})"`;
            }
        }
        
        // Render connector path
        if (shape.geometry && shape.geometry.length > 0) {
            return this.renderConnectorWithGeometry(shape, pageH, strokeWidth, lineColor, dashArray, markerAttrs);
        } else {
            // QUALITY FIX 4: Professional orthogonal routing
            const dx = endX - beginX;
            const dy = endY - beginY;
            const adx = Math.abs(dx);
            const ady = Math.abs(dy);
            
            if (adx > 10 && ady > 10) {
                // Smart orthogonal routing: prefer direction with larger displacement
                let pathD;
                if (adx > ady) {
                    // Horizontal-first (L-shape)
                    const midX = beginX + dx * 0.5;
                    pathD = `M ${beginX.toFixed(2)} ${beginY.toFixed(2)} L ${midX.toFixed(2)} ${beginY.toFixed(2)} L ${midX.toFixed(2)} ${endY.toFixed(2)} L ${endX.toFixed(2)} ${endY.toFixed(2)}`;
                } else {
                    // Vertical-first (L-shape)
                    const midY = beginY + dy * 0.5;
                    pathD = `M ${beginX.toFixed(2)} ${beginY.toFixed(2)} L ${beginX.toFixed(2)} ${midY.toFixed(2)} L ${endX.toFixed(2)} ${midY.toFixed(2)} L ${endX.toFixed(2)} ${endY.toFixed(2)}`;
                }
                
                elements.push(
                    `<polyline points="${this.pathToPoints(pathD)}" fill="none" stroke="${lineColor}" stroke-width="${strokeWidth.toFixed(2)}"` +
                    (dashArray ? ` stroke-dasharray="${dashArray}"` : '') +
                    markerAttrs + ' stroke-linejoin="round"/>'
                );
            } else {
                elements.push(
                    `<line x1="${beginX.toFixed(2)}" y1="${beginY.toFixed(2)}" x2="${endX.toFixed(2)}" y2="${endY.toFixed(2)}" ` +
                    `stroke="${lineColor}" stroke-width="${strokeWidth.toFixed(2)}"` +
                    (dashArray ? ` stroke-dasharray="${dashArray}"` : '') +
                    markerAttrs + '/>'
                );
            }
        }
        
        // FIX: Add connector text labels with background
        if (shape.text && shape.text.trim()) {
            const midX = (beginX + endX) / 2;
            const midY = (beginY + endY) / 2;
            const text = shape.text.trim();
            const fontSize = 8;
            const padding = 2;
            const textWidth = text.length * fontSize * 0.5;
            const textHeight = fontSize + padding * 2;
            
            // Semi-transparent background
            elements.push(
                '<rect x="' + (midX - textWidth/2 - padding) + '" y="' + (midY - textHeight/2) + 
                '" width="' + (textWidth + padding*2) + '" height="' + textHeight + 
                '" fill="white" fill-opacity="0.55" rx="1"/>'
            );
            
            // Text
            elements.push(
                '<text x="' + midX.toFixed(2) + '" y="' + midY.toFixed(2) + 
                '" text-anchor="middle" dominant-baseline="central" font-family="Noto Sans, sans-serif" font-size="' + 
                fontSize + '" fill="#000000">' + this.escapeXml(text) + '</text>'
            );
        }
        
        return elements;
    }

    /**
     * Render 2D shape with geometry
     */
    render2DShape(shape, pageH, wInch, hInch, wPx, hPx, usedMarkers, gradients, hasShadow, textLayer, pageRels) {
        const elements = [];
        const transform = this.computeTransform(shape, pageH);
        const style = this.getShapeStyle(shape, gradients, hasShadow);
        
        const masterW = shape._master_w || 0.0;
        const masterH = shape._master_h || 0.0;
        
        for (const geo of shape.geometry) {
            if (geo.no_show) continue;
            
            let geoFill = style.fill;
            let geoStroke = style.stroke;
            
            if (geo.no_fill) geoFill = "none";
            if (geo.no_line) geoStroke = "none";
            
            const pathD = this.geometryToPath(geo, wInch, hInch, masterW, masterH);
            if (!pathD) continue;
            
            const geoStyle = this.buildStyleString({
                fill: geoFill,
                stroke: geoStroke,
                strokeWidth: style.strokeWidth,
                fillOpacity: style.fillOpacity,
                strokeOpacity: style.strokeOpacity,
                strokeDasharray: style.strokeDasharray
            });
            
            elements.push(`<path d="${pathD}" ${geoStyle} ${style.shadowAttr} transform="${transform}"/>`);
        }
        
        // Render embedded image
        if (shape.foreign_data && this.media) {
            elements.push(...this.renderEmbeddedImage(shape, pageRels, wPx, hPx, transform));
        }
        
        // Collect text
        if (shape.text) {
            this.appendTextSvg(textLayer, shape, pageH, wPx, hPx);
        }
        
        return elements;
    }

    /**
     * Render fallback rectangle shape
     */
    renderFallbackShape(shape, pageH, wPx, hPx, usedMarkers, gradients, hasShadow, textLayer, pageRels) {
        const elements = [];
        const transform = this.computeTransform(shape, pageH);
        const style = this.getShapeStyle(shape, gradients, hasShadow);
        
        const rounding = this.getCellFloat(shape, "Rounding") * this.INCH_TO_PX;
        const rx = Math.max(rounding, 4.0);
        const rxAttr = ` rx="${rx.toFixed(2)}"`;
        
        // Prefer outlined rect for fallback shapes
        const fallbackFill = style.fill !== "none" ? style.fill : "#FAFAFA";
        const fallbackStroke = style.stroke !== "none" ? style.stroke : (this.resolveColor(this.getCellVal(shape, "LineColor")) || "#CCCCCC");
        
        const fallbackStyle = this.buildStyleString({
            fill: fallbackFill,
            stroke: fallbackStroke,
            strokeWidth: Math.max(style.strokeWidth, 0.75),
            fillOpacity: style.fillOpacity,
            strokeOpacity: style.strokeOpacity,
            strokeDasharray: style.strokeDasharray
        });
        
        elements.push(
            `<rect x="0" y="0" width="${wPx.toFixed(2)}" height="${hPx.toFixed(2)}" ` +
            `${fallbackStyle}${rxAttr}${style.shadowAttr} transform="${transform}"/>`
        );
        
        // Auto-add shape name as text if no text present
        if (!shape.text && !shape._has_text_elem) {
            const shapeLabel = shape.name_u || shape.name || "";
            if (shapeLabel && !shapeLabel.startsWith("Sheet.") && !shapeLabel.match(/^\d+$/)) {
                const cleanLabel = shapeLabel.includes(".") ? shapeLabel.split(".").pop() : shapeLabel;
                if (cleanLabel && cleanLabel.length < 30) {
                    shape.text = cleanLabel;
                }
            }
        }
        
        if (shape.text) {
            this.appendTextSvg(textLayer, shape, pageH, wPx, hPx);
        }
        
        return elements;
    }

    /**
     * Get shape style properties
     */
    getShapeStyle(shape, gradients, hasShadow) {
        let lineWeight = this.getCellFloat(shape, "LineWeight", 0.01) * this.INCH_TO_PX;
        if (lineWeight < 0.5) lineWeight = 1.5;
        else if (lineWeight > 20) lineWeight = 20;
        
        let lineColor = this.resolveColor(this.getCellVal(shape, "LineColor")) || "#333333";
        let fillForegnd = this.resolveColor(this.getCellVal(shape, "FillForegnd"));
        let fillBkgnd = this.resolveColor(this.getCellVal(shape, "FillBkgnd"));
        
        // Theme color resolution
        const resolvedColors = this.resolveThemeColors(shape, lineColor, fillForegnd, fillBkgnd);
        lineColor = resolvedColors.lineColor;
        fillForegnd = resolvedColors.fillForegnd;
        fillBkgnd = resolvedColors.fillBkgnd;
        
        const fillPattern = this.getCellVal(shape, "FillPattern", "1");
        const linePattern = parseInt(this.safeFloat(this.getCellVal(shape, "LinePattern", "1")));
        
        // Determine fill
        let fill = this.determineFill(fillPattern, fillForegnd, fillBkgnd, shape, gradients);
        
        // Stroke
        const stroke = linePattern !== 0 ? lineColor : "none";
        const strokeWidth = lineWeight;
        const dashArray = this.getDashArray(linePattern, strokeWidth);
        
        // Opacity
        const fillTrans = this.getCellFloat(shape, "FillForegndTrans");
        const fillOpacity = fillTrans > 0 ? (fillTrans > 1 ? 1.0 - fillTrans / 100.0 : 1.0 - fillTrans) : 1.0;
        
        const lineTrans = this.getCellFloat(shape, "LineColorTrans");
        const strokeOpacity = lineTrans > 0 ? (lineTrans > 1 ? 1.0 - lineTrans / 100.0 : 1.0 - lineTrans) : 1.0;
        
        // Container styling
        const isContainer = this.isContainer(shape);
        if (isContainer) {
            if (fill === "none" && !fillForegnd) {
                fill = "#F8F8F8";
                fillOpacity = 0.5;
            } else if (fill !== "none" && fillOpacity > 0.9) {
                fillOpacity = Math.max(0.3, fillOpacity * 0.5);
            }
        }
        
        // QUALITY FIX 1: Add professional shadows to ALL shapes (not just those with ShdwPattern)
        let shadowAttr = ' filter="url(#shadow)"';
        const shdwPattern = this.getCellVal(shape, "ShdwPattern");
        if (shdwPattern && shdwPattern !== "0") {
            // Use custom shadow if specified
            const shdwOffsetX = this.getCellFloat(shape, "ShdwOffsetX", 0.0278) * this.INCH_TO_PX;
            const shdwOffsetY = -this.getCellFloat(shape, "ShdwOffsetY", -0.0278) * this.INCH_TO_PX;
            const shdwColor = this.resolveColor(this.getCellVal(shape, "ShdwForegnd")) || "#000000";
            const shdwTrans = this.getCellFloat(shape, "ShdwForegndTrans", 0.5);
            const shdwOpacity = Math.max(0.05, 1.0 - (shdwTrans > 1 ? shdwTrans / 100.0 : shdwTrans));
            const shdwBlur = Math.max(0.5, Math.min(Math.abs(shdwOffsetX), Math.abs(shdwOffsetY), 4.0));
            
            const shdwId = `shadow_${shape.id}`;
            hasShadow.add(`${shdwId}|${shdwOffsetX.toFixed(1)}|${shdwOffsetY.toFixed(1)}|${shdwColor}|${shdwOpacity.toFixed(2)}|${shdwBlur.toFixed(1)}`);
            shadowAttr = ` filter="url(#${shdwId})"`;
        }
        
        return {
            fill, stroke, strokeWidth, fillOpacity, strokeOpacity,
            strokeDasharray: dashArray, shadowAttr
        };
    }

    /**
     * Determine fill style (solid, gradient, pattern)
     */
    determineFill(fillPattern, fillForegnd, fillBkgnd, shape, gradients) {
        const fillPatInt = parseInt(this.safeFloat(fillPattern, 1));
        
        if (fillPatInt === 0) {
            return "none";
        } else if (fillPatInt === 1) {
            const baseFill = fillForegnd || fillBkgnd;
            if (!baseFill || baseFill === "none") {
                return "none";
            }
            
            // QUALITY FIX 5: Add subtle gradient to solid fills for professional look
            const gradId = `grad_${shape.id}_auto`;
            const lighterColor = this.lightenColor(baseFill, 0.15);
            
            gradients[gradId] = {
                start: lighterColor,
                end: baseFill,
                dir: 0, // Top to bottom
                radial: false
            };
            
            return `url(#${gradId})`;
        } else if (fillPatInt >= 25 && fillPatInt <= 40) {
            // Gradient fill
            const startColor = fillBkgnd || "#FFFFFF";
            const endColor = fillForegnd || fillBkgnd || "#CCCCCC";
            
            if (startColor.toUpperCase() === endColor.toUpperCase()) {
                return startColor;
            }
            
            const gradId = `grad_${shape.id}_${fillPatInt}`;
            const gradAngle = this.getGradientAngle(shape, fillPatInt);
            const isRadial = [29, 30, 31, 32, 37, 38, 39].includes(fillPatInt);
            
            const gradInfo = {
                start: startColor,
                end: endColor,
                dir: gradAngle,
                radial: isRadial
            };
            
            // Multi-stop gradients
            if (shape._gradient_stops?.[0]) {
                const resolvedStops = [];
                for (const [pos, col] of shape._gradient_stops[0]) {
                    const resolvedCol = this.resolveColor(col) || col;
                    if (resolvedCol.startsWith("#")) {
                        resolvedStops.push([pos, resolvedCol]);
                    }
                }
                if (resolvedStops.length >= 2) {
                    gradInfo.stops = resolvedStops;
                }
            }
            
            if (isRadial) {
                gradInfo.fx = this.getCellFloat(shape, "FillGradientFocusX", 0.5) * 100;
                gradInfo.fy = this.getCellFloat(shape, "FillGradientFocusY", 0.5) * 100;
            }
            
            gradients[gradId] = gradInfo;
            return `url(#${gradId})`;
        } else if (fillPatInt >= 2 && fillPatInt <= 24) {
            // Pattern fill
            const patId = `fpat_${shape.id}_${fillPatInt}`;
            const fgColor = fillForegnd || "#333333";
            const bgColor = fillBkgnd || "#FFFFFF";
            
            gradients[`__pat__${patId}`] = {
                fg: fgColor,
                bg: bgColor,
                type: fillPatInt,
                _is_pattern: true
            };
            
            return `url(#${patId})`;
        }
        
        return "none";
    }

    /**
     * Get gradient angle for fill pattern
     */
    getGradientAngle(shape, fillPatInt) {
        const gradDir = this.getCellFloat(shape, "FillGradientDir");
        const gradAngleRaw = this.getCellFloat(shape, "FillGradientAngle");
        
        if (gradAngleRaw) {
            return Math.abs(gradAngleRaw) < 7 ? gradAngleRaw * 180 / Math.PI : gradAngleRaw;
        } else if (gradDir) {
            return gradDir * 45;
        } else {
            const patternAngles = {
                25: 0, 26: 90, 27: 45, 28: 315, 29: 0, 30: 90,
                33: 0, 34: 90, 35: 45, 36: 315, 40: 0
            };
            return patternAngles[fillPatInt] || 0;
        }
    }

    /**
     * Convert geometry to SVG path
     */
    geometryToPath(geo, w, h, masterW = 0, masterH = 0) {
        if (geo.no_show) return "";
        
        const absW = Math.abs(w);
        const absH = Math.abs(h);
        
        const absMW = Math.abs(masterW);
        const absMH = Math.abs(masterH);
        const sx = (absMW > 1e-6 && Math.abs(absMW - absW) > 1e-6) ? absW / absMW : 1.0;
        const sy = (absMH > 1e-6 && Math.abs(absMH - absH) > 1e-6) ? absH / absMH : 1.0;
        
        const dParts = [];
        let cx = 0, cy = 0;
        
        for (let rowIdx = 0; rowIdx < geo.rows.length; rowIdx++) {
            const row = geo.rows[rowIdx];
            const rt = row.type;
            const cells = row.cells;
            
            // Skip geometry rows with empty coordinates
            if (["LineTo", "ArcTo"].includes(rt)) {
                const hasAny = ["X", "Y"].some(cn => {
                    const cv = cells[cn]?.V;
                    return cv !== null && cv !== undefined && cv !== "";
                });
                if (!hasAny) {
                    const hasFormula = ["X", "Y"].some(cn => {
                        const cf = cells[cn]?.F;
                        return cf && cf !== "Inh";
                    });
                    if (!hasFormula) continue;
                }
            }
            
            if (rt === "MoveTo") {
                const x = this.safeFloat(cells.X?.V) * sx;
                const y = this.safeFloat(cells.Y?.V) * sy;
                dParts.push(`M ${(x * this.INCH_TO_PX).toFixed(2)} ${((absH - y) * this.INCH_TO_PX).toFixed(2)}`);
                cx = x; cy = y;
            } else if (rt === "RelMoveTo") {
                // Relative coordinates (0-1) scale to shape dimensions
                const x = this.safeFloat(cells.X?.V) * absW;
                const y = this.safeFloat(cells.Y?.V) * absH;
                dParts.push(`M ${(x * this.INCH_TO_PX).toFixed(2)} ${((absH - y) * this.INCH_TO_PX).toFixed(2)}`);
                cx = x / absW; // Normalized for next operations
                cy = y / absH;
                
                // Detect oval pattern
                const remaining = geo.rows.slice(rowIdx + 1);
                const remainingTypes = remaining.map(r => r.type).filter(t => t);
                if (remainingTypes.length >= 3 && remainingTypes.every(t => t === "ArcTo")) {
                    const arcPoints = [[x, y]];
                    for (const ar of remaining) {
                        if (ar.type !== "ArcTo") break;
                        const ax = this.safeFloat(ar.cells.X?.V) * sx;
                        const ay = this.safeFloat(ar.cells.Y?.V) * sy;
                        arcPoints.push([ax, ay]);
                    }
                    
                    if (arcPoints.length >= 4) {
                        const first = arcPoints[0];
                        const last = arcPoints[arcPoints.length - 1];
                        const dist = Math.sqrt((first[0] - last[0]) ** 2 + (first[1] - last[1]) ** 2);
                        
                        if (dist < 0.01) {
                            const allX = arcPoints.map(p => p[0]);
                            const allY = arcPoints.map(p => p[1]);
                            const ecx = (Math.min(...allX) + Math.max(...allX)) / 2 * this.INCH_TO_PX;
                            const ecy = (absH - (Math.min(...allY) + Math.max(...allY)) / 2) * this.INCH_TO_PX;
                            const erx = (Math.max(...allX) - Math.min(...allX)) / 2 * this.INCH_TO_PX;
                            const ery = (Math.max(...allY) - Math.min(...allY)) / 2 * this.INCH_TO_PX;
                            
                            if (erx > 0.5 && ery > 0.5) {
                                dParts.length = 0;
                                dParts.push(`M ${(ecx - erx).toFixed(2)} ${ecy.toFixed(2)}`);
                                dParts.push(`A ${erx.toFixed(2)} ${ery.toFixed(2)} 0 1 0 ${(ecx + erx).toFixed(2)} ${ecy.toFixed(2)}`);
                                dParts.push(`A ${erx.toFixed(2)} ${ery.toFixed(2)} 0 1 0 ${(ecx - erx).toFixed(2)} ${ecy.toFixed(2)}`);
                                dParts.push("Z");
                                break;
                            }
                        }
                    }
                }
            } else if (rt === "LineTo") {
                const x = this.safeFloat(cells.X?.V) * sx;
                const y = this.safeFloat(cells.Y?.V) * sy;
                dParts.push(`L ${(x * this.INCH_TO_PX).toFixed(2)} ${((absH - y) * this.INCH_TO_PX).toFixed(2)}`);
                cx = x; cy = y;
            } else if (rt === "RelLineTo") {
                // Relative coordinates (0-1) scale to shape dimensions
                const x = this.safeFloat(cells.X?.V) * absW;
                const y = this.safeFloat(cells.Y?.V) * absH;
                dParts.push(`L ${(x * this.INCH_TO_PX).toFixed(2)} ${((absH - y) * this.INCH_TO_PX).toFixed(2)}`);
                cx = x / absW; // Normalized for next operations
                cy = y / absH;
            } else if (rt === "ArcTo") {
                const x = this.safeFloat(cells.X?.V) * sx;
                const y = this.safeFloat(cells.Y?.V) * sy;
                const a = this.safeFloat(cells.A?.V) * sy;
                this.appendArc(dParts, cx, cy, x, y, a, absH);
                cx = x; cy = y;
            } else if (rt === "Ellipse") {
                const ex = this.safeFloat(cells.X?.V) * sx;
                const ey = this.safeFloat(cells.Y?.V) * sy;
                const ea = this.safeFloat(cells.A?.V) * sx;
                const eb = this.safeFloat(cells.B?.V) * sy;
                const ec = this.safeFloat(cells.C?.V) * sx;
                const ed = this.safeFloat(cells.D?.V) * sy;
                
                const rx = Math.sqrt((ea - ex) ** 2 + (eb - ey) ** 2);
                const ry = Math.sqrt((ec - ex) ** 2 + (ed - ey) ** 2);
                const cpx = ex * this.INCH_TO_PX;
                const cpy = (absH - ey) * this.INCH_TO_PX;
                const rpx = Math.max(0.001, rx) * this.INCH_TO_PX;
                const rpy = Math.max(0.001, ry) * this.INCH_TO_PX;
                
                dParts.push(
                    `M ${(cpx - rpx).toFixed(2)} ${cpy.toFixed(2)} ` +
                    `A ${rpx.toFixed(2)} ${rpy.toFixed(2)} 0 1 0 ${(cpx + rpx).toFixed(2)} ${cpy.toFixed(2)} ` +
                    `A ${rpx.toFixed(2)} ${rpy.toFixed(2)} 0 1 0 ${(cpx - rpx).toFixed(2)} ${cpy.toFixed(2)} Z`
                );
            }
            // Additional geometry types (NURBSTo, PolylineTo, etc.) would be implemented here
        }
        
        let result = dParts.join(" ");
        
        // Ensure path starts with M
        if (result && !result.startsWith("M")) {
            result = `M 0.00 0.00 ${result}`;
        }
        
        // Auto-close path if needed
        if (result && !result.includes("Z") && dParts.length >= 3) {
            const firstM = result.match(/M\s+([-+]?[\d.]+)\s+([-+]?[\d.]+)/);
            if (firstM) {
                const lastPart = dParts[dParts.length - 1];
                const lastCoords = lastPart.match(/([-+]?[\d.]+)/g);
                if (lastCoords && lastCoords.length >= 2) {
                    const sx = parseFloat(firstM[1]);
                    const sy = parseFloat(firstM[2]);
                    const ex = parseFloat(lastCoords[lastCoords.length - 2]);
                    const ey = parseFloat(lastCoords[lastCoords.length - 1]);
                    
                    if (Math.abs(sx - ex) < 0.5 && Math.abs(sy - ey) < 0.5) {
                        result += " Z";
                    }
                }
            }
        }
        
        return result;
    }

    /**
     * Append arc segment to path
     */
    appendArc(dParts, cx, cy, x, y, bulge, h) {
        if (Math.abs(bulge) < 1e-6) {
            dParts.push(`L ${(x * this.INCH_TO_PX).toFixed(2)} ${((h - y) * this.INCH_TO_PX).toFixed(2)}`);
            return;
        }
        
        const dx = x - cx;
        const dy = y - cy;
        const chord = Math.sqrt(dx * dx + dy * dy);
        if (chord < 1e-10) return;
        
        const sagitta = Math.abs(bulge);
        const radius = (chord * chord / 4 + sagitta * sagitta) / (2 * sagitta);
        const maxRadius = chord * 5.0;
        const finalRadius = Math.min(radius, maxRadius);
        const radiusPx = finalRadius * this.INCH_TO_PX;
        
        const largeArc = sagitta > chord / 2 ? 1 : 0;
        const sweep = bulge > 0 ? 0 : 1;
        
        dParts.push(
            `A ${radiusPx.toFixed(2)} ${radiusPx.toFixed(2)} 0 ${largeArc} ${sweep} ` +
            `${(x * this.INCH_TO_PX).toFixed(2)} ${((h - y) * this.INCH_TO_PX).toFixed(2)}`
        );
    }

    /**
     * Compute SVG transform for shape positioning
     */
    computeTransform(shape, pageH) {
        const pinX = this.getCellFloat(shape, "PinX") * this.INCH_TO_PX;
        const pinY = (pageH - this.getCellFloat(shape, "PinY")) * this.INCH_TO_PX;
        const w = this.getCellFloat(shape, "Width");
        const h = this.getCellFloat(shape, "Height");
        
        const locPinXVal = this.getCellVal(shape, "LocPinX");
        const locPinX = (locPinXVal ? this.safeFloat(locPinXVal) : Math.abs(w) * 0.5) * this.INCH_TO_PX;
        const locPinYVal = this.getCellVal(shape, "LocPinY");
        const locPinYRaw = locPinYVal ? this.safeFloat(locPinYVal) : Math.abs(h) * 0.5;
        const locPinY = (Math.abs(h) - locPinYRaw) * this.INCH_TO_PX;
        
        const angle = this.getCellFloat(shape, "Angle");
        const flipX = this.getCellVal(shape, "FlipX") === "1";
        const flipY = this.getCellVal(shape, "FlipY") === "1";
        
        const parts = [];
        
        const tx = pinX - locPinX;
        const ty = pinY - locPinY;
        parts.push(`translate(${tx.toFixed(2)},${ty.toFixed(2)})`);
        
        if (Math.abs(angle) > 1e-6) {
            const angleDeg = -angle * 180 / Math.PI;
            parts.push(`rotate(${angleDeg.toFixed(2)},${locPinX.toFixed(2)},${locPinY.toFixed(2)})`);
        }
        
        if (flipX || flipY) {
            const sx = flipX ? -1 : 1;
            const sy = flipY ? -1 : 1;
            parts.push(`translate(${locPinX.toFixed(2)},${locPinY.toFixed(2)})`);
            parts.push(`scale(${sx},${sy})`);
            parts.push(`translate(${(-locPinX).toFixed(2)},${(-locPinY).toFixed(2)})`);
        }
        
        return parts.join(" ");
    }

    /**
     * Merge shape with master data
     */
    mergeShapeWithMaster(shape, masters, parentMasterId = "") {
        const masterId = shape.master || parentMasterId;
        const masterShapeId = shape.master_shape;
        
        if (!masterId || !masters[masterId]) {
            return shape;
        }
        
        const masterShapes = masters[masterId];
        let masterSd = null;
        
        if (masterShapeId && masterShapes[masterShapeId]) {
            masterSd = masterShapes[masterShapeId];
        } else if (Object.keys(masterShapes).length > 0) {
            // Prefer shapes with geometry when MasterShape is not specified
            const shapesWithGeometry = Object.values(masterShapes).filter(s => s.geometry && s.geometry.length > 0);
            if (shapesWithGeometry.length > 0) {
                masterSd = shapesWithGeometry[0];
            } else {
                masterSd = Object.values(masterShapes)[0];
            }
        }
        
        if (!masterSd) return shape;
        
        // Merge cells
        const mergedCells = { ...masterSd.cells };
        for (const [k, v] of Object.entries(shape.cells)) {
            if (v.V || v.F) {
                mergedCells[k] = v;
            }
        }
        shape.cells = mergedCells;
        
        // Merge geometry
        if (!shape.geometry.length && masterSd.geometry?.length) {
            shape.geometry = masterSd.geometry;
            if (masterSd.cells.Width?.V) shape._master_w = this.safeFloat(masterSd.cells.Width.V);
            if (masterSd.cells.Height?.V) shape._master_h = this.safeFloat(masterSd.cells.Height.V);
        }
        
        // Merge text
        if (!shape.text && !shape._has_text_elem && masterSd.text && shape.type !== "Group") {
            const txt = masterSd.text;
            if (!["Label", "Abc", "Table", "Entity", "Class"].includes(txt)) {
                shape.text = txt;
                if (!shape.text_parts.length && masterSd.text_parts?.length) {
                    shape.text_parts = masterSd.text_parts;
                }
            }
        }
        
        // Merge formatting
        if (!Object.keys(shape.char_formats).length && masterSd.char_formats) {
            shape.char_formats = masterSd.char_formats;
        }
        if (!Object.keys(shape.para_formats).length && masterSd.para_formats) {
            shape.para_formats = masterSd.para_formats;
        }
        
        // Merge other properties
        if (!Object.keys(shape.controls).length && masterSd.controls) {
            shape.controls = masterSd.controls;
        }
        if (!Object.keys(shape.connections).length && masterSd.connections) {
            shape.connections = masterSd.connections;
        }
        if (!Object.keys(shape.user).length && masterSd.user) {
            shape.user = masterSd.user;
        }
        if (!shape.foreign_data && masterSd.foreign_data) {
            shape.foreign_data = masterSd.foreign_data;
        }
        if (!shape._gradient_stops && masterSd._gradient_stops) {
            shape._gradient_stops = masterSd._gradient_stops;
        }
        
        return shape;
    }

    /**
     * Append text SVG to text layer
     */
    appendTextSvg(textLayer, shape, pageH, wPx, hPx) {
        const text = shape.text;
        if (!text) return;
        
        // Get text position
        const pinX = this.getCellFloat(shape, "PinX") * this.INCH_TO_PX;
        const pinY = (pageH - this.getCellFloat(shape, "PinY")) * this.INCH_TO_PX;
        const wInch = this.getCellFloat(shape, "Width");
        const hInch = this.getCellFloat(shape, "Height");
        const locPinX = this.getCellFloat(shape, "LocPinX") * this.INCH_TO_PX || Math.abs(wInch) * 0.5 * this.INCH_TO_PX;
        const locPinY = this.getCellFloat(shape, "LocPinY") * this.INCH_TO_PX || Math.abs(hInch) * 0.5 * this.INCH_TO_PX;
        
        // Calculate text center
        const shapeLeft = pinX - locPinX;
        const shapeTop = pinY - (Math.abs(hInch) * this.INCH_TO_PX - locPinY);
        let tx = shapeLeft + Math.abs(wInch) * this.INCH_TO_PX * 0.5;
        let ty = shapeTop + Math.abs(hInch) * this.INCH_TO_PX * 0.5;
        
        // Get text formatting
        const charFormats = shape.char_formats || {};
        const charFmt = charFormats["0"] || {};
        // QUALITY FIX 3: Smart font sizing based on shape dimensions
        let fontSize = this.safeFloat(charFmt.Size || "0.1111") * this.INCH_TO_PX;
        if (fontSize < 6 || fontSize > 72) {
            // Calculate optimal font size based on shape size
            const shapeWidthPx = Math.abs(wInch) * this.INCH_TO_PX;
            const shapeHeightPx = Math.abs(hInch) * this.INCH_TO_PX;
            const textLength = text ? text.length : 10;
            
            // Adaptive sizing: aim for 60-80% of shape width utilization
            const targetWidth = shapeWidthPx * 0.7;
            const scaleFactor = 0.6; // Approximate character width factor
            fontSize = Math.min(18, Math.max(8, targetWidth / (textLength * scaleFactor)));
            
            // Also consider height constraint
            const maxFontForHeight = shapeHeightPx * 0.4; // 40% of shape height
            fontSize = Math.min(fontSize, maxFontForHeight);
        }
        
        // QUALITY FIX 2: Smart text contrast based on shape background
        let textColor = this.resolveColor(charFmt.Color);
        if (!textColor) {
            // Auto-determine text color based on shape background
            const shapeFill = this.resolveColor(this.getCellVal(shape, "FillForegnd")) || "#FFFFFF";
            textColor = this.getContrastTextColor(shapeFill);
        }
        const fontName = charFmt.Font || "";
        const fontFamily = this.mapFontFamily(fontName);
        
        const styleBits = parseInt(this.safeFloat(charFmt.Style || "0"));
        const isBold = !!(styleBits & 1);
        const isItalic = !!(styleBits & 2);
        
        // Paragraph alignment
        const paraFmt = shape.para_formats?.["0"] || {};
        const halign = parseInt(this.safeFloat(paraFmt.HorzAlign || "1"));
        const textAnchor = ["start", "middle", "end"][halign] || "middle";
        
        // Text rotation
        const txtAngle = this.getCellFloat(shape, "TxtAngle");
        let txtRotate = "";
        if (Math.abs(txtAngle) > 1e-6) {
            const txtAngleDeg = -txtAngle * 180 / Math.PI;
            txtRotate = ` transform="rotate(${txtAngleDeg.toFixed(1)},${tx.toFixed(2)},${ty.toFixed(2)})"`;
        }
        
        // Font attributes
        const fw = isBold ? ' font-weight="bold"' : "";
        const fs = isItalic ? ' font-style="italic"' : "";
        
        // Handle multi-line text
        const textLines = text.split("\n");
        
        if (textLines.length === 1) {
            // Single line
            const escaped = this.escapeXml(textLines[0]);
            textLayer.push(
                `<text x="${tx.toFixed(2)}" y="${ty.toFixed(2)}" text-anchor="${textAnchor}" dominant-baseline="central" ` +
                `font-family="${fontFamily}" font-size="${fontSize.toFixed(1)}" ` +
                `fill="${textColor}"${fw}${fs}${txtRotate}>${escaped}</text>`
            );
        } else {
            // Multi-line
            const totalHeight = textLines.length * fontSize * 1.2;
            let startY = ty - totalHeight / 2 + fontSize * 0.6;
            
            for (let j = 0; j < textLines.length; j++) {
                const tline = textLines[j];
                if (!tline.trim()) continue;
                
                const escaped = this.escapeXml(tline);
                const ly = startY + j * fontSize * 1.2;
                
                textLayer.push(
                    `<text x="${tx.toFixed(2)}" y="${ly.toFixed(2)}" ` +
                    `text-anchor="${textAnchor}" font-family="${fontFamily}" ` +
                    `font-size="${fontSize.toFixed(1)}" fill="${textColor}"${fw}${fs}${txtRotate}>${escaped}</text>`
                );
            }
        }
    }

    /**
     * Avoid text collisions
     */
    avoidTextCollisions(textElements) {
        const textRe = /<text\s+x="([^"]+)"\s+y="([^"]+)"[^>]*font-size="([^"]+)"[^>]*>(.*?)<\/text>/g;
        
        const parsed = [];
        for (const elem of textElements) {
            const match = textRe.exec(elem);
            textRe.lastIndex = 0; // Reset regex
            
            if (!match || elem.includes('data-noclip="1"')) {
                parsed.push([elem, null]);
                continue;
            }
            
            const tx = parseFloat(match[1]);
            const ty = parseFloat(match[2]);
            const fs = parseFloat(match[3]);
            const txt = match[4];
            
            const cleanTxt = txt.replace(/<[^>]+>/g, '');
            if (!cleanTxt.trim()) {
                parsed.push([elem, null]);
                continue;
            }
            
            const estW = cleanTxt.length * fs * 0.55;
            const estH = fs * 1.3;
            
            let boxX;
            if (elem.includes('text-anchor="start"')) {
                boxX = tx - 1;
            } else if (elem.includes('text-anchor="end"')) {
                boxX = tx - estW - 1;
            } else {
                boxX = tx - estW / 2 - 1;
            }
            
            const boxY = ty - estH * 0.55;
            
            parsed.push([elem, {
                tx, ty, fs, origY: match[2], cleanTxt, estW, estH, boxX, boxY
            }]);
        }
        
        const placedBoxes = [];
        const result = [];
        
        for (let idx = 0; idx < parsed.length; idx++) {
            const [elem, data] = parsed[idx];
            
            if (data === null) {
                result.push(elem);
                continue;
            }
            
            let { tx, ty, estW, estH, boxX, boxY, origY } = data;
            
            // Simple collision avoidance - shift down if overlapping
            for (let attempt = 0; attempt < 3; attempt++) {
                let collision = false;
                for (const [px, py, pw, ph] of placedBoxes) {
                    const overlapX = Math.min(boxX + estW + 2, px + pw) - Math.max(boxX, px);
                    const overlapY = Math.min(boxY + estH, py + ph) - Math.max(boxY, py);
                    if (overlapX > 0 && overlapY > 0 && overlapY > estH * 0.2) {
                        collision = true;
                        ty += estH + 2;
                        boxY = ty - estH * 0.55;
                        break;
                    }
                }
                if (!collision) break;
            }
            
            placedBoxes.push([boxX, boxY, estW + 2, estH]);
            
            let updatedElem = elem;
            if (Math.abs(ty - parseFloat(origY)) > 0.5) {
                updatedElem = updatedElem.replace(`y="${origY}"`, `y="${ty.toFixed(2)}"`);
            }
            
            result.push(updatedElem);
        }
        
        return result;
    }

    /**
     * Build shape index for connections
     */
    buildShapeIndex(shapes) {
        const idx = {};
        for (const s of shapes) {
            idx[s.id] = s;
            if (s.sub_shapes) {
                for (const sub of s.sub_shapes) {
                    idx[sub.id] = sub;
                    if (sub.sub_shapes) {
                        for (const subsub of sub.sub_shapes) {
                            idx[subsub.id] = subsub;
                        }
                    }
                }
            }
        }
        return idx;
    }

    /**
     * Render connections as SVG
     */
    renderConnectionsSvg(connects, shapeIndex, pageH) {
        const lines = [];
        const connectorSheets = new Set();
        
        for (const conn of connects) {
            const fromCell = conn.from_cell;
            if (["BeginX", "EndX"].includes(fromCell)) {
                connectorSheets.add(conn.from_sheet);
            }
        }
        
        for (const conn of connects) {
            const fromShape = shapeIndex[conn.from_sheet];
            const toShape = shapeIndex[conn.to_sheet];
            
            if (!fromShape || !toShape || connectorSheets.has(conn.from_sheet)) {
                continue;
            }
            
            const fromPt = this.resolveConnectionPoint(fromShape, conn.from_cell, pageH);
            const toPt = this.resolveConnectionPoint(toShape, conn.to_cell, pageH);
            
            if (fromPt && toPt) {
                lines.push(
                    `<line x1="${fromPt[0].toFixed(2)}" y1="${fromPt[1].toFixed(2)}" ` +
                    `x2="${toPt[0].toFixed(2)}" y2="${toPt[1].toFixed(2)}" ` +
                    `stroke="#555555" stroke-width="1.50" marker-end="url(#arrow_end_2_555555)"/>`
                );
            }
        }
        
        return lines;
    }

    /**
     * Resolve connection point to page coordinates
     */
    resolveConnectionPoint(shape, cellRef, pageH) {
        const pinX = this.getCellFloat(shape, "PinX");
        const pinY = this.getCellFloat(shape, "PinY");
        const locPinX = this.getCellFloat(shape, "LocPinX");
        const locPinY = this.getCellFloat(shape, "LocPinY");
        
        if (cellRef.startsWith("Controls.")) {
            const rowKey = cellRef.split(".")[1];
            const ctrl = shape.controls?.[rowKey];
            if (ctrl) {
                const lx = this.safeFloat(ctrl.X);
                const ly = this.safeFloat(ctrl.Y);
                const px = (pinX - locPinX + lx) * this.INCH_TO_PX;
                const py = (pageH - (pinY - locPinY + ly)) * this.INCH_TO_PX;
                return [px, py];
            }
        } else if (cellRef.startsWith("Connections.")) {
            const suffix = cellRef.split(".")[1];
            const match = suffix.match(/X(\d+)/);
            if (match) {
                const rowIx = (parseInt(match[1]) - 1).toString();
                const conn = shape.connections?.[rowIx];
                if (conn) {
                    const lx = this.safeFloat(conn.X);
                    const ly = this.safeFloat(conn.Y);
                    const px = (pinX - locPinX + lx) * this.INCH_TO_PX;
                    const py = (pageH - (pinY - locPinY + ly)) * this.INCH_TO_PX;
                    return [px, py];
                }
            }
        }
        
        return null;
    }

    /**
     * Generate arrow marker definitions
     */
    arrowMarkerDefs(usedMarkers) {
        if (usedMarkers.size === 0) return [];
        
        const lines = ["<defs>"];
        
        for (const markerId of Array.from(usedMarkers).sort()) {
            const parts = markerId.split("_");
            const direction = parts[1] || "end";
            const sizeIdx = parseInt(parts[2]) || 3;
            const color = parts[3] ? `#${parts[3]}` : "#333333";
            
            const scale = this.ARROW_SIZES[sizeIdx] || 1.0;
            const markerW = 10 * scale;
            const markerH = 7 * scale;
            
            if (direction === "start") {
                lines.push(
                    `<marker id="${markerId}" markerWidth="${markerW.toFixed(1)}" ` +
                    `markerHeight="${markerH.toFixed(1)}" refX="0" refY="${(markerH/2).toFixed(1)}" ` +
                    `orient="auto" markerUnits="userSpaceOnUse">` +
                    `<polygon points="${markerW.toFixed(1)} 0, 0 ${(markerH/2).toFixed(1)}, ` +
                    `${markerW.toFixed(1)} ${markerH.toFixed(1)}" fill="${color}"/>` +
                    `</marker>`
                );
            } else {
                lines.push(
                    `<marker id="${markerId}" markerWidth="${markerW.toFixed(1)}" ` +
                    `markerHeight="${markerH.toFixed(1)}" refX="${markerW.toFixed(1)}" ` +
                    `refY="${(markerH/2).toFixed(1)}" orient="auto" markerUnits="userSpaceOnUse">` +
                    `<polygon points="0 0, ${markerW.toFixed(1)} ${(markerH/2).toFixed(1)}, ` +
                    `0 ${markerH.toFixed(1)}" fill="${color}"/>` +
                    `</marker>`
                );
            }
        }
        
        lines.push("</defs>");
        return lines;
    }

    /**
     * Generate fill pattern definitions
     */
    fillPatternDefs(patterns) {
        if (Object.keys(patterns).length === 0) return [];
        
        const lines = [];
        
        for (const [pid, p] of Object.entries(patterns)) {
            const fg = p.fg || "#000000";
            const bg = p.bg || "#FFFFFF";
            const patType = p.type || 2;
            const spacing = 6;
            const strokeW = 1.0;
            
            if ([2, 3, 4, 5].includes(patType)) {
                if (patType === 2) {
                    // Horizontal lines
                    lines.push(
                        `<pattern id="${pid}" patternUnits="userSpaceOnUse" ` +
                        `width="${spacing}" height="${spacing}">` +
                        `<rect width="${spacing}" height="${spacing}" fill="${bg}"/>` +
                        `<line x1="0" y1="${spacing/2}" x2="${spacing}" y2="${spacing/2}" ` +
                        `stroke="${fg}" stroke-width="${strokeW}"/>` +
                        `</pattern>`
                    );
                } else if (patType === 3) {
                    // Vertical lines
                    lines.push(
                        `<pattern id="${pid}" patternUnits="userSpaceOnUse" ` +
                        `width="${spacing}" height="${spacing}">` +
                        `<rect width="${spacing}" height="${spacing}" fill="${bg}"/>` +
                        `<line x1="${spacing/2}" y1="0" x2="${spacing/2}" y2="${spacing}" ` +
                        `stroke="${fg}" stroke-width="${strokeW}"/>` +
                        `</pattern>`
                    );
                }
                // Additional pattern types would be implemented here
            }
        }
        
        return lines;
    }

    /**
     * Generate gradient definitions
     */
    gradientDefs(gradients) {
        if (Object.keys(gradients).length === 0) return [];
        
        const lines = [];
        
        for (const [gid, g] of Object.entries(gradients)) {
            const stops = g.stops || [[0, g.start], [100, g.end]];
            const stopXml = stops.map(([off, col]) => 
                `<stop offset="${off}%" stop-color="${col}"/>`
            ).join("");
            
            if (g.radial) {
                const gcx = g.cx || 50;
                const gcy = g.cy || 50;
                const gfx = g.fx || gcx;
                const gfy = g.fy || gcy;
                const gr = g.r || 50;
                
                lines.push(
                    `<radialGradient id="${gid}" cx="${gcx}%" cy="${gcy}%" ` +
                    `r="${gr}%" fx="${gfx}%" fy="${gfy}%">${stopXml}</radialGradient>`
                );
            } else {
                const angle = g.dir || 0;
                const rad = angle * Math.PI / 180;
                const x1 = 50 - 50 * Math.cos(rad);
                const y1 = 50 + 50 * Math.sin(rad);
                const x2 = 50 + 50 * Math.cos(rad);
                const y2 = 50 - 50 * Math.sin(rad);
                
                lines.push(
                    `<linearGradient id="${gid}" ` +
                    `x1="${x1.toFixed(1)}%" y1="${y1.toFixed(1)}%" x2="${x2.toFixed(1)}%" y2="${y2.toFixed(1)}%">` +
                    `${stopXml}</linearGradient>`
                );
            }
        }
        
        return lines;
    }

    /**
     * Generate shadow filter definition
     */
    shadowFilterDef(shadowId = "shadow", dx = 2.0, dy = 2.0, color = "#000000", opacity = 0.25, blur = 1.5) {
        return (
            `<filter id="${shadowId}" x="-20%" y="-20%" width="150%" height="150%">` +
            `<feDropShadow dx="${dx.toFixed(1)}" dy="${dy.toFixed(1)}" stdDeviation="${blur.toFixed(1)}" ` +
            `flood-color="${color}" flood-opacity="${opacity.toFixed(2)}"/>` +
            `</filter>`
        );
    }

    /**
     * QUALITY FIX 5: Lighten a hex color by a percentage for gradients
     */
    lightenColor(hex, percent) {
        if (!hex || hex === "none" || !hex.startsWith("#")) {
            return hex;
        }

        let color = hex.replace("#", "");
        if (color.length === 3) {
            color = color.split('').map(c => c + c).join('');
        }

        const r = parseInt(color.substr(0, 2), 16);
        const g = parseInt(color.substr(2, 2), 16);
        const b = parseInt(color.substr(4, 2), 16);

        const newR = Math.min(255, Math.round(r + (255 - r) * percent));
        const newG = Math.min(255, Math.round(g + (255 - g) * percent));
        const newB = Math.min(255, Math.round(b + (255 - b) * percent));

        return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
    }

    /**
     * QUALITY FIX 4: Convert path commands to polyline points
     */
    pathToPoints(pathD) {
        const points = [];
        const parts = pathD.split(/[ML]/);
        for (let i = 1; i < parts.length; i++) {
            const coords = parts[i].trim().split(' ');
            if (coords.length >= 2) {
                points.push(`${coords[0]},${coords[1]}`);
            }
        }
        return points.join(' ');
    }

    /**
     * QUALITY FIX 2: Calculate optimal text color for readability
     */
    getContrastTextColor(backgroundColor) {
        if (!backgroundColor || backgroundColor === "none") {
            return "#000000";
        }

        // Parse hex color
        let hex = backgroundColor.replace("#", "");
        if (hex.length === 3) {
            hex = hex.split('').map(c => c + c).join('');
        }

        const r = parseInt(hex.substr(0, 2), 16) / 255.0;
        const g = parseInt(hex.substr(2, 2), 16) / 255.0;
        const b = parseInt(hex.substr(4, 2), 16) / 255.0;

        // Calculate luminance (relative luminance formula)
        const luminance = 0.299 * r + 0.587 * g + 0.114 * b;

        // Dark background (luminance < 0.5) needs white text
        // Light background (luminance >= 0.5) needs dark text
        return luminance < 0.5 ? "#FFFFFF" : "#333333";
    }

    // Utility functions

    /**
     * Resolve color from Visio format to CSS
     */
    resolveColor(val, themeColors = null) {
        if (!val) return "";
        
        val = val.toString().trim();
        
        // THEMEVAL formulas
        if (val.includes("THEMEVAL") || val.includes("THEMEGUARD")) {
            if (themeColors || this.themeColors) {
                const colors = themeColors || this.themeColors;
                const match = val.match(/THEMEVAL\s*\(\s*"?(\w+)"?/i);
                if (match) {
                    const key = match[1].toLowerCase();
                    return colors[key] || colors[parseInt(key)] || "";
                }
            }
            return "";
        }
        
        if (val === "Inh" || val.startsWith("=") || val.includes("THEME")) {
            return "";
        }
        
        // Hex color
        if (val.startsWith("#")) return val;
        
        // RGB function
        const rgbMatch = val.match(/RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (rgbMatch) {
            const r = parseInt(rgbMatch[1]);
            const g = parseInt(rgbMatch[2]);
            const b = parseInt(rgbMatch[3]);
            return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
        }
        
        // HSL function
        const hslMatch = val.match(/HSL\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (hslMatch) {
            return this.hslToRgb(parseInt(hslMatch[1]), parseInt(hslMatch[2]), parseInt(hslMatch[3]));
        }
        
        // Color index
        const index = parseInt(val);
        if (!isNaN(index) && this.VISIO_COLORS[index]) {
            return this.VISIO_COLORS[index];
        }
        
        return "";
    }

    /**
     * Convert HSL to RGB
     */
    hslToRgb(h, s, l) {
        const hf = (h / 255.0) * 360.0;
        const sf = s / 255.0;
        const lf = l / 255.0;
        
        let r, g, b;
        
        if (sf === 0) {
            r = g = b = lf;
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1/6) return p + (q - p) * 6 * t;
                if (t < 1/2) return q;
                if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
                return p;
            };
            
            const q = lf < 0.5 ? lf * (1 + sf) : lf + sf - lf * sf;
            const p = 2 * lf - q;
            const hn = hf / 360.0;
            
            r = hue2rgb(p, q, hn + 1/3);
            g = hue2rgb(p, q, hn);
            b = hue2rgb(p, q, hn - 1/3);
        }
        
        const rInt = Math.round(r * 255);
        const gInt = Math.round(g * 255);
        const bInt = Math.round(b * 255);
        
        return `#${rInt.toString(16).padStart(2, '0')}${gInt.toString(16).padStart(2, '0')}${bInt.toString(16).padStart(2, '0')}`;
    }

    /**
     * Get dash array for line pattern
     */
    getDashArray(pattern, weight) {
        if (pattern === 0) return "none";
        
        let p = this.LINE_PATTERNS[pattern] || "";
        if (!p || p === "none") {
            if (pattern >= 2 && pattern <= 23) {
                if (pattern % 3 === 0) p = "1,2";
                else if (pattern % 3 === 1) p = "6,3";
                else p = "6,3,1,3";
            } else {
                return "";
            }
        }
        
        const scale = Math.max(weight, 0.5);
        const parts = p.split(",").map(x => (parseFloat(x) * scale).toFixed(1));
        return parts.join(",");
    }

    /**
     * Check if shape is a container
     */
    isContainer(shape) {
        const user = shape.user || {};
        const structureType = user.msvStructureType?.Value || "";
        if (structureType === "Container") return true;
        
        const shapeName = (shape.name_u || shape.name || "").toLowerCase();
        return ["dash square", "container", "swimlane"].some(kw => shapeName.includes(kw));
    }

    /**
     * Get shape z-order for sorting
     */
    shapeZOrder(shape) {
        return this.isContainer(shape) ? 0 : 1;
    }

    /**
     * Map font family
     */
    mapFontFamily(fontName) {
        if (!fontName || fontName === "Themed") {
            return "Noto Sans, sans-serif";
        }
        const key = fontName.toLowerCase().trim();
        return this.FONT_MAP[key] || `${fontName}, Noto Sans, sans-serif`;
    }

    /**
     * Build style string
     */
    buildStyleString(style) {
        const parts = [];
        
        parts.push(`fill="${style.fill}"`);
        parts.push(`stroke="${style.stroke}"`);
        parts.push(`stroke-width="${style.strokeWidth.toFixed(2)}"`);
        
        if (style.fillOpacity < 0.99) {
            parts.push(`fill-opacity="${style.fillOpacity.toFixed(2)}"`);
        }
        if (style.strokeOpacity < 0.99) {
            parts.push(`stroke-opacity="${style.strokeOpacity.toFixed(2)}"`);
        }
        if (style.strokeDasharray) {
            parts.push(`stroke-dasharray="${style.strokeDasharray}"`);
        }
        
        return parts.join(" ");
    }

    /**
     * Get cell value
     */
    getCellVal(shape, name, defaultVal = "") {
        return shape.cells?.[name]?.V || defaultVal;
    }

    /**
     * Get cell value as float
     */
    getCellFloat(shape, name, defaultVal = 0.0) {
        return this.safeFloat(this.getCellVal(shape, name), defaultVal);
    }

    /**
     * Safe float conversion
     */
    safeFloat(val, defaultVal = 0.0) {
        if (val === null || val === undefined) return defaultVal;
        const num = parseFloat(val);
        return isNaN(num) ? defaultVal : num;
    }

    /**
     * Resolve theme colors for shape styling
     */
    resolveThemeColors(shape, lineColor, fillForegnd, fillBkgnd) {
        // QuickStyle color resolution
        const qsFillColorVal = this.getCellVal(shape, "QuickStyleFillColor");
        if (this.themeColors && qsFillColorVal) {
            const qsFillColor = parseInt(this.safeFloat(qsFillColorVal, -1));
            const themeFill = this.resolveQuickStyleColor(qsFillColor);
            
            // Resolve THEMEVAL formulas
            const ffFormula = shape.cells?.FillForegnd?.F || "";
            const fbFormula = shape.cells?.FillBkgnd?.F || "";
            
            if (ffFormula.includes("THEMEVAL") && ffFormula.includes("FillColor") && themeFill) {
                fillForegnd = themeFill;
            }
            if (fbFormula.includes("THEMEVAL") && fbFormula.includes("FillColor2") && themeFill) {
                fillBkgnd = this.lightenColor(themeFill, 0.85);
            }
            
            // If FillForegnd is absent but QuickStyleFillColor exists
            if (!fillForegnd && !ffFormula && themeFill && !this.isBlack(themeFill)) {
                fillForegnd = themeFill;
            }
        }
        
        // GUARD placeholders - replace magenta with theme accent
        const defaultAccent = "#5B9BD5";
        const ffFormula = shape.cells?.FillForegnd?.F || "";
        const fbFormula = shape.cells?.FillBkgnd?.F || "";
        const lcFormula = shape.cells?.LineColor?.F || "";
        
        if (ffFormula.includes("GUARD") && fillForegnd === "#FF00FF") {
            fillForegnd = this.themeColors.accent1 || defaultAccent;
        }
        if (fbFormula.includes("GUARD") && fillBkgnd === "#FF00FF") {
            fillBkgnd = this.themeColors.accent1 || defaultAccent;
        }
        
        // THEMEVAL resolving to black - use accent instead
        if (ffFormula.includes("THEMEVAL") && this.isBlack(fillForegnd)) {
            fillForegnd = this.themeColors.accent1 || defaultAccent;
        }
        if (fbFormula.includes("THEMEVAL") && this.isBlack(fillBkgnd)) {
            fillBkgnd = this.themeColors.accent1 || defaultAccent;
        }
        
        // GUARD(0) in stencils - replace with accent
        if (ffFormula.includes("GUARD") && this.isBlack(fillForegnd)) {
            fillForegnd = this.themeColors.accent1 || defaultAccent;
        }
        if (fbFormula.includes("GUARD") && this.isBlack(fillBkgnd)) {
            fillBkgnd = this.themeColors.accent1 || defaultAccent;
        }
        
        // Line color theme resolution
        const qsLineColorVal = this.getCellVal(shape, "QuickStyleLineColor");
        if (this.themeColors && qsLineColorVal) {
            const qsLineColor = parseInt(this.safeFloat(qsLineColorVal, -1));
            if (qsLineColor >= 0) {
                const themeLine = this.resolveQuickStyleColor(qsLineColor);
                if (themeLine && lcFormula.includes("THEMEVAL")) {
                    lineColor = themeLine;
                }
            }
        }
        
        return { lineColor, fillForegnd, fillBkgnd };
    }

    /**
     * Resolve QuickStyle color index to theme color
     */
    resolveQuickStyleColor(qsIndex) {
        const qsMap = {
            0: "dk1", 1: "lt1", 2: "dk2", 3: "lt2",
            4: "accent1", 5: "accent2", 6: "accent3",
            7: "accent4", 8: "accent5", 9: "accent6",
            100: "dk1", 101: "lt1", 102: "dk2",
            103: "accent1", 104: "accent2", 105: "accent3",
            106: "accent4", 107: "accent5", 108: "accent6"
        };
        
        const name = qsMap[qsIndex];
        return (name && this.themeColors[name]) || this.themeColors.accent1 || "";
    }

    /**
     * Lighten a color by blending towards white
     */
    lightenColor(hexColor, factor = 0.7) {
        hexColor = hexColor.trim().replace('#', '');
        if (hexColor.length !== 6) return "#E8E8E8";
        
        try {
            let r = parseInt(hexColor.substring(0, 2), 16);
            let g = parseInt(hexColor.substring(2, 4), 16);
            let b = parseInt(hexColor.substring(4, 6), 16);
            
            r = Math.round(r + (255 - r) * factor);
            g = Math.round(g + (255 - g) * factor);
            b = Math.round(b + (255 - b) * factor);
            
            return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
        } catch (e) {
            return "#E8E8E8";
        }
    }

    /**
     * Check if color is black or near-black
     */
    isBlack(color) {
        if (!color) return false;
        const c = color.trim().toUpperCase();
        return ["#000000", "#000", "0"].includes(c);
    }

    /**
     * Check if color is dark
     */
    isDarkColor(color) {
        if (!color || color === "none") return false;
        const c = color.trim().replace('#', '');
        if (c.length === 6) {
            try {
                const r = parseInt(c.substring(0, 2), 16);
                const g = parseInt(c.substring(2, 4), 16);
                const b = parseInt(c.substring(4, 6), 16);
                const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0;
                return lum < 0.4;
            } catch (e) {}
        }
        return false;
    }

    /**
     * Render group shape
     */
    renderGroupShape(shape, pageH, usedMarkers, gradients, hasShadow, textLayer, pageRels) {
        const elements = [];
        const transform = this.computeTransform(shape, pageH);
        const wInch = this.getCellFloat(shape, "Width");
        const hInch = this.getCellFloat(shape, "Height");
        let groupH = Math.abs(hInch);
        
        // Estimate group height from sub-shapes if needed
        if (groupH < 1e-6 && shape.sub_shapes?.length > 0) {
            let maxSubY = 0;
            for (const sub of shape.sub_shapes) {
                const subPy = this.safeFloat(sub.cells?.PinY?.V);
                const subH = Math.abs(this.safeFloat(sub.cells?.Height?.V));
                maxSubY = Math.max(maxSubY, subPy + subH / 2);
            }
            if (maxSubY > 0) {
                groupH = maxSubY;
            }
        }
        
        const groupMasterId = shape.master || "";
        elements.push(`<g transform="${transform}">`);
        
        // Render group's own geometry if present
        if (shape.geometry?.length > 0) {
            const style = this.getShapeStyle(shape, gradients, hasShadow);
            const masterW = shape._master_w || 0;
            const masterH = shape._master_h || 0;
            
            for (const geo of shape.geometry) {
                const pathD = this.geometryToPath(geo, wInch, hInch, masterW, masterH);
                if (!pathD) continue;
                
                let geoFill = style.fill;
                let geoStroke = style.stroke;
                
                if (geo.no_fill) geoFill = "none";
                if (geo.no_line) geoStroke = "none";
                
                const geoStyle = this.buildStyleString({
                    fill: geoFill, stroke: geoStroke, strokeWidth: style.strokeWidth,
                    fillOpacity: style.fillOpacity, strokeOpacity: style.strokeOpacity,
                    strokeDasharray: style.strokeDasharray
                });
                
                elements.push(`<path d="${pathD}" ${geoStyle}${style.shadowAttr}/>`);
            }
        }
        
        // Render embedded image for group
        if (shape.foreign_data && this.media) {
            const wPx = Math.abs(wInch) * this.INCH_TO_PX;
            const hPx = Math.abs(hInch) * this.INCH_TO_PX;
            elements.push(...this.renderEmbeddedImage(shape, pageRels, wPx, hPx, ""));
        }
        
        // Render sub-shapes
        for (const sub of shape.sub_shapes || []) {
            // Ensure sub-shape inherits master ID from parent group if not set
            if (!sub.master && shape.master) {
                sub.master = shape.master;
            }
            
            const subElements = this.renderShapeToSvg(
                sub, groupH, usedMarkers, gradients, hasShadow, textLayer, pageRels
            );
            elements.push(...subElements);
        }
        
        elements.push('</g>');
        
        // Render group text
        if (shape.text) {
            const wPx = Math.abs(wInch) * this.INCH_TO_PX;
            const hPx = groupH * this.INCH_TO_PX;
            this.appendTextSvg(textLayer, shape, pageH, wPx, hPx);
        }
        
        return elements;
    }

    /**
     * Render embedded image
     */
    renderEmbeddedImage(shape, pageRels, wPx, hPx, transform) {
        const elements = [];
        const fd = shape.foreign_data;
        if (!fd || !this.media) return elements;
        
        let imgHref = null;
        
        if (fd.rel_id && pageRels[fd.rel_id]) {
            const target = pageRels[fd.rel_id];
            const imgName = target.split("/").pop();
            if (imgName && this.media[imgName]) {
                imgHref = this.imageToDataUri(this.media[imgName], imgName);
            }
        } else if (fd.data) {
            // Inline data
            const extMap = { "PNG": ".png", "JPEG": ".jpeg", "BMP": ".bmp", "GIF": ".gif", "TIFF": ".tiff" };
            const comp = (fd.compression || "PNG").toUpperCase();
            const fakeExt = extMap[comp] || ".png";
            
            try {
                // Assuming fd.data is base64
                const raw = this.base64ToArrayBuffer(fd.data);
                const fname = `inline_${shape.id}${fakeExt}`;
                imgHref = this.imageToDataUri(new Uint8Array(raw), fname);
            } catch (e) {}
        }
        
        if (imgHref) {
            const imgW = this.getCellFloat(shape, "ImgWidth") || this.getCellFloat(shape, "Width");
            const imgH = this.getCellFloat(shape, "ImgHeight") || this.getCellFloat(shape, "Height");
            const imgOffX = this.getCellFloat(shape, "ImgOffsetX");
            const imgOffY = this.getCellFloat(shape, "ImgOffsetY");
            
            let imgWPx = imgW * this.INCH_TO_PX;
            let imgHPx = imgH * this.INCH_TO_PX;
            
            // Enforce minimum size
            if (imgWPx > 0 && imgHPx > 0 && (imgWPx < 24 || imgHPx < 24)) {
                const scale = Math.max(24.0 / imgWPx, 24.0 / imgHPx);
                imgWPx *= scale;
                imgHPx *= scale;
            }
            
            const imgXPx = imgOffX * this.INCH_TO_PX;
            const imgYPx = imgOffY * this.INCH_TO_PX;
            
            elements.push(
                `<image x="${imgXPx.toFixed(2)}" y="${imgYPx.toFixed(2)}" ` +
                `width="${imgWPx.toFixed(2)}" height="${imgHPx.toFixed(2)}" ` +
                `xlink:href="${imgHref}" preserveAspectRatio="xMidYMid meet" ` +
                `${transform ? `transform="${transform}"` : ""}/>`
            );
        }
        
        return elements;
    }

    /**
     * Convert image bytes to data URI
     */
    imageToDataUri(data, filename) {
        const ext = filename.split('.').pop()?.toLowerCase() || "png";
        let mime = this.IMAGE_MIMETYPES[`.${ext}`] || "image/png";
        
        // Convert Uint8Array to base64
        const base64 = this.arrayBufferToBase64(data.buffer || data);
        return `data:${mime};base64,${base64}`;
    }

    /**
     * Convert ArrayBuffer to base64
     */
    arrayBufferToBase64(buffer) {
        const bytes = new Uint8Array(buffer);
        let binary = '';
        for (let i = 0; i < bytes.length; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }

    /**
     * Convert base64 to ArrayBuffer
     */
    base64ToArrayBuffer(base64) {
        const binary = atob(base64);
        const buffer = new ArrayBuffer(binary.length);
        const bytes = new Uint8Array(buffer);
        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }
        return buffer;
    }

    /**
     * Render connector with geometry
     */
    renderConnectorWithGeometry(shape, pageH, strokeWidth, lineColor, dashArray, markerAttrs) {
        const elements = [];
        
        if (!shape.geometry || shape.geometry.length === 0) {
            return elements;
        }
        
        // Get connector transform
        const transform = this.computeTransform(shape, pageH);
        
        // Convert each geometry section to path
        for (const geo of shape.geometry) {
            if (geo.no_show) continue;
            
            const wInch = this.getCellFloat(shape, "Width");
            const hInch = this.getCellFloat(shape, "Height");
            const pathD = this.geometryToPath(geo, wInch, hInch, 0, 0);
            
            if (!pathD) continue;
            
            elements.push(
                `<path d="${pathD}" fill="none" stroke="${lineColor}" stroke-width="${strokeWidth.toFixed(2)}"` +
                (dashArray ? ` stroke-dasharray="${dashArray}"` : '') +
                markerAttrs + ` transform="${transform}"/>`
            );
        }
        
        // FIX: Add connector text labels with background for geometry connectors
        if (shape.text && shape.text.trim()) {
            // Calculate midpoint from connector endpoints
            const beginX = this.safeFloat(this.getCellVal(shape, "BeginX")) * this.INCH_TO_PX;
            const beginY = (pageH - this.safeFloat(this.getCellVal(shape, "BeginY"))) * this.INCH_TO_PX;
            const endX = this.safeFloat(this.getCellVal(shape, "EndX")) * this.INCH_TO_PX;
            const endY = (pageH - this.safeFloat(this.getCellVal(shape, "EndY"))) * this.INCH_TO_PX;
            
            const midX = (beginX + endX) / 2;
            const midY = (beginY + endY) / 2;
            const text = shape.text.trim();
            const fontSize = 8;
            const padding = 2;
            const textWidth = text.length * fontSize * 0.5;
            const textHeight = fontSize + padding * 2;
            
            // Semi-transparent background
            elements.push(
                '<rect x="' + (midX - textWidth/2 - padding) + '" y="' + (midY - textHeight/2) + 
                '" width="' + (textWidth + padding*2) + '" height="' + textHeight + 
                '" fill="white" fill-opacity="0.55" rx="1"/>'
            );
            
            // Text
            elements.push(
                '<text x="' + midX.toFixed(2) + '" y="' + midY.toFixed(2) + 
                '" text-anchor="middle" dominant-baseline="central" font-family="Noto Sans, sans-serif" font-size="' + 
                fontSize + '" fill="#000000">' + this.escapeXml(text) + '</text>'
            );
        }
        
        return elements;
    }

    /**
     * Escape XML characters
     */
    escapeXml(text) {
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#x27;');
    }
}

// Export for browser use
if (typeof window !== 'undefined') {
    window.VisioConverter = VisioConverter;
}

// Export for Node.js use
if (typeof module !== 'undefined' && module.exports) {
    module.exports = VisioConverter;
}