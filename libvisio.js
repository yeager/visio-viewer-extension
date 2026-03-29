/**
 * Pure JavaScript VSDX to SVG converter
 * Ported from libvisio-ng Python implementation
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
        
        this.INCH_TO_PX = 72.0;
        
        // XML namespaces
        this.NS = {
            v: "http://schemas.microsoft.com/office/visio/2012/main",
            r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        };
    }

    /**
     * Load VSDX file from ArrayBuffer
     */
    async loadFromArrayBuffer(buffer) {
        this.zipFile = await JSZip.loadAsync(buffer);
        
        // Parse file structure
        await this.parseMasterShapes();
        await this.parseTheme();
        await this.extractMedia();
        
        // Parse pages
        const pageFiles = this.getPageFiles();
        this.pages = [];
        this.pageNames = [];
        
        for (let i = 0; i < pageFiles.length; i++) {
            const pageFile = pageFiles[i];
            try {
                const pageXml = await this.zipFile.file(pageFile).async("text");
                const shapes = this.parseShapes(pageXml);
                const { width, height, scale } = this.parsePageDimensions(pageXml);
                
                this.pages.push({
                    shapes,
                    width,
                    height,
                    scale,
                    file: pageFile
                });
                
                // Extract page name (default to "Page N")
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
        return this.shapesToSvg(page.shapes, page.width, page.height);
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
        return files.sort(); // Ensure correct order
    }

    /**
     * Parse master shapes (templates/stencils)
     */
    async parseMasterShapes() {
        this.masters = {};
        
        Object.keys(this.zipFile.files).forEach(async name => {
            if (name.match(/^visio\/masters\/master\d+\.xml$/)) {
                try {
                    const xml = await this.zipFile.file(name).async("text");
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(xml, "text/xml");
                    // TODO: Parse master shape details
                    // For now, just store the XML for reference
                    const masterId = name.match(/master(\d+)/)?.[1];
                    if (masterId) {
                        this.masters[masterId] = { xml, doc };
                    }
                } catch (error) {
                    console.warn(`Failed to parse master ${name}:`, error);
                }
            }
        });
    }

    /**
     * Parse theme colors
     */
    async parseTheme() {
        this.themeColors = {};
        
        try {
            const themeFile = this.zipFile.file("visio/theme/theme1.xml");
            if (themeFile) {
                const xml = await themeFile.async("text");
                const parser = new DOMParser();
                const doc = parser.parseFromString(xml, "text/xml");
                
                // Extract theme colors (simplified)
                const colorScheme = doc.querySelector("colorScheme");
                if (colorScheme) {
                    const colors = colorScheme.querySelectorAll("*[val]");
                    colors.forEach((color, index) => {
                        const val = color.getAttribute("val");
                        if (val && val.length === 6) {
                            this.themeColors[index.toString()] = "#" + val;
                        }
                    });
                }
            }
        } catch (error) {
            console.warn("Failed to parse theme:", error);
        }
    }

    /**
     * Extract media files (images, etc.)
     */
    async extractMedia() {
        this.media = {};
        
        Object.keys(this.zipFile.files).forEach(async name => {
            if (name.startsWith("visio/media/")) {
                try {
                    const data = await this.zipFile.file(name).async("base64");
                    const filename = name.split("/").pop();
                    this.media[filename] = `data:image/png;base64,${data}`;
                } catch (error) {
                    console.warn(`Failed to extract media ${name}:`, error);
                }
            }
        });
    }

    /**
     * Parse shapes from page XML
     */
    parseShapes(pageXml) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(pageXml, "text/xml");
        
        const shapes = [];
        
        // Find the Shapes container (namespace-aware)
        const shapesContainer = doc.querySelector("Shapes");
        if (!shapesContainer) {
            return shapes;
        }
        
        const shapeNodes = shapesContainer.querySelectorAll("Shape");
        
        shapeNodes.forEach(node => {
            const shape = this.parseShape(node);
            if (shape) {
                shapes.push(shape);
            }
        });
        
        return shapes;
    }

    /**
     * Parse single shape node
     */
    parseShape(shapeNode) {
        const shape = {
            id: shapeNode.getAttribute("ID") || "",
            name: shapeNode.getAttribute("Name") || "",
            name_u: shapeNode.getAttribute("NameU") || "",
            type: shapeNode.getAttribute("Type") || "",
            master: shapeNode.getAttribute("Master") || "",
            cells: {},
            sections: {}
        };

        // Parse all sections recursively
        this.parseShapeSections(shapeNode, shape);

        // Extract basic text content if available
        const textNode = shapeNode.querySelector("Text");
        if (textNode) {
            shape.text = textNode.textContent || "";
        }

        return shape;
    }

    /**
     * Parse shape sections (Cell, Geom, etc.)
     */
    parseShapeSections(element, shape) {
        // Parse direct cells
        const cells = element.querySelectorAll(":scope > Cell");
        cells.forEach(cell => {
            const name = cell.getAttribute("N");
            const value = cell.getAttribute("V") || cell.getAttribute("F") || "";
            if (name) {
                shape.cells[name] = { V: value };
            }
        });

        // Parse geometry sections
        const geoms = element.querySelectorAll(":scope > Geom");
        geoms.forEach((geom, index) => {
            const geomKey = `Geom${index}`;
            shape.sections[geomKey] = this.parseGeometry(geom);
        });

        // Parse text sections
        const textSections = element.querySelectorAll(":scope > Text, :scope > Char, :scope > Para");
        textSections.forEach(section => {
            const tagName = section.tagName;
            if (!shape.sections[tagName]) {
                shape.sections[tagName] = [];
            }
            
            const sectionData = {};
            const sectionCells = section.querySelectorAll("Cell");
            sectionCells.forEach(cell => {
                const name = cell.getAttribute("N");
                const value = cell.getAttribute("V") || cell.getAttribute("F") || "";
                if (name) {
                    sectionData[name] = { V: value };
                }
            });
            
            shape.sections[tagName].push(sectionData);
        });

        // Parse nested shapes (for groups)
        const nestedShapes = element.querySelectorAll(":scope > Shapes > Shape");
        if (nestedShapes.length > 0) {
            shape.shapes = [];
            nestedShapes.forEach(nestedShape => {
                shape.shapes.push(this.parseShape(nestedShape));
            });
        }
    }

    /**
     * Parse geometry section (MoveTo, LineTo, etc.)
     */
    parseGeometry(geomNode) {
        const geometry = [];
        
        geomNode.childNodes.forEach(node => {
            if (node.nodeType === 1 /* Node.ELEMENT_NODE */) {
                const type = node.nodeName;
                const row = { type };
                
                const cells = node.querySelectorAll("Cell");
                cells.forEach(cell => {
                    const name = cell.getAttribute("N");
                    const value = cell.getAttribute("V") || cell.getAttribute("F") || "0";
                    if (name) {
                        row[name] = this.safeFloat(value);
                    }
                });
                
                geometry.push(row);
            }
        });
        
        return geometry;
    }

    /**
     * Parse page dimensions and properties
     */
    parsePageDimensions(pageXml) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(pageXml, "text/xml");
        
        const pageSheet = doc.querySelector("PageSheet");
        let width = 8.5, height = 11.0, scale = 1.0; // defaults
        
        if (pageSheet) {
            const cells = pageSheet.querySelectorAll("Cell");
            cells.forEach(cell => {
                const name = cell.getAttribute("N");
                const value = parseFloat(cell.getAttribute("V") || "0");
                if (name === "PageWidth") width = value;
                if (name === "PageHeight") height = value;
                if (name === "DrawingScale") scale = value || 1.0;
            });
        }
        
        return { width, height, scale };
    }

    /**
     * Extract page name
     */
    extractPageName(pageXml) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(pageXml, "text/xml");
        
        const pageSheet = doc.querySelector("PageSheet");
        if (pageSheet) {
            const nameCell = pageSheet.querySelector('Cell[N="PageName"]');
            if (nameCell) {
                return nameCell.getAttribute("V") || "";
            }
        }
        
        return null;
    }

    /**
     * Convert shapes to SVG
     */
    shapesToSvg(shapes, pageW, pageH) {
        const pageWPx = pageW * this.INCH_TO_PX;
        const pageHPx = pageH * this.INCH_TO_PX;
        
        // Calculate viewBox (simplified - use full page for now)
        const vbX = 0, vbY = 0, vbW = pageWPx, vbH = pageHPx;
        
        // Initialize gradients collection
        this.customGradients = [];
        
        let svg = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            `<svg xmlns="http://www.w3.org/2000/svg" `,
            `xmlns:xlink="http://www.w3.org/1999/xlink" `,
            `width="${pageWPx}" height="${pageHPx}" `,
            `viewBox="${vbX} ${vbY} ${vbW} ${vbH}">`,
            // Arrow markers and filters
            `<defs>`,
            `  <marker id="arrowEnd" markerWidth="10" markerHeight="7" `,
            `          refX="9" refY="3.5" orient="auto">`,
            `    <polygon points="0 0, 10 3.5, 0 7" fill="#666666"/>`,
            `  </marker>`,
            `  <marker id="arrowStart" markerWidth="10" markerHeight="7" `,
            `          refX="1" refY="3.5" orient="auto">`,
            `    <polygon points="10 0, 0 3.5, 10 7" fill="#666666"/>`,
            `  </marker>`,
            `  <filter id="dropShadow" x="-50%" y="-50%" width="200%" height="200%">`,
            `    <feDropShadow dx="2" dy="2" stdDeviation="2" flood-opacity="0.3"/>`,
            `  </filter>`,
            `  <linearGradient id="defaultGradient" x1="0%" y1="0%" x2="0%" y2="100%">`,
            `    <stop offset="0%" style="stop-color:#E8E8E8;stop-opacity:1" />`,
            `    <stop offset="100%" style="stop-color:#C8C8C8;stop-opacity:1" />`,
            `  </linearGradient>`,
            ...this.customGradients,
            `</defs>`,
            `<rect x="${vbX}" y="${vbY}" width="${vbW}" height="${vbH}" fill="white"/>`,
        ];

        // Render shapes
        shapes.forEach(shape => {
            const svgElements = this.renderShapeToSvg(shape, pageH);
            svg.push(...svgElements);
        });

        svg.push('</svg>');
        return svg.join('\n');
    }

    /**
     * Render single shape to SVG elements
     */
    renderShapeToSvg(shape, pageH) {
        const elements = [];
        
        // Get basic properties
        const pinX = this.safeFloat(shape.cells.PinX?.V) * this.INCH_TO_PX;
        const pinY = (pageH - this.safeFloat(shape.cells.PinY?.V)) * this.INCH_TO_PX;
        const width = this.safeFloat(shape.cells.Width?.V) * this.INCH_TO_PX;
        const height = this.safeFloat(shape.cells.Height?.V) * this.INCH_TO_PX;
        
        // Check for line/connector shapes (BeginX/EndX instead of Width/Height)
        const beginX = shape.cells.BeginX?.V;
        const endX = shape.cells.EndX?.V;
        if (beginX !== undefined && endX !== undefined) {
            let bx = this.safeFloat(beginX) * this.INCH_TO_PX;
            let by = (pageH - this.safeFloat(shape.cells.BeginY?.V)) * this.INCH_TO_PX;
            let ex = this.safeFloat(endX) * this.INCH_TO_PX;
            let ey = (pageH - this.safeFloat(shape.cells.EndY?.V)) * this.INCH_TO_PX;
            
            const stroke = this.resolveColor(shape.cells.LineColor?.V) || "#666666";
            const strokeWidth = this.safeFloat(shape.cells.LineWeight?.V) * this.INCH_TO_PX || 1;
            
            // Check for arrows
            let markerAttrs = "";
            const beginArrow = this.safeFloat(shape.cells.BeginArrow?.V);
            const endArrow = this.safeFloat(shape.cells.EndArrow?.V);
            
            if (beginArrow > 0) {
                markerAttrs += ` marker-start="url(#arrowStart)"`;
            }
            if (endArrow > 0) {
                markerAttrs += ` marker-end="url(#arrowEnd)"`;
            }
            
            // Add default arrows on flow-like connectors (vertical downward lines in diagrams)
            if (beginArrow === 0 && endArrow === 0) {
                const isVertical = Math.abs(bx - ex) < 5; // Roughly vertical
                const isDownward = by < ey; // Going down (positive Y direction)
                
                if (isVertical && isDownward) {
                    markerAttrs += ` marker-end="url(#arrowEnd)"`;
                }
            }
            
            // TODO: Improve connector routing to avoid shapes
            // For now, use simple orthogonal routing for better visual appeal
            const isOrthogonal = shape.cells.ShapeRouteStyle?.V === "16" || 
                               shape.cells.ConnectorLayout?.V === "1";
            
            if (isOrthogonal && Math.abs(bx - ex) > 20 && Math.abs(by - ey) > 20) {
                // Create orthogonal path (L-shaped)
                const midX = (bx + ex) / 2;
                elements.push(
                    `<path d="M ${bx.toFixed(1)} ${by.toFixed(1)} L ${midX.toFixed(1)} ${by.toFixed(1)} L ${midX.toFixed(1)} ${ey.toFixed(1)} L ${ex.toFixed(1)} ${ey.toFixed(1)}" ` +
                    `fill="none" stroke="${stroke}" stroke-width="${strokeWidth}"${markerAttrs}/>`
                );
            } else {
                // Standard straight line
                elements.push(
                    `<line x1="${bx.toFixed(1)}" y1="${by.toFixed(1)}" x2="${ex.toFixed(1)}" y2="${ey.toFixed(1)}" ` +
                    `stroke="${stroke}" stroke-width="${strokeWidth}"${markerAttrs}/>`
                );
            }
            return elements;
        }

        if (width <= 0 || height <= 0) {
            return elements; // Skip invalid shapes
        }

        // Check for geometry sections first (custom paths)
        const hasGeometry = Object.keys(shape.sections || {}).some(key => key.startsWith('Geom'));
        
        // Check for ellipse geometry specifically
        let isEllipse = false;
        if (hasGeometry) {
            Object.values(shape.sections).forEach(section => {
                if (Array.isArray(section)) {
                    const hasEllipticalArc = section.some(row => row.type === 'EllipticalArcTo' || row.type === 'Ellipse');
                    if (hasEllipticalArc) isEllipse = true;
                }
            });
        }
        
        if (isEllipse) {
            // Render as ellipse
            const cx = pinX;
            const cy = pinY;
            const rx = width / 2;
            const ry = height / 2;
            
            let fill = this.resolveColor(shape.cells.FillForegnd?.V) || "#E0E0E0";
            const fillPattern = this.safeFloat(shape.cells.FillPattern?.V);
            
            if (fillPattern >= 25 || shape.cells.FillBkgnd?.V) {
                const bgColor = this.resolveColor(shape.cells.FillBkgnd?.V);
                if (bgColor && bgColor !== fill) {
                    const gradientId = `grad_${shape.id || Math.random().toString(36).substr(2, 9)}`;
                    this.addGradientDefinition(gradientId, fill, bgColor);
                    fill = `url(#${gradientId})`;
                }
            }
            
            const stroke = this.resolveColor(shape.cells.LineColor?.V) || "#000000";
            const strokeWidth = this.safeFloat(shape.cells.LineWeight?.V) * this.INCH_TO_PX || 1;
            
            // Handle drop shadow
            let filterAttrs = "";
            if (shape.cells.ShdwForegnd?.V || shape.cells.ShdwOffsetX?.V || shape.cells.ShdwOffsetY?.V) {
                const shdwOffsetX = this.safeFloat(shape.cells.ShdwOffsetX?.V) * this.INCH_TO_PX;
                const shdwOffsetY = this.safeFloat(shape.cells.ShdwOffsetY?.V) * this.INCH_TO_PX;
                
                if (Math.abs(shdwOffsetX) > 0.5 || Math.abs(shdwOffsetY) > 0.5) {
                    filterAttrs = ` filter="url(#dropShadow)"`;
                }
            }
            
            elements.push(
                `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" ` +
                `fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"${filterAttrs}/>`
            );
        } else if (hasGeometry) {
            // Render geometry-based shape
            Object.keys(shape.sections).forEach(key => {
                if (key.startsWith('Geom')) {
                    const geometry = shape.sections[key];
                    const pathData = this.geometryToPath(geometry, pinX, pinY, width, height);
                    if (pathData) {
                        const fill = this.resolveColor(shape.cells.FillForegnd?.V) || "#E0E0E0";
                        const stroke = this.resolveColor(shape.cells.LineColor?.V) || "#000000";
                        const strokeWidth = this.safeFloat(shape.cells.LineWeight?.V) * this.INCH_TO_PX || 1;
                        
                        elements.push(
                            `<path d="${pathData}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"/>`
                        );
                    }
                }
            });
        } else {
            // Default rectangle shape
            const x = pinX - width / 2;
            const y = pinY - height / 2;
            
            // Handle fill - check for gradients
            let fill = this.resolveColor(shape.cells.FillForegnd?.V) || "#E0E0E0";
            const fillPattern = this.safeFloat(shape.cells.FillPattern?.V);
            
            // Use gradient for certain fill patterns (25+ typically indicates gradients in Visio)
            if (fillPattern >= 25 || shape.cells.FillBkgnd?.V) {
                const bgColor = this.resolveColor(shape.cells.FillBkgnd?.V);
                if (bgColor && bgColor !== fill) {
                    // Create custom gradient
                    const gradientId = `grad_${shape.id || Math.random().toString(36).substr(2, 9)}`;
                    this.addGradientDefinition(gradientId, fill, bgColor);
                    fill = `url(#${gradientId})`;
                }
            }
            
            const stroke = this.resolveColor(shape.cells.LineColor?.V) || "#000000";
            const strokeWidth = this.safeFloat(shape.cells.LineWeight?.V) * this.INCH_TO_PX || 1;
            
            // Handle rounded corners
            let roundingAttrs = "";
            if (shape.cells.Rounding?.V) {
                const rounding = this.safeFloat(shape.cells.Rounding.V) * this.INCH_TO_PX;
                if (rounding > 0) {
                    roundingAttrs = ` rx="${rounding}" ry="${rounding}"`;
                }
            }
            
            // Handle drop shadow
            let filterAttrs = "";
            if (shape.cells.ShdwForegnd?.V || shape.cells.ShdwOffsetX?.V || shape.cells.ShdwOffsetY?.V) {
                const shdwOffsetX = this.safeFloat(shape.cells.ShdwOffsetX?.V) * this.INCH_TO_PX;
                const shdwOffsetY = this.safeFloat(shape.cells.ShdwOffsetY?.V) * this.INCH_TO_PX;
                
                if (Math.abs(shdwOffsetX) > 0.5 || Math.abs(shdwOffsetY) > 0.5) {
                    filterAttrs = ` filter="url(#dropShadow)"`;
                }
            }
            
            elements.push(
                `<rect x="${x}" y="${y}" width="${width}" height="${height}"${roundingAttrs} ` +
                `fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"${filterAttrs}/>`
            );
        }

        // Render text
        const textElements = this.renderShapeText(shape, pinX, pinY, width, height);
        elements.push(...textElements);

        // Render nested shapes (groups)
        if (shape.shapes && shape.shapes.length > 0) {
            // Create group element
            const groupAttrs = [];
            
            // Apply transformation if needed
            const angle = this.safeFloat(shape.cells.Angle?.V);
            if (Math.abs(angle) > 0.001) {
                const degrees = (angle * 180) / Math.PI;
                groupAttrs.push(`transform="rotate(${degrees.toFixed(2)} ${pinX} ${pinY})"`);
            }
            
            if (groupAttrs.length > 0) {
                elements.push(`<g ${groupAttrs.join(' ')}>`);
            } else {
                elements.push(`<g>`);
            }
            
            // Render nested shapes
            shape.shapes.forEach(nestedShape => {
                const nestedElements = this.renderShapeToSvg(nestedShape, pageH);
                elements.push(...nestedElements);
            });
            
            elements.push(`</g>`);
        }
        
        return elements;
    }

    /**
     * Render text for a shape
     */
    renderShapeText(shape, pinX, pinY, width, height) {
        const elements = [];
        
        // Try to get text content from various sources
        let textContent = "";
        
        // Check for direct text in shape attribute
        if (shape.text) {
            textContent = shape.text;
        }
        // Check sections for text
        else if (shape.sections && shape.sections.Text && shape.sections.Text.length > 0) {
            const textSection = shape.sections.Text[0];
            textContent = textSection.Text?.V || textContent;
        }
        // Fallback to cell text
        else if (shape.cells && shape.cells.Text) {
            textContent = shape.cells.Text.V || "";
        }
        
        // Clean up text content
        if (!textContent || typeof textContent !== 'string') {
            return elements;
        }
        
        textContent = textContent.trim();
        if (!textContent) {
            return elements;
        }

        // Get text formatting
        let fontSize = Math.max(10, Math.min(height / 3, width / 6)); // Adaptive font size
        let fontFamily = "Arial, sans-serif";
        let textColor = "#000000";
        let fontWeight = "normal";
        let fontStyle = "normal";
        
        if (shape.sections && shape.sections.Char && shape.sections.Char.length > 0) {
            const charSection = shape.sections.Char[0];
            
            // Font size
            if (charSection.Size && charSection.Size.V) {
                fontSize = this.safeFloat(charSection.Size.V) * this.INCH_TO_PX * 0.75 || fontSize;
            }
            
            // Font color
            if (charSection.Color && charSection.Color.V) {
                textColor = this.resolveColor(charSection.Color.V) || "#000000";
            }
            
            // Font style (bold/italic)
            if (charSection.Style && charSection.Style.V) {
                const style = this.safeFloat(charSection.Style.V);
                if ((style & 1) !== 0) fontWeight = "bold"; // Bit 0 = bold
                if ((style & 2) !== 0) fontStyle = "italic"; // Bit 1 = italic
            }
        }
        
        // Auto-contrast: white text on dark backgrounds
        if (shape.cells && shape.cells.FillForegnd?.V) {
            const bgColor = this.resolveColor(shape.cells.FillForegnd.V);
            if (bgColor && this.isDarkColor(bgColor)) {
                textColor = "#FFFFFF";
            }
        }

        // Text positioning - default to center of shape
        let txtPinX = 0;
        let txtPinY = 0;
        
        if (shape.cells) {
            txtPinX = this.safeFloat(shape.cells.TxtPinX?.V) || 0;
            txtPinY = this.safeFloat(shape.cells.TxtPinY?.V) || 0;
        }
        
        const textX = pinX + (txtPinX * this.INCH_TO_PX);
        const textY = pinY - (txtPinY * this.INCH_TO_PX);
        
        // Handle multi-line text
        const textLines = textContent.split(/\r?\n/);
        const lineHeight = fontSize * 1.2;
        
        textLines.forEach((line, index) => {
            if (!line.trim()) return; // Skip empty lines
            
            const yOffset = (index - (textLines.length - 1) / 2) * lineHeight;
            elements.push(
                `<text x="${textX.toFixed(2)}" y="${(textY + yOffset).toFixed(2)}" text-anchor="middle" dominant-baseline="middle" ` +
                `font-family="${fontFamily}" font-size="${fontSize.toFixed(1)}" font-weight="${fontWeight}" font-style="${fontStyle}" fill="${textColor}">` +
                `${this.escapeXml(line.trim())}</text>`
            );
        });
        
        return elements;
    }

    /**
     * Convert geometry to SVG path
     */
    geometryToPath(geometry, pinX, pinY, width, height) {
        if (!geometry || geometry.length === 0) {
            return null;
        }
        
        const path = [];
        let currentX = 0, currentY = 0;
        
        geometry.forEach(row => {
            // Coordinates are relative to shape bounds (0,0 to 1,1 typically)
            let x = this.safeFloat(row.X);
            let y = this.safeFloat(row.Y);
            
            // Convert to shape coordinate system
            const shapeX = pinX - width/2 + (x * width);
            const shapeY = pinY - height/2 + (y * height);
            
            switch (row.type) {
                case 'MoveTo':
                    path.push(`M ${shapeX.toFixed(2)} ${shapeY.toFixed(2)}`);
                    currentX = shapeX;
                    currentY = shapeY;
                    break;
                    
                case 'LineTo':
                    path.push(`L ${shapeX.toFixed(2)} ${shapeY.toFixed(2)}`);
                    currentX = shapeX;
                    currentY = shapeY;
                    break;
                    
                case 'ArcTo':
                    // For now, approximate with straight line
                    // TODO: Proper arc calculation with A parameter
                    path.push(`L ${shapeX.toFixed(2)} ${shapeY.toFixed(2)}`);
                    currentX = shapeX;
                    currentY = shapeY;
                    break;
                    
                case 'EllipticalArcTo':
                    // Simplified elliptical arc
                    path.push(`L ${shapeX.toFixed(2)} ${shapeY.toFixed(2)}`);
                    currentX = shapeX;
                    currentY = shapeY;
                    break;
                    
                case 'Ellipse':
                    // Draw ellipse using arc commands
                    // Ellipse geometry uses X,Y,A,B where A,B are semi-axes
                    const cx = pinX - width/2 + (this.safeFloat(row.X) * width);
                    const cy = pinY - height/2 + (this.safeFloat(row.Y) * height);
                    const a = this.safeFloat(row.A) * width || width / 2;
                    const b = this.safeFloat(row.B) * height || height / 2;
                    
                    path.push(`M ${cx - a} ${cy}`);
                    path.push(`A ${a} ${b} 0 1 1 ${cx + a} ${cy}`);
                    path.push(`A ${a} ${b} 0 1 1 ${cx - a} ${cy}`);
                    path.push('Z'); // Close path
                    break;
                    
                case 'InfiniteLine':
                    // Draw a line across the shape
                    path.push(`M ${pinX - width/2} ${pinY - height/2}`);
                    path.push(`L ${pinX + width/2} ${pinY + height/2}`);
                    break;
                    
                case 'SplineStart':
                case 'SplineKnot':
                    // Simplified spline rendering as line segments
                    path.push(`L ${shapeX.toFixed(2)} ${shapeY.toFixed(2)}`);
                    currentX = shapeX;
                    currentY = shapeY;
                    break;
            }
        });
        
        if (path.length > 0) {
            // Close path if it's a polygon/shape (not a line)
            const firstPoint = path[0];
            const lastPoint = path[path.length - 1];
            if (firstPoint && lastPoint && 
                firstPoint.startsWith('M') && 
                (lastPoint.startsWith('L') || lastPoint.startsWith('A')) &&
                geometry.length > 2) {
                path.push('Z');
            }
            return path.join(' ');
        }
        
        return null;
    }

    /**
     * Resolve color from Visio format to CSS
     */
    resolveColor(val) {
        if (!val) return null;
        
        val = val.toString().trim();
        
        // Skip inherited values and formulas for now
        if (val === "Inh" || val.startsWith("=") || val.includes("THEME")) {
            return null;
        }
        
        // Already a hex color
        if (val.startsWith("#")) return val;
        
        // RGB function
        const rgbMatch = val.match(/RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (rgbMatch) {
            const r = parseInt(rgbMatch[1]);
            const g = parseInt(rgbMatch[2]);
            const b = parseInt(rgbMatch[3]);
            return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
        }
        
        // HSL function - Visio uses 0-255 range
        const hslMatch = val.match(/HSL\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (hslMatch) {
            return this.hslToRgb(parseInt(hslMatch[1]), parseInt(hslMatch[2]), parseInt(hslMatch[3]));
        }
        
        // Color index (try both int and float)
        const index = parseInt(val);
        if (!isNaN(index) && this.VISIO_COLORS[index]) {
            return this.VISIO_COLORS[index];
        }
        
        // Try float index
        const floatIndex = parseInt(parseFloat(val));
        if (!isNaN(floatIndex) && this.VISIO_COLORS[floatIndex]) {
            return this.VISIO_COLORS[floatIndex];
        }
        
        return null;
    }

    /**
     * Convert HSL (0-255 range) to RGB hex
     */
    hslToRgb(h, s, l) {
        // Normalize to 0-1 range
        const hf = (h / 255.0) * 360.0;
        const sf = s / 255.0;
        const lf = l / 255.0;
        
        let r, g, b;
        
        if (sf === 0) {
            r = g = b = lf; // achromatic
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
        
        return `#${Math.round(r*255).toString(16).padStart(2, '0')}${Math.round(g*255).toString(16).padStart(2, '0')}${Math.round(b*255).toString(16).padStart(2, '0')}`;
    }

    /**
     * Safe float conversion
     */
    safeFloat(val) {
        if (val === null || val === undefined) return 0;
        const num = parseFloat(val);
        return isNaN(num) ? 0 : num;
    }

    /**
     * Check if a color is dark (requires white text)
     */
    isDarkColor(hexColor) {
        if (!hexColor || !hexColor.startsWith('#')) return false;
        
        const rgb = hexColor.substring(1);
        const r = parseInt(rgb.substring(0, 2), 16);
        const g = parseInt(rgb.substring(2, 4), 16);
        const b = parseInt(rgb.substring(4, 6), 16);
        
        // Calculate luminance (simplified)
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b);
        return luminance < 128;
    }

    /**
     * Add gradient definition for dynamic gradients
     */
    addGradientDefinition(gradientId, startColor, endColor) {
        this.customGradients.push(
            `  <linearGradient id="${gradientId}" x1="0%" y1="0%" x2="0%" y2="100%">`,
            `    <stop offset="0%" style="stop-color:${startColor};stop-opacity:1" />`,
            `    <stop offset="100%" style="stop-color:${endColor};stop-opacity:1" />`,
            `  </linearGradient>`
        );
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

// Export for use
window.VisioConverter = VisioConverter;