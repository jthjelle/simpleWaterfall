/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*/
"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import * as d3 from "d3";

// Types
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ITooltipService = powerbi.extensibility.ITooltipService;
import IColorPalette = powerbi.extensibility.IColorPalette;

// Interfaces
interface WaterfallDataPoint {
    category: string;
    delta: number;
    start: number;
    end: number;
    py: number;
    cy: number;
    runningTotal: number;
    color: string;
    isTotal: boolean;
    selectionId: powerbi.visuals.ISelectionId;
    tooltipValues?: { displayName: string, value: string }[];
    isReference?: boolean;
    varianceRef?: number; // Value to compare against (End Total)
    highlight?: boolean; // For highlight data support
    isSpacer?: boolean; // For spacing keys
    originalIndex?: number;
    referenceValue?: number; // Value for Reference Mark
    referenceY?: number; // Calculated Chart Coordinate for Reference Mark
    groupName?: string; // Name of the parent category
    isGroupStart?: boolean; // Is this the first item in the group?
    isGroupEnd?: boolean; // Is this the last item in the group?
    isSingletGroup?: boolean; // If true, this group has only 1 item
    isSubtotal?: boolean; // Is this a Group Subtotal bar?
    visualHeight?: number; // Visual height override for huge bars (Axis Break)
    isBroken?: boolean; // Is this bar broken? (Clamped height)
}

interface WaterfallChartData {
    title: string;
    dataPoints: WaterfallDataPoint[];
    maxValue: number;
    minValue: number;
    pyTotal: number;
    startName: string;
    endName: string;
}

// To avoid TS errors if types are missing
interface CustomVisualDrillAction {
    action: string;
    selectionId: powerbi.visuals.ISelectionId;
}

interface WaterfallViewModel {
    charts: WaterfallChartData[];
    globalMaxValue: number;
    globalMinValue: number;
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private svg: d3.Selection<SVGSVGElement, any, any, any>;
    private mainGroup: d3.Selection<SVGGElement, any, any, any>;
    private xAxisGroup: d3.Selection<SVGGElement, any, any, any>;
    private yAxisGroup: d3.Selection<SVGGElement, any, any, any>;
    private settings: VisualFormattingSettingsModel;
    private currentSelection: powerbi.visuals.ISelectionId[] = [];
    private collapsedGroups: Set<string> = new Set();
    private currentOptions: VisualUpdateOptions; // Store for re-render
    private canDrillDown: boolean = false;
    private events: IVisualEventService;
    private isHighContrast: boolean = false;
    private colorPalette: IColorPalette;
    private licenseOverlay: d3.Selection<HTMLDivElement, any, any, any>;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = this.host.createSelectionManager();
        this.tooltipService = this.host.tooltipService;
        this.target = options.element;

        // Initialize event service for rendering events
        this.events = this.host.eventService;

        // Initialize color palette for high contrast detection
        this.colorPalette = this.host.colorPalette;

        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, null);

        this.svg = d3.select(this.target)
            .append("svg")
            .classed("waterfall-visual", true);



        this.mainGroup = this.svg.append('g').classed('mainGroup', true);
        this.xAxisGroup = this.mainGroup.append('g').classed('xAxis', true);
        this.yAxisGroup = this.mainGroup.append('g').classed('yAxis', true);

        // Add Overlay Container
        this.licenseOverlay = d3.select(options.element)
            .append('div')
            .classed('license-overlay', true);

        this.licenseOverlay.append('p').text("This visual requires a license to view in Power BI Service.");



        // Define gradients/masks
        const defs = this.svg.append("defs");
        const gradient = defs.append("linearGradient")
            .attr("id", "axis-break-gradient")
            .attr("x1", "0%")
            .attr("y1", "100%") // Bottom
            .attr("x2", "0%")
            .attr("y2", "0%"); // Top

        // Bottom - transparent (black in mask)
        gradient.append("stop").attr("offset", "0%").attr("stop-color", "black").attr("stop-opacity", 1);
        // Fade to opaque quickly
        gradient.append("stop").attr("offset", "15%").attr("stop-color", "white").attr("stop-opacity", 1);
        gradient.append("stop").attr("offset", "100%").attr("stop-color", "white").attr("stop-opacity", 1);

        const mask = defs.append("mask")
            .attr("id", "axis-break-mask")
            .attr("maskContentUnits", "objectBoundingBox");

        mask.append("rect")
            .attr("x", 0).attr("y", 0).attr("width", 1).attr("height", 1)
            .attr("fill", "url(#axis-break-gradient)");

        // Handle right-click context menu (required for AppSource certification)
        this.handleContextMenu();
    }

    private handleContextMenu() {
        this.target.addEventListener('contextmenu', (event: MouseEvent) => {
            const target = event.target as Element;
            const datum = d3.select(target).datum() as WaterfallDataPoint | undefined;
            const selectionId = datum && datum.selectionId ? datum.selectionId : {};
            this.selectionManager.showContextMenu(selectionId, {
                x: event.clientX,
                y: event.clientY
            });
            event.preventDefault();
        });
    }

    public update(options: VisualUpdateOptions) {
        this.currentOptions = options; // Save options for re-render
        // Always render the visual first
        this.runUpdate(options);

        // Then Check License (Async)
        this.checkLicensing().then((isAllowed) => {
            if (!isAllowed) {
                this.renderOverlay(true);
            } else {
                this.renderOverlay(false);
            }
        });
    }

    private runUpdate(options: VisualUpdateOptions) {
        // Signal rendering started
        this.events.renderingStarted(options);

        // Landing Page Logic (Zero State)
        if (!options.dataViews || !options.dataViews[0] || !options.dataViews[0].categorical || !options.dataViews[0].categorical.categories || !options.dataViews[0].categorical.values) {
            this.mainGroup.selectAll("*").remove();
            this.mainGroup.append("text")
                .classed("message-text", true)
                .attr("x", options.viewport.width / 2)
                .attr("y", options.viewport.height / 2)
                .attr("text-anchor", "middle")
                .style("font-size", "20px")
                .style("fill", "#666")
                .text("Please add data fields");
            this.events.renderingFinished(options);
            return;
        }

        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

        // Validation: Start & End Columns cannot be the same
        let startCol = this.settings.totalSettings.startTotalColumn.value.value;
        if (!startCol || startCol === "auto" || startCol === "none") startCol = "py";

        let endCol = this.settings.totalSettings.endTotalColumn.value.value;
        if (!endCol || endCol === "auto" || endCol === "none") endCol = "cy";

        if (startCol === endCol) {
            this.mainGroup.selectAll("*").remove();
            this.mainGroup.append("text")
                .attr("x", options.viewport.width / 2)
                .attr("y", options.viewport.height / 2)
                .attr("text-anchor", "middle")
                .style("font-size", "14px")
                .style("fill", "red")
                .text("Start Column and End Column can't be the same value");
            this.events.renderingFinished(options);
            return;
        }


        // Detect high contrast mode (cast to any for compatibility)
        this.isHighContrast = (this.colorPalette as any).isHighContrast || false;

        // Clean main group
        this.mainGroup.selectAll("*").remove();

        const viewModel = this.transformData(options);

        if (!viewModel || viewModel.charts.length === 0) {
            this.host.displayWarningIcon("Data Requirement", "Please drag fields into Category, PY, and TY to visualize data.");
            return;
        } else {
            // Clear warning on success
            this.host.displayWarningIcon("", "");
        }

        const width = options.viewport.width;
        const height = options.viewport.height;

        this.svg.attr("width", width).attr("height", height);



        // Small Multiples Layout
        const smSettings = this.settings.smallMultiplesSettings;
        const isAuto = smSettings.layoutMode.value.value === "auto";
        const numCharts = viewModel.charts.length;

        let rows = 1;
        let cols = 1;

        if (isAuto) {
            if (numCharts > 1) {
                cols = Math.ceil(Math.sqrt(numCharts));
                rows = Math.ceil(numCharts / cols);
            }
        } else {
            rows = Math.max(1, smSettings.rows.value);
            cols = Math.max(1, smSettings.columns.value);
        }

        const uniformYAxis = smSettings.uniformYAxis.value;
        const globalMinMax = uniformYAxis ? { min: viewModel.globalMinValue, max: viewModel.globalMaxValue } : null;

        const gap = 15;
        const cellWidth = (width - (cols - 1) * gap) / cols;
        const cellHeight = (height - (rows - 1) * gap) / rows;

        if (cellWidth < 1 || cellHeight < 1) return;

        viewModel.charts.forEach((chart, i) => {
            const row = Math.floor(i / cols);
            const col = i % cols;

            if (!isAuto && row >= rows) return;

            const x = col * (cellWidth + gap);
            const y = row * (cellHeight + gap);

            try {
                const cellGroup = this.mainGroup.append("g")
                    .attr("transform", `translate(${x}, ${y})`)
                    .classed("sm-cell", true);

                let titleHeight = 0;
                if (chart.title) {
                    titleHeight = 20;
                    cellGroup.append("text")
                        .attr("x", cellWidth / 2)
                        .attr("y", 15)
                        .attr("text-anchor", "middle")
                        .style("font-weight", "bold")
                        .style("font-size", "12px")
                        .text(chart.title);
                }

                // Adjust chart area for title
                const chartGroup = cellGroup.append("g")
                    .attr("transform", `translate(0, ${titleHeight})`);

                const chartHeight = Math.max(0, cellHeight - titleHeight);
                this.renderChart(chartGroup, chart, cellWidth, chartHeight, globalMinMax);
            } catch (e) {

                this.host.displayWarningIcon("Rendering Error", "An error occurred while rendering the visual. Please check your data.");
            }
        });

        // Signal rendering finished
        this.events.renderingFinished(options);
    }

    private formatNumber(value: number, decimalPlaces: number, useThousandsSeparator: boolean, numberScale?: string, thousandsAbbrev?: string, millionsAbbrev?: string): string {
        let scaledValue = value;
        let suffix = "";

        if (numberScale === "auto") {
            const abs = Math.abs(value);
            if (abs >= 1000000) {
                scaledValue = value / 1000000;
                suffix = millionsAbbrev !== undefined ? millionsAbbrev : "M";
            } else if (abs >= 1000) {
                scaledValue = value / 1000;
                suffix = thousandsAbbrev !== undefined ? thousandsAbbrev : "K";
            }
        } else if (numberScale === "thousands") {
            scaledValue = value / 1000;
            suffix = thousandsAbbrev !== undefined ? thousandsAbbrev : "K";
        } else if (numberScale === "millions") {
            scaledValue = value / 1000000;
            suffix = millionsAbbrev !== undefined ? millionsAbbrev : "M";
        }

        const fixed = scaledValue.toFixed(decimalPlaces);
        let result = fixed;

        if (useThousandsSeparator) {
            const parts = fixed.split('.');
            parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
            result = parts.join('.');
        }

        return result + suffix;
    }

    private lightenColor(color: string, percent: number): string {
        let hex = color.replace('#', '');
        if (hex.length === 3) {
            hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        }

        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);

        const newR = Math.min(255, Math.round(r + (255 - r) * (percent / 100)));
        const newG = Math.min(255, Math.round(g + (255 - g) * (percent / 100)));
        const newB = Math.min(255, Math.round(b + (255 - b) * (percent / 100)));

        return '#' + [newR, newG, newB].map(c => c.toString(16).padStart(2, '0')).join('');
    }

    private getSafeEnum(selection: any): string {
        if (!selection) return "";
        if (typeof selection === "object" && selection.value !== undefined) return String(selection.value);
        return String(selection);
    }

    private transformData(options: VisualUpdateOptions): WaterfallViewModel {
        const dataViews = options.dataViews;
        if (!dataViews || !dataViews[0] || !dataViews[0].categorical) {
            return { charts: [], globalMaxValue: 0, globalMinValue: 0 };
        }

        const categorical = dataViews[0].categorical;
        let categoryCols: any[] = [];
        let groupCol = null;

        if (categorical.categories) {
            for (const cat of categorical.categories) {
                if (cat.source.roles["smallMultiples"]) {
                    groupCol = cat;
                } else if (cat.source.roles["category"]) {
                    categoryCols.push(cat);
                }
            }
        }

        // Fallback or Force selection of at least one category if available
        if (categoryCols.length === 0 && categorical.categories && categorical.categories.length > 0) {
            // If no explicitly marked 'category' role found (rare), but categories exist and aren't small multiples
            for (const cat of categorical.categories) {
                if (!cat.source.roles["smallMultiples"]) {
                    categoryCols.push(cat);
                    // If we found one, that's enough for fallback? 
                    // Or should we take all? Standard is to take all grouping levels.
                }
            }
        }

        if (categoryCols.length === 0) return { charts: [], globalMaxValue: 0, globalMinValue: 0 };

        // Use the primary (first) category column for length check
        const numPoints = categoryCols[0].values.length;
        const values = categorical.values;

        let pyValues: any[] = [];
        let pyHighlights: any[] = null;
        let cyValues: any[] = [];
        let cyHighlights: any[] = null;
        let budgetValues: any[] = [];
        let budgetHighlights: any[] = null;
        let tooltipCols: { source: any, values: any[] }[] = [];

        if (values) {
            values.forEach(v => {
                if (v.source.roles["py"]) {
                    pyValues = v.values;
                    pyHighlights = v.highlights;
                }
                if (v.source.roles["cy"]) {
                    cyValues = v.values;
                    cyHighlights = v.highlights;
                }
                if (v.source.roles["budget"]) {
                    budgetValues = v.values;
                    budgetHighlights = v.highlights;
                }
                if (v.source.roles["tooltips"]) tooltipCols.push({ source: v.source, values: v.values });
            });
        }

        const charts: WaterfallChartData[] = [];

        if (groupCol) {
            const groups = new Map<string, number[]>();
            for (let i = 0; i < numPoints; i++) {
                const gVal = groupCol.values[i] ? groupCol.values[i].toString() : "Undefined";
                if (!groups.has(gVal)) groups.set(gVal, []);
                groups.get(gVal).push(i);
            }

            groups.forEach((indices, name) => {
                charts.push(this.generateChartData(indices, categoryCols, groupCol, pyValues, pyHighlights, cyValues, cyHighlights, budgetValues, budgetHighlights, tooltipCols, name));
            });
        } else {
            const indices = Array.from({ length: numPoints }, (_, i) => i);
            charts.push(this.generateChartData(indices, categoryCols, null, pyValues, pyHighlights, cyValues, cyHighlights, budgetValues, budgetHighlights, tooltipCols, ""));
        }

        // Check Drill Availability
        this.canDrillDown = this.checkDrillAvailability(options.dataViews[0]);

        // Small Multiples Sorting
        const smSortBy = this.settings.smallMultiplesSettings.sortBy.value.value;
        const smSortDesc = this.settings.smallMultiplesSettings.sortDirection.value.value === "desc";

        if (charts.length > 1) {
            charts.sort((a, b) => {
                let valA: string | number = 0;
                let valB: string | number = 0;

                if (smSortBy === "name") {
                    valA = a.title || "";
                    valB = b.title || "";
                } else if (smSortBy === "start") {
                    // Use the first data point if it's a Total, otherwise 0
                    const firstA = a.dataPoints[0];
                    valA = (firstA && firstA.isTotal) ? firstA.end : 0;

                    const firstB = b.dataPoints[0];
                    valB = (firstB && firstB.isTotal) ? firstB.end : 0;

                } else if (smSortBy === "end") {
                    // End Value is usually the last bar if it's a total.
                    // Or the cumulative total of the last delta bar.
                    const last = a.dataPoints[a.dataPoints.length - 1];
                    // If Ref column is on, last might be Ref. End Total is before it.
                    // If showEndTotal is on, find it.
                    // Let's look for the *End Total* specifically?
                    // Or just the max accumulated value?
                    // "End" usually implies the final result.
                    // Let's find the data point overlapping with "End Label"?
                    // Or just take the last data point's 'runningTotal'.
                    valA = last ? last.runningTotal : 0;

                    const lastB = b.dataPoints[b.dataPoints.length - 1];
                    valB = lastB ? lastB.runningTotal : 0;
                }

                if (typeof valA === "string" && typeof valB === "string") {
                    return smSortDesc ? valB.localeCompare(valA) : valA.localeCompare(valB);
                } else {
                    return smSortDesc ? Number(valB) - Number(valA) : Number(valA) - Number(valB);
                }
            });
        }

        let globalMax = 0;
        let globalMin = 0;
        charts.forEach(c => {
            globalMax = Math.max(globalMax, c.maxValue);
            globalMin = Math.min(globalMin, c.minValue);
        });

        return {
            charts,
            globalMaxValue: globalMax,
            globalMinValue: globalMin
        };
    }

    private renderChart(
        targetGroup: d3.Selection<SVGGElement, any, any, any>,
        chartData: WaterfallChartData,
        width: number,
        height: number,
        globalMinMax: { min: number, max: number } | null
    ) {
        let xAxisGroup = targetGroup.select<SVGGElement>(".x-axis");
        if (xAxisGroup.empty()) {
            xAxisGroup = targetGroup.append("g").classed("x-axis", true);
        }
        let yAxisGroup = targetGroup.select<SVGGElement>(".y-axis");
        if (yAxisGroup.empty()) {
            yAxisGroup = targetGroup.append("g").classed("y-axis", true);
        }

        // Clean
        targetGroup.selectAll(".bar").remove();
        targetGroup.selectAll(".connector").remove();
        targetGroup.selectAll(".label").remove();
        targetGroup.selectAll(".summary-indicator").remove();
        targetGroup.selectAll(".axis-break").remove();

        const viewModel = chartData;

        // Synced Font Settings
        const labelFontSize = this.settings.dataLabelSettings.fontSize.value + "px";
        const labelFontFamily = this.settings.dataLabelSettings.fontFamily.value;

        if (!viewModel || viewModel.dataPoints.length === 0) return;

        // Calc Visible Scale Bounds locally to respect showRefMark
        const showRefMark = this.settings.totalSettings.showRefMark.value;
        let visualMin = 0;
        let visualMax = 0;

        viewModel.dataPoints.forEach(d => {
            visualMin = Math.min(visualMin, d.start, d.end);
            visualMax = Math.max(visualMax, d.start, d.end);
            if (showRefMark && d.referenceY !== undefined) {
                visualMin = Math.min(visualMin, d.referenceY);
                visualMax = Math.max(visualMax, d.referenceY);
            }
        });

        const showConnectors = this.settings.layoutSettings.showConnectors.value;
        const connectorColor = this.settings.layoutSettings.connectorColor.value.value;

        const labelOrientation = String(this.settings.xAxisSettings.labelOrientation.value.value);

        const orientation = this.settings.layoutSettings.orientation.value.value; // "vertical" or "horizontal"
        const isHorizontal = orientation === "horizontal";
        const barShape = this.settings.layoutSettings.barShape.value.value;

        // Logic: Map Settings to Physical Axes based on Orientation
        // Horizontal Mode: Left Axis = Category (X-Settings), Bottom Axis = Value (Y-Settings)
        // Vertical Mode: Left Axis = Value (Y-Settings), Bottom Axis = Category (X-Settings)

        const showXAxisLine = this.settings.xAxisSettings.show.value; // "Show Axis" Toggle (Line only)
        const showXAxisLabels = true; // ALWAYS SHOW LABELS per user request
        const showYAxisLabels = this.settings.yAxisSettings.showLabels.value; // Value Axis mainly controlled by Show Labels

        // Determine specific visibility flags for Physical Axes

        // Left Axis
        const showLeftAxisLine = isHorizontal ? showXAxisLine : showYAxisLabels;
        const showLeftAxisLabels = isHorizontal ? showXAxisLabels : showYAxisLabels;

        // Bottom Axis
        const showBottomAxisLine = isHorizontal ? showYAxisLabels : showXAxisLine;
        const showBottomAxisLabels = isHorizontal ? showYAxisLabels : showXAxisLabels;

        // Dynamic Left Margin for Horizontal Mode (Categories)
        let leftMargin = isHorizontal ? 100 : 40;

        if (showLeftAxisLabels) {
            if (isHorizontal) {
                // Measure Max Category Width
                let maxCatWidth = 0;
                // Need context to measure. Re-create or reuse. 
                const context = document.createElement("canvas").getContext("2d");
                const paramsFont = "12px wf_standard-font, helvetica, arial, sans-serif"; // Approximate
                if (context) context.font = paramsFont;

                viewModel.dataPoints.forEach(d => {
                    const w = context ? context.measureText(d.category).width : d.category.length * 7;
                    if (w > maxCatWidth) maxCatWidth = w;
                });
                // Limit max width to 35% of Visual to prevent overtaking
                const maxAllowed = Math.min(maxCatWidth + 20, width * 0.35);
                leftMargin = Math.max(60, maxAllowed);
            } else {
                leftMargin = Math.max(40, Math.min(60, width * 0.15));
            }
        } else {
            leftMargin = 10;
        }

        let bottomMargin = 30; // Increased default base
        if (labelOrientation === "vertical") {
            bottomMargin = Math.min(80, height * 0.4);
        } else if (labelOrientation === "angled") {
            bottomMargin = Math.min(60, height * 0.4);
        }

        const minChartHeight = 20;
        if (height - bottomMargin - 10 < minChartHeight) {
            bottomMargin = Math.max(0, height - minChartHeight - 10);
        }

        const showSummaryIndicator = this.settings.totalSettings.showSummaryIndicator.value;
        const showVariance = this.settings.totalSettings.showVariance.value;
        const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;

        const showRefMarkLabel = this.settings.totalSettings.showRefMarkLabel.value;

        // Increase margins for indicators
        let rightMargin = 10;
        let topMargin = 10;

        if (isHorizontal) {
            // Horizontal Mode: Indicators at Top
            topMargin = 20;

            if (showSummaryIndicator) topMargin += 25;
            if (showVariance && referenceColumn !== "none") topMargin += 35;
            if (showRefMarkLabel) topMargin += 20;

        } else {
            // Vertical Mode: Indicators at Right
            // The previous margin calculation was stealing too much horizontal space for labels.
            // Let's rely on much tighter borders, and only reserve exactly what we need for the variance bracket itself.
            // A bracket is only 10px wide. The text is drawn near the bracket.
            // To prevent clipping but avoid giant whitespace, we just need enough space for the text width.
            rightMargin = 5;

            if (showSummaryIndicator) rightMargin = Math.max(rightMargin, 40);

            if (showVariance && referenceColumn !== "none") {
                const format = this.settings.totalSettings.varianceLabelFormat.value.value;
                const stackLabels = this.settings.totalSettings.stackVarianceLabels.value;

                let requiredWidth = 50;
                if (format === "both") {
                    if (stackLabels) requiredWidth = 50;
                    else requiredWidth = 85;
                } else {
                    requiredWidth = 50;
                }
                rightMargin = Math.max(rightMargin, requiredWidth);
            }

            if (showRefMarkLabel) {
                rightMargin = Math.max(rightMargin, 35);
            }
        }

        // Add bottom margin for Group Axis
        const hasGroups = viewModel.dataPoints.some(d => d.groupName);
        if ((hasGroups || this.canDrillDown) && !isHorizontal) {
            bottomMargin += 50; // Increased from 30 to ensure visibility
        }

        const margin = { top: topMargin, right: rightMargin, bottom: bottomMargin, left: leftMargin };
        const innerWidth = Math.max(10, width - margin.left - margin.right);
        const innerHeight = Math.max(10, height - margin.top - margin.bottom);


        let contentGroup = targetGroup.select<SVGGElement>(".content-group");
        if (contentGroup.empty()) {
            contentGroup = targetGroup.append("g").classed("content-group", true);
        }
        contentGroup.attr("transform", `translate(${margin.left}, ${margin.top})`);

        // ... (rest of code)


        xAxisGroup = contentGroup.select(".x-axis");
        if (xAxisGroup.empty()) xAxisGroup = contentGroup.append("g").classed("x-axis", true);
        yAxisGroup = contentGroup.select(".y-axis");
        if (yAxisGroup.empty()) yAxisGroup = contentGroup.append("g").classed("y-axis", true);

        const enableAxisBreak = this.settings.yAxisSettings.enableAxisBreak.value;
        const axisBreakPercent = this.settings.yAxisSettings.axisBreakPercent.value;
        let breakAmount = 0;

        let xScale: any, yScale: any;

        if (isHorizontal) {
            // Horizontal: Y is Categories (Band), X is Values (Linear)
            yScale = d3.scaleBand()
                .domain(viewModel.dataPoints.map(d => d.category))
                .range([0, innerHeight])
                .padding(0.2);

            // X-Axis logic (was Y-Axis logic)
            let xMin = globalMinMax ? globalMinMax.min : Math.min(0, visualMin);
            let xMax = globalMinMax ? globalMinMax.max : Math.max(0, visualMax);

            const settingsMin = this.settings.yAxisSettings.min.value; // Reuse min/max settings for value axis
            const settingsMax = this.settings.yAxisSettings.max.value;
            if (settingsMin != null) xMin = settingsMin;
            if (settingsMax != null) xMax = settingsMax;

            let xDomainPadding = (xMax - xMin) * 0.1;
            const minStartZero = xMin === 0;

            // Axis Break Logic for Horizontal (Same logic as Vertical but for X)
            breakAmount = 0;
            if (enableAxisBreak && viewModel.dataPoints.length > 0) {
                const deltaPoints = viewModel.dataPoints.filter(d => !d.isTotal);
                if (deltaPoints.length > 0) {
                    // Min value of Start/End for bars (Left side of bars in horizontal)
                    const minStart = Math.min(...deltaPoints.map(d => Math.min(d.start, d.end)));
                    const baselineRange = minStart - xMin;

                    // breakAmount: How much of the empty space we want to "skip"
                    breakAmount = baselineRange * (axisBreakPercent / 100);

                    if (breakAmount > 0) {
                        xMin = xMin + breakAmount;
                        xDomainPadding = (xMax - xMin) * 0.1;
                    }
                }
            }

            const invert = this.settings.yAxisSettings.invert.value;
            // For Horizontal, normally 0 is left, Max is right. Invert would swap.
            // Standard: [0, innerWidth]
            xScale = d3.scaleLinear()
                .domain([minStartZero && !enableAxisBreak ? xMin : xMin - xDomainPadding, xMax + xDomainPadding])
                .range(invert ? [innerWidth, 0] : [0, innerWidth]);

        } else {
            // Vertical (Standard): X is Categories (Band), Y is Values (Linear)
            xScale = d3.scaleBand()
                .domain(viewModel.dataPoints.map(d => d.category))
                .range([0, innerWidth])
                .padding(0.2);

            let yMin = globalMinMax ? globalMinMax.min : Math.min(0, visualMin);
            let yMax = globalMinMax ? globalMinMax.max : Math.max(0, visualMax);

            const settingsMin = this.settings.yAxisSettings.min.value;
            const settingsMax = this.settings.yAxisSettings.max.value;
            if (settingsMin != null) yMin = settingsMin;
            if (settingsMax != null) yMax = settingsMax;

            let yDomainPadding = (yMax - yMin) * 0.1;
            const minStartZero = yMin === 0;

            // Axis Break Logic for Vertical
            breakAmount = 0;
            if (enableAxisBreak && viewModel.dataPoints.length > 0) {
                const deltaPoints = viewModel.dataPoints.filter(d => !d.isTotal);
                if (deltaPoints.length > 0) {
                    const minStart = Math.min(...deltaPoints.map(d => Math.min(d.start, d.end)));
                    const baselineRange = minStart - yMin;
                    breakAmount = baselineRange * (axisBreakPercent / 100);
                    if (breakAmount > 0) {
                        yMin = yMin + breakAmount;
                        yDomainPadding = (yMax - yMin) * 0.1;
                    }
                }
            }

            const invert = this.settings.yAxisSettings.invert.value;
            yScale = d3.scaleLinear()
                .domain([minStartZero && !enableAxisBreak ? yMin : yMin - yDomainPadding, yMax + yDomainPadding])
                .range(invert ? [0, innerHeight] : [innerHeight, 0]);
        }

        // Secondary Scale Logic
        // For Waterfall, we want Reference Marks/Columns to share the same scale as Primary.
        // Independent scaling causes misalignment and ignores Axis Break.
        let xScale2: any = xScale;
        let yScale2: any = yScale;

        const numberScale = String(this.settings.dataLabelSettings.numberScale.value.value);
        const thousandsAbbrev = this.settings.dataLabelSettings.thousandsAbbrev.value;
        const millionsAbbrev = this.settings.dataLabelSettings.millionsAbbrev.value;
        const decimalPlaces = this.settings.dataLabelSettings.decimalPlaces.value;
        const percentDecimalPlaces = this.settings.dataLabelSettings.percentDecimalPlaces.value;
        const useThousandsSeparator = this.settings.dataLabelSettings.useThousandsSeparator.value;

        const labelMaxChars = this.settings.xAxisSettings.labelMaxChars.value;
        const xAxis = d3.axisBottom(xScale);
        const yAxis = d3.axisLeft(yScale);

        // Apply formatting to the value axis and truncation to category axis
        if (isHorizontal) {
            xAxis.tickFormat(d => this.formatNumber(Number(d), decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev));
        } else {
            yAxis.tickFormat(d => this.formatNumber(Number(d), decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev));

            if (labelMaxChars > 0) {
                xAxis.tickFormat((d: string) => {
                    const text = String(d);
                    // Trim " Total" suffix for clean display if desired, or keep specific "Kid Total"
                    // User requested separation. "Kid Total" is clear. 
                    // If user wants just "Total" displayed, we can strip it here, but it might be ambiguous visually?
                    // Given the user said "Both... are under column Total... should be separated", 
                    // Showing "Kid Total" and "Hemtex Total" solves separation clearly.
                    // Let's implement smart truncation if maxChars is hit, but prioritize showing unique name.
                    // Actually, let's STRIP the Group Name prefix if it's a Total?
                    // e.g. "Kid Total" -> "Total"? No, that brings back the visual confusion (two "Total" columns).
                    // The safest bet is to display the full unique name unless it's too long.

                    if (labelMaxChars > 0 && text.length > labelMaxChars) return text.substring(0, labelMaxChars) + "...";
                    return text;
                });
            }
        }


        // --- Axis Configuration & Rendering ---

        // Helper for dash styles
        const getDashArray = (style: string) => {
            switch (style) {
                case "dotted": return "1, 4";
                case "dashed": return "4, 4";
                default: return "none"; // solid
            }
        };

        // 1. X-Axis Configuration
        // Note: We use the boolean flags calculated above for Rendering, but D3 Axis setup needs specific parameters.
        const showXGrid = this.settings.xAxisSettings.showGrid.value;
        const xGridStyle = String(this.settings.xAxisSettings.gridLineStyle.value.value);
        const xGridColor = this.settings.xAxisSettings.gridLineColor.value.value;

        if (showXGrid) {
            xAxis.tickSizeInner(-innerHeight);
        }

        if (!showXAxisLine) {
            xAxis.tickSizeOuter(0); // Hide outer ticks if Line is hidden
            if (!showXGrid) {
                // If grid is off AND line is hidden, hide inner ticks too.
                xAxis.tickSizeInner(0);
            }
        }

        // Render X-Axis
        const xAxisG = xAxisGroup.attr("transform", `translate(0, ${innerHeight})`).call(xAxis);

        // Styling X-Axis Grid
        if (showXGrid) {
            xAxisG.selectAll(".tick line")
                .attr("stroke", xGridColor)
                .attr("stroke-dasharray", getDashArray(xGridStyle))
                .style("opacity", 1);
        }

        if (!showXAxisLine) {
            // Hide Axis Line (Domain) explicitly if settings say so
            // Note: Styling block below will handle "display: none" for domain/ticks more robustly
        }

        // 2. Y-Axis Configuration
        let showYGrid = this.settings.yAxisSettings.showGrid.value;
        const showYLabels = this.settings.yAxisSettings.showLabels.value;

        // Auto-disable grid if labels are hidden (User Request)
        if (!showYLabels) {
            showYGrid = false;
        }

        const yGridStyle = String(this.settings.yAxisSettings.gridLineStyle.value.value);
        const yGridColor = this.settings.yAxisSettings.gridLineColor.value.value;

        if (showYGrid) {
            yAxis.tickSizeInner(-innerWidth);
        }

        // Render Y-Axis (yAxis settings like showLabels already handled by basic d3, but let's check)
        const yAxisG = yAxisGroup.call(yAxis);

        // Styling Y-Axis Grid
        if (showYGrid) {
            yAxisG.selectAll(".tick line")
                .attr("stroke", yGridColor)
                .attr("stroke-dasharray", getDashArray(yGridStyle))
                .style("opacity", 1);
        }

        // Make Axis Labels Interactive
        xAxisG.selectAll(".tick text")
            .style("cursor", "pointer")
            .on("click", (event, d) => {
                const cat = String(d);
                // Find matching data point
                // Note: category might be truncated in label, but d here is usually the full domain value from D3
                const dp = viewModel.dataPoints.find(p => p.category === cat);
                if (dp && dp.selectionId) {
                    const isMultiSelect = event.ctrlKey || event.metaKey;
                    this.selectionManager.select(dp.selectionId, isMultiSelect).then((ids: powerbi.visuals.ISelectionId[]) => {
                        this.currentSelection = ids;
                        this.syncSelectionState(ids);
                    });
                    event.stopPropagation();
                }
            });

        // --- Group Axis Rendering ---
        if (hasGroups && !isHorizontal) {
            contentGroup.selectAll(".group-axis").remove();
            const groupAxis = contentGroup.append("g").classed("group-axis", true).attr("transform", `translate(0, ${innerHeight})`);

            const bandwidth = xScale.bandwidth();
            let currentGroup = null;
            let groupStartX = 0;

            // Iterate points to find group boundaries
            for (let i = 0; i < viewModel.dataPoints.length; i++) {
                const d = viewModel.dataPoints[i];
                // Skip if no group name (shouldn't happen if hasGroups is true, but safe check)
                if (!d.groupName) continue;

                const x = xScale(d.category);

                // Start of new group
                if (d.groupName !== currentGroup) {
                    currentGroup = d.groupName;
                    groupStartX = x;
                }

                // Check for end of group
                // Either it's the last point, or next point has different group
                const next = viewModel.dataPoints[i + 1];
                if (!next || next.groupName !== currentGroup) {
                    const groupEndX = x + bandwidth;
                    const width = groupEndX - groupStartX;

                    // Draw Group Label
                    const labelX = groupStartX + width / 2;
                    const labelY = 40; // Moved down to 40 for better clearance

                    groupAxis.append("text")
                        .attr("x", labelX)
                        .attr("y", labelY) // Position below standard labels
                        .attr("text-anchor", "middle")
                        .style("font-weight", "bold")
                        .style("font-size", "11px")
                        .style("fill", "#666")
                        .text(currentGroup);

                    // Draw Separator Line (if not last group)
                    if (next) {
                        const nextX = xScale(next.category);
                        const midGap = (groupEndX + nextX) / 2;

                        groupAxis.append("line")
                            .attr("x1", midGap)
                            .attr("y1", -5) // Start slightly inside chart?
                            .attr("x2", midGap)
                            .attr("y2", 35) // Extend down
                            .attr("stroke", "#d1d1d1")
                            .attr("stroke-width", 1);
                    }
                }
            }
        }

        // Secondary Axis Rendering - Hidden but scale preserved
        contentGroup.select(".secondary-axis").remove();


        // Axis styling
        // Left Axis visibility (yAxisGroup)
        yAxisGroup.selectAll("text").style("display", showLeftAxisLabels ? null : "none");
        yAxisGroup.selectAll("line").style("display", showLeftAxisLine ? null : "none"); // Line covers Ticks too
        yAxisGroup.select("path.domain").style("display", showLeftAxisLine ? null : "none"); // The main axis line

        // Bottom Axis visibility (xAxisGroup)
        xAxisGroup.selectAll("text").style("display", showBottomAxisLabels ? null : "none");
        xAxisGroup.selectAll("line").style("display", showBottomAxisLine ? null : "none"); // Ticks
        xAxisGroup.select("path.domain").style("display", showBottomAxisLine ? null : "none"); // Axis Line

        // Truncate logic... (applies to Band axis mainly)

        // Orientation handling for labels
        if (!isHorizontal) {
            // Vertical Mode: X-Axis labels might need rotation
            if (labelOrientation === "vertical") {
                xAxisGroup.selectAll("text").attr("transform", "rotate(-90)").attr("text-anchor", "end").attr("dx", "-0.8em").attr("dy", "-0.5em");
            } else if (labelOrientation === "angled") {
                xAxisGroup.selectAll("text").attr("transform", "rotate(-45)").attr("text-anchor", "end").attr("dx", "-0.5em").attr("dy", "0.15em");
            }
        }

        if (enableAxisBreak && breakAmount > 0) {
            contentGroup.selectAll(".axis-break").remove(); // Clean up old breaks

            if (isHorizontal) {
                // Horizontal Mode: Break is on the X-Axis (Bottom), near the start (Left)
                // We want a vertical zigzag near x=0 (after margins)
                // X Axis is at y = innerHeight

                // Position logic: The break is effectively skipping data *before* the visible start.
                // The visual representation (zigzag) is usually placed on the axis line itself.

                const breakX = -5; // Start of axis
                const yPos = innerHeight;

                contentGroup.append("path")
                    .classed("axis-break", true)
                    .attr("d", `M ${breakX} ${yPos - 5} L ${breakX + 5} ${yPos} L ${breakX} ${yPos + 5} L ${breakX + 5} ${yPos + 10}`)
                    // Drawing a vertical zigzag on the horizontal axis is tricky.
                    // Usually it's a break *in* the axis line. Like //
                    // Standard approach: Two slanted lines chopping the axis.
                    // Let's do // style on the horizontal line at the start.

                    .attr("d", `M 0 ${yPos + 5} L 5 ${yPos - 5} M 5 ${yPos + 5} L 10 ${yPos - 5}`)
                    // Slanted strikes at X=0 to X=10
                    .attr("stroke", "#999").attr("stroke-width", 1.5).attr("fill", "none");

            } else {
                // Vertical Mode: Break is on Y-Axis (Left), near the bottom
                const breakY = innerHeight - 5;
                contentGroup.append("path")
                    .classed("axis-break", true)
                    .attr("d", `M -10 ${breakY + 6} L -5 ${breakY + 3} L -10 ${breakY} L -5 ${breakY - 3} L -10 ${breakY - 6}`)
                    .attr("stroke", "#999").attr("stroke-width", 1.5).attr("fill", "none");
            }
        }

        const showLabels = this.settings.dataLabelSettings.show.value;

        // Slot Layout Calculation (Simplified for Waterfall only)
        // No side-by-side columns anymore.
        const pySlot = -1, wfSlot = 0, cySlot = -1;
        const usedCount = 1;

        const getSlotInfo = (bandwidth: number, slotIndex: number) => {
            // Simple full width (minus padding)
            const padding = Math.min(bandwidth * 0.1, 10);
            const width = bandwidth - padding;
            const offset = padding / 2;
            return { offset, width };
        };

        // --- Zero Baseline Line ---
        // Draw a thin line at y=0 when data crosses zero (like Zebra BI)
        contentGroup.selectAll(".zero-baseline").remove();
        if (!isHorizontal) {
            const domain = yScale.domain();
            if (domain[0] < 0 && domain[1] > 0) {
                contentGroup.append("line")
                    .classed("zero-baseline", true)
                    .attr("x1", 0)
                    .attr("x2", innerWidth)
                    .attr("y1", yScale(0))
                    .attr("y2", yScale(0))
                    .attr("stroke", "#999")
                    .attr("stroke-width", 1)
                    .attr("stroke-opacity", 0.7);
            }
        } else {
            const domain = xScale.domain();
            if (domain[0] < 0 && domain[1] > 0) {
                contentGroup.append("line")
                    .classed("zero-baseline", true)
                    .attr("x1", xScale(0))
                    .attr("x2", xScale(0))
                    .attr("y1", 0)
                    .attr("y2", innerHeight)
                    .attr("stroke", "#999")
                    .attr("stroke-width", 1)
                    .attr("stroke-opacity", 0.7);
            }
        }

        if (showConnectors) {
            for (let i = 0; i < viewModel.dataPoints.length - 1; i++) {
                const current = viewModel.dataPoints[i];
                const next = viewModel.dataPoints[i + 1];
                const line = contentGroup.append("line")
                    .classed("connector", true)
                    .attr("stroke", connectorColor).attr("stroke-width", 1).attr("stroke-dasharray", "4,2");

                if (isHorizontal) {
                    const bandwidth = yScale.bandwidth();
                    const slot = getSlotInfo(bandwidth, wfSlot);

                    line.attr("x1", xScale(current.end))
                        .attr("x2", xScale(current.end))
                        .attr("y1", yScale(current.category) + slot.offset + slot.width)
                        .attr("y2", yScale(next.category) + slot.offset);

                } else {
                    const bandwidth = xScale.bandwidth();
                    const slot = getSlotInfo(bandwidth, wfSlot);

                    line.attr("x1", xScale(current.category) + slot.offset + slot.width)
                        .attr("y1", yScale(current.end))
                        .attr("x2", xScale(next.category) + slot.offset)
                        .attr("y2", yScale(current.end));
                }
            }
        }

        // remove old supplemental columns
        contentGroup.selectAll(".supplemental-col").remove();



        // Connectors Logic Removed


        // Draw Bars (Waterfall)
        // Draw Bars (Waterfall)
        const bars = contentGroup.selectAll(".bar")
            .data(viewModel.dataPoints.filter(d => !d.isSpacer))
            .enter().append("path")
            .classed("bar", true)
            .attr("tabindex", 0)
            .attr("role", "button")
            .attr("aria-label", d => `${d.category}. Value: ${this.formatNumber(d.delta, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev)}`);

        bars.append("title").text(d => `${d.category}: ${this.formatNumber(d.delta, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev)}`);

        bars.attr("fill", d => {
            if (d.isReference) return "none";
            return d.color;
        })
            .attr("d", d => {
                let x = 0, y = 0, w = 0, h = 0;
                let direction = "none";

                // Standard Variance Mode - Using Secondary Axis (indicatorXScale/YScale)
                // We use xScale/yScale for Waterfall Bars (Primary Axis)
                // This ensures they match the visible axis and work even if PY/TY are missing.
                const activeXScale = xScale;
                const activeYScale = yScale;



                if (isHorizontal) {
                    const bw = yScale.bandwidth();
                    const firstSlot = getSlotInfo(bw, 0);
                    const lastSlot = getSlotInfo(bw, usedCount - 1);
                    const clusterOffset = firstSlot.offset;
                    const clusterWidth = (lastSlot.offset + lastSlot.width) - clusterOffset;

                    y = yScale(d.category) + clusterOffset;
                    h = clusterWidth;

                    const axisBase = activeXScale.domain()[0];
                    const base = axisBase; // Use actual domain minimum, not Math.max(0, ...) which clips negative bars
                    const startX = (d.isTotal && !d.isSubtotal) ? activeXScale(0) : activeXScale(d.start);
                    const endX = activeXScale(d.end);

                    x = Math.min(startX, endX);
                    w = Math.abs(startX - endX);

                    if (!d.isTotal) {
                        if (d.delta >= 0) direction = "right";
                        else direction = "left";
                    }
                } else {
                    const bw = xScale.bandwidth();
                    // Horizontal alignment relies on Category Axis (xScale) which is Band.
                    // xScale is correct here (Categories). Value scale is Y.

                    const firstSlot = getSlotInfo(bw, 0);
                    const lastSlot = getSlotInfo(bw, usedCount - 1);
                    const clusterOffset = firstSlot.offset;
                    const clusterWidth = (lastSlot.offset + lastSlot.width) - clusterOffset;

                    x = xScale(d.category) + clusterOffset;
                    w = clusterWidth;

                    const axisBottom = activeYScale.domain()[0];
                    const base = axisBottom; // Use actual domain minimum, not Math.max(0, ...) which clips negative bars
                    const startY = (d.isTotal && !d.isSubtotal) ? activeYScale(0) : activeYScale(d.start);
                    const endY = activeYScale(d.end);

                    let yTop = Math.min(startY, endY);
                    let yBot = Math.max(startY, endY);

                    // Clamp to Visible Chart Area (Axis Bottom)
                    // The visual bottom is activeYScale(axisBottom).
                    const visualBottom = activeYScale(base); // base was Math.max(0, axisBottom) -> axisBottom usually > 0? No axisBottom is Value.
                    // activeYScale(axisBottom) gives the PIXEL coordinate of the bottom.
                    // Wait, base calculation: const base = Math.max(0, axisBottom);
                    // If axisBottom is 3000. Base is 3000.
                    // activeYScale(3000) is the bottom pixel.

                    const pixelBottom = activeYScale(base);

                    // Initialize broken status for Mask logic
                    d.isBroken = false;

                    const isInverted = this.settings.yAxisSettings.invert.value;

                    if (!isInverted) {
                        // Standard: Baseline is at Bottom (High Pixel Value)
                        // Bars usually grow Up (Lower Pixel Values)
                        // Break if Bar extends "Below" Baseline (Higher Pixel > Baseline)
                        if (yBot > pixelBottom) {
                            yBot = pixelBottom;
                            d.isBroken = true;
                        }
                        if (yTop > pixelBottom) yTop = pixelBottom;
                    } else {
                        // Inverted: Baseline is at Top (Low Pixel Value, usually 0)
                        // Bars grow Down (Higher Pixel Values)
                        // Break if Bar extends "Above" Baseline (Lower Pixel < Baseline)
                        if (yTop < pixelBottom) {
                            yTop = pixelBottom;
                            d.isBroken = true;
                        }
                        if (yBot < pixelBottom) yBot = pixelBottom;
                    }

                    y = yTop;
                    h = Math.abs(yBot - yTop);

                    // Safety for 0 height
                    if (h < 1 && Math.abs(d.delta) > 0 && !d.isBroken) h = 1;

                    if (!d.isTotal) {
                        if (d.delta >= 0) direction = "up";
                        else direction = "down";
                    }
                }

                return this.getBarPath(String(barShape), x, y, w, h, direction);
            })
            // Apply Mask AFTER calculation of isBroken
            .attr("mask", d => {
                const enableAxisBreak = this.settings.yAxisSettings.enableAxisBreak.value;
                if (enableAxisBreak) {
                    // Break ONLY if marked as broken (Calculated Limit)
                    // Also check if isTotal (Start/End columns) are broken?
                    // User requested Start/End specifically. 
                    // If they are broken, mask applies.
                    if (d.isBroken && !isHorizontal) return "url(#axis-break-mask)";
                }
                return null;
            });





        // Apply highlight opacity - dim non-highlighted bars when highlights are present
        const hasHighlights = viewModel.dataPoints.some(d => d.highlight === true);
        if (hasHighlights) {
            bars.style("opacity", d => d.highlight ? 1 : 0.4);
        }

        const showRunningTotal = this.settings.dataLabelSettings.showRunningTotal.value;
        const fontSize = this.settings.dataLabelSettings.fontSize.value;
        const labelColor = this.settings.dataLabelSettings.color.value.value;


        if (showLabels) {
            const deltaDisplayMode = String(this.settings.dataLabelSettings.deltaDisplayMode.value.value);
            const fontFamily = this.settings.dataLabelSettings.fontFamily.value;
            const showBackground = this.settings.dataLabelSettings.showBackground.value;
            const backgroundColor = this.settings.dataLabelSettings.backgroundColor.value.value;
            const transparency = this.settings.dataLabelSettings.backgroundTransparency.value;
            const opacity = (100 - transparency) / 100;

            const badgeShape = this.settings.dataLabelSettings.badgeShape.value.value;

            // Start fresh: Remove old label groups to ensure structure (rect + text) is correct
            contentGroup.selectAll(".label-group").remove();

            const labels = contentGroup.selectAll(".label-group")
                .data(viewModel.dataPoints.filter(d => !d.isSpacer));

            const labelsEnter = labels.enter().append("g").classed("label-group", true);
            labelsEnter.append("rect").classed("label-bg", true);
            labelsEnter.append("text").classed("label-text", true);

            const labelsMerge = labels.merge(labelsEnter);

            labelsMerge.each((d, i, nodes) => {
                const group = d3.select(nodes[i]);
                const textEl = group.select("text.label-text");
                const rectEl = group.select("rect.label-bg");

                // If it is a Subtotal (Group), show DELTA (+9M).
                // If it is a Start/End/Global Total, show RUNNING TOTAL (1471M).
                const useDelta = !d.isTotal || (d.isTotal && d.isSubtotal);
                const val = showRunningTotal ? d.runningTotal : d.delta;

                const absStr = this.formatNumber(val, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);

                let mainText = absStr;
                let subText = "";

                if (!d.isTotal || d.isSubtotal) {
                    let pVal = (d.py !== 0) ? (d.delta / d.py) * 100 : 0;
                    const pStr = "(" + pVal.toFixed(percentDecimalPlaces) + "%)";
                    if (deltaDisplayMode === "percent") mainText = pStr;
                    else if (deltaDisplayMode === "both") subText = pStr;
                }

                // Apply text content and font
                // Clear existing content
                textEl.selectAll("*").remove();
                textEl.text(null);

                // Stack labels in vertical orientation when mode is "both"
                if (!isHorizontal && subText) {
                    // Use tspan for stacking
                    textEl.append("tspan")
                        .attr("x", 0)
                        .attr("dy", "0em")
                        .text(mainText);
                    textEl.append("tspan")
                        .attr("x", 0)
                        .attr("dy", "1.1em")
                        .text(subText);
                } else {
                    textEl.text(subText ? `${mainText} ${subText}` : mainText);
                }

                textEl.style("font-family", fontFamily)
                    .style("font-size", Math.max(9, fontSize) + "px")
                    .attr("fill", labelColor)
                    .attr("text-anchor", "middle");

                // Determine Position
                let x = 0;
                let y = 0;

                if (isHorizontal) {
                    y = yScale(d.category) + yScale.bandwidth() / 2 + 5;
                    const barEndX = xScale(d.end);
                    if (d.delta >= 0) x = barEndX + 15;
                    else x = barEndX - 15;
                } else {
                    x = xScale(d.category) + xScale.bandwidth() / 2;
                    const barEndY = yScale(d.end);
                    // Adjust Y offset to account for stacked labels
                    const stackOffset = subText ? 8 : 0;
                    if (d.delta >= 0) y = barEndY - 12 - stackOffset;
                    else y = barEndY + 15;
                }

                group.attr("transform", `translate(${x}, ${y})`);

                // Update Background Rect
                const node = textEl.node() as SVGTextElement;
                if (node && showBackground) {
                    const bbox = node.getBBox();
                    const paddingX = 4;
                    const paddingY = 2;
                    const rectH = bbox.height + paddingY * 2;

                    let rx = 0;
                    if (badgeShape === "rounded") rx = 3;
                    else if (badgeShape === "pill") rx = rectH / 2;

                    rectEl.style("display", null)
                        .attr("x", bbox.x - paddingX)
                        .attr("y", bbox.y - paddingY)
                        .attr("width", bbox.width + paddingX * 2)
                        .attr("height", rectH)
                        .attr("rx", rx)
                        .attr("ry", rx) // Standard SVG: ry defaults to rx if omitted, but explicit is safe
                        .attr("fill", backgroundColor)
                        .attr("opacity", opacity);
                } else {
                    rectEl.style("display", "none");
                }
            });

            labels.exit().remove();
        }

        bars.on("click", (event, d) => {
            const isMultiSelect = event.ctrlKey || event.metaKey;
            // Guard against null selectionId (e.g. Totals)
            if (d.selectionId) {
                this.selectionManager.select(d.selectionId, isMultiSelect).then((ids: powerbi.visuals.ISelectionId[]) => {
                    this.currentSelection = ids;
                    this.syncSelectionState(ids);
                });
            }
            event.stopPropagation();
        })
            .on("contextmenu", (event, d) => {
                // Show context menu on right-click
                const mouseEvent: MouseEvent = event as MouseEvent;
                this.selectionManager.showContextMenu(d.selectionId ? d.selectionId : {}, {
                    x: mouseEvent.clientX,
                    y: mouseEvent.clientY
                });
                mouseEvent.preventDefault();
                mouseEvent.stopPropagation();
            })
            .on("keydown", (event, d) => {
                if (event.key === "Enter" || event.key === " ") {
                    const isMultiSelect = event.ctrlKey || event.metaKey;
                    if (d.selectionId) {
                        this.selectionManager.select(d.selectionId, isMultiSelect).then((ids: powerbi.visuals.ISelectionId[]) => {
                            this.currentSelection = ids;
                            this.syncSelectionState(ids);
                        });
                    }
                    event.preventDefault();
                    event.stopPropagation();
                }
            })
            .on("mouseover", (event, d) => {
                // Hover Highlight: Dim others
                this.mainGroup.selectAll<SVGRectElement, WaterfallDataPoint>(".bar")
                    .style("opacity", (barData) => {
                        return barData.category === d.category ? 1 : 0.4;
                    });

                let tooltips: VisualTooltipDataItem[] = [];

                // Exclusive Tooltips: If user provided custom tooltips, show ONLY those.
                if (d.tooltipValues && d.tooltipValues.length > 0) {
                    d.tooltipValues.forEach(tv => {
                        tooltips.push({ displayName: tv.displayName, value: tv.value });
                    });
                } else {
                    // Default Tooltips (formatted with specific Tooltip Settings)
                    const ttUnits = this.settings.tooltipSettings.tooltipNumberScale.value.value as string;
                    const ttDecimals = this.settings.tooltipSettings.tooltipDecimalPlaces.value;
                    const useSeparator = this.settings.tooltipSettings.useThousandsSeparator.value;
                    const tAbbrev = this.settings.dataLabelSettings.thousandsAbbrev.value; // Reuse abbreviations
                    const mAbbrev = this.settings.dataLabelSettings.millionsAbbrev.value;

                    const deltaStr = this.formatNumber(d.delta, ttDecimals, useSeparator, ttUnits, tAbbrev, mAbbrev);
                    const runningTotalStr = this.formatNumber(d.runningTotal, ttDecimals, useSeparator, ttUnits, tAbbrev, mAbbrev);

                    tooltips.push(
                        { displayName: "Category", value: d.category },
                        { displayName: "Delta", value: deltaStr },
                        { displayName: "Running Total", value: runningTotalStr }
                    );

                    if (!d.isTotal) {
                        const refStr = d.referenceValue !== undefined ? this.formatNumber(d.referenceValue, ttDecimals, useSeparator, ttUnits, tAbbrev, mAbbrev) : "N/A";

                        /* PY/CY tooltips removed */
                        if (d.referenceValue !== undefined) {
                            tooltips.push({ displayName: "Budget/Ref", value: refStr });
                        }
                    }
                }

                this.tooltipService.show({
                    coordinates: [event.clientX, event.clientY],
                    isTouchEvent: false,
                    dataItems: tooltips,
                    identities: d.selectionId ? [d.selectionId] : []
                });
            })
            .on("mousemove", (event) => {
                this.tooltipService.move({
                    coordinates: [event.clientX, event.clientY],
                    isTouchEvent: false,
                    identities: [],
                    dataItems: []
                });
            })
            .on("mouseout", () => {
                this.tooltipService.hide({ isTouchEvent: false, immediately: true });
                this.syncSelectionState(this.currentSelection);
            });



        // Variance Indicator Logic
        const showVarianceLabel = this.settings.totalSettings.showVarianceLabel.value;
        const showSummaryIndicatorLabel = this.settings.totalSettings.showSummaryIndicatorLabel.value;
        const varianceLabelFormat = this.settings.totalSettings.varianceLabelFormat.value.value;
        const summaryLabelFormat = this.settings.totalSettings.summaryLabelFormat.value.value;

        // Flags for dynamic layout
        let summaryIndicatorVisible = false;
        let summaryIndicatorMaxPos = 0; // Max extent (Y in horizontal, X in vertical)
        let summaryIndicatorLabelY = -9999; // Store Y position of first label for collision detection

        // 1. Summary Indicator (End Column Variance to Start Column)
        if (showSummaryIndicator && viewModel.pyTotal !== 0) {
            summaryIndicatorVisible = true;
            // ... (Calculation logic remains same)

            const totalDelta = viewModel.dataPoints.filter(d => !d.isTotal).reduce((sum, d) => sum + d.delta, 0);
            const totalPercent = (totalDelta / viewModel.pyTotal) * 100;
            const absStr = this.formatNumber(totalDelta, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
            const pStr = "(" + (totalPercent >= 0 ? "+" : "") + totalPercent.toFixed(percentDecimalPlaces) + "%)";

            let valueText = "";
            if (summaryLabelFormat === "data") valueText = absStr;
            else if (summaryLabelFormat === "percent") valueText = pStr;
            else valueText = absStr + " " + pStr;

            let labelText = valueText;
            if (showSummaryIndicatorLabel) {
                labelText = "∆" + viewModel.startName + ": " + valueText;
            }

            const indicatorColor = totalDelta >= 0 ? this.settings.totalSettings.summaryPositiveColor.value.value : this.settings.totalSettings.summaryNegativeColor.value.value;

            if (isHorizontal) {
                // ... Horizontal Logic (Keep existing)
                const xStart = xScale(viewModel.pyTotal);
                const xEnd = xScale(viewModel.pyTotal + totalDelta);
                const bracketY = -15; // Closer to axis
                summaryIndicatorMaxPos = -30; // Reduced height usage

                contentGroup.append("line").attr("x1", xStart).attr("y1", 0).attr("x2", xStart).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xEnd).attr("y1", 0).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xStart).attr("y1", bracketY).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", indicatorColor).attr("stroke-width", 2);

                contentGroup.append("circle").attr("cx", xStart).attr("cy", bracketY).attr("r", 3).attr("fill", indicatorColor);
                contentGroup.append("circle").attr("cx", xEnd).attr("cy", bracketY).attr("r", 3).attr("fill", indicatorColor);

                const labelX = (xStart + xEnd) / 2;
                const labelY = bracketY - 8;
                contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(labelText)
                    .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "middle");

            } else {
                // Vertical Mode
                // Improved: Search by category name instead of index
                const startPoint = viewModel.dataPoints.find(d => d.isTotal && d.category === viewModel.startName);
                const endPoint = viewModel.dataPoints.find(d => d.isTotal && d.category === viewModel.endName);

                let startVal = viewModel.pyTotal;
                if (startPoint) startVal = startPoint.end;

                let endVal = startVal + totalDelta;
                if (endPoint) endVal = endPoint.end;

                const yStart = yScale(startVal);
                const yEnd = yScale(endVal);

                let bracketX = innerWidth + 20;

                if (endPoint) {
                    const endBarX = xScale(endPoint.category) + xScale.bandwidth();
                    bracketX = endBarX + 10;
                }

                summaryIndicatorMaxPos = bracketX + 80;

                let xStartOrigin = 0;
                if (startPoint) xStartOrigin = xScale(startPoint.category) + xScale.bandwidth();

                let xEndOrigin = 0;
                if (endPoint) xEndOrigin = xScale(endPoint.category) + xScale.bandwidth();

                contentGroup.append("line").attr("x1", xStartOrigin).attr("y1", yStart).attr("x2", bracketX).attr("y2", yStart)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");

                contentGroup.append("line").attr("x1", xEndOrigin).attr("y1", yEnd).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");

                contentGroup.append("line").attr("x1", bracketX).attr("y1", yStart).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", indicatorColor).attr("stroke-width", 2);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yStart).attr("r", 3).attr("fill", indicatorColor);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yEnd).attr("r", 3).attr("fill", indicatorColor);

                const labelX = bracketX + 5;
                const labelY = (yStart + yEnd) / 2;
                summaryIndicatorLabelY = labelY; // Store for collision check

                // ... Text Rendering Logic (Keep lines 1494-1530 logic roughly same but inside this block context if needed, wait, I am replacing the surrounding block)
                // Need to reproduce the text rendering logic here or it will be lost.
                // Re-using the text logic from lines 1497-1530:

                if (summaryLabelFormat === "both") {
                    const stackLabels = this.settings.totalSettings.stackVarianceLabels.value;
                    if (stackLabels) {
                        if (showSummaryIndicatorLabel) {
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY - 11).text("∆" + viewModel.startName)
                                .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(absStr)
                                .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY + 11).text(pStr)
                                .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).attr("text-anchor", "start");
                        } else {
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY - 5).text(absStr)
                                .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY + 6).text(pStr)
                                .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).attr("text-anchor", "start");
                        }
                    } else {
                        let line1 = absStr + " " + pStr;
                        if (showSummaryIndicatorLabel) line1 = "∆" + viewModel.startName + ": " + line1;
                        contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(line1)
                            .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                    }
                } else {
                    contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(labelText)
                        .attr("fill", indicatorColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                }
            }
        }



        const refStrokeColor = this.settings.totalSettings.refMarkColor.value.value;
        const refMarkShape = this.settings.totalSettings.refMarkShape.value.value;
        // showRefMarkOnColumns removed

        // Find End Point with Reference Value (Needed for Variance Indicator)
        const refPoint = viewModel.dataPoints.find(d => d.isTotal && d.referenceValue !== undefined);

        contentGroup.selectAll(".ref-mark").remove();

        // Loop through all data points
        // Determine active scales for indicators
        // Use Secondary Axis for Reference Marks as requested.
        const indicatorXScale = xScale2;
        const indicatorYScale = yScale2;

        viewModel.dataPoints.forEach(d => {
            if (d.referenceValue !== undefined) {
                // Determine if we should draw
                const shouldDraw = (d.isTotal && showRefMark);
                if (!shouldDraw) return;

                // FIX: Use referenceValue (Absolute) 
                let refVal = d.referenceValue;

                // Ensure we have a value
                if (refVal === undefined || refVal === null) return;

                // FIX: Select Scale based on Type
                // Totals (Start/End) are on Primary Axis (Waterfall Bars).
                // Deltas (Columns) are on Secondary Axis.
                let usePrimary = d.isTotal;

                let activeXScale, activeYScale;

                if (isHorizontal) {
                    activeXScale = usePrimary ? xScale : xScale2;
                    activeYScale = yScale; // Category Scale
                } else {
                    activeXScale = xScale; // Category Scale
                    activeYScale = usePrimary ? yScale : yScale2;
                }

                if (isHorizontal) {
                    const xRef = activeXScale(refVal);
                    const bw = yScale.bandwidth();
                    const slot = getSlotInfo(bw, wfSlot);
                    const yPos = activeYScale(d.category) + slot.offset;
                    const h = slot.width;

                    if (refMarkShape === "line") {
                        contentGroup.append("line")
                            .classed("ref-mark", true)
                            .attr("x1", xRef).attr("y1", yPos)
                            .attr("x2", xRef).attr("y2", yPos + h)
                            .attr("stroke", refStrokeColor).attr("stroke-width", 2);
                    } else if (refMarkShape === "dashed") {
                        contentGroup.append("line")
                            .classed("ref-mark", true)
                            .attr("x1", xRef).attr("y1", yPos)
                            .attr("x2", xRef).attr("y2", yPos + h)
                            .attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("stroke-dasharray", "4,2");
                    } else if (refMarkShape === "circle") {
                        contentGroup.append("circle")
                            .classed("ref-mark", true)
                            .attr("cx", xRef).attr("cy", yPos + h / 2)
                            .attr("r", 4)
                            .attr("fill", refStrokeColor);
                    } else if (refMarkShape === "cross") {
                        const cx = xRef;
                        const cy = yPos + h / 2;
                        const r = 4;
                        contentGroup.append("line")
                            .classed("ref-mark", true).attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("x1", cx - r).attr("y1", cy - r).attr("x2", cx + r).attr("y2", cy + r);
                        contentGroup.append("line")
                            .classed("ref-mark", true).attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("x1", cx - r).attr("y1", cy + r).attr("x2", cx + r).attr("y2", cy - r);
                    }

                    // Render Reference Label (Horizontal Mode)
                    const showRefLabel = this.settings.totalSettings.showRefMarkLabel.value;
                    if (showRefLabel) {
                        const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
                        let refAbbrev = "Ref";
                        const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";
                        const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
                        const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";

                        if (referenceColumn === "budget") refAbbrev = budgetAbbrev;
                        else if (referenceColumn === "py") refAbbrev = pyAbbrev;
                        else if (referenceColumn === "cy") refAbbrev = cyAbbrev;

                        contentGroup.append("text")
                            .attr("x", xRef) // Centered horizontally on line
                            .attr("y", yPos - 5) // Above the line
                            .text(refAbbrev)
                            .attr("fill", refStrokeColor)
                            .style("font-size", "10px")
                            .style("font-weight", "bold")
                            .attr("text-anchor", "middle");
                    }

                } else {
                    const yRef = activeYScale(refVal);
                    const bw = xScale.bandwidth();
                    const slot = getSlotInfo(bw, wfSlot);
                    const xPos = activeXScale(d.category) + slot.offset;
                    const w = slot.width;

                    if (refMarkShape === "line") {
                        contentGroup.append("line")
                            .classed("ref-mark", true)
                            .attr("x1", xPos).attr("y1", yRef)
                            .attr("x2", xPos + w).attr("y2", yRef)
                            .attr("stroke", refStrokeColor).attr("stroke-width", 2);
                    } else if (refMarkShape === "dashed") {
                        contentGroup.append("line")
                            .classed("ref-mark", true)
                            .attr("x1", xPos).attr("y1", yRef)
                            .attr("x2", xPos + w).attr("y2", yRef)
                            .attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("stroke-dasharray", "4,2");
                    } else if (refMarkShape === "circle") {
                        contentGroup.append("circle")
                            .classed("ref-mark", true)
                            .attr("cx", xPos + w / 2).attr("cy", yRef)
                            .attr("r", 4)
                            .attr("fill", refStrokeColor);
                    } else if (refMarkShape === "cross") {
                        const cx = xPos + w / 2;
                        const cy = yRef;
                        const r = 4;
                        contentGroup.append("line")
                            .classed("ref-mark", true).attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("x1", cx - r).attr("y1", cy + r).attr("x2", cx + r).attr("y2", cy - r);
                    }

                    // Render Reference Label
                    const showRefLabel = this.settings.totalSettings.showRefMarkLabel.value;
                    if (showRefLabel) {
                        const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
                        let refAbbrev = "Ref";
                        const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";
                        const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
                        const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";

                        if (referenceColumn === "budget") refAbbrev = budgetAbbrev;
                        else if (referenceColumn === "py") refAbbrev = pyAbbrev;
                        else if (referenceColumn === "cy") refAbbrev = cyAbbrev;

                        contentGroup.append("text")
                            .attr("x", xPos + w + 15)
                            .attr("y", yRef + 4) // Centered vertically relative to line (approx)
                            .text(refAbbrev)
                            .attr("fill", refStrokeColor)
                            .style("font-size", "11px")
                            .attr("text-anchor", "start");
                    }
                }
            }
        });




        // 2. Variance Indicator (Reference Variance to End Column)
        if (showVariance && showRefMark && refPoint) {
            const endVal = refPoint.end;
            const refVal = refPoint.referenceValue;
            const variance = endVal - refVal;

            let variancePct = 0;
            if (refVal !== 0) variancePct = (variance / refVal) * 100;

            const vAbsStr = this.formatNumber(variance, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
            const vPStr = "(" + (variancePct >= 0 ? "+" : "") + variancePct.toFixed(percentDecimalPlaces) + "%)";

            let vValueText = "";
            if (varianceLabelFormat === "data") vValueText = vAbsStr;
            else if (varianceLabelFormat === "percent") vValueText = vPStr;
            else vValueText = vAbsStr + " " + vPStr;

            let vLabelText = vValueText;
            if (showVarianceLabel) {
                // Get Reference Name
                // We should use referenceColumn setting to map to abbrev?
                // Or just use viewModel.refName? 
                // Let's check viewModel.refName (it was replaced earlier).
                // Or "∆" + viewModel.refName (Ref Name is usually Ref/Column Name). 
                // But user wants ∆BUD. 
                // Let's use simple logic: Abbrev of "Budget" is BUD. 
                // We already have 'refName' calc in generateChartData. 
                // Let's use that.
                // However, refName is not on viewModel. 
                // Let's infer from settings again.
                const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
                let refAbbrev = "Ref";
                const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";
                const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
                const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";

                if (referenceColumn === "budget") refAbbrev = budgetAbbrev;
                else if (referenceColumn === "py") refAbbrev = pyAbbrev;
                else if (referenceColumn === "cy") refAbbrev = cyAbbrev;

                vLabelText = "∆" + refAbbrev + ": " + vLabelText;
            }

            const vColor = variance >= 0 ? this.settings.totalSettings.summaryPositiveColor.value.value : this.settings.totalSettings.summaryNegativeColor.value.value;

            // Use Primary Scale
            const indXScale = xScale;
            const indYScale = yScale;

            if (isHorizontal) {
                // Horizontal Mode
                const xStart = indXScale(endVal); // Current (End) Value
                const xEnd = indXScale(refVal);   // Reference Value (Mark)

                // Determine Y Offset
                // If summary visible, stack ABOVE. (Negative Y)
                let bracketY = -35; // Default if summary not there
                if (summaryIndicatorVisible) {
                    // Summary used -30. Go above.
                    bracketY = summaryIndicatorMaxPos - 25; // -30 - 25 = -55.
                }

                contentGroup.append("line").attr("x1", xStart).attr("y1", 0).attr("x2", xStart).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xEnd).attr("y1", 0).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xStart).attr("y1", bracketY).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", vColor).attr("stroke-width", 2);
                contentGroup.append("circle").attr("cx", xStart).attr("cy", bracketY).attr("r", 3).attr("fill", vColor);
                contentGroup.append("circle").attr("cx", xEnd).attr("cy", bracketY).attr("r", 3).attr("fill", vColor);

                const labelX = (xStart + xEnd) / 2;
                const labelY = bracketY - 8;
                contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(vLabelText)
                    .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "middle");

            } else {
                // Vertical Mode
                // Position Variance Indicator to the right of End Column
                const endBarX = indXScale(refPoint.category) + indXScale.bandwidth();

                // Reverting Horizontal Shift. User prefers vertical offset for values.
                let bracketX = endBarX + 10;

                const yStart = indYScale(endVal); // Current Value
                const yEnd = indYScale(refVal);   // Reference Value

                // Line 1: From End Column Value to Bracket
                contentGroup.append("line").attr("x1", endBarX).attr("y1", yStart).attr("x2", bracketX).attr("y2", yStart)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");

                // Line 2: From End Column Reference Mark to Bracket
                contentGroup.append("line").attr("x1", endBarX).attr("y1", yEnd).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2"); // from mark

                contentGroup.append("line").attr("x1", bracketX).attr("y1", yStart).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", vColor).attr("stroke-width", 2);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yStart).attr("r", 3).attr("fill", vColor);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yEnd).attr("r", 3).attr("fill", vColor);

                const labelX = bracketX + 5;
                let labelY = (yStart + yEnd) / 2;

                // Vertical Offset Logic for Collision
                // If Summary Indicator is visible, check if our label overlaps with its label.
                if (summaryIndicatorVisible && summaryIndicatorLabelY !== -9999) {
                    const stackLabels = this.settings.totalSettings.stackVarianceLabels.value;
                    const collisionThreshold = stackLabels ? 45 : 18;
                    const shiftAmount = stackLabels ? 25 : 15;

                    const diff = Math.abs(summaryIndicatorLabelY - labelY);
                    if (diff < collisionThreshold) {
                        // Overlap detected. 
                        if (yEnd < yStart) {
                            labelY = labelY - shiftAmount;
                        } else {
                            labelY = labelY + shiftAmount;
                        }
                    }
                }

                // Stacked Logic for "Both" format (or if label is forced to stack)
                // Stacked Logic for "Both" format (or if label is forced to stack)
                if (varianceLabelFormat === "both") {
                    const stackLabels = this.settings.totalSettings.stackVarianceLabels.value;

                    if (stackLabels) {
                        if (showVarianceLabel) {
                            // 3-Line Mode: Label, Value, Percent
                            // Reconstruct prefix
                            const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
                            let refAbbrev = "Ref";
                            const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";
                            const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
                            const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";

                            if (referenceColumn === "budget") refAbbrev = budgetAbbrev;
                            else if (referenceColumn === "py") refAbbrev = pyAbbrev;
                            else if (referenceColumn === "cy") refAbbrev = cyAbbrev;

                            contentGroup.append("text").attr("x", labelX).attr("y", labelY - 11).text("∆" + refAbbrev)
                                .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(vAbsStr)
                                .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY + 11).text(vPStr)
                                .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).attr("text-anchor", "start");
                        } else {
                            // 2-Line Mode (Stack ON, Label OFF): Value, Percent
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY - 5).text(vAbsStr)
                                .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                            contentGroup.append("text").attr("x", labelX).attr("y", labelY + 6).text(vPStr)
                                .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).attr("text-anchor", "start");
                        }

                    } else {
                        // 1-Line Mode (Stack OFF): "Label: Value (Percent)"
                        let line1 = vAbsStr + " " + vPStr;
                        if (showVarianceLabel) {
                            const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
                            let refAbbrev = "Ref";
                            const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";
                            const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
                            const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";

                            if (referenceColumn === "budget") refAbbrev = budgetAbbrev;
                            else if (referenceColumn === "py") refAbbrev = pyAbbrev;
                            else if (referenceColumn === "cy") refAbbrev = cyAbbrev;

                            line1 = "∆" + refAbbrev + ": " + line1;
                        }

                        contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(line1)
                            .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                    }

                } else if (showVarianceLabel) {
                    contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(vLabelText)
                        .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                } else {
                    // Single line, no prefix, not both (so either only data or only percent)
                    contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(vValueText)
                        .attr("fill", vColor).style("font-size", labelFontSize).style("font-family", labelFontFamily).style("font-weight", "bold").attr("text-anchor", "start");
                }
            }
        }

    }


    private syncSelectionState(selectionIds: powerbi.visuals.ISelectionId[]) {
        if (!selectionIds || selectionIds.length === 0) {
            this.mainGroup.selectAll(".bar").style("opacity", 1);
            return;
        }

        this.mainGroup.selectAll<SVGRectElement, WaterfallDataPoint>(".bar")
            .style("opacity", d => {
                // Check if selectionId exists (Totals might be null)
                if (!d.selectionId) return 0.5;
                const isSelected = selectionIds.some(id => id.equals(d.selectionId));
                return isSelected ? 1 : 0.5;
            });
    }

    private generateChartData(indices: number[], categoryColumns: powerbi.DataViewCategoryColumn[], groupColumn: powerbi.DataViewCategoryColumn | null, pyValues: any[], pyHighlights: any[], cyValues: any[], cyHighlights: any[], budgetValues: any[], budgetHighlights: any[], tooltipCols: { source: any, values: any[] }[], title: string): WaterfallChartData {
        const increaseColor = this.settings.sentimentColors.increaseColor.value.value;
        const decreaseColor = this.settings.sentimentColors.decreaseColor.value.value;
        const totalColor = "#808080"; // Default total color (setting removed)
        // const showStartTotal removed
        // const showEndTotal removed
        // Default to Standard Variance (PY -> TY) if not specified or set to "none"
        // Fix: Allow "zero"
        // Default to Standard Variance (PY -> TY) if not specified or set to "none"
        // Fix: Allow "zero"
        let startTotalColumn = this.getSafeEnum(this.settings.totalSettings.startTotalColumn.value);
        if (!startTotalColumn || startTotalColumn === "auto" || startTotalColumn === "none") startTotalColumn = "py";

        let endTotalColumn = this.getSafeEnum(this.settings.totalSettings.endTotalColumn.value);
        if (!endTotalColumn || endTotalColumn === "auto" || endTotalColumn === "none") endTotalColumn = "cy";
        let referenceColumn = this.getSafeEnum(this.settings.totalSettings.referenceColumn.value);
        // Auto-detect reference column when showRefMark is enabled but column is "none"
        const showRefMarkSetting = this.settings.totalSettings.showRefMark.value;
        if (showRefMarkSetting && (referenceColumn === "none" || !referenceColumn)) {
            if (budgetValues && budgetValues.length > 0) referenceColumn = "budget";
            else if (pyValues && pyValues.length > 0) referenceColumn = "py";
        }
        const sortBy = this.getSafeEnum(this.settings.sortingSettings.sortBy.value);
        const sortDesc = this.getSafeEnum(this.settings.sortingSettings.sortDirection.value) === "desc";
        // Abbrev Replacements
        const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
        const cyAbbrev = this.settings.localizationSettings.cyAbbrev.value || "CY";
        const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";

        const replaceAbbrev = (name: string, type: string) => {
            if (type === "py") return pyAbbrev;
            if (type === "cy") return cyAbbrev;
            if (type === "budget") return budgetAbbrev;
            return name;
        };

        // Start and End names are now derived from abbreviations only
        let startName = replaceAbbrev("Start", String(startTotalColumn));
        let endName = replaceAbbrev("End", String(endTotalColumn));

        // Reference Name
        let refName = replaceAbbrev("Reference", String(referenceColumn));
        const enableTopN = this.settings.rankingSettings.enable.value;
        const topNCount = this.settings.rankingSettings.count.value;
        const othersLabel = this.settings.rankingSettings.othersLabel.value;

        // Helper to get value (prioritize highlights)
        const getValue = (source: any[], highlights: any[], idx: number) => {
            if (highlights) {
                // If highlights array exists, we are in highlighting/cross-filtering mode
                // If the specific index is null, it means it's not highlighted -> return 0
                return highlights[idx] !== null ? Number(highlights[idx]) : 0;
            }
            if (!source || source.length === 0) return 0;
            return Number(source[idx] || 0);
        };
        // ... (intermediate code skipped by tool - need to be careful with range)
        // Actually, I can't skip code in ReplaceFileContent unless I use separate chunks.
        // I will target the top definitions first.

        // ...


        let rawPoints = [];
        const isHighlighting = (pyHighlights || cyHighlights || budgetHighlights);

        // Subtotal Loop Vars
        let lastParent = null;
        let currentGroupRunningTotal = 0;
        let currentGroupPY = 0;
        let showSubtotals = this.settings.totalSettings.showSubtotals.value;


        // --- Hierarchical Sorting ---
        let processedIndices = [...indices];

        // If hierarchy exists, group indices by Parent Category first
        if (categoryColumns.length > 1) {
            const parentCol = categoryColumns[0];
            const groups = new Map<string, number[]>();
            const groupTotals = new Map<string, number>();
            // Capture Group Sizes for Singlet Detection
            const groupSizes = new Map<string, number>();

            // 1. Group Indices
            processedIndices.forEach(valIndex => {
                const parent = parentCol.values[valIndex] ? String(parentCol.values[valIndex]) : "(Blank)";
                if (!groups.has(parent)) {
                    groups.set(parent, []);
                    groupTotals.set(parent, 0);
                    groupSizes.set(parent, 0);
                }
                groups.get(parent)?.push(valIndex);
                groupSizes.set(parent, (groupSizes.get(parent) || 0) + 1);

                // Calculate Contribution (Absolute Delta for sorting)
                const sVal = getValue(pyValues, pyHighlights, valIndex); // Using helper defined above
                const eVal = getValue(cyValues, cyHighlights, valIndex); // Using helper defined above
                // Note: Delta logic duplicated here for sorting pre-calculation
                // Standard Delta = End - Start.
                // But wait, startValue logic is complex below (switches based on settings).
                // Let's use simple abs(cy - py) for sorting weight approximation, or replicate full logic?
                // Full logic is safer.

                let startV = 0;
                if (startTotalColumn === "py") startV = Number(pyValues ? pyValues[valIndex] : 0);
                else if (startTotalColumn === "cy") startV = Number(cyValues ? cyValues[valIndex] : 0);
                else if (startTotalColumn === "budget") startV = Number(budgetValues ? budgetValues[valIndex] : 0);
                else if (startTotalColumn === "zero") startV = 0;

                let endV = 0;
                if (endTotalColumn === "py") endV = Number(pyValues ? pyValues[valIndex] : 0);
                else if (endTotalColumn === "cy") endV = Number(cyValues ? cyValues[valIndex] : 0);
                else if (endTotalColumn === "budget") endV = Number(budgetValues ? budgetValues[valIndex] : 0);

                const d = endV - startV;
                groupTotals.set(parent, (groupTotals.get(parent) || 0) + d); // Group Total is Algebraic Sum? Or Abs Sum? 
                // Usually for waterfall, largest *movement* is interesting. But standard sort is often by Value.
                // Let's use Algebraic Sum for now (net impact).
            });

            // FIX: If there is only ONE parent group (e.g. drilled down to specific category), 
            // hide the subtotal to avoid redundancy with the main columns.
            if (groups.size <= 1) {
                showSubtotals = false;
            }

            // 2. Sort Groups
            const sortedGroups = Array.from(groups.keys()).sort((a, b) => {
                const totA = groupTotals.get(a) || 0;
                const totB = groupTotals.get(b) || 0;

                // --- Group Sorting Logic ---
                if (sortBy === "default") {
                    // Respect Data Source Order: Sort by the minimum index found in the group
                    const minIndA = Math.min(...(groups.get(a) || [Infinity]));
                    const minIndB = Math.min(...(groups.get(b) || [Infinity]));
                    return minIndA - minIndB;
                }

                if (sortBy === "category") {
                    return sortDesc ? b.localeCompare(a) : a.localeCompare(b);
                }

                // For Value-based Sorts (Delta, CY, PY)
                let valA = 0;
                let valB = 0;

                if (sortBy === "delta") {
                    valA = totA;
                    valB = totB;
                } else if (sortBy === "cy" || sortBy === "py") {
                    // Calculate Aggregated Value for the Group
                    const indsA = groups.get(a) || [];
                    const indsB = groups.get(b) || [];
                    const sum = (inds: number[], col: any[]) => inds.reduce((acc, idx) => acc + (col && col[idx] ? Number(col[idx]) : 0), 0);

                    const targetCol = sortBy === "cy" ? cyValues : pyValues;
                    if (targetCol) {
                        valA = sum(indsA, targetCol);
                        valB = sum(indsB, targetCol);
                    }
                }

                return sortDesc ? valB - valA : valA - valB;
            });

            // --- Apply Top N to Parent Groups ---
            let finalGroups = sortedGroups;
            let othersParentGroupIndices: number[] = []; // Indices from excluded parent groups
            const othersChildrenPerGroup = new Map<string, number[]>(); // Track excluded children per group

            if (enableTopN && topNCount > 0 && sortedGroups.length > topNCount) {
                // For Top N on groups, sort by absolute delta
                const groupsByAbsDelta = [...sortedGroups].sort((a, b) => {
                    return Math.abs(groupTotals.get(b) || 0) - Math.abs(groupTotals.get(a) || 0);
                });
                const topNGroups = groupsByAbsDelta.slice(0, topNCount);
                const otherGroups = groupsByAbsDelta.slice(topNCount);

                // Collect all indices from "other" parent groups
                otherGroups.forEach(gk => {
                    othersParentGroupIndices.push(...(groups.get(gk) || []));
                });

                // Only keep top N groups (in original sort order)
                finalGroups = sortedGroups.filter(g => topNGroups.includes(g));
            }

            // 3. Sort Children within Groups & Apply Child-Level Top N & Flatten
            const childSortCol = categoryColumns[categoryColumns.length - 1]; // Leaf
            processedIndices = [];

            // Helper to calculate delta for an index
            const getChildDelta = (idx: number) => {
                let s = 0, e = 0;
                if (startTotalColumn === "py") s = Number(pyValues ? pyValues[idx] : 0);
                else if (startTotalColumn === "cy") s = Number(cyValues ? cyValues[idx] : 0);
                else if (startTotalColumn === "budget") s = Number(budgetValues ? budgetValues[idx] : 0);
                else if (startTotalColumn === "zero") s = 0;

                if (endTotalColumn === "py") e = Number(pyValues ? pyValues[idx] : 0);
                else if (endTotalColumn === "cy") e = Number(cyValues ? cyValues[idx] : 0);
                else if (endTotalColumn === "budget") e = Number(budgetValues ? budgetValues[idx] : 0);
                return e - s;
            };

            // Helper to sort child indices
            const sortChildIndices = (childIndices: number[]) => {
                return [...childIndices].sort((iA, iB) => {
                    if (sortBy === "default") {
                        return iA - iB;
                    }
                    if (sortBy === "category") {
                        const catA = String(childSortCol.values[iA] || "");
                        const catB = String(childSortCol.values[iB] || "");
                        return sortDesc ? catB.localeCompare(catA) : catA.localeCompare(catB);
                    }
                    let vA = 0, vB = 0;
                    if (sortBy === "delta") {
                        vA = getChildDelta(iA);
                        vB = getChildDelta(iB);
                    } else if (sortBy === "cy") {
                        vA = Number(cyValues ? cyValues[iA] : 0);
                        vB = Number(cyValues ? cyValues[iB] : 0);
                    } else if (sortBy === "py") {
                        vA = Number(pyValues ? pyValues[iA] : 0);
                        vB = Number(pyValues ? pyValues[iB] : 0);
                    }
                    return sortDesc ? vB - vA : vA - vB;
                });
            };

            // Process each top N parent group - show ALL children (no child-level Top N)
            finalGroups.forEach(groupKey => {
                const childIndices = sortChildIndices(groups.get(groupKey) || []);
                processedIndices.push(...childIndices);
            });

            // Process synthetic "Others" parent group (from excluded parent groups)
            // Show ALL children from excluded parent groups - no child-level Top N
            if (enableTopN && othersParentGroupIndices.length > 0) {
                // Sort all children from excluded parent groups
                const othersGroupChildren = sortChildIndices(othersParentGroupIndices);

                // Add sentinel to mark start of "Others" parent group
                processedIndices.push(-998); // Sentinel: Start of "Others" parent group
                processedIndices.push(...othersGroupChildren);
            }
        }
        // Else: Not hierarchical, processedIndices remains as 'indices' (already sorted by Power BI or we accept it)
        // Ensure groupSizes is available even if empty
        const groupCounts = new Map<string, number>();
        if (categoryColumns.length > 1) {
            processedIndices.forEach(i => {
                const p = categoryColumns[0].values[i] ? String(categoryColumns[0].values[i]) : "(Blank)";
                groupCounts.set(p, (groupCounts.get(p) || 0) + 1);
            });
        }

        // Debug Log for Metadata
        console.log("Visual Update: Metadata Check", this.host, categoryColumns);

        // Track current group for "Others" parent handling
        let inOthersParentGroup = false;

        // Helper to aggregate indices into an "Others" row
        const createOthersPoint = (indices: number[], groupName: string | null, isGroupStart: boolean = false) => {
            let othersPy = 0, othersCy = 0, othersRef = 0, othersStart = 0, othersEnd = 0;

            indices.forEach(idx => {
                const py = getValue(pyValues, pyHighlights, idx);
                const cy = getValue(cyValues, cyHighlights, idx);
                const budgetV = getValue(budgetValues, budgetHighlights, idx);

                othersPy += py;
                othersCy += cy;

                let sv = 0;
                if (startTotalColumn === "py") sv = py;
                else if (startTotalColumn === "cy") sv = cy;
                else if (startTotalColumn === "budget") sv = budgetV;

                let ev = 0;
                if (endTotalColumn === "py") ev = py;
                else if (endTotalColumn === "cy") ev = cy;
                else if (endTotalColumn === "budget") ev = budgetV;
                else ev = cy;

                othersStart += sv;
                othersEnd += ev;

                if (referenceColumn === "py") othersRef += py;
                else if (referenceColumn === "cy") othersRef += cy;
                else if (referenceColumn === "budget") othersRef += budgetV;
            });

            const othersDelta = othersEnd - othersStart;

            return {
                category: othersLabel,
                groupName: groupName,
                py: othersPy,
                cy: othersCy,
                ref: othersRef,
                delta: othersDelta,
                originalIndex: -1,
                selectionId: null,
                tooltipValues: [],
                isOthers: true,
                isGroupStart: isGroupStart
            };
        };

        for (const i of processedIndices) {
            // Handle "Others" parent group start sentinel (-998)
            if (i === -998) {
                inOthersParentGroup = true;
                continue;
            }

            // Skip any other negative sentinel values
            if (i < 0) {
                continue;
            }

            const rawPy = getValue(pyValues, pyHighlights, i);
            const rawCy = getValue(cyValues, cyHighlights, i);
            let startVal = 0;
            if (startTotalColumn === "py") startVal = rawPy;
            else if (startTotalColumn === "cy") startVal = rawCy;
            else if (startTotalColumn === "budget") startVal = getValue(budgetValues, budgetHighlights, i);

            let endVal = 0;
            if (endTotalColumn === "py") endVal = rawPy;
            else if (endTotalColumn === "cy") endVal = rawCy;
            else if (endTotalColumn === "budget") endVal = getValue(budgetValues, budgetHighlights, i);
            else endVal = rawCy; // fallback

            let refVal = undefined; // Capture reference value
            if (referenceColumn === "py") refVal = rawPy;
            else if (referenceColumn === "cy") refVal = rawCy;
            else if (referenceColumn === "budget") refVal = getValue(budgetValues, budgetHighlights, i);
            else if (referenceColumn !== "none") {
                // Auto-detect: If user hasn't selected a specific column, but dragged 'Budget', use it.
                if (budgetValues && budgetValues.length > 0) refVal = getValue(budgetValues, budgetHighlights, i);
            }

            // Filter out empty points if highlighting is active 
            // (If all relevant values are 0, and we are highlighting, don't show the flat line)
            if (isHighlighting && rawPy === 0 && rawCy === 0 && (refVal === undefined || refVal === 0)) {
                // Check if start/end match (delta 0)
                const d = endVal - startVal;
                if (d === 0) continue;
            }

            const delta = endVal - startVal;

            const tooltips = tooltipCols.map(tc => {
                const val = tc.values[i];
                let formatted = "";
                if (val != null) {
                    if (tc.source.format) {
                        try {
                            formatted = valueFormatter.create({ format: tc.source.format }).format(val);
                        } catch (e) {
                            formatted = val.toString();
                        }
                    } else {
                        formatted = val.toString();
                    }
                }
                return { displayName: tc.source.displayName, value: formatted };
            });

            let builder = this.host.createSelectionIdBuilder();
            if (groupColumn) {
                builder = builder.withCategory(groupColumn, i);
            }

            // Build Category Label and Selection ID
            let categoryLabelParts = [];
            categoryColumns.forEach(col => {
                builder = builder.withCategory(col, i);
                const val = col.values[i] ? col.values[i].toString() : "(Blank)";
                categoryLabelParts.push(val);
            });
            const categoryLabel = categoryLabelParts.join(" ");

            const catSelectionId = builder.createSelectionId();

            // --- Subtotal Logic ---
            // For children in "Others" parent group, use othersLabel as the parent
            const actualParent = categoryColumns.length > 1 ? (categoryColumns[0].values[i] ? categoryColumns[0].values[i].toString() : "(Blank)") : null;
            const currentParent = inOthersParentGroup ? othersLabel : actualParent;

            if (showSubtotals && currentParent && lastParent && currentParent !== lastParent) {
                // Parent Changed: Insert Subtotal OR Summary for the PREVIOUS group (lastParent)
                const isCollapsed = this.collapsedGroups.has(lastParent);

                if (isCollapsed) {
                    // Collapsed: Push Single Summary Bar (Delta)
                    rawPoints.push({
                        category: lastParent, // Use Group Name
                        groupName: lastParent,
                        py: currentGroupPY,
                        cy: 0, // Not relevant for delta-based
                        ref: undefined,
                        delta: currentGroupRunningTotal,
                        originalIndex: -1,
                        selectionId: null,
                        tooltipValues: [],
                        color: currentGroupRunningTotal >= 0 ? increaseColor : decreaseColor, // Use sentiment color
                        isTotal: false, // It acts as a normal bar
                        isGroupStart: true,
                        isGroupEnd: true,
                        isSingletGroup: groupCounts.get(lastParent) === 1
                    });
                } else {
                    // Expanded: Push Standard Subtotal
                    // SKIP if Singlet Group (Count == 1), to avoid redundancy
                    const gCount = groupCounts.get(lastParent) || 0;
                    if (gCount > 1) {
                        rawPoints.push({
                            category: lastParent + " Total",
                            groupName: lastParent,
                            py: currentGroupPY,
                            cy: currentGroupRunningTotal,
                            ref: undefined,
                            delta: currentGroupRunningTotal,
                            originalIndex: -1,
                            selectionId: null,
                            tooltipValues: [],
                            color: currentGroupRunningTotal >= 0 ? this.settings.sentimentColors.subtotalIncreaseColor.value.value : this.settings.sentimentColors.subtotalDecreaseColor.value.value,
                            isTotal: true,
                            isGroupEnd: true,
                            isSingletGroup: false
                        });
                    }
                }
                currentGroupRunningTotal = 0;
                currentGroupPY = 0;
            }

            let isGroupStart = false;

            if (categoryColumns.length > 1) {
                if (currentParent !== lastParent) {
                    isGroupStart = true;
                }
                currentGroupRunningTotal += delta;
                currentGroupPY += rawPy;
                lastParent = currentParent;
            }

            const leafLabel = categoryColumns.length > 1 ? String(categoryColumns[categoryColumns.length - 1].values[i] || "(Blank)") : categoryLabel;

            // Only push child points if group is NOT collapsed
            if (!this.collapsedGroups.has(currentParent)) {
                rawPoints.push({
                    category: leafLabel,
                    groupName: currentParent,
                    isGroupStart: isGroupStart,
                    py: rawPy,
                    cy: rawCy,
                    ref: refVal, // Store ref value
                    delta,
                    originalIndex: i,
                    selectionId: catSelectionId,

                    tooltipValues: tooltips,
                    isSingletGroup: groupCounts.get(currentParent) === 1
                });
            }
        }

        // Push final subtotal if needed
        // Push final subtotal if needed
        if (showSubtotals && lastParent && categoryColumns.length > 1) {
            const isCollapsed = this.collapsedGroups.has(lastParent);
            if (isCollapsed) {
                rawPoints.push({
                    category: lastParent,
                    groupName: lastParent,
                    py: currentGroupPY,
                    cy: 0,
                    ref: undefined,
                    delta: currentGroupRunningTotal,
                    originalIndex: -1,
                    selectionId: null,
                    tooltipValues: [],
                    color: currentGroupRunningTotal >= 0 ? increaseColor : decreaseColor,
                    isTotal: false,
                    isGroupStart: true,
                    isGroupEnd: true,
                    isSingletGroup: groupCounts.get(lastParent) === 1
                });
            } else {
                const gCount = groupCounts.get(lastParent) || 0;
                if (gCount > 1) {
                    const subtotalIncrease = this.settings.sentimentColors.subtotalIncreaseColor.value.value;
                    const subtotalDecrease = this.settings.sentimentColors.subtotalDecreaseColor.value.value;

                    rawPoints.push({
                        category: lastParent + " Total",
                        groupName: lastParent,
                        py: currentGroupPY,
                        cy: currentGroupRunningTotal,
                        ref: undefined,
                        delta: currentGroupRunningTotal, // Group Subtotal uses Net Delta for color
                        originalIndex: -1,
                        selectionId: null,
                        tooltipValues: [],
                        color: currentGroupRunningTotal >= 0 ? subtotalIncrease : subtotalDecrease,
                        isTotal: true,
                        isGroupEnd: true,
                        isSingletGroup: false,
                        isSubtotal: true // Mark as Subtotal for labeling
                    });
                }
            }
        }

        let processedPoints = [...rawPoints];
        // Global Top N (Only for Flat Data)
        // For Hierarchical Data, Top N is already applied at parent and child levels during sorting
        if (categoryColumns.length <= 1 && enableTopN && topNCount > 0 && processedPoints.length > topNCount) {
            const sortedByDelta = [...processedPoints].sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
            const topNItems = sortedByDelta.slice(0, topNCount);
            const otherItems = sortedByDelta.slice(topNCount);
            const othersDelta = otherItems.reduce((sum, p) => sum + p.delta, 0);
            const othersPy = otherItems.reduce((sum, p) => sum + p.py, 0);
            const othersCy = otherItems.reduce((sum, p) => sum + p.cy, 0);
            const othersRef = otherItems.reduce((sum, p) => sum + (p.ref || 0), 0);
            const othersPoint = {
                category: othersLabel, py: othersPy, cy: othersCy, ref: othersRef, delta: othersDelta, originalIndex: -1, selectionId: null
            };
            processedPoints = [...topNItems, othersPoint];
            processedPoints.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
            rawPoints = processedPoints;
        }
        // Global Sort (Only for Flat Data)
        // For Hierarchical Data, we already sorted indices into Groups -> Children order.
        // Re-sorting here would destroy the hierarchy (interleaving groups based on value).
        if (categoryColumns.length <= 1 && sortBy !== "default") {
            if (sortBy !== "category") {
                rawPoints.sort((a, b) => {
                    let valA = a[sortBy] ?? 0;
                    let valB = b[sortBy] ?? 0;
                    return sortDesc ? valB - valA : valA - valB;
                });
            } else {
                rawPoints.sort((a, b) => {
                    if (a.category === othersLabel) return 1;
                    if (b.category === othersLabel) return -1;

                    const catA = a.category || "";
                    const catB = b.category || "";

                    return sortDesc
                        ? catB.localeCompare(catA)
                        : catA.localeCompare(catB);
                });
            }
        }


        // Variance Bridge Mode Logic
        // 1. Start with Total Start Value (PY, Zero, etc.).
        // 2. Accumulate Deltas.
        // 3. End with Total End Value (TY, etc.).

        // FIX: For Small Multiples (when groupColumn is present), calculate totals from ONLY the indices for this chart
        // For single chart mode (no groupColumn), calculate from entire value arrays to keep Start/End constant during drill-down
        const useFilteredIndices = groupColumn !== null; // Small Multiples mode

        let startTotalValue = 0;
        if (startTotalColumn === "py" && pyValues) {
            if (useFilteredIndices) {
                // Small Multiples: Sum only values for this chart's indices
                for (const idx of indices) {
                    startTotalValue += Number(pyValues[idx] || 0);
                }
            } else {
                // Single chart: Sum all values for consistent drill-down behavior
                for (let i = 0; i < pyValues.length; i++) {
                    startTotalValue += Number(pyValues[i] || 0);
                }
            }
        } else if (startTotalColumn === "cy" && cyValues) {
            if (useFilteredIndices) {
                for (const idx of indices) {
                    startTotalValue += Number(cyValues[idx] || 0);
                }
            } else {
                for (let i = 0; i < cyValues.length; i++) {
                    startTotalValue += Number(cyValues[i] || 0);
                }
            }
        } else if (startTotalColumn === "budget" && budgetValues) {
            if (useFilteredIndices) {
                for (const idx of indices) {
                    startTotalValue += Number(budgetValues[idx] || 0);
                }
            } else {
                for (let i = 0; i < budgetValues.length; i++) {
                    startTotalValue += Number(budgetValues[i] || 0);
                }
            }
        }


        const singleMeasureMode = this.settings.totalSettings.singleMeasureMode.value;
        const hasOnlyCy = cyValues && cyValues.length > 0 && (!pyValues || pyValues.length === 0) && (!budgetValues || budgetValues.length === 0);
        const useSingleMeasureBridge = singleMeasureMode && hasOnlyCy && rawPoints.length >= 2;

        const startColumnColor = this.settings.totalSettings.startColumnColor.value.value;
        const endColumnColor = this.settings.totalSettings.endColumnColor.value.value;

        let dataPoints: WaterfallDataPoint[] = [];

        // Add Start Column
        if (!useSingleMeasureBridge) {
            dataPoints.push({
                category: startName,
                delta: startTotalValue,
                start: 0,
                end: startTotalValue,
                runningTotal: startTotalValue,
                py: 0,
                cy: 0,
                color: startColumnColor, // Use Start Column Color setting
                isTotal: true,
                selectionId: null,
                referenceValue: undefined
            });
        }

        let runningTotal = startTotalValue;

        if (useSingleMeasureBridge) {
            runningTotal = 0; // Starts at 0
            rawPoints[0].isMainTotal = true;
            rawPoints[0].isTotal = true;
            rawPoints[0].colorOverride = startColumnColor;

            rawPoints[rawPoints.length - 1].isMainTotal = true;
            rawPoints[rawPoints.length - 1].isTotal = true;
            rawPoints[rawPoints.length - 1].colorOverride = endColumnColor;
        }

        rawPoints.forEach(p => {
            let start, end;

            const isSubtotal = p.isTotal === true && !p.isMainTotal;
            const isMainTotal = p.isMainTotal === true;

            if (isMainTotal) {
                // Main Total Bar: Starts at 0, goes to delta (value)
                start = 0;
                end = p.delta;
                runningTotal = end; // Reset running total
            } else if (isSubtotal) {
                // Subtotal Bar: Floating bar spanning the group's net change
                end = runningTotal;
                start = runningTotal - p.delta;
            } else {
                // Normal Step Bar
                start = runningTotal;
                end = start + p.delta;
                runningTotal = end;
            }

            dataPoints.push({
                category: p.category,
                delta: p.delta,
                start: start,
                end: end,
                runningTotal: end, // For display
                py: p.py,
                cy: p.cy,
                color: p.colorOverride || (isSubtotal ? p.color : (p.delta >= 0 ? increaseColor : decreaseColor)),
                isTotal: isSubtotal || isMainTotal,
                isSubtotal: isSubtotal,
                selectionId: p.selectionId,
                referenceValue: p.ref,
                // Formula: Start + (Budget - PY)
                // This shows where the bar WOULD have ended if it met the budget.
                // STRICT CHECK: Only calculate if p.ref is valid and NOT zero/null.
                referenceY: (p.ref !== undefined && p.ref !== null && p.py !== undefined)
                    ? start + (p.ref - p.py)
                    : undefined,
                tooltipValues: p.tooltipValues,

                groupName: p.groupName, // Propagate Group Name
                isSingletGroup: p.isSingletGroup
            });
        });


        // Add End Column (Total TY) - Always show in Bridge Mode (unless strictly disabled, but bridge implies it)
        // We defer to showEndTotal if user wants to hide it, but ignore showTY (which is for the background column).
        // Add End Column (Total TY) - Always show in Bridge Mode
        if (!useSingleMeasureBridge) {
            let totalRef = undefined;
            // FIX: For Small Multiples (when groupColumn is present), calculate from ONLY the indices for this chart
            // For single chart mode, calculate from all rows for consistent drill-down behavior
            let refSourceArray: any[] | null = null;
            if (referenceColumn === "py" && pyValues && pyValues.length > 0) refSourceArray = pyValues;
            else if (referenceColumn === "cy" && cyValues && cyValues.length > 0) refSourceArray = cyValues;
            else if (referenceColumn === "budget" && budgetValues && budgetValues.length > 0) refSourceArray = budgetValues;

            if (refSourceArray) {
                totalRef = 0;
                if (useFilteredIndices) {
                    // Small Multiples: Sum only values for this chart's indices
                    for (const idx of indices) {
                        totalRef += Number(refSourceArray[idx] || 0);
                    }
                } else {
                    // Single chart: Sum all values for consistent drill-down behavior
                    for (let i = 0; i < refSourceArray.length; i++) {
                        totalRef += Number(refSourceArray[i] || 0);
                    }
                }

            }

            dataPoints.push({
                category: endName,
                delta: runningTotal,
                start: 0,
                end: runningTotal,
                runningTotal: runningTotal,
                py: 0,
                cy: 0, // Ignored by renderer for waterfall bar
                color: endColumnColor, // Use End Column Color setting
                isTotal: true,
                selectionId: null,
                referenceValue: totalRef,
                referenceY: totalRef // For Totals, reference is absolute
            });
        }






        let minVal = 0;
        let maxVal = 0;
        const showRefMarkInDomain = this.settings.totalSettings.showRefMark.value;
        dataPoints.forEach(d => {
            minVal = Math.min(minVal, d.start, d.end);
            maxVal = Math.max(maxVal, d.start, d.end);
            // Only include referenceY in domain when ref marks are visible
            if (showRefMarkInDomain && d.referenceY !== undefined) {
                minVal = Math.min(minVal, d.referenceY);
                maxVal = Math.max(maxVal, d.referenceY);
            }
        });

        // --- Axis Break / Visual Limit Logic (Limit Column Height) ---
        // Rule: Delta Columns should be max 75% of Min(Start, End).
        // Only if axis break is enabled in settings (re-checked in render, but layout needs domain).
        // We'll calculate the visual limit regardless, but only apply to domain if we decide to break.
        // Actually, we must decide NOW if we want the scale to be small.
        // Let's rely on stored setting? 
        // We can access 'this.settings' inside this method! (It's an instance method).
        const enableBreak = this.settings.yAxisSettings.enableAxisBreak.value;

        if (enableBreak) {
            // Find Start and End Heights (Absolute)
            // Start Total is 'startTotalValue'. End is 'runningTotal'.
            const startH = Math.abs(startTotalValue);
            const endH = Math.abs(runningTotal);
            // Ignore 0 values (e.g. if start is 0)
            const minMain = Math.min(startH || Infinity, endH || Infinity);

            if (minMain !== Infinity && minMain > 0) {
                // --- Axis Break / Moved Baseline Logic (Build 1.4.2.122) ---
                const breakPercent = this.settings.yAxisSettings.axisBreakPercent.value;
                const cutOff = (minMain * breakPercent) / 100;

                // Safety: Ensure we don't hide any actual data points that dip below this cutoff.
                let lowestDataPoint = Infinity;
                dataPoints.forEach(d => {
                    if (d.category === startName || d.category === endName) return;

                    const low = Math.min(d.start, d.end);
                    lowestDataPoint = Math.min(lowestDataPoint, low);

                    // Also check Reference Value if present (e.g. for Variance Line)
                    if (d.referenceValue !== undefined) {
                        lowestDataPoint = Math.min(lowestDataPoint, d.referenceValue);
                    }
                });

                if (lowestDataPoint === Infinity) lowestDataPoint = 0;

                // The Axis Minimum (minVal) can be raised to 'cutOff', BUT limited by 'lowestDataPoint'.
                const finalMin = Math.min(cutOff, lowestDataPoint);

                if (finalMin > 0) {
                    minVal = finalMin;
                }
            }
        }
        return { title, dataPoints, maxValue: maxVal, minValue: minVal, pyTotal: startTotalValue, startName, endName };


    }



    private getBarPath(shape: string, x: number, y: number, w: number, h: number, direction: string): string {
        if (shape !== "arrow" && shape !== "rounded") {
            return `M ${x},${y} h ${w} v ${h} h ${-w} Z`;
        }

        if (shape === "rounded") {
            const r = 5;
            const rx = Math.min(r, w / 2, h / 2);
            if (w < 1 || h < 1) return "";
            return `M ${x + rx},${y} h ${w - 2 * rx} a ${rx},${rx} 0 0 1 ${rx},${rx} v ${h - 2 * rx} a ${rx},${rx} 0 0 1 -${rx},${rx} h -${w - 2 * rx} a ${rx},${rx} 0 0 1 -${rx},-${rx} v -${h - 2 * rx} a ${rx},${rx} 0 0 1 ${rx},-${rx} z`;
        }

        // Arrow
        // For Up/Down, head height is limited by h. Head Width is w.
        // For Left/Right, head width is limited by w. Head Height is h.

        if (direction === "up") {
            const headH = Math.min(h, w * 0.5);
            return `M ${x},${y + headH} L ${x + w / 2},${y} L ${x + w},${y + headH} L ${x + w},${y + h} L ${x},${y + h} L ${x},${y + headH} Z`;
        }
        else if (direction === "down") {
            const headH = Math.min(h, w * 0.5);
            return `M ${x},${y} L ${x + w},${y} L ${x + w},${y + h - headH} L ${x + w / 2},${y + h} L ${x},${y + h - headH} L ${x},${y + h - headH} Z`;
        }
        else if (direction === "right") {
            const headW = Math.min(w, h * 0.5);
            return `M ${x},${y} L ${x + w - headW},${y} L ${x + w},${y + h / 2} L ${x + w - headW},${y + h} L ${x},${y + h} L ${x},${y} Z`;
        }
        else if (direction === "left") {
            const headW = Math.min(w, h * 0.5);
            return `M ${x + headW},${y} L ${x + w},${y} L ${x + w},${y + h} L ${x + headW},${y + h} L ${x},${y + h / 2} L ${x + headW},${y} Z`;
        }

        return `M ${x},${y} h ${w} v ${h} h ${-w} Z`;
    }
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.settings);
    }

                                private async checkLicensing(): Promise<boolean> {
        // 1. Desktop Check (Free)
        if (this.host.hostEnv === powerbi.common.CustomVisualHostEnv.Desktop) {
            return true;
        }

        // 2. Service Check (AppSource)
        try {
            const licenseInfo = await this.host.licenseManager.getAvailableServicePlans();
            if (!licenseInfo || !licenseInfo.plans) {
                return false;
            }
            
            // Check for any active plan
            const hasActivePlan = licenseInfo.plans.some(plan =>
                plan.state === powerbi.ServicePlanState.Active ||
                plan.state === powerbi.ServicePlanState.Warning
            );
            return hasActivePlan;
            
        } catch (err) {
            console.error('License check failed', err);
            return false;
        }
    }

    private debugBoundCount: number = 0;
    private debugBoundNames: string = "";
    private debugLoadedCount: number = 0;

    private checkDrillAvailability(dataView: powerbi.DataView): boolean {
        if (!dataView || !dataView.metadata || !dataView.categorical || !dataView.categorical.categories) return false;

        const categoryRoleName = "category"; // Role name for category fields
        // Find how many columns are bound to Category role in metadata
        const boundColumns = dataView.metadata.columns.filter(c => c.roles && c.roles[categoryRoleName]);
        // Find loaded levels
        const loadedLevels = dataView.categorical.categories.filter(c => c.source.roles && c.source.roles[categoryRoleName]).length;

        this.debugBoundCount = boundColumns.length;
        this.debugBoundNames = boundColumns.map(c => c.displayName).join(", ");
        this.debugLoadedCount = loadedLevels;

        // If defined structure > loaded structure, drill down is possible
        return boundColumns.length > loadedLevels;
    }
    private renderOverlay(show: boolean) {
        if (show) {
            this.licenseOverlay.style("display", "flex");
            // Do not hide SVG, just show overlay on top
        } else {
            this.licenseOverlay.style("display", "none");
        }
    }
}