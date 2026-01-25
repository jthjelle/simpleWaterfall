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
import * as d3 from "d3";

// Types
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IColorPalette = powerbi.extensibility.IColorPalette;

// Interfaces
interface WaterfallDataPoint {
    category: string;
    delta: number;
    start: number;
    end: number;
    py: number;
    ty: number;
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
}

interface WaterfallChartData {
    title: string;
    dataPoints: WaterfallDataPoint[];
    maxValue: number;
    minValue: number;
    pyTotal: number;
    startName?: string;
    endName?: string;
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
    private svg: d3.Selection<SVGSVGElement, any, any, any>;
    private mainGroup: d3.Selection<SVGGElement, any, any, any>;
    private xAxisGroup: d3.Selection<SVGGElement, any, any, any>;
    private yAxisGroup: d3.Selection<SVGGElement, any, any, any>;
    private settings: VisualFormattingSettingsModel;
    private currentSelection: powerbi.visuals.ISelectionId[] = [];
    private events: IVisualEventService;
    private isHighContrast: boolean = false;
    private colorPalette: IColorPalette;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.selectionManager = this.host.createSelectionManager();
        this.target = options.element;

        // Initialize event service for rendering events
        this.events = this.host.eventService;

        // Initialize color palette for high contrast detection
        this.colorPalette = this.host.colorPalette;

        this.settings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, null);

        this.svg = d3.select(this.target)
            .append("svg")
            .classed("waterfall-visual", true);

        this.mainGroup = this.svg.append("g");
        // Legacy groups (kept for structure, though update clears mainGroup and creates new cells)
        this.xAxisGroup = this.mainGroup.append("g").classed("x-axis", true);
        this.yAxisGroup = this.mainGroup.append("g").classed("y-axis", true);

        this.svg.on("click", () => {
            this.selectionManager.clear().then(() => {
                this.currentSelection = [];
                this.syncSelectionState([]);
            });
        });

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
        // Middle - opaque (white in mask)
        gradient.append("stop").attr("offset", "50%").attr("stop-color", "white").attr("stop-opacity", 1);
        gradient.append("stop").attr("offset", "100%").attr("stop-color", "white").attr("stop-opacity", 1);

        const mask = defs.append("mask")
            .attr("id", "axis-break-mask")
            .attr("maskContentUnits", "objectBoundingBox");

        mask.append("rect")
            .attr("x", 0).attr("y", 0).attr("width", 1).attr("height", 1)
            .attr("fill", "url(#axis-break-gradient)");
    }

    public update(options: VisualUpdateOptions) {
        // Signal rendering started
        this.events.renderingStarted(options);

        // Landing Page Logic (Zero State)
        if (!options.dataViews || !options.dataViews[0] || !options.dataViews[0].categorical || !options.dataViews[0].categorical.categories || !options.dataViews[0].categorical.values) {
            this.mainGroup.selectAll("*").remove();
            this.mainGroup.append("text")
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
                console.error("Error rendering chart for " + chart.title, e);
                this.host.displayWarningIcon("Rendering Error", "An error occurred while rendering the visual. Please check your data.");
            }
        });

        // Signal rendering finished
        this.events.renderingFinished(options);
    }

    private formatNumber(value: number, decimalPlaces: number, useThousandsSeparator: boolean, numberScale?: string, thousandsAbbrev?: string, millionsAbbrev?: string): string {
        let scaledValue = value;
        let suffix = "";

        if (numberScale === "thousands") {
            scaledValue = value / 1000;
            suffix = thousandsAbbrev || "K";
        } else if (numberScale === "millions") {
            scaledValue = value / 1000000;
            suffix = millionsAbbrev || "M";
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

    private transformData(options: VisualUpdateOptions): WaterfallViewModel {
        const dataViews = options.dataViews;
        if (!dataViews || !dataViews[0] || !dataViews[0].categorical) {
            return { charts: [], globalMaxValue: 0, globalMinValue: 0 };
        }

        const categorical = dataViews[0].categorical;
        let categoryCol = null;
        let groupCol = null;

        if (categorical.categories) {
            for (const cat of categorical.categories) {
                if (cat.source.roles["smallMultiples"]) groupCol = cat;
                if (cat.source.roles["category"]) categoryCol = cat;
            }
        }
        if (!categoryCol && categorical.categories && categorical.categories.length > 0) {
            categoryCol = categorical.categories[0];
        }

        if (!categoryCol) return { charts: [], globalMaxValue: 0, globalMinValue: 0 };

        const values = categorical.values;
        let pyValues: any[] = [];
        let tyValues: any[] = [];
        let budgetValues: any[] = [];
        let tooltipCols: { source: any, values: any[] }[] = [];

        if (values) {
            values.forEach(v => {
                if (v.source.roles["py"]) pyValues = v.values;
                if (v.source.roles["ty"]) tyValues = v.values;
                if (v.source.roles["budget"]) budgetValues = v.values;
                if (v.source.roles["tooltips"]) tooltipCols.push({ source: v.source, values: v.values });
            });
        }

        const charts: WaterfallChartData[] = [];
        const numPoints = categoryCol.values.length;

        if (groupCol) {
            const groups = new Map<string, number[]>();
            for (let i = 0; i < numPoints; i++) {
                const gVal = groupCol.values[i] ? groupCol.values[i].toString() : "Undefined";
                if (!groups.has(gVal)) groups.set(gVal, []);
                groups.get(gVal).push(i);
            }

            groups.forEach((indices, name) => {
                charts.push(this.generateChartData(indices, categoryCol, groupCol, pyValues, tyValues, budgetValues, tooltipCols, name));
            });
        } else {
            const indices = Array.from({ length: numPoints }, (_, i) => i);
            charts.push(this.generateChartData(indices, categoryCol, null, pyValues, tyValues, budgetValues, tooltipCols, ""));
        }

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

        if (!viewModel || viewModel.dataPoints.length === 0) return;

        const showConnectors = this.settings.layoutSettings.showConnectors.value;
        const connectorColor = this.settings.layoutSettings.connectorColor.value.value;
        const showPYLabel = this.settings.columnSettings.showPYLabel.value;
        const showTYLabel = this.settings.columnSettings.showTYLabel.value;


        const showYAxisLabels = this.settings.yAxisSettings.showLabels.value;
        const labelOrientation = String(this.settings.xAxisSettings.labelOrientation.value.value);

        const orientation = this.settings.layoutSettings.orientation.value.value; // "vertical" or "horizontal"
        const isHorizontal = orientation === "horizontal";
        const barShape = this.settings.layoutSettings.barShape.value.value;

        let leftMargin = showYAxisLabels ? Math.max(40, Math.min(60, width * 0.15)) : 10;
        let bottomMargin = 20;
        if (labelOrientation === "vertical") {
            bottomMargin = Math.min(150, height * 0.4);
        } else if (labelOrientation === "angled") {
            bottomMargin = Math.min(100, height * 0.4);
        }

        const minChartHeight = 20;
        if (height - bottomMargin - 10 < minChartHeight) {
            bottomMargin = Math.max(0, height - minChartHeight - 10);
        }

        const showSummaryIndicator = this.settings.totalSettings.showSummaryIndicator.value;
        const showVariance = this.settings.totalSettings.showVariance.value;
        const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;

        // Increase margins for indicators
        let rightMargin = 10;
        let topMargin = 10;

        if (isHorizontal) {
            // Horizontal Mode: Indicators at Top
            if (showSummaryIndicator) topMargin = 50;
            if (showVariance && referenceColumn !== "none") topMargin = 70;
        } else {
            // Vertical Mode: Indicators at Right
            // Variance Indicator (End vs Ref) needs ~120px (offset 70px + text).
            // Summary Indicator (Start vs End) needs ~60px (offset 10px + text).

            if (showVariance && referenceColumn !== "none") {
                rightMargin = 120;
            } else if (showSummaryIndicator) {
                rightMargin = 60;
            }
        }

        const margin = { top: topMargin, right: rightMargin, bottom: bottomMargin, left: leftMargin };
        const innerWidth = Math.max(10, width - margin.left - margin.right);
        const innerHeight = Math.max(10, height - margin.top - margin.bottom);

        let contentGroup = targetGroup.select<SVGGElement>(".content-group");
        if (contentGroup.empty()) {
            contentGroup = targetGroup.append("g").classed("content-group", true);
        }
        contentGroup.attr("transform", `translate(${margin.left}, ${margin.top})`);

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
            let xMin = globalMinMax ? globalMinMax.min : Math.min(0, viewModel.minValue);
            let xMax = globalMinMax ? globalMinMax.max : Math.max(0, viewModel.maxValue);

            const settingsMin = this.settings.yAxisSettings.min.value; // Reuse min/max settings for value axis
            const settingsMax = this.settings.yAxisSettings.max.value;
            if (settingsMin != null) xMin = settingsMin;
            if (settingsMax != null) xMax = settingsMax;

            let xDomainPadding = (xMax - xMin) * 0.1;
            const minStartZero = xMin === 0;

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

            let yMin = globalMinMax ? globalMinMax.min : Math.min(0, viewModel.minValue);
            let yMax = globalMinMax ? globalMinMax.max : Math.max(0, viewModel.maxValue); // Fix typo globalMinMax.max

            const settingsMin = this.settings.yAxisSettings.min.value;
            const settingsMax = this.settings.yAxisSettings.max.value;
            if (settingsMin != null) yMin = settingsMin;
            if (settingsMax != null) yMax = settingsMax;

            let yDomainPadding = (yMax - yMin) * 0.2;
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
                        yDomainPadding = (yMax - yMin) * 0.2;
                    }
                }
            }

            const invert = this.settings.yAxisSettings.invert.value;
            yScale = d3.scaleLinear()
                .domain([minStartZero && !enableAxisBreak ? yMin : yMin - yDomainPadding, yMax + yDomainPadding])
                .range(invert ? [0, innerHeight] : [innerHeight, 0]);
        }

        // Secondary Scale Logic
        let xScale2: any = xScale;
        let yScale2: any = yScale;


        // Always calculate secondary scale (for PY/TY columns) regardless of axis visibility
        // This ensures they are scaled relative to the Start/End totals constraint
        const secMinSetting = null;
        const secMaxSetting = null;

        // Calculate Domain from Columns
        const colValues = viewModel.dataPoints.flatMap(d => [d.py, d.ty, d.referenceValue]).filter(v => v !== undefined && v !== null);
        let secMin = colValues.length > 0 ? Math.min(0, ...colValues) : 0;
        let secMax = colValues.length > 0 ? Math.max(0, ...colValues) : 0;

        // Scaling Constraint: Columns not higher than lowest of start/end
        const endName = viewModel.endName;
        const startPoint = viewModel.dataPoints.find(d => d.isTotal && !d.isReference && d.originalIndex === undefined && d.category !== endName);
        const endPoint = viewModel.dataPoints.find(d => d.category === endName);

        let limitValue = Infinity;
        if (startPoint) limitValue = Math.min(limitValue, Math.abs(startPoint.end));
        if (endPoint) limitValue = Math.min(limitValue, Math.abs(endPoint.end));

        if (limitValue !== Infinity && limitValue > 0) {
            // Calculate required secMax to satisfy constraint
            // We must compare VISUAL HEIGHTS.
            // If primary axis does not start at 0 (e.g. broken axis or zoomed), the visual height of "LimitValue"
            // is not proportional to LimitValue itself, but to (LimitValue - AxisMin).

            const maxColVal = Math.max(0, ...colValues);
            if (maxColVal > 0) {
                let primeRange = 0;
                let primeMin = 0;

                if (isHorizontal) {
                    const domain = xScale.domain();
                    primeMin = Math.min(domain[0], domain[1]);
                    // Assuming value axis might be inverted in logic, but domain has min/max.
                    // Usually domain is [min, max] or [max, min].
                    primeRange = Math.abs(domain[1] - domain[0]);
                } else {
                    const domain = yScale.domain();
                    primeMin = Math.min(domain[0], domain[1]);
                    primeRange = Math.abs(domain[1] - domain[0]);
                }

                // Visible Height (in value units) of the Limit Bar
                // LimitBar usually starts at 0. So visible part is [PrimeMin, LimitValue].
                // Length = LimitValue - PrimeMin. (Clamped to 0).
                // If PrimeMin < 0, Length = LimitValue - PrimeMin (full bar).
                // Usually waterfall values are positive. If PrimeMin > 0, we lose the bottom part.

                // Note: limitValue is Absolute.
                // We use Math.max(0, ...) to ensure we don't divide by 0 or negative.
                let visibleLimitLength = limitValue;
                if (primeMin > 0) {
                    visibleLimitLength = Math.max(0, limitValue - primeMin);
                }

                if (visibleLimitLength > 0) {
                    // Formula: VisualColHeight <= VisualLimitHeight
                    // (MaxCol / SecRange) * PlotSize <= (VisibleLimit / PrimeRange) * PlotSize
                    // SecRange >= (MaxCol * PrimeRange) / VisibleLimit

                    const requiredMinRange = (maxColVal * primeRange) / visibleLimitLength;

                    const currentRange = secMax - secMin;
                    if (currentRange < requiredMinRange) {
                        // Adjust secMax (assuming secMin is usually 0, or we just expand range)
                        secMax = secMin + requiredMinRange;
                    }
                }
            }
        }

        if (secMinSetting != null) secMin = secMinSetting;
        if (secMaxSetting != null) secMax = secMaxSetting;

        // Paddings
        const padding = (secMax - secMin) * 0.1;

        if (isHorizontal) {
            // Secondary Axis determines X (Value)
            const invert = this.settings.yAxisSettings.invert.value;
            xScale2 = d3.scaleLinear()
                .domain([secMin, secMax + padding])
                .range(invert ? [innerWidth, 0] : [0, innerWidth]);
        } else {
            // Secondary Axis determines Y (Value)
            const invert = this.settings.yAxisSettings.invert.value;
            yScale2 = d3.scaleLinear()
                .domain([secMin, secMax + padding])
                .range(invert ? [0, innerHeight] : [innerHeight, 0]);
        }

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
                    if (text.length > labelMaxChars) return text.substring(0, labelMaxChars) + "...";
                    return text;
                });
            }
        }

        const xAxisG = xAxisGroup.attr("transform", `translate(0, ${innerHeight})`).call(xAxis);
        const yAxisG = yAxisGroup.call(yAxis);

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

        // Secondary Axis Rendering - Hidden but scale preserved
        contentGroup.select(".secondary-axis").remove();


        // Axis styling
        yAxisGroup.selectAll("text").style("display", showYAxisLabels ? null : "none");
        yAxisGroup.selectAll("line").style("display", showYAxisLabels ? null : "none");
        yAxisGroup.select("path.domain").style("display", showYAxisLabels ? null : "none");

        // Rotate labels if needed (mostly for Band scale)
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
            const breakY = innerHeight - 5;
            contentGroup.append("path")
                .classed("axis-break", true)
                .attr("d", `M -10 ${breakY + 6} L -5 ${breakY + 3} L -10 ${breakY} L -5 ${breakY - 3} L -10 ${breakY - 6}`)
                .attr("stroke", "#999").attr("stroke-width", 1.5).attr("fill", "none");
        }

        const showLabels = this.settings.dataLabelSettings.show.value;

        // Column Settings
        const showPY = this.settings.columnSettings.showPY.value;
        const showTY = this.settings.columnSettings.showTY.value;
        const pyColor = this.settings.columnSettings.pyColumnColor.value.value;
        const tyColor = this.settings.columnSettings.tyColumnColor.value.value;
        const overlapPercent = 90; // Fixed 90% overlap
        const flipOverlap = this.settings.columnSettings.flipOverlap.value;
        // showSecondaryAxis removed


        // Slot Layout Calculation
        let slotCount = 1;
        if (showPY) slotCount++;
        if (showTY) slotCount++;

        let pySlot = -1, wfSlot = -1, tySlot = -1;
        let currentSlot = 0;

        // Variance Mode: Waterfall Bar overlays the reference columns.
        wfSlot = currentSlot; // Use current slot but don't increment (floats on top)

        if (!flipOverlap) {
            if (showPY) pySlot = currentSlot++;
            if (showTY) tySlot = currentSlot++;
        } else {
            if (showTY) tySlot = currentSlot++;
            if (showPY) pySlot = currentSlot++;
        }

        const usedCount = currentSlot;

        const getSlotInfo = (bandwidth: number, slotIndex: number) => {
            const overlapPercent = 90;
            const overlapFactor = Math.max(0, Math.min(100, overlapPercent)) / 100;
            const padding = Math.min(bandwidth * 0.05, 10);

            // If usedCount is 0 (e.g. no columns, just waterfall), then 1?
            const count = Math.max(1, usedCount);

            const availableWidth = bandwidth - (count - 1) * padding;
            const width = availableWidth / (1 + (count - 1) * (1 - overlapFactor));
            const step = width * (1 - overlapFactor) + padding;
            const offset = slotIndex * step;

            return { offset, width };
        };

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
                    const centerOffset = slot.offset + slot.width / 2;

                    line.attr("x1", xScale(current.end))
                        .attr("y1", yScale(current.category) + centerOffset)
                        .attr("x2", xScale(current.end))
                        .attr("y2", yScale(next.category)); // Need next category center offset? 
                    // Connectors usually go from End of Prev to Start of Next? 
                    // Vertical: goes from Prev Bar End-Y to Next Bar Start-Y?
                    // Simple Waterfall: Line connects Prev End Value to Next Start Value.
                    // Ideally: Prev Bar (X, Y-end) -> Next Bar (X, Y-start??). 
                    // Actually it's a step line.
                    // Existing logic was: x1=xScale(end), y1=yScale(cat)+bw, x2=xScale(end), y2=yScale(nextCat).
                    // This implies vertical line segment between horizontal bars.
                    // It stays at "end" value.
                    // So I just need to center it in the Band (Offset).
                    // y2 should also be centered?
                    // y2 is Top of Next bar block. But we want it to connect to the bar?
                    // Standard connector connects corners.
                    // Here it connects End Value across categories.
                    // So y1 is centered in Prev. y2 is centered in Next?
                    // Wait, existing logic `y2 = yScale(next.category)` which is TOP of next band.
                    // It draws a vertical line from Center of Prev Band? No.
                    // `y1 = yScale(current.category) + bandwidth`. Bottom of current band.
                    // `y2 = yScale(next.category)`. Top of next band.
                    // So it draws line in the GAP between bands.
                    // Ah. And `x1=x2=xScale(current.end)`.
                    // So it's a straight line at Value X.
                    // Correct.
                    // BUT if we have slots, the waterfall bar is SHIFTED.
                    // Does that affect X (value)? No. X is Value.
                    // Does it affect Y?
                    // If Connector connects Bottom of Prev Bar to Top of Next Bar...
                    // Yes. But bars are now thinner.
                    // `y1 = yScale(current.cat) + offset + width` (Bottom of bar).
                    // `y2 = yScale(next.cat) + offset` (Top of bar).

                    line.attr("y1", yScale(current.category) + slot.offset + slot.width)
                        .attr("y2", yScale(next.category) + slot.offset);

                } else {
                    const bandwidth = xScale.bandwidth();
                    const slot = getSlotInfo(bandwidth, wfSlot);
                    const centerOffset = slot.offset + slot.width / 2;

                    // Vertical: Connector at Y value (Value axis).
                    // x1 = Center of Prev Bar? No, End of Prev Bar layout-wise?
                    // Existing: `x1 = xScale(cat) + bandwidth`. Right edge of Prev Band.
                    // `x2 = xScale(next.cat)`. Left edge of Next Band.
                    // Connected at `yScale(current.end)`.
                    // So it draws horizontal line across gap.

                    line.attr("x1", xScale(current.category) + slot.offset + slot.width)
                        .attr("y1", yScale(current.end))
                        .attr("x2", xScale(next.category) + slot.offset)
                        .attr("y2", yScale(current.end));
                }
            }
        }

        // remove old supplemental columns
        contentGroup.selectAll(".supplemental-col").remove();

        // Draw PY Columns
        // Helper functions for drawing columns to allow reordering
        const drawPY = () => {
            if (!showPY) return;
            contentGroup.selectAll(".py-col")
                .data(viewModel.dataPoints.filter(d => !d.isSpacer && !d.isTotal))
                .enter().append("rect")
                .classed("supplemental-col py-col", true)
                .attr("tabindex", 0)
                .attr("fill", pyColor)
                .attr("width", d => {
                    if (isHorizontal) {
                        const base = Math.max(0, xScale2.domain()[0]);
                        return Math.abs(xScale2(d.py) - xScale2(base || 0));
                    }
                    const bw = xScale.bandwidth();
                    return getSlotInfo(bw, pySlot).width;
                })
                .attr("height", d => {
                    if (isHorizontal) {
                        const bw = yScale.bandwidth();
                        return getSlotInfo(bw, pySlot).width;
                    }
                    const base = Math.max(0, yScale2.domain()[0]);
                    return Math.abs(yScale2(d.py) - yScale2(base || 0));
                })
                .attr("x", d => {
                    if (isHorizontal) {
                        const base = Math.max(0, xScale2.domain()[0]);
                        return Math.min(xScale2(d.py), xScale2(base || 0));
                    }
                    const bw = xScale.bandwidth();
                    return xScale(d.category) + getSlotInfo(bw, pySlot).offset;
                })
                .attr("y", d => {
                    if (isHorizontal) {
                        const bw = yScale.bandwidth();
                        return yScale(d.category) + getSlotInfo(bw, pySlot).offset;
                    }
                    const base = Math.max(0, yScale2.domain()[0]);
                    return Math.min(yScale2(d.py), yScale2(base || 0));
                })
                .attr("rx", barShape === "rounded" ? 3 : 0)
                .attr("ry", barShape === "rounded" ? 3 : 0);

            // PY Labels
            if (showPYLabel) {
                contentGroup.selectAll(".py-label").remove();
                contentGroup.selectAll(".py-label")
                    .data(viewModel.dataPoints.filter(d => !d.isSpacer && !d.isTotal))
                    .enter().append("text")
                    .classed("py-label", true)
                    .attr("x", d => {
                        if (isHorizontal) {
                            const base = Math.max(0, xScale2.domain()[0]);
                            const w = Math.abs(xScale2(d.py) - xScale2(base || 0));
                            const xVal = Math.min(xScale2(d.py), xScale2(base || 0));
                            return xVal + w / 2;
                        }
                        const bw = xScale.bandwidth();
                        return xScale(d.category) + getSlotInfo(bw, pySlot).offset + getSlotInfo(bw, pySlot).width / 2;
                    })
                    .attr("y", d => {
                        if (isHorizontal) {
                            const bw = yScale.bandwidth();
                            return yScale(d.category) + getSlotInfo(bw, pySlot).offset + getSlotInfo(bw, pySlot).width / 2 + 4;
                        }
                        const base = Math.max(0, yScale2.domain()[0]); // Match column mapping
                        const yVal = Math.min(yScale2(d.py), yScale2(base || 0));
                        return yVal - 5;
                    })
                    .text(d => this.formatNumber(d.py, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev))
                    .style("font-size", "10px")
                    .attr("text-anchor", "middle")
                    .attr("fill", "black");
            }
        };

        const drawTY = () => {
            if (!showTY) return;
            contentGroup.selectAll(".ty-col")
                .data(viewModel.dataPoints.filter(d => !d.isSpacer && !d.isTotal))
                .enter().append("rect")
                .classed("supplemental-col ty-col", true)
                .attr("tabindex", 0)
                .attr("fill", tyColor)
                .attr("width", d => {
                    if (isHorizontal) {
                        const base = Math.max(0, xScale2.domain()[0]);
                        return Math.abs(xScale2(d.ty) - xScale2(base || 0));
                    }
                    const bw = xScale.bandwidth();
                    return getSlotInfo(bw, tySlot).width;
                })
                .attr("height", d => {
                    if (isHorizontal) {
                        const bw = yScale.bandwidth();
                        return getSlotInfo(bw, tySlot).width;
                    }
                    const base = Math.max(0, yScale2.domain()[0]);
                    return Math.abs(yScale2(d.ty) - yScale2(base || 0));
                })
                .attr("x", d => {
                    if (isHorizontal) {
                        const base = Math.max(0, xScale2.domain()[0]);
                        return Math.min(xScale2(d.ty), xScale2(base || 0));
                    }
                    const bw = xScale.bandwidth();
                    return xScale(d.category) + getSlotInfo(bw, tySlot).offset;
                })
                .attr("y", d => {
                    if (isHorizontal) {
                        const bw = yScale.bandwidth();
                        return yScale(d.category) + getSlotInfo(bw, tySlot).offset;
                    }
                    const base = Math.max(0, yScale2.domain()[0]);
                    return Math.min(yScale2(d.ty), yScale2(base || 0));
                })
                .attr("rx", barShape === "rounded" ? 3 : 0)
                .attr("ry", barShape === "rounded" ? 3 : 0);

            // TY Labels
            if (showTYLabel) {
                contentGroup.selectAll(".ty-label").remove();
                contentGroup.selectAll(".ty-label")
                    .data(viewModel.dataPoints.filter(d => !d.isSpacer && !d.isTotal))
                    .enter().append("text")
                    .classed("ty-label", true)
                    .attr("x", d => {
                        if (isHorizontal) {
                            const base = Math.max(0, xScale2.domain()[0]);
                            const w = Math.abs(xScale2(d.ty) - xScale2(base || 0));
                            const xVal = Math.min(xScale2(d.ty), xScale2(base || 0));
                            return xVal + w / 2;
                        }
                        const bw = xScale.bandwidth();
                        return xScale(d.category) + getSlotInfo(bw, tySlot).offset + getSlotInfo(bw, tySlot).width / 2;
                    })
                    .attr("y", d => {
                        if (isHorizontal) {
                            const bw = yScale.bandwidth();
                            return yScale(d.category) + getSlotInfo(bw, tySlot).offset + getSlotInfo(bw, tySlot).width / 2 + 4;
                        }
                        const base = Math.max(0, yScale2.domain()[0]);
                        const yVal = Math.min(yScale2(d.ty), yScale2(base || 0));
                        return yVal - 5;
                    })
                    .text(d => this.formatNumber(d.ty, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev))
                    .style("font-size", "10px")
                    .attr("text-anchor", "middle")
                    .attr("fill", "black");
            }
        };

        // Execute drawing based on Flip Order
        // If flipped, we want the "Latest" slot (usually Ty or Wf) to be on *bottom* if we want to see what's behind?
        // No, standard Painter's algo: Last drawn is on top.
        // If Flip Overlap: We want the "First" slot (which was PY) to be on Top?
        // Or just reverse the sequence?

        if (flipOverlap) {
            // Draw TY first, then PY (So PY is on top of TY)
            drawTY();
            drawPY();
        } else {
            // Default: Draw PY first, then TY (So TY is on top of PY)
            drawPY();
            drawTY();
        }

        // Connectors Logic Removed


        // Draw Bars (Waterfall)
        const bars = contentGroup.selectAll(".bar")
            .data(viewModel.dataPoints.filter(d => !d.isSpacer))
            .enter().append("path")
            .classed("bar", true)
            .attr("tabindex", 0);

        bars.attr("fill", d => {
            if (d.isReference) return "none";
            return d.color;
        })
            .attr("mask", d => {
                const enableAxisBreak = this.settings.yAxisSettings.enableAxisBreak.value;
                if (enableAxisBreak && d.isTotal && !d.isReference && !isHorizontal) {
                    return "url(#axis-break-mask)";
                }
                return null;
            })
            .attr("d", d => {
                let x = 0, y = 0, w = 0, h = 0;
                let direction = "none";

                // Standard Variance Mode - Using Secondary Axis (indicatorXScale/YScale)
                // We use xScale/yScale for Waterfall Bars (Primary Axis)
                // This ensures they match the visible axis and work even if PY/TY are missing.
                const activeXScale = xScale;
                const activeYScale = yScale;

                const usedCount = Math.max(1, currentSlot);

                if (isHorizontal) {
                    const bw = yScale.bandwidth();
                    const firstSlot = getSlotInfo(bw, 0);
                    const lastSlot = getSlotInfo(bw, usedCount - 1);
                    const clusterOffset = firstSlot.offset;
                    const clusterWidth = (lastSlot.offset + lastSlot.width) - clusterOffset;

                    y = yScale(d.category) + clusterOffset;
                    h = clusterWidth;

                    const axisBase = activeXScale.domain()[0];
                    const base = Math.max(0, axisBase);
                    const startX = d.isTotal ? activeXScale(base) : activeXScale(d.start);
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
                    const base = Math.max(0, axisBottom);
                    const startY = d.isTotal ? activeYScale(base) : activeYScale(d.start);
                    const endY = activeYScale(d.end);

                    y = Math.min(startY, endY);
                    h = Math.abs(startY - endY);

                    if (!d.isTotal) {
                        if (d.delta >= 0) direction = "up";
                        else direction = "down";
                    }
                }

                return this.getBarPath(String(barShape), x, y, w, h, direction);
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

            // Remove old text-only labels logic to avoid duplicates if switching modes
            contentGroup.selectAll("text.label").remove();

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

                const val = showRunningTotal ? d.runningTotal : d.delta;
                const absStr = this.formatNumber(val, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);

                let mainText = absStr;
                let subText = "";

                if (!d.isTotal) {
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

                const deltaStr = this.formatNumber(d.delta, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
                const runningTotalStr = this.formatNumber(d.runningTotal, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);

                const tooltips: VisualTooltipDataItem[] = [
                    { displayName: "Category", value: d.category },
                    { displayName: "Delta", value: deltaStr },
                    { displayName: "Running Total", value: runningTotalStr }
                ];

                if (!d.isTotal) {
                    const pyStr = this.formatNumber(d.py, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
                    const tyStr = this.formatNumber(d.ty, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
                    const refStr = d.referenceValue !== undefined ? this.formatNumber(d.referenceValue, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev) : "N/A";

                    tooltips.push(
                        { displayName: "PY", value: pyStr },
                        { displayName: "TY", value: tyStr }
                    );
                    if (d.referenceValue !== undefined) {
                        tooltips.push({ displayName: "Budget/Ref", value: refStr });
                    }
                }

                if (d.tooltipValues) {
                    d.tooltipValues.forEach(tv => {
                        tooltips.push({ displayName: tv.displayName, value: tv.value });
                    });
                }

                this.host.tooltipService.show({
                    coordinates: [event.clientX, event.clientY],
                    isTouchEvent: false,
                    dataItems: tooltips,
                    identities: d.selectionId ? [d.selectionId] : []
                });
            }).on("mouseout", () => {
                this.host.tooltipService.hide({ isTouchEvent: false, immediately: true });
                this.syncSelectionState(this.currentSelection);
            });

        // Apply interactions to Supplemental Columns
        contentGroup.selectAll(".supplemental-col")
            .on("click", (event, d: WaterfallDataPoint) => {
                const isMultiSelect = event.ctrlKey || event.metaKey;
                if (d.selectionId) {
                    this.selectionManager.select(d.selectionId, isMultiSelect).then((ids: powerbi.visuals.ISelectionId[]) => {
                        this.currentSelection = ids;
                        this.syncSelectionState(ids);
                    });
                }
                event.stopPropagation();
            })
            .on("contextmenu", (event, d: WaterfallDataPoint) => {
                const mouseEvent: MouseEvent = event as MouseEvent;
                this.selectionManager.showContextMenu(d.selectionId ? d.selectionId : {}, {
                    x: mouseEvent.clientX,
                    y: mouseEvent.clientY
                });
                mouseEvent.preventDefault();
                mouseEvent.stopPropagation();
            })
            .on("keydown", (event, d: WaterfallDataPoint) => {
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
            .on("mouseover", (event, d: WaterfallDataPoint) => {
                // Reuse tooltip logic? Or simplify.
                // Let's reuse basic tooltip logic logic for consistency or just call standard tooltip
                // For now, minimal tooltip or copy logic?
                // Copy logic locally for simplicity
                this.mainGroup.selectAll(".supplemental-col").style("opacity", 1); // Reset

                const tooltips: VisualTooltipDataItem[] = [
                    { displayName: "Category", value: d.category }
                ];
                // If it's PY Col, show PY. If TY, show TY.
                // d has py/ty props.
                const decimalPlaces = this.settings.dataLabelSettings.decimalPlaces.value;
                // We need access to formatNumber... it's a private method on class 'this'.
                // We are inside renderChart => this is Visual instance.

                // However, tooltips logic above was complex (built list).
                // Use simple tooltip for columns
                const pyStr = this.formatNumber(d.py, 0, true); // Simplified formatting or reuse settings?
                const tyStr = this.formatNumber(d.ty, 0, true);

                // Identify if clicked element is PY or TY? 
                // d is same for both.
                // We can check class list of event target?
                const target = d3.select(event.currentTarget);
                if (target.classed("py-col")) {
                    tooltips.push({ displayName: "PY", value: pyStr });
                } else if (target.classed("ty-col")) {
                    tooltips.push({ displayName: "TY", value: tyStr });
                }

                this.host.tooltipService.show({
                    coordinates: [event.clientX, event.clientY],
                    isTouchEvent: false,
                    dataItems: tooltips,
                    identities: d.selectionId ? [d.selectionId] : []
                });
            })
            .on("mouseout", () => {
                this.host.tooltipService.hide({ isTouchEvent: false, immediately: true });
            });

        if (showSummaryIndicator && viewModel.pyTotal !== 0) {
            const totalDelta = viewModel.dataPoints.filter(d => !d.isTotal).reduce((sum, d) => sum + d.delta, 0);
            const totalPercent = (totalDelta / viewModel.pyTotal) * 100;
            const absStr = this.formatNumber(totalDelta, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
            const pStr = "(" + (totalPercent >= 0 ? "+" : "") + totalPercent.toFixed(percentDecimalPlaces) + "%)";

            const indicatorColor = totalDelta >= 0 ? this.settings.totalSettings.summaryPositiveColor.value.value : this.settings.totalSettings.summaryNegativeColor.value.value;

            // Summary Indicator Logic
            if (isHorizontal) {
                const xStart = xScale(viewModel.pyTotal);
                const xEnd = xScale(viewModel.pyTotal + totalDelta);
                const bracketY = -25; // Top Margin

                // Draw Horizontal Bracket
                contentGroup.append("line").attr("x1", xStart).attr("y1", 0).attr("x2", xStart).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xEnd).attr("y1", 0).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");
                contentGroup.append("line").attr("x1", xStart).attr("y1", bracketY).attr("x2", xEnd).attr("y2", bracketY)
                    .attr("stroke", indicatorColor).attr("stroke-width", 2);

                contentGroup.append("circle").attr("cx", xStart).attr("cy", bracketY).attr("r", 3).attr("fill", indicatorColor);
                contentGroup.append("circle").attr("cx", xEnd).attr("cy", bracketY).attr("r", 3).attr("fill", indicatorColor);

                // Labels
                const labelX = (xStart + xEnd) / 2;
                const labelY = bracketY - 8;
                contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(absStr + " " + pStr)
                    .attr("fill", indicatorColor).style("font-size", "10px").style("font-weight", "bold").attr("text-anchor", "middle");

            } else {
                // Vertical Mode
                const yStart = yScale(viewModel.pyTotal);
                const yEnd = yScale(viewModel.pyTotal + totalDelta);
                let bracketX = innerWidth + 20;

                const startPoint = viewModel.dataPoints.find((d, idx) => d.isTotal && !d.isReference && idx === 0);
                const endPoint = viewModel.dataPoints.find((d, idx) => d.isTotal && !d.isReference && idx > 0);

                if (endPoint) {
                    const endBarX = xScale(endPoint.category) + xScale.bandwidth();
                    bracketX = endBarX + 10;
                }

                let xStartOrigin = 0;
                if (startPoint) {
                    xStartOrigin = xScale(startPoint.category) + xScale.bandwidth();
                }

                let xEndOrigin = 0;
                if (endPoint) {
                    xEndOrigin = xScale(endPoint.category) + xScale.bandwidth();
                }

                // Line 1: From Start Column to Bracket
                contentGroup.append("line").attr("x1", xStartOrigin).attr("y1", yStart).attr("x2", bracketX).attr("y2", yStart)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");

                // Line 2: From End Column to Bracket
                contentGroup.append("line").attr("x1", xEndOrigin).attr("y1", yEnd).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", "#aaa").attr("stroke-dasharray", "2,2");

                contentGroup.append("line").attr("x1", bracketX).attr("y1", yStart).attr("x2", bracketX).attr("y2", yEnd)
                    .attr("stroke", indicatorColor).attr("stroke-width", 2);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yStart).attr("r", 3).attr("fill", indicatorColor);
                contentGroup.append("circle").attr("cx", bracketX).attr("cy", yEnd).attr("r", 3).attr("fill", indicatorColor);

                const labelX = bracketX + 5;
                const labelY = (yStart + yEnd) / 2;
                contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(absStr)
                    .attr("fill", indicatorColor).style("font-size", "10px").style("font-weight", "bold").attr("text-anchor", "start");
                contentGroup.append("text").attr("x", labelX).attr("y", labelY + 10).text(pStr)
                    .attr("fill", indicatorColor).style("font-size", "9px").attr("text-anchor", "start");
            }
        }

        const refStrokeColor = "#333";
        const refMarkShape = this.settings.totalSettings.refMarkShape.value.value;
        const showRefMarkOnColumns = this.settings.columnSettings.showRefMarkOnColumns.value;

        // Find End Point with Reference Value (Needed for Variance Indicator)
        const refPoint = viewModel.dataPoints.find(d => d.isTotal && d.referenceValue !== undefined);

        contentGroup.selectAll(".ref-mark").remove();

        // Loop through all data points
        // Determine active scales for indicators
        // Use Secondary Axis for Reference Marks as requested.
        const indicatorXScale = xScale2;
        const indicatorYScale = yScale2;

        // Loop through all data points
        viewModel.dataPoints.forEach(d => {
            if (d.referenceValue !== undefined) {
                // Determine if we should draw
                const shouldDraw = d.isTotal || showRefMarkOnColumns;
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

                    // For Totals, spread across full band or slot? 
                    // Totals are full band usually, but here they are just bars.
                    // Let's use the same slot logic as the bar it marks.
                    // But Totals are usually "wfSlot" (Waterfall Slot).
                    const slot = getSlotInfo(bw, wfSlot);

                    // Center Y and Height depends on slot
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
                            .attr("x1", cx - r).attr("y1", cy - r).attr("x2", cx + r).attr("y2", cy + r);
                        contentGroup.append("line")
                            .classed("ref-mark", true).attr("stroke", refStrokeColor).attr("stroke-width", 2)
                            .attr("x1", cx - r).attr("y1", cy + r).attr("x2", cx + r).attr("y2", cy - r);
                    }
                }
            }
        });

        // Variance Indicator (End vs Reference)
        // Uses Primary Scale (xScale / yScale) because End Column and Total Reference are on Primary Axis.
        if (showVariance && refPoint && refPoint.referenceValue !== undefined) {
            const endVal = refPoint.end;
            const refVal = refPoint.referenceValue;
            const variance = endVal - refVal;

            let variancePct = 0;
            if (refVal !== 0) variancePct = (variance / refVal) * 100;

            const vAbsStr = this.formatNumber(variance, decimalPlaces, useThousandsSeparator, numberScale, thousandsAbbrev, millionsAbbrev);
            const vPStr = "(" + (variancePct >= 0 ? "+" : "") + variancePct.toFixed(percentDecimalPlaces) + "%)";
            const vColor = variance >= 0 ? this.settings.totalSettings.summaryPositiveColor.value.value : this.settings.totalSettings.summaryNegativeColor.value.value;

            // Use Primary Scale
            const indXScale = xScale;
            const indYScale = yScale;

            if (isHorizontal) {
                // Horizontal Mode
                const xStart = indXScale(endVal); // Current (End) Value
                const xEnd = indXScale(refVal);   // Reference Value (Mark)
                const bracketY = -50;

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
                contentGroup.append("text").attr("x", labelX).attr("y", labelY).text(vAbsStr + " " + vPStr)
                    .attr("fill", vColor).style("font-size", "10px").style("font-weight", "bold").attr("text-anchor", "middle");

            } else {
                // Vertical Mode
                let bracketX = 0;

                // Position Variance Indicator to the right of End Column
                const endBarX = indXScale(refPoint.category) + indXScale.bandwidth();
                bracketX = endBarX + 70;

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
                const labelY = (yStart + yEnd) / 2;
                contentGroup.append("text").attr("x", labelX).attr("y", labelY - 2).text(vAbsStr)
                    .attr("fill", vColor).style("font-size", "10px").style("font-weight", "bold").attr("text-anchor", "start");
                contentGroup.append("text").attr("x", labelX).attr("y", labelY + 10).text(vPStr)
                    .attr("fill", vColor).style("font-size", "9px").attr("text-anchor", "start");
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

    private generateChartData(indices: number[], categoryColumn: powerbi.DataViewCategoryColumn, groupColumn: powerbi.DataViewCategoryColumn | null, pyValues: any[], tyValues: any[], budgetValues: any[], tooltipCols: { source: any, values: any[] }[], title: string): WaterfallChartData {
        const increaseColor = this.settings.sentimentColors.increaseColor.value.value;
        const decreaseColor = this.settings.sentimentColors.decreaseColor.value.value;
        const totalColor = this.settings.totalSettings.totalColor.value.value;
        const showStartTotal = this.settings.totalSettings.showStartTotal.value;
        const showEndTotal = this.settings.totalSettings.showEndTotal.value;
        // Default to Standard Variance (PY -> TY) if not specified or set to "none"/"zero"
        // (User cannot opt-out of these in standard Variance "Bridge" mode or math breaks)
        let startTotalColumn = this.settings.totalSettings.startTotalColumn.value.value;
        if (!startTotalColumn || startTotalColumn === "auto" || startTotalColumn === "none" || startTotalColumn === "zero") startTotalColumn = "py";

        let endTotalColumn = this.settings.totalSettings.endTotalColumn.value.value;
        if (!endTotalColumn || endTotalColumn === "auto" || endTotalColumn === "none") endTotalColumn = "ty";
        const referenceColumn = this.settings.totalSettings.referenceColumn.value.value;
        const sortBy = this.settings.sortingSettings.sortBy.value.value;
        const sortDesc = this.settings.sortingSettings.sortDirection.value.value === "desc";

        // Abbrev Replacements
        const pyAbbrev = this.settings.localizationSettings.pyAbbrev.value || "PY";
        const tyAbbrev = this.settings.localizationSettings.tyAbbrev.value || "TY";
        const budgetAbbrev = this.settings.localizationSettings.budgetAbbrev.value || "BUD";

        const replaceAbbrev = (name: string, type: string) => {
            // If name is user-provided constant "Start" or "End", we might want to override it IF the user hasn't customized it?
            // Requirement: "The label of 'Start' and 'End' should then automatically be the abbreviations based on the measures"
            // This suggests we ignore the text input if it matches default or if we just want to force it.
            // But usually visual allows override.
            // However, logic says "automatically be the abbreviations".
            // Let's assume if it is "Start" or "End" default, we replace.
            // Better yet, just use the type to determine label, unless user explicitly typed something else?
            // Re-reading: "The label of 'Start' and 'End' should then automatically be the abbreviations"
            // I'll implement logic: If type is 'py', label is pyAbbrev. If 'ty', label is tyAbbrev.
            if (type === "py") return pyAbbrev;
            if (type === "ty") return tyAbbrev;
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

        // Helper to get value
        const getValue = (source: any[], idx: number) => {
            if (!source || source.length === 0) return 0;
            return Number(source[idx] || 0);
        };

        let rawPoints = [];
        for (const i of indices) {
            const rawPy = getValue(pyValues, i);
            const rawTy = getValue(tyValues, i);
            let startVal = 0;
            if (startTotalColumn === "py") startVal = rawPy;
            else if (startTotalColumn === "ty") startVal = rawTy;
            else if (startTotalColumn === "budget") startVal = getValue(budgetValues, i);

            let endVal = 0;
            if (endTotalColumn === "py") endVal = rawPy;
            else if (endTotalColumn === "ty") endVal = rawTy;
            else if (endTotalColumn === "budget") endVal = getValue(budgetValues, i);
            else endVal = rawTy; // fallback

            let refVal = undefined; // Capture reference value
            if (referenceColumn === "py") refVal = rawPy;
            else if (referenceColumn === "ty") refVal = rawTy;
            else if (referenceColumn === "budget") refVal = getValue(budgetValues, i);
            else {
                // Auto-detect: If user hasn't selected a specific column, but dragged 'Budget', use it.
                if (budgetValues && budgetValues.length > 0) refVal = getValue(budgetValues, i);
            }

            const delta = endVal - startVal;

            const tooltips = tooltipCols.map(tc => {
                const val = tc.values[i];
                const formatted = val != null ? val.toString() : "";
                return { displayName: tc.source.displayName, value: formatted };
            });

            let builder = this.host.createSelectionIdBuilder();
            if (groupColumn) {
                builder = builder.withCategory(groupColumn, i);
            }
            const catSelectionId = builder
                .withCategory(categoryColumn, i)
                .createSelectionId();

            rawPoints.push({
                category: categoryColumn.values[i].toString(),
                py: rawPy,
                ty: rawTy,
                ref: refVal, // Store ref value
                delta,
                originalIndex: i,
                selectionId: catSelectionId,
                tooltipValues: tooltips
            });
        }

        let processedPoints = [...rawPoints];
        if (enableTopN && topNCount > 0 && processedPoints.length > topNCount) {
            const sortedByDelta = [...processedPoints].sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
            const topNItems = sortedByDelta.slice(0, topNCount);
            const otherItems = sortedByDelta.slice(topNCount);
            const othersDelta = otherItems.reduce((sum, p) => sum + p.delta, 0);
            const othersPy = otherItems.reduce((sum, p) => sum + p.py, 0);
            const othersTy = otherItems.reduce((sum, p) => sum + p.ty, 0);
            const othersRef = otherItems.reduce((sum, p) => sum + (p.ref || 0), 0);
            const othersPoint = {
                category: othersLabel, py: othersPy, ty: othersTy, ref: othersRef, delta: othersDelta, originalIndex: -1, selectionId: null
            };
            processedPoints = [...topNItems, othersPoint];
            processedPoints.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
            rawPoints = processedPoints;
        }

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

        // Variance Bridge Mode Logic
        // 1. Start with Total PY.
        // 2. Accumulate Deltas.
        // 3. End with Total TY.

        let pyTotal = rawPoints.reduce((sum, p) => sum + p.py, 0);
        let dataPoints: WaterfallDataPoint[] = [];

        const showPY = this.settings.columnSettings.showPY.value;
        const showTY = this.settings.columnSettings.showTY.value;
        const pyColor = this.settings.columnSettings.pyColumnColor.value.value;
        const tyColor = this.settings.columnSettings.tyColumnColor.value.value;

        const startColumnColor = this.settings.totalSettings.startColumnColor.value.value;
        const endColumnColor = this.settings.totalSettings.endColumnColor.value.value;

        // Add Start Column (Total PY) - Always show in Bridge Mode
        dataPoints.push({
            category: startName,
            delta: pyTotal,
            start: 0,
            end: pyTotal,
            runningTotal: pyTotal,
            py: 0,
            ty: 0,
            color: startColumnColor, // Use Start Column Color setting
            isTotal: true,
            selectionId: null,
            referenceValue: undefined
        });

        let runningTotal = pyTotal;

        rawPoints.forEach(p => {
            // For each category, we step from previous runningTotal
            let start = runningTotal;
            let end = start + p.delta;

            // Update running total
            runningTotal = end;

            dataPoints.push({
                category: p.category,
                delta: p.delta,
                start: start,
                end: end,
                runningTotal: end,
                py: p.py,
                ty: p.ty,
                color: p.delta >= 0 ? increaseColor : decreaseColor,
                isTotal: false,
                selectionId: p.selectionId,
                referenceValue: p.ref,
                // Formula: Start + (Budget - PY)
                // This shows where the bar WOULD have ended if it met the budget.
                // STRICT CHECK: Only calculate if p.ref is valid and NOT zero/null.
                referenceY: (p.ref !== undefined && p.ref !== null && p.py !== undefined)
                    ? start + (p.ref - p.py)
                    : undefined
            });
        });

        // Add End Column (Total TY) - Always show in Bridge Mode (unless strictly disabled, but bridge implies it)
        // We defer to showEndTotal if user wants to hide it, but ignore showTY (which is for the background column).
        if (showEndTotal) {
            let totalRef = undefined;
            // Calculate total ref if explicit column selected OR if auto-detected values exist
            const hasRefValues = rawPoints.some(p => p.ref !== undefined && p.ref !== null);
            if (referenceColumn !== "none" || hasRefValues) {
                totalRef = rawPoints.reduce((sum, p) => sum + (p.ref || 0), 0);
            }

            dataPoints.push({
                category: endName,
                delta: runningTotal,
                start: 0,
                end: runningTotal,
                runningTotal: runningTotal,
                py: 0,
                ty: 0, // Ignored by renderer for waterfall bar
                color: endColumnColor, // Use End Column Color setting
                isTotal: true,
                selectionId: null,
                referenceValue: totalRef,
                referenceY: totalRef // For Totals, reference is absolute
            });
        }




        let minVal = 0;
        let maxVal = 0;
        dataPoints.forEach(d => {
            minVal = Math.min(minVal, d.start, d.end);
            maxVal = Math.max(maxVal, d.start, d.end);
            // Include referenceY (calculated position) in domain
            if (d.referenceY !== undefined) {
                minVal = Math.min(minVal, d.referenceY);
                maxVal = Math.max(maxVal, d.referenceY);
            }
        });
        return { title, dataPoints, maxValue: maxVal, minValue: minVal, pyTotal, startName, endName };


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
}