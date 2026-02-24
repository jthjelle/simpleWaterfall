"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// Layout Settings (was VisualSettingsCard)
class LayoutSettingsCard extends FormattingSettingsCard {
    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "Orientation",
        value: { value: "vertical", displayName: "Vertical" },
        items: [
            { value: "vertical", displayName: "Vertical" },
            { value: "horizontal", displayName: "Horizontal" }
        ]
    });

    barShape = new formattingSettings.ItemDropdown({
        name: "barShape",
        displayName: "Bar Shape",
        value: { value: "rounded", displayName: "Rounded" },
        items: [
            { value: "rectangle", displayName: "Rectangle" },
            { value: "rounded", displayName: "Rounded" },
            { value: "arrow", displayName: "Arrow" }
        ]
    });

    showConnectors = new formattingSettings.ToggleSwitch({
        name: "showConnectors",
        displayName: "Show Connectors",
        value: true
    });

    connectorColor = new formattingSettings.ColorPicker({
        name: "connectorColor",
        displayName: "Connector Color",
        value: { value: "#808080" }
    });

    name: string = "layoutSettings";
    displayName: string = "Layout";
    slices: Array<FormattingSettingsSlice> = [this.orientation, this.barShape, this.showConnectors, this.connectorColor];
}

// Localization Settings (was GeneralSettingsCard)
class LocalizationSettingsCard extends FormattingSettingsCard {
    pyAbbrev = new formattingSettings.TextInput({
        name: "pyAbbrev",
        displayName: "Previous Year Abbreviation",
        value: "PY",
        placeholder: "PY"
    });

    cyAbbrev = new formattingSettings.TextInput({
        name: "cyAbbrev",
        displayName: "Current Year Abbreviation",
        value: "CY",
        placeholder: "CY"
    });

    budgetAbbrev = new formattingSettings.TextInput({
        name: "budgetAbbrev",
        displayName: "Budget Abbreviation",
        value: "BUD",
        placeholder: "BUD"
    });

    name: string = "localizationSettings";
    displayName: string = "Naming conventions";
    slices: Array<FormattingSettingsSlice> = [this.pyAbbrev, this.cyAbbrev, this.budgetAbbrev];
}






// Sentiment Colors (Restructured ColorSettingsCard)
class SentimentColorsCard extends FormattingSettingsCard {
    increaseColor = new formattingSettings.ColorPicker({
        name: "increaseColor",
        displayName: "Increase Color",
        value: { value: "#7ACA00" } // RGB(122, 202, 0)
    });

    decreaseColor = new formattingSettings.ColorPicker({
        name: "decreaseColor",
        displayName: "Decrease Color",
        value: { value: "#FF0000" } // RGB(255, 0, 0)
    });

    subtotalIncreaseColor = new formattingSettings.ColorPicker({
        name: "subtotalIncreaseColor",
        displayName: "Subtotal Increase Color",
        value: { value: "#497900" } // 40% darker than #7ACA00
    });

    subtotalDecreaseColor = new formattingSettings.ColorPicker({
        name: "subtotalDecreaseColor",
        displayName: "Subtotal Decrease Color",
        value: { value: "#990000" } // 40% darker than #FF0000
    });

    name: string = "sentimentColors";
    displayName: string = "Sentiment Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.increaseColor, this.decreaseColor,
        this.subtotalIncreaseColor, this.subtotalDecreaseColor
    ];
}

// Total Settings Card (Expanded)
class TotalSettingsCard extends FormattingSettingsCard {
    startTotalColumn = new formattingSettings.ItemDropdown({
        name: "startTotalColumn",
        displayName: "Start Column Value",
        value: { value: "py", displayName: "PY" },
        items: [
            { value: "py", displayName: "Previous Year (PY)" },
            { value: "cy", displayName: "Current Year (CY)" },
            { value: "budget", displayName: "Budget" }
        ]
    });

    startColumnColor = new formattingSettings.ColorPicker({
        name: "startColumnColor",
        displayName: "Start Column Color",
        value: { value: "#999999" }
    });

    endTotalColumn = new formattingSettings.ItemDropdown({
        name: "endTotalColumn",
        displayName: "End Column Value",
        value: { value: "cy", displayName: "Current Year (CY)" },
        items: [

            { value: "py", displayName: "Previous Year (PY)" },
            { value: "cy", displayName: "Current Year (CY)" },
            { value: "budget", displayName: "Budget" },
        ]
    });

    endColumnColor = new formattingSettings.ColorPicker({
        name: "endColumnColor",
        displayName: "End Column Color",
        value: { value: "#000000" }
    });

    singleMeasureMode = new formattingSettings.ToggleSwitch({
        name: "singleMeasureMode",
        displayName: "Single Measure Mode (First & Last as Totals)",
        value: true
    });

    showRefMark = new formattingSettings.ToggleSwitch({
        name: "showRefMark",
        displayName: "Show Reference Mark",
        value: false
    });



    referenceColumn = new formattingSettings.ItemDropdown({
        name: "referenceColumn",
        displayName: "Reference Column",
        value: { value: "none", displayName: "None" },
        items: [
            { value: "none", displayName: "None" },
            { value: "py", displayName: "Previous Year (PY)" },
            { value: "cy", displayName: "Current Year (CY)" },
            { value: "budget", displayName: "Budget" }
        ]
    });

    refMarkShape = new formattingSettings.ItemDropdown({
        name: "refMarkShape",
        displayName: "Ref Mark Shape",
        value: { value: "dashed", displayName: "Dashed Line" },
        items: [
            { value: "line", displayName: "Line" },
            { value: "dashed", displayName: "Dashed Line" },
            { value: "circle", displayName: "Circle" },
            { value: "cross", displayName: "Cross" }
        ]
    });

    showRefMarkLabel = new formattingSettings.ToggleSwitch({
        name: "showRefMarkLabel",
        displayName: "Show Reference Label",
        value: false
    });

    refMarkColor = new formattingSettings.ColorPicker({
        name: "refMarkColor",
        displayName: "Ref Mark Color",
        value: { value: "#D9B300" }
    });

    showSubtotals = new formattingSettings.ToggleSwitch({
        name: "showSubtotals",
        displayName: "Show Category Subtotals",
        value: true
    });





    showVariance = new formattingSettings.ToggleSwitch({
        name: "showVariance",
        displayName: "Show Reference Variance to End Column",
        value: false
    });

    showVarianceLabel = new formattingSettings.ToggleSwitch({
        name: "showVarianceLabel",
        displayName: "Show Label on Reference Variance",
        value: false
    });

    varianceLabelFormat = new formattingSettings.ItemDropdown({
        name: "varianceLabelFormat",
        displayName: "Reference Variance Label Format",
        value: { value: "both", displayName: "Both" },
        items: [
            { value: "both", displayName: "Both" },
            { value: "data", displayName: "Data" },
            { value: "percent", displayName: "Percent" }
        ]
    });

    showSummaryIndicator = new formattingSettings.ToggleSwitch({
        name: "showSummaryIndicator",
        displayName: "Show End Column Variance to Start Column",
        value: true
    });

    showSummaryIndicatorLabel = new formattingSettings.ToggleSwitch({
        name: "showSummaryIndicatorLabel",
        displayName: "Show Label on End Column Reference",
        value: false
    });

    summaryLabelFormat = new formattingSettings.ItemDropdown({
        name: "summaryLabelFormat",
        displayName: "End Column Variance Label Format",
        value: { value: "both", displayName: "Both" },
        items: [
            { value: "both", displayName: "Both" },
            { value: "data", displayName: "Data" },
            { value: "percent", displayName: "Percent" }
        ]
    });

    stackVarianceLabels = new formattingSettings.ToggleSwitch({
        name: "stackVarianceLabels",
        displayName: "Stack Variance Labels",
        value: true
    });

    summaryPositiveColor = new formattingSettings.ColorPicker({
        name: "summaryPositiveColor",
        displayName: "Summary Pos Color",
        value: { value: "#6B8E23" }
    });

    summaryNegativeColor = new formattingSettings.ColorPicker({
        name: "summaryNegativeColor",
        displayName: "Summary Neg Color",
        value: { value: "#CD5C5C" }
    });

    name: string = "totalSettings";
    displayName: string = "Totals";
    slices: Array<FormattingSettingsSlice> = [
        this.startTotalColumn,
        this.startColumnColor,
        this.endTotalColumn,
        this.endColumnColor,
        this.singleMeasureMode,
        this.showRefMark,
        this.showRefMarkLabel,
        this.referenceColumn, // Intentionally kept for now
        this.refMarkShape,
        this.refMarkColor,
        this.showVariance,
        this.showVarianceLabel,
        this.varianceLabelFormat,
        this.showSummaryIndicator,
        this.showSummaryIndicatorLabel,
        this.summaryLabelFormat,
        this.stackVarianceLabels,
        this.summaryPositiveColor,
        this.summaryNegativeColor,
        this.showSubtotals
    ];
}

// Data Labels Card (Consolidated)
class DataLabelSettingsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show Labels",
        value: true
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Color",
        value: { value: "#000000" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size",
        value: 11
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font Family",
        value: "Segoe UI, wf_segoe-ui_normal, helvetica, arial, sans-serif"
    });

    // Formatting options from NumFormat
    numberScale = new formattingSettings.ItemDropdown({
        name: "numberScale",
        displayName: "Display Units",
        value: { value: "auto", displayName: "Auto" },
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "none", displayName: "None" },
            { value: "thousands", displayName: "Thousands" },
            { value: "millions", displayName: "Millions" }
        ]
    });

    decimalPlaces = new formattingSettings.NumUpDown({
        name: "decimalPlaces",
        displayName: "Decimal Places",
        value: 0
    });

    useThousandsSeparator = new formattingSettings.ToggleSwitch({
        name: "useThousandsSeparator",
        displayName: "Thousands Separator",
        value: true
    });

    thousandsAbbrev = new formattingSettings.TextInput({
        name: "thousandsAbbrev",
        displayName: "Thousands Abbreviation",
        value: "K",
        placeholder: ""
    });

    millionsAbbrev = new formattingSettings.TextInput({
        name: "millionsAbbrev",
        displayName: "Millions Abbreviation",
        value: "M",
        placeholder: ""
    });

    // Delta Options
    deltaDisplayMode = new formattingSettings.ItemDropdown({
        name: "deltaDisplayMode",
        displayName: "Delta Display Mode",
        value: { value: "absolute", displayName: "Absolute" },
        items: [
            { value: "absolute", displayName: "Absolute" },
            { value: "percent", displayName: "Percent" },
            { value: "both", displayName: "Both" }
        ]
    });

    percentDecimalPlaces = new formattingSettings.NumUpDown({
        name: "percentDecimalPlaces",
        displayName: "Percent Decimal Places",
        value: 1
    });

    // Running Total
    showRunningTotal = new formattingSettings.ToggleSwitch({
        name: "showRunningTotal",
        displayName: "Show Running Total",
        value: false
    });

    // Background (Badge)
    showBackground = new formattingSettings.ToggleSwitch({
        name: "showBackground",
        displayName: "Show Background",
        value: false
    });

    badgeShape = new formattingSettings.ItemDropdown({
        name: "badgeShape",
        displayName: "Badge Shape",
        value: { value: "rounded", displayName: "Rounded" },
        items: [
            { value: "rounded", displayName: "Rounded" },
            { value: "rectangle", displayName: "Rectangle" },
            { value: "pill", displayName: "Pill" }
        ]
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background Color",
        value: { value: "#E6E6E6" }
    });

    backgroundTransparency = new formattingSettings.Slider({
        name: "backgroundTransparency",
        displayName: "Transparency",
        value: 0
    });

    name: string = "dataLabelSettings";
    displayName: string = "Data Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.show, this.color, this.fontSize, this.fontFamily,
        this.deltaDisplayMode,
        this.numberScale, this.decimalPlaces, this.percentDecimalPlaces, this.useThousandsSeparator,
        this.thousandsAbbrev, this.millionsAbbrev,
        this.showRunningTotal,
        this.showBackground, this.badgeShape, this.backgroundColor, this.backgroundTransparency
    ];
}


// Y-Axis Settings Card
class YAxisSettingsCard extends FormattingSettingsCard {
    showLabels = new formattingSettings.ToggleSwitch({
        name: "showLabels",
        displayName: "Show Y-Axis Labels",
        value: false
    });

    enableAxisBreak = new formattingSettings.ToggleSwitch({
        name: "enableAxisBreak",
        displayName: "Enable Axis Break",
        value: false
    });

    axisBreakPercent = new formattingSettings.NumUpDown({
        name: "axisBreakPercent",
        displayName: "Axis Break (%)",
        value: 50
    });

    invert = new formattingSettings.ToggleSwitch({
        name: "invert",
        displayName: "Invert Axis",
        value: false
    });

    min = new formattingSettings.NumUpDown({
        name: "min",
        displayName: "Minimum",
        value: null
    });

    max = new formattingSettings.NumUpDown({
        name: "max",
        displayName: "Maximum",
        value: null
    });

    showGrid = new formattingSettings.ToggleSwitch({
        name: "showGrid",
        displayName: "Show Grid",
        value: false
    });

    gridLineStyle = new formattingSettings.ItemDropdown({
        name: "gridLineStyle",
        displayName: "Grid Line Style",
        value: { value: "dashed", displayName: "Dashed" },
        items: [
            { value: "solid", displayName: "Solid" },
            { value: "dashed", displayName: "Dashed" },
            { value: "dotted", displayName: "Dotted" }
        ]
    });

    gridLineColor = new formattingSettings.ColorPicker({
        name: "gridLineColor",
        displayName: "Grid Color",
        value: { value: "#d3d3d3" }
    });

    name: string = "yAxisSettings";
    displayName: string = "Y-Axis";
    slices: Array<FormattingSettingsSlice> = [this.showLabels, this.enableAxisBreak, this.axisBreakPercent, this.invert, this.min, this.max, this.showGrid, this.gridLineStyle, this.gridLineColor];
}

// X-Axis Settings Card
class XAxisSettingsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show Axis Line", // Updated display name for clarity? Or keep "Show Axis" but behavior implies line only. User said "hide the axis line".
        value: true
    });

    labelOrientation = new formattingSettings.ItemDropdown({
        name: "labelOrientation",
        displayName: "Label Orientation",
        value: { value: "angled", displayName: "Angled (45°)" },
        items: [
            { value: "horizontal", displayName: "Horizontal" },
            { value: "vertical", displayName: "Vertical" },
            { value: "angled", displayName: "Angled (45°)" }
        ]
    });

    labelMaxChars = new formattingSettings.NumUpDown({
        name: "labelMaxChars",
        displayName: "Max Characters",
        value: 10
    });


    showGrid = new formattingSettings.ToggleSwitch({
        name: "showGrid",
        displayName: "Show Grid",
        value: false
    });

    gridLineStyle = new formattingSettings.ItemDropdown({
        name: "gridLineStyle",
        displayName: "Grid Line Style",
        value: { value: "dashed", displayName: "Dashed" },
        items: [
            { value: "solid", displayName: "Solid" },
            { value: "dashed", displayName: "Dashed" },
            { value: "dotted", displayName: "Dotted" }
        ]
    });

    gridLineColor = new formattingSettings.ColorPicker({
        name: "gridLineColor",
        displayName: "Grid Color",
        value: { value: "#d3d3d3" }
    });

    name: string = "xAxisSettings";
    displayName: string = "X-Axis";
    slices: Array<FormattingSettingsSlice> = [this.show, this.labelOrientation, this.labelMaxChars, this.showGrid, this.gridLineStyle, this.gridLineColor];
}

// Sorting Settings Card


// Sorting Settings Card
class SortingSettingsCard extends FormattingSettingsCard {
    sortBy = new formattingSettings.ItemDropdown({
        name: "sortBy",
        displayName: "Sort By",
        value: { value: "default", displayName: "Default (Data Source)" },
        items: [
            { value: "default", displayName: "Default (Data Source)" },
            { value: "category", displayName: "Category" },
            { value: "delta", displayName: "Delta" },
            { value: "cy", displayName: "Current Year" },
            { value: "py", displayName: "Previous Year" }
        ]
    });

    sortDirection = new formattingSettings.ItemDropdown({
        name: "sortDirection",
        displayName: "Direction",
        value: { value: "desc", displayName: "Descending" },
        items: [
            { value: "asc", displayName: "Ascending" },
            { value: "desc", displayName: "Descending" }
        ]
    });

    name: string = "sortingSettings";
    displayName: string = "Sorting";
    slices: Array<FormattingSettingsSlice> = [this.sortBy, this.sortDirection];
}

// Tooltip Settings Card
export class TooltipSettingsCard extends FormattingSettingsCard {
    tooltipNumberScale = new formattingSettings.ItemDropdown({
        name: "tooltipNumberScale",
        displayName: "Display Units",
        value: { value: "none", displayName: "None" },
        items: [
            { value: "none", displayName: "None" },
            { value: "thousands", displayName: "Thousands" },
            { value: "millions", displayName: "Millions" }
        ]
    });

    tooltipDecimalPlaces = new formattingSettings.NumUpDown({
        name: "tooltipDecimalPlaces",
        displayName: "Decimal Places",
        value: 0
    });

    useThousandsSeparator = new formattingSettings.ToggleSwitch({
        name: "useThousandsSeparator",
        displayName: "Thousands Separator",
        value: true
    });

    name: string = "tooltipSettings";
    displayName: string = "Tooltip Settings";
    slices: Array<FormattingSettingsSlice> = [this.tooltipNumberScale, this.tooltipDecimalPlaces, this.useThousandsSeparator];
}

// Ranking Settings Card
class RankingSettingsCard extends FormattingSettingsCard {
    enable = new formattingSettings.ToggleSwitch({
        name: "enable",
        displayName: "Enable Top N",
        value: false
    });

    count = new formattingSettings.NumUpDown({
        name: "count",
        displayName: "Top N Count",
        value: 5
    });

    othersLabel = new formattingSettings.TextInput({
        name: "othersLabel",
        displayName: "Others Label",
        value: "Others",
        placeholder: "Others"
    });

    name: string = "rankingSettings";
    displayName: string = "Top N";
    slices: Array<FormattingSettingsSlice> = [this.enable, this.count, this.othersLabel];
}

// Small Multiples Settings Card
class SmallMultiplesSettingsCard extends FormattingSettingsCard {
    layoutMode = new formattingSettings.ItemDropdown({
        name: "layoutMode",
        displayName: "Layout Mode",
        value: { value: "auto", displayName: "Auto" },
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "fixed", displayName: "Fixed" }
        ]
    });

    rows = new formattingSettings.NumUpDown({
        name: "rows",
        displayName: "Rows",
        value: 2
    });

    columns = new formattingSettings.NumUpDown({
        name: "columns",
        displayName: "Columns",
        value: 2
    });

    name: string = "smallMultiplesSettings";
    displayName: string = "Small Multiples";
    uniformYAxis = new formattingSettings.ToggleSwitch({
        name: "uniformYAxis",
        displayName: "Uniform Y-Axis",
        value: false
    });

    sortBy = new formattingSettings.ItemDropdown({
        name: "sortBy",
        displayName: "Sort By",
        value: { value: "name", displayName: "Name" },
        items: [
            { value: "name", displayName: "Name" },
            { value: "start", displayName: "Start Value" },
            { value: "end", displayName: "End Value" }
        ]
    });

    sortDirection = new formattingSettings.ItemDropdown({
        name: "sortDirection",
        displayName: "Direction",
        value: { value: "asc", displayName: "Ascending" },
        items: [
            { value: "asc", displayName: "Ascending" },
            { value: "desc", displayName: "Descending" }
        ]
    });

    slices: Array<FormattingSettingsSlice> = [this.layoutMode, this.uniformYAxis, this.rows, this.columns, this.sortBy, this.sortDirection];
}



/**
* Visual Formatting Settings Model Class
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    layoutSettings = new LayoutSettingsCard();
    localizationSettings = new LocalizationSettingsCard();
    sentimentColors = new SentimentColorsCard();
    totalSettings = new TotalSettingsCard();
    dataLabelSettings = new DataLabelSettingsCard();
    yAxisSettings = new YAxisSettingsCard();
    xAxisSettings = new XAxisSettingsCard();
    sortingSettings = new SortingSettingsCard();

    tooltipSettings = new TooltipSettingsCard();
    rankingSettings = new RankingSettingsCard();
    smallMultiplesSettings = new SmallMultiplesSettingsCard();

    cards = [
        this.layoutSettings,
        this.totalSettings,
        this.dataLabelSettings,
        this.yAxisSettings,
        this.xAxisSettings,
        this.sentimentColors,
        this.localizationSettings,
        this.tooltipSettings,
        this.sortingSettings,
        this.rankingSettings,
        this.smallMultiplesSettings
    ];
}
