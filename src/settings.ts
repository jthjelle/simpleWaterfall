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
        value: { value: "rectangle", displayName: "Rectangle" },
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
        displayName: "PY Abbreviation",
        value: "PY",
        placeholder: "PY"
    });

    tyAbbrev = new formattingSettings.TextInput({
        name: "tyAbbrev",
        displayName: "TY Abbreviation",
        value: "TY",
        placeholder: "TY"
    });

    budgetAbbrev = new formattingSettings.TextInput({
        name: "budgetAbbrev",
        displayName: "Budget Abbreviation",
        value: "BUD",
        placeholder: "BUD"
    });

    name: string = "localizationSettings";
    displayName: string = "Localization";
    slices: Array<FormattingSettingsSlice> = [this.pyAbbrev, this.tyAbbrev, this.budgetAbbrev];
}




// Column Settings Card (Unchanged mostly)
class ColumnSettingsCard extends FormattingSettingsCard {
    showPY = new formattingSettings.ToggleSwitch({
        name: "showPY",
        displayName: "Show PY Column",
        value: false
    });

    showTY = new formattingSettings.ToggleSwitch({
        name: "showTY",
        displayName: "Show TY Column",
        value: false
    });

    pyColumnColor = new formattingSettings.ColorPicker({
        name: "pyColumnColor",
        displayName: "PY Column Color",
        value: { value: "#A6C9EC" } // Light Blue
    });

    tyColumnColor = new formattingSettings.ColorPicker({
        name: "tyColumnColor",
        displayName: "TY Column Color",
        value: { value: "#5F5FC4" } // Dark Blue
    });

    flipOverlap = new formattingSettings.ToggleSwitch({
        name: "flipOverlap",
        displayName: "Flip Overlap Order",
        value: false
    });

    showPYLabel = new formattingSettings.ToggleSwitch({
        name: "showPYLabel",
        displayName: "Show PY Label",
        value: false
    });

    showTYLabel = new formattingSettings.ToggleSwitch({
        name: "showTYLabel",
        displayName: "Show TY Label",
        value: false
    });

    showRefMarkOnColumns = new formattingSettings.ToggleSwitch({
        name: "showRefMarkOnColumns",
        displayName: "Show reference marks on columns",
        value: false
    });

    name: string = "columnSettings";
    displayName: string = "Column Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.showPY, this.pyColumnColor, this.showPYLabel,
        this.showTY, this.tyColumnColor, this.showTYLabel,
        this.flipOverlap,
        this.showRefMarkOnColumns
    ];
}

// Sentiment Colors (Restructured ColorSettingsCard)
class SentimentColorsCard extends FormattingSettingsCard {
    increaseColor = new formattingSettings.ColorPicker({
        name: "increaseColor",
        displayName: "Increase Color",
        value: { value: "#6B8E23" } // muted green
    });

    decreaseColor = new formattingSettings.ColorPicker({
        name: "decreaseColor",
        displayName: "Decrease Color",
        value: { value: "#CD5C5C" } // muted red
    });

    name: string = "sentimentColors";
    displayName: string = "Sentiment Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.increaseColor, this.decreaseColor
    ];
}

// Total Settings Card (Expanded)
class TotalSettingsCard extends FormattingSettingsCard {
    showStartTotal = new formattingSettings.ToggleSwitch({
        name: "showStartTotal",
        displayName: "Show Start Total",
        value: true
    });

    startTotalColumn = new formattingSettings.ItemDropdown({
        name: "startTotalColumn",
        displayName: "Start Column Value",
        value: { value: "py", displayName: "PY" },
        items: [
            { value: "zero", displayName: "0 (Zero)" },
            { value: "py", displayName: "Previous Year (PY)" },
            { value: "ty", displayName: "This Year (TY)" },
            { value: "budget", displayName: "Budget" }
        ]
    });

    startColumnColor = new formattingSettings.ColorPicker({
        name: "startColumnColor",
        displayName: "Start Column Color",
        value: { value: "#999999" }
    });

    showEndTotal = new formattingSettings.ToggleSwitch({
        name: "showEndTotal",
        displayName: "Show End Total",
        value: true
    });

    endTotalColumn = new formattingSettings.ItemDropdown({
        name: "endTotalColumn",
        displayName: "End Column Value",
        value: { value: "ty", displayName: "TY" },
        items: [
            { value: "calculated", displayName: "Calculated (Sum of Deltas)" },
            { value: "py", displayName: "Previous Year (PY)" },
            { value: "ty", displayName: "This Year (TY)" },
            { value: "budget", displayName: "Budget" }
        ]
    });

    endColumnColor = new formattingSettings.ColorPicker({
        name: "endColumnColor",
        displayName: "End Column Color",
        value: { value: "#000000" }
    });

    totalColor = new formattingSettings.ColorPicker({
        name: "totalColor",
        displayName: "Default Total Color",
        description: "Fallback color if not start/end specific",
        value: { value: "#808080" }
    });

    referenceColumn = new formattingSettings.ItemDropdown({
        name: "referenceColumn",
        displayName: "Reference Column",
        value: { value: "none", displayName: "None" },
        items: [
            { value: "none", displayName: "None" },
            { value: "py", displayName: "Previous Year (PY)" },
            { value: "ty", displayName: "This Year (TY)" },
            { value: "budget", displayName: "Budget" }
        ]
    });

    refMarkShape = new formattingSettings.ItemDropdown({
        name: "refMarkShape",
        displayName: "Ref Mark Shape",
        value: { value: "line", displayName: "Line" },
        items: [
            { value: "line", displayName: "Line" },
            { value: "dashed", displayName: "Dashed Line" },
            { value: "circle", displayName: "Circle" },
            { value: "cross", displayName: "Cross" }
        ]
    });



    showVariance = new formattingSettings.ToggleSwitch({
        name: "showVariance",
        displayName: "Show Variance",
        value: false
    });

    showSummaryIndicator = new formattingSettings.ToggleSwitch({
        name: "showSummaryIndicator",
        displayName: "Show Summary Indicator",
        value: false
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
        this.showStartTotal, this.startTotalColumn, this.startColumnColor,
        this.showEndTotal, this.endTotalColumn, this.endColumnColor,
        this.totalColor,
        this.referenceColumn, this.refMarkShape,
        this.showVariance,
        this.showSummaryIndicator, this.summaryPositiveColor, this.summaryNegativeColor
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
        value: "Segoe UI, wf_segoe-ui, helvetica, arial, sans-serif"
    });

    // Formatting options from NumFormat
    numberScale = new formattingSettings.ItemDropdown({
        name: "numberScale",
        displayName: "Display Units",
        value: { value: "none", displayName: "None" },
        items: [
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
        placeholder: "K"
    });

    millionsAbbrev = new formattingSettings.TextInput({
        name: "millionsAbbrev",
        displayName: "Millions Abbreviation",
        value: "M",
        placeholder: "M"
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
        value: { value: "#FFFFFF" }
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
        value: true
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

    name: string = "yAxisSettings";
    displayName: string = "Y-Axis";
    slices: Array<FormattingSettingsSlice> = [this.showLabels, this.enableAxisBreak, this.axisBreakPercent, this.invert, this.min, this.max];
}

// X-Axis Settings Card
class XAxisSettingsCard extends FormattingSettingsCard {
    labelOrientation = new formattingSettings.ItemDropdown({
        name: "labelOrientation",
        displayName: "Label Orientation",
        value: { value: "horizontal", displayName: "Horizontal" },
        items: [
            { value: "horizontal", displayName: "Horizontal" },
            { value: "vertical", displayName: "Vertical" },
            { value: "angled", displayName: "Angled (45°)" }
        ]
    });

    labelMaxChars = new formattingSettings.NumUpDown({
        name: "labelMaxChars",
        displayName: "Max Characters",
        value: 0  // 0 means no limit
    });

    name: string = "xAxisSettings";
    displayName: string = "X-Axis";
    slices: Array<FormattingSettingsSlice> = [this.labelOrientation, this.labelMaxChars];
}

// Sorting Settings Card
class SortingSettingsCard extends FormattingSettingsCard {
    sortBy = new formattingSettings.ItemDropdown({
        name: "sortBy",
        displayName: "Sort By",
        value: { value: "category", displayName: "Category (Default)" },
        items: [
            { value: "category", displayName: "Category (Default)" },
            { value: "delta", displayName: "Delta" },
            { value: "ty", displayName: "This Year" },
            { value: "py", displayName: "Previous Year" }
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

    name: string = "sortingSettings";
    displayName: string = "Sorting";
    slices: Array<FormattingSettingsSlice> = [this.sortBy, this.sortDirection];
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
    columnSettings = new ColumnSettingsCard();
    sentimentColors = new SentimentColorsCard();
    totalSettings = new TotalSettingsCard();
    dataLabelSettings = new DataLabelSettingsCard();
    yAxisSettings = new YAxisSettingsCard();
    xAxisSettings = new XAxisSettingsCard();
    sortingSettings = new SortingSettingsCard();
    rankingSettings = new RankingSettingsCard();
    smallMultiplesSettings = new SmallMultiplesSettingsCard();

    cards = [
        this.layoutSettings,
        this.localizationSettings,
        this.columnSettings,
        this.sentimentColors,
        this.totalSettings,
        this.dataLabelSettings,
        this.yAxisSettings,
        this.xAxisSettings,
        this.sortingSettings,
        this.rankingSettings,
        this.smallMultiplesSettings
    ];
}
