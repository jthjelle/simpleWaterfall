"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// Visual Settings Card
class VisualSettingsCard extends FormattingSettingsCard {
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

    name: string = "visualSettings";
    displayName: string = "Visual Settings";
    slices: Array<FormattingSettingsSlice> = [this.orientation, this.showConnectors, this.connectorColor, this.barShape];
}

// Column Settings Card
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



    name: string = "columnSettings";
    displayName: string = "Column Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.showPY, this.pyColumnColor, this.showPYLabel,
        this.showTY, this.tyColumnColor, this.showTYLabel,
        this.flipOverlap
    ];
}

// General Settings Card (Abbreviations)
class GeneralSettingsCard extends FormattingSettingsCard {
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

    name: string = "generalSettings";
    displayName: string = "General";
    slices: Array<FormattingSettingsSlice> = [this.pyAbbrev, this.tyAbbrev, this.budgetAbbrev];
}

// Color Settings Card
class ColorSettingsCard extends FormattingSettingsCard {
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

    totalColor = new formattingSettings.ColorPicker({
        name: "totalColor",
        displayName: "Total Color",
        value: { value: "#808080" } // neutral gray
    });

    startColumnColor = new formattingSettings.ColorPicker({
        name: "startColumnColor",
        displayName: "Start Column Color",
        value: { value: "#333333" } // 20% lighter (approx)
    });

    endColumnColor = new formattingSettings.ColorPicker({
        name: "endColumnColor",
        displayName: "End Column Color",
        value: { value: "#000000" } // Black
    });

    name: string = "colorSettings";
    displayName: string = "Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.increaseColor, this.decreaseColor, this.totalColor,
        this.startColumnColor, this.endColumnColor
    ];
}

// Label Settings Card
class LabelSettingsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show Labels",
        value: true
    });

    showRunningTotal = new formattingSettings.ToggleSwitch({
        name: "showRunningTotal",
        displayName: "Show Running Total",
        value: false
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size",
        value: 11
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "Color",
        value: { value: "#000000" }
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font Family",
        value: "Segoe UI, wf_segoe-ui, helvetica, arial, sans-serif"
    });

    showBackground = new formattingSettings.ToggleSwitch({
        name: "showBackground",
        displayName: "Show Background",
        value: false
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

    name: string = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.show,
        this.showRunningTotal,
        this.fontFamily,
        this.fontSize,
        this.color,
        this.showBackground,
        this.backgroundColor,
        this.backgroundTransparency
    ];
}

class DeltaFormatSettingsCard extends FormattingSettingsCard {
    deltaDisplayMode = new formattingSettings.ItemDropdown({
        name: "deltaDisplayMode",
        displayName: "Display Mode",
        value: { value: "absolute", displayName: "Absolute" },
        items: [
            { value: "absolute", displayName: "Absolute" },
            { value: "percent", displayName: "Percent" },
            { value: "both", displayName: "Both" }
        ]
    });

    name: string = "deltaFormatSettings";
    displayName: string = "Delta Format";
    slices: Array<FormattingSettingsSlice> = [this.deltaDisplayMode];
}

class NumberFormatSettingsCard extends FormattingSettingsCard {
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

    decimalPlaces = new formattingSettings.NumUpDown({
        name: "decimalPlaces",
        displayName: "Decimal Places",
        value: 0
    });

    percentDecimalPlaces = new formattingSettings.NumUpDown({
        name: "percentDecimalPlaces",
        displayName: "Percent Decimal Places",
        value: 1
    });

    useThousandsSeparator = new formattingSettings.ToggleSwitch({
        name: "useThousandsSeparator",
        displayName: "Thousands Separator",
        value: true
    });

    name: string = "numberFormatSettings";
    displayName: string = "Number Format";
    slices: Array<FormattingSettingsSlice> = [
        this.numberScale,
        this.decimalPlaces,
        this.percentDecimalPlaces,
        this.useThousandsSeparator,
        this.thousandsAbbrev,
        this.millionsAbbrev
    ];
}

// Total Settings Card
class TotalSettingsCard extends FormattingSettingsCard {
    showStartTotal = new formattingSettings.ToggleSwitch({
        name: "showStartTotal",
        displayName: "Show Start Total",
        value: true
    });

    showEndTotal = new formattingSettings.ToggleSwitch({
        name: "showEndTotal",
        displayName: "Show End Total",
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

    showRefMarkOnColumns = new formattingSettings.ToggleSwitch({
        name: "showRefMarkOnColumns",
        displayName: "Show Mark on Columns",
        value: false
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
        this.showStartTotal,
        this.showEndTotal,
        this.startTotalColumn,
        this.endTotalColumn,
        this.referenceColumn,
        this.refMarkShape,
        this.showRefMarkOnColumns,
        this.showVariance,
        this.showSummaryIndicator,
        this.summaryPositiveColor,
        this.summaryNegativeColor
    ];
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
    generalSettings = new GeneralSettingsCard();
    visualSettings = new VisualSettingsCard();
    columnSettings = new ColumnSettingsCard();
    colorSettings = new ColorSettingsCard();
    labelSettings = new LabelSettingsCard();
    deltaFormatSettings = new DeltaFormatSettingsCard();
    numberFormatSettings = new NumberFormatSettingsCard();
    totalSettings = new TotalSettingsCard();
    yAxisSettings = new YAxisSettingsCard();
    xAxisSettings = new XAxisSettingsCard();
    sortingSettings = new SortingSettingsCard();
    rankingSettings = new RankingSettingsCard();
    smallMultiplesSettings = new SmallMultiplesSettingsCard();

    cards = [
        this.generalSettings,
        this.visualSettings,
        this.columnSettings,
        this.colorSettings,
        this.labelSettings,
        this.deltaFormatSettings,
        this.numberFormatSettings,
        this.totalSettings,
        this.yAxisSettings,
        this.xAxisSettings,
        this.sortingSettings,
        this.rankingSettings,
        this.smallMultiplesSettings
    ];
}
