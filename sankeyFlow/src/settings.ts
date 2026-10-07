"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

const paletteItems = [
    { value: "corporate", displayName: "Corporate" },
    { value: "vibrant", displayName: "Vibrant" },
    { value: "pastel", displayName: "Pastel" },
    { value: "earth", displayName: "Earth" },
    { value: "ocean", displayName: "Ocean" },
    { value: "theme", displayName: "Report theme" },
    { value: "custom", displayName: "Custom palette" }
];
const linkItems = [
    { value: "gradient", displayName: "Gradient" },
    { value: "solid", displayName: "Solid (source color)" },
    { value: "gray", displayName: "Gray" }
];
const alignItems = [
    { value: "justify", displayName: "Justify" },
    { value: "left", displayName: "Left" },
    { value: "right", displayName: "Right" },
    { value: "center", displayName: "Center" }
];

class ColorsCard extends FormattingSettingsCard {
    palette = new formattingSettings.ItemDropdown({
        name: "palette", displayName: "Palette", items: paletteItems, value: paletteItems[0]
    });
    customPalette = new formattingSettings.TextInput({
        name: "customPalette", displayName: "Custom palette (hex list)",
        description: "Used when Palette = Custom palette. Example: #E4002B, #1F4E79, #F2A900",
        placeholder: "#E4002B, #1F4E79, ...", value: ""
    });
    nodeColors = new formattingSettings.TextInput({
        name: "nodeColors", displayName: "Node colors",
        description: "Override single nodes. Example: Salary=#E4002B; Rent=#1F4E79",
        placeholder: "Name=#hex; Name=#hex", value: ""
    });
    linkStyle = new formattingSettings.ItemDropdown({
        name: "linkStyle", displayName: "Flow style", items: linkItems, value: linkItems[0]
    });
    linkOpacity = new formattingSettings.NumUpDown({
        name: "linkOpacity", displayName: "Flow opacity (%)", value: 50
    });
    name: string = "colors";
    displayName: string = "Colors";
    slices: Array<FormattingSettingsSlice> = [this.palette, this.customPalette, this.nodeColors, this.linkStyle, this.linkOpacity];
}

class LayoutCard extends FormattingSettingsCard {
    nodeWidth = new formattingSettings.NumUpDown({ name: "nodeWidth", displayName: "Node width", value: 14 });
    nodePadding = new formattingSettings.NumUpDown({ name: "nodePadding", displayName: "Node spacing", value: 12 });
    align = new formattingSettings.ItemDropdown({
        name: "align", displayName: "Alignment", items: alignItems, value: alignItems[0]
    });
    name: string = "layout";
    displayName: string = "Layout";
    slices: Array<FormattingSettingsSlice> = [this.nodeWidth, this.nodePadding, this.align];
}

class LabelsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Show labels", value: true });
    showValues = new formattingSettings.ToggleSwitch({ name: "showValues", displayName: "Show values", value: true });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Text size", value: 11 });
    unit = new formattingSettings.TextInput({
        name: "unit", displayName: "Unit", description: "Text shown after values, e.g. EUR, kg",
        placeholder: "e.g. EUR", value: ""
    });
    topLevelSlice = this.show;
    name: string = "labels";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [this.showValues, this.fontSize, this.unit];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    colorsCard = new ColorsCard();
    layoutCard = new LayoutCard();
    labelsCard = new LabelsCard();
    cards = [this.colorsCard, this.layoutCard, this.labelsCard];
}
