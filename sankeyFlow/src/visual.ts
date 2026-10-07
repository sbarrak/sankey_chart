"use strict";

import powerbi from "powerbi-visuals-api";
import { select, Selection } from "d3";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataViewTable = powerbi.DataViewTable;

import { VisualFormattingSettingsModel } from "./settings";
import { render, LinkRow, SankeySettings } from "./render";

export class Visual implements IVisual {
    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private svg: Selection<SVGSVGElement, unknown, null, undefined>;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // state of the last update (needed to redraw after a selection change)
    private links: LinkRow[] = [];
    private linkIds: ISelectionId[][] = [];
    private width = 0;
    private height = 0;
    private format: (v: number) => string = (v) => String(v);

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.svg = select(options.element).append("svg").classed("sankey", true)
            .attr("role", "img").attr("aria-label", "Sankey diagram") as Selection<SVGSVGElement, unknown, null, undefined>;
        this.selectionManager.registerOnSelectCallback(() => this.draw());
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            const dv = options.dataViews && options.dataViews[0];
            this.formattingSettings = this.formattingSettingsService
                .populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
            this.width = Math.max(50, options.viewport.width);
            this.height = Math.max(50, options.viewport.height);
            this.links = [];
            this.linkIds = [];

            if (dv && dv.table) {
                this.readTable(dv.table);
            }
            this.draw();
            this.events.renderingFinished(options);
        } catch (error) {
            this.events.renderingFailed(options, String(error));
        }
    }

    /** Table rows (Source, Target, Value) -> flows. Rows with the same pair are summed. */
    private readTable(table: DataViewTable): void {
        const col = (role: string): number =>
            table.columns.findIndex(c => c.roles && c.roles[role]);
        const si = col("source"), ti = col("target"), vi = col("value");
        if (si < 0 || ti < 0 || vi < 0) { return; }

        const map = new Map<string, LinkRow>();
        let maxVal = 0;
        table.rows.forEach((row, r) => {
            const s = row[si] == null ? "" : String(row[si]).trim();
            const t = row[ti] == null ? "" : String(row[ti]).trim();
            const v = Number(row[vi]);
            if (!s || !t || s === t || !isFinite(v) || v <= 0) { return; }
            const key = s + "\u0000" + t;
            let l = map.get(key);
            if (!l) { l = { source: s, target: t, value: 0, rows: [] }; map.set(key, l); }
            l.value += v;
            l.rows.push(r);
            maxVal = Math.max(maxVal, l.value);
        });
        this.links = Array.from(map.values());
        this.linkIds = this.links.map(l => l.rows.map(r =>
            this.host.createSelectionIdBuilder().withTable(table, r).createSelectionId()));

        const f = valueFormatter.create({ format: table.columns[vi].format, value: maxVal });
        this.format = (v: number) => f.format(v);
    }

    private settings(): SankeySettings {
        const m = this.formattingSettings;
        return {
            palette: String(m.colorsCard.palette.value.value),
            customPalette: m.colorsCard.customPalette.value,
            nodeColors: m.colorsCard.nodeColors.value,
            linkStyle: String(m.colorsCard.linkStyle.value.value),
            linkOpacity: Number(m.colorsCard.linkOpacity.value),
            nodeWidth: Number(m.layoutCard.nodeWidth.value),
            nodePadding: Number(m.layoutCard.nodePadding.value),
            align: String(m.layoutCard.align.value.value),
            showLabels: !!m.labelsCard.show.value,
            showValues: !!m.labelsCard.showValues.value,
            fontSize: Number(m.labelsCard.fontSize.value),
            unit: m.labelsCard.unit.value
        };
    }

    private draw(): void {
        if (!this.formattingSettings) { return; }
        const S = this.settings();
        const unit = S.unit ? " " + S.unit : "";
        const selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];
        let selected: Set<LinkRow> | null = null;
        if (selectedIds.length) {
            selected = new Set<LinkRow>();
            this.links.forEach((l, i) => {
                if (this.linkIds[i].some(id => selectedIds.some(s => s.equals(id)))) { selected!.add(l); }
            });
        }
        render({
            svg: this.svg,
            width: this.width,
            height: this.height,
            links: this.links,
            settings: S,
            formatValue: (v: number) => this.format(v) + unit,
            themeColor: (key: string) => this.host.colorPalette.getColor(key).value,
            selected,
            onLinkClick: (link, ev) => {
                const i = this.links.indexOf(link);
                if (i >= 0) {
                    this.selectionManager.select(this.linkIds[i], ev.ctrlKey || ev.metaKey)
                        .then(() => this.draw());
                }
            },
            onBackgroundClick: () => {
                this.selectionManager.clear().then(() => this.draw());
            }
        });
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
