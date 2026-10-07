"use strict";

import { Selection } from "d3";
import { sankey, sankeyLinkHorizontal, sankeyJustify, sankeyLeft, sankeyRight, sankeyCenter } from "d3-sankey";

/** One flow (rows with the same Source/Target are already summed). */
export interface LinkRow {
    source: string;
    target: string;
    value: number;
    rows: number[];          // table row indexes that make up this flow (for selection)
}

export interface SankeySettings {
    palette: string;
    customPalette: string;
    nodeColors: string;
    linkStyle: string;
    linkOpacity: number;
    nodeWidth: number;
    nodePadding: number;
    align: string;
    showLabels: boolean;
    showValues: boolean;
    fontSize: number;
    unit: string;
}

export interface RenderArgs {
    svg: Selection<SVGSVGElement, unknown, null, undefined>;
    width: number;
    height: number;
    links: LinkRow[];
    settings: SankeySettings;
    formatValue: (v: number) => string;
    themeColor?: (key: string) => string;
    selected: Set<LinkRow> | null;
    onLinkClick?: (link: LinkRow, ev: MouseEvent) => void;
    onBackgroundClick?: () => void;
}

export const PALETTES: Record<string, string[]> = {
    corporate: ["#1F4E79", "#2E75B6", "#0096A0", "#70AD47", "#ED7D31", "#FFC000", "#5B9BD5", "#A5A5A5", "#44546A", "#C0504D"],
    vibrant: ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC"],
    pastel: ["#8CAABE", "#F096A0", "#96C8EB", "#F5C896", "#96E1AF", "#CDA0D7", "#EBE18C", "#AFCD8C", "#E6A5A5", "#BEBEB4"],
    earth: ["#8C5632", "#BF9000", "#70803C", "#A0522D", "#5E7279", "#CD853F", "#556B2F", "#8B7355", "#B26646", "#787864"],
    ocean: ["#005F73", "#0A9396", "#6EBEAA", "#D2BE78", "#EE9B00", "#CA6702", "#BB3E03", "#AE2012", "#9B2226", "#003C50"]
};

export function parseHexList(s: string): string[] {
    const out: string[] = [];
    if (!s) { return out; }
    for (const m of s.matchAll(/(?<![0-9a-fA-F])#?([0-9a-fA-F]{6})(?![0-9a-fA-F])/g)) {
        out.push("#" + m[1].toUpperCase());
    }
    return out;
}

/** "Name=#hex; Name2=#hex" -> map (lower-case name -> #HEX) */
export function parseNodeColors(s: string): Map<string, string> {
    const map = new Map<string, string>();
    if (!s) { return map; }
    for (const part of s.split(/[;\n]/)) {
        const p = part.lastIndexOf("=");
        if (p < 1) { continue; }
        const hex = parseHexList(part.substring(p + 1))[0];
        if (hex) { map.set(part.substring(0, p).trim().toLowerCase(), hex); }
    }
    return map;
}

const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v));

function message(a: RenderArgs, text: string): void {
    a.svg.append("text")
        .attr("x", a.width / 2).attr("y", a.height / 2)
        .attr("text-anchor", "middle").attr("fill", "#6B7280")
        .attr("font-size", 13).text(text);
}

let uid = 0;

/** Draws the diagram. Returns an error text if nothing could be drawn. */
export function render(a: RenderArgs): string | null {
    const S = a.settings;
    const W = a.width, H = a.height;
    a.svg.attr("width", W).attr("height", H).attr("viewBox", `0 0 ${W} ${H}`);
    a.svg.selectAll("*").remove();

    if (!a.links.length) {
        message(a, "Add Source, Target and Value fields (values must be positive).");
        return "no data";
    }

    // ---- nodes / links ---------------------------------------------------------------
    const nameIdx = new Map<string, number>();
    const names: string[] = [];
    const idOf = (n: string): number => {
        let i = nameIdx.get(n);
        if (i === undefined) { i = names.length; nameIdx.set(n, i); names.push(n); }
        return i;
    };
    const raw = a.links.map(l => ({ source: idOf(l.source), target: idOf(l.target), value: l.value, ref: l }));
    const hasIn: boolean[] = names.map(() => false);
    const hasOut: boolean[] = names.map(() => false);
    const inSum: number[] = names.map(() => 0);
    const outSum: number[] = names.map(() => 0);
    raw.forEach(l => {
        hasOut[l.source] = true; hasIn[l.target] = true;
        outSum[l.source] += l.value; inSum[l.target] += l.value;
    });
    const nodeTotal = (i: number): number => Math.max(inSum[i], outSum[i]);

    // ---- margins for labels on the first / last column ---------------------------------------
    const fs = clamp(S.fontSize, 6, 40);
    let mL = 12, mR = 12;
    if (S.showLabels) {
        names.forEach((n, i) => {
            const txt = n + (S.showValues ? "  " + a.formatValue(nodeTotal(i)) : "");
            const w = txt.length * fs * 0.58 + 12;
            if (!hasIn[i]) { mL = Math.max(mL, w); } else if (!hasOut[i]) { mR = Math.max(mR, w); }
        });
        mL = Math.min(mL, W * 0.3);
        mR = Math.min(mR, W * 0.3);
    }

    // ---- layout ---------------------------------------------------------------------------
    const alignFns: Record<string, unknown> = {
        justify: sankeyJustify, left: sankeyLeft, right: sankeyRight, center: sankeyCenter
    };
    const gen: any = sankey()
        .nodeWidth(clamp(S.nodeWidth, 2, 100))
        .nodePadding(clamp(S.nodePadding, 0, 100))
        .nodeAlign((alignFns[S.align] || sankeyJustify) as never)
        .extent([[mL, 8], [Math.max(mL + 60, W - mR), Math.max(60, H - 8)]]);
    let graph: any;
    try {
        graph = gen({
            nodes: names.map(name => ({ name })),
            links: raw.map(l => ({ source: l.source, target: l.target, value: l.value, ref: l.ref }))
        });
    } catch (e) {
        message(a, "Circular flows (A > B > A) are not supported. Check your data.");
        return "circular";
    }

    // ---- colours -----------------------------------------------------------------------------
    const overrides = parseNodeColors(S.nodeColors);
    let pal: string[] = PALETTES[S.palette] || PALETTES.corporate;
    if (S.palette === "custom") {
        const custom = parseHexList(S.customPalette);
        pal = custom.length ? custom : PALETTES.corporate;
    }
    const colorOf = (n: any): string => {
        const o = overrides.get(String(n.name).toLowerCase());
        if (o) { return o; }
        if (S.palette === "theme" && a.themeColor) { return a.themeColor(String(n.name)); }
        return pal[n.index % pal.length];
    };

    const opacity = clamp(S.linkOpacity, 5, 100) / 100;
    const id = "sk" + (uid++);
    const root = a.svg.append("g");
    a.svg.on("click", () => { if (a.onBackgroundClick) { a.onBackgroundClick(); } });

    // ---- flows --------------------------------------------------------------------------------
    if (S.linkStyle === "gradient") {
        const defs = a.svg.append("defs");
        graph.links.forEach((l: any, i: number) => {
            const g = defs.append("linearGradient").attr("id", `${id}-${i}`)
                .attr("gradientUnits", "userSpaceOnUse").attr("x1", l.source.x1).attr("x2", l.target.x0);
            g.append("stop").attr("offset", "0%").attr("stop-color", colorOf(l.source));
            g.append("stop").attr("offset", "100%").attr("stop-color", colorOf(l.target));
        });
    }
    const hasSel = !!a.selected && a.selected.size > 0;
    root.append("g").attr("fill", "none")
        .selectAll("path").data(graph.links as unknown[]).join("path")
        .attr("class", "link")
        .attr("d", sankeyLinkHorizontal() as never)
        .attr("stroke", (l: any, i: number) =>
            S.linkStyle === "gradient" ? `url(#${id}-${i})` : S.linkStyle === "gray" ? "#B8BEC7" : colorOf(l.source))
        .attr("stroke-width", (l: any) => Math.max(1, l.width))
        .attr("stroke-opacity", (l: any) => hasSel && !a.selected!.has(l.ref) ? opacity * 0.25 : opacity)
        .on("click", (ev: MouseEvent, l: any) => {
            ev.stopPropagation();
            if (a.onLinkClick) { a.onLinkClick(l.ref, ev); }
        })
        .append("title")
        .text((l: any) => {
            const pct = l.source.value > 0 ? ` (${(100 * l.value / Math.max(outSum[l.source.index], 1e-12)).toFixed(1)}% of ${l.source.name})` : "";
            return `${l.source.name} → ${l.target.name}\n${a.formatValue(l.value)}${pct}`;
        });

    // ---- nodes ----------------------------------------------------------------------------------
    const nodeG = root.append("g");
    nodeG.selectAll("rect").data(graph.nodes as unknown[]).join("rect")
        .attr("class", "node")
        .attr("x", (n: any) => n.x0).attr("y", (n: any) => n.y0)
        .attr("width", (n: any) => n.x1 - n.x0).attr("height", (n: any) => Math.max(1, n.y1 - n.y0))
        .attr("fill", (n: any) => colorOf(n))
        .append("title")
        .text((n: any) => `${n.name}: ${a.formatValue(nodeTotal(n.index))}`);

    // ---- labels ------------------------------------------------------------------------------------
    if (S.showLabels) {
        const lab = root.append("g").attr("font-family", "Segoe UI, sans-serif").attr("font-size", fs)
            .attr("pointer-events", "none");
        graph.nodes.forEach((n: any) => {
            const left = !hasIn[n.index];
            const t = lab.append("text")
                .attr("x", left ? n.x0 - 6 : n.x1 + 6)
                .attr("y", (n.y0 + n.y1) / 2)
                .attr("dy", "0.35em")
                .attr("text-anchor", left ? "end" : "start")
                .attr("fill", "#374151")
                .style("paint-order", "stroke").style("stroke", "#ffffff")
                .style("stroke-width", "3px").style("stroke-linejoin", "round");
            t.append("tspan").attr("font-weight", 600).text(n.name);
            if (S.showValues) {
                t.append("tspan").attr("dx", 6).attr("fill", "#6B7280").text(a.formatValue(nodeTotal(n.index)));
            }
        });
    }
    return null;
}
