"use strict";

// Leaflet & Esri Leaflet
import * as L from "leaflet";
import "leaflet/dist/leaflet.css";
import * as LEsri from "esri-leaflet";
type EsriFeatureLayer = ReturnType<typeof LEsri.featureLayer>;

// Grouped Layer Control
import "leaflet-groupedlayercontrol";
import "./../style/leaflet.groupedlayercontrol.css";

// Power BI
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

// For reading tabular data
import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;

// County GeoJSON + FIPS lookup
import counties from "./data/nc_counties";
import fipsToCounty from "./data/fipsToCounty";

import "./../style/visual.less";

export class Visual implements IVisual {
  private target!: HTMLElement;
  private map!: L.Map;

  private inited = false; // <-- defer vector layers until map has size

  // Basemaps
  private darkMap!: L.TileLayer;
  private lightMap!: L.TileLayer;

  // Overlays
  private countiesLayer!: L.GeoJSON;
  private layerControl?: L.Control.Layers;

  // Work zone layers
  private truckClosureLayer?: EsriFeatureLayer;
  private constructionLayer?: EsriFeatureLayer;
  private nightConstructionLayer?: EsriFeatureLayer;
  private maintenanceLayer?: EsriFeatureLayer;
  private nightMaintenanceLayer?: EsriFeatureLayer;
  private emergencyLayer?: EsriFeatureLayer;
  private obstructionLayer?: EsriFeatureLayer;
  private weatherLayer?: EsriFeatureLayer;
  private specialLayer?: EsriFeatureLayer;
  private otherLayer?: EsriFeatureLayer;

  private autoRefreshTimer: number | null = null;

  constructor(options: VisualConstructorOptions | undefined) {
    this.target = options?.element!;
    this.createMapContainer();
    this.initMap(); // ONLY map + panes + basemaps here
  }

  public update(options: VisualUpdateOptions): void {
    if (!this.inited) {
      this.inited = true;
      // Defer one frame so Leaflet computes renderer bounds after size is applied
      requestAnimationFrame(() => this.initLayersAndUI());
    }

    // just invalidate size; no need to set px width/height
    requestAnimationFrame(() => this.map.invalidateSize());

    // For reading tabular data (if needed in future)
    const dataView: DataView | undefined = options.dataViews && options.dataViews[0];
    if (!dataView || !dataView.categorical) {
      console.log("No data");
      return;
    }

    const categorical = dataView.categorical;

    // Categories = roles of kind "Grouping" (tables)
    if (categorical.categories) {
      categorical.categories.forEach((cat: DataViewCategoryColumn) => {
        console.log("Category role:", cat.source.roles);
        console.log("DisplayName:", cat.source.displayName);
        console.log("Values:", cat.values);
      });
    }

    // Values = roles of kind "Measure" (ignore if you only want groupings)
    if (categorical.values) {
      categorical.values.forEach(val => {
        console.log("Measure role:", val.source.roles);
        console.log("DisplayName:", val.source.displayName);
        console.log("Values:", val.values);
      });
    }    
  }

  public destroy(): void {
    if (this.autoRefreshTimer) clearInterval(this.autoRefreshTimer);
    document.removeEventListener("visibilitychange", this.onVisibilityRefresh);
    this.map.remove();
  }

  // ---------------- internals ----------------

  private createMapContainer() {
    const existing = document.getElementById("mapid");
    if (existing) existing.remove();
    const div = document.createElement("div");
    div.id = "mapid";
    div.style.width = "100%";
    div.style.height = "100%";
    this.target.appendChild(div);
  }

  private initMap() {
    this.map = L.map("mapid", { center: [35.54, -79.24], zoom: 7, maxZoom: 20, minZoom: 3 });

    // panes with explicit z-index (tiles < counties < workzones)
    this.map.createPane("counties");
    this.map.getPane("counties")!.style.zIndex = "400";

    this.map.createPane("workzones");
    this.map.getPane("workzones")!.style.zIndex = "450";

    // basemaps
    this.darkMap = L.tileLayer("https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png", {
      attribution: '&copy; <a href="https://carto.com/">CARTO</a>',
      subdomains: "abcd",
      maxZoom: 20
    }).addTo(this.map);

    this.lightMap = L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    });
  }

  private initLayersAndUI() {
    // -------- County overlay --------
    this.countiesLayer = L.geoJSON(counties as any, {
      pane: "counties",
      style: { fill: true, fillOpacity: 0.0, fillColor: "#000", weight: 1, color: "gray", dashArray: "3" },
      onEachFeature: (feature, layer) => {
        const fips = (feature.properties as any)?.FIPS;
        if (fips != null) {
          const name = fipsToCounty[Number(fips)] ?? "Unknown";
          (layer as L.Path).bindTooltip(`${name} County`, { sticky: true });
          layer.on("mouseover", () => (layer as L.Path).setStyle({ weight: 3, color: "yellow" }));
          layer.on("mouseout", () => (layer as L.Path).setStyle({ weight: 1, color: "gray" }));
        }
      }
    }).addTo(this.map);

    // -------- Work zone style / tooltip --------
    const workzoneLinePolyStyle: L.PathOptions = {
      pane: "workzones",
      interactive: true,
      weight: 8,
      color: "blue",
      opacity: 0.5
    };

    const workzoneEachFeature = (feature: GeoJSON.Feature, layer: L.Layer) => {
      const p = (feature.properties as any) || {};
      const sev = p.Severity === 1 ? "Low" : p.Severity === 2 ? "Medium" : p.Severity === 3 ? "High" : "Not Provided";
      const dirMap: Record<string, string> = { N: "North", S: "South", E: "East", W: "West", A: "All", O: "Outer" };
      const direction = dirMap[p.Direction] || "Not Provided";
      const label = `<strong>Type:</strong> ${p.IncidentType || "Not Provided"}<br/>
                     <strong>Impact Level:</strong> ${sev}<br/>
                     <strong>Condition:</strong> ${p.Condition || "Not Provided"}<br/>
                     <strong>Place:</strong> ${p.City || "Unknown City"}, ${p.CountyName || "Unknown"} County<br/>
                     <strong>Road:</strong> ${p.Road || "Not Provided"}<br/>
                     <strong>Direction:</strong> ${direction}<br/>
                     <strong>Reason:</strong> ${p.Reason || "Not Provided"}<br/>
                     <strong>Until:</strong> ${p.EndDateET || "Not Provided"}`;
      (layer as L.Path).bindTooltip(label, { sticky: true });
      layer.on("mouseover", () => (layer as L.Path).setStyle({ color: "yellow" }));
      layer.on("mouseout", () => (layer as L.Path).setStyle({ color: "blue" }));
    };

    // -------- Work zones (Esri Feature Service) --------
    const tims =
      "https://services.arcgis.com/NuWFvHYDMVmmxMeM/ArcGIS/rest/services/NCDOT_TIMSIncidentsByIncidentType/FeatureServer/1";

    const wz = (where: string): EsriFeatureLayer =>
      LEsri.featureLayer({
        url: tims,
        where,
        pane: "workzones",
        style: workzoneLinePolyStyle,
        onEachFeature: workzoneEachFeature
      });

    this.truckClosureLayer = wz("IncidentType = 'Truck Closure'");
    this.constructionLayer = wz("IncidentType = 'Construction'");
    this.nightConstructionLayer = wz("IncidentType = 'Night Time Construction'");
    this.maintenanceLayer = wz("IncidentType = 'Maintenance'");
    this.nightMaintenanceLayer = wz("IncidentType = 'Night Time Maintenance'");
    this.emergencyLayer = wz("IncidentType = 'Emergency Road Work'");
    this.obstructionLayer = wz("IncidentType = 'Road Obstruction'");
    this.weatherLayer = wz("IncidentType = 'Weather Event'");
    this.specialLayer = wz("IncidentType = 'Special Event'");
    this.otherLayer = wz("IncidentType = 'Other'");

    // Show at least one category by default so users see data immediately
    this.constructionLayer.addTo(this.map);

    // -------- Layer control --------
    const baseMaps = {
      "Dark Basemap": this.darkMap,
      "Light Basemap": this.lightMap
    };

    const groupedOverlays: Record<string, Record<string, L.Layer>> = {
      "Boundaries": {
        "Counties": this.countiesLayer
      },
      "Work Zones": {
        "Truck Closure": this.truckClosureLayer,
        "Construction": this.constructionLayer,
        "Night Construction": this.nightConstructionLayer,
        "Maintenance": this.maintenanceLayer,
        "Night Time Maintenance": this.nightMaintenanceLayer,
        "Emergency Road Work": this.emergencyLayer,
        "Road Obstruction": this.obstructionLayer,
        "Weather Event": this.weatherLayer,
        "Special Event": this.specialLayer,
        "Other": this.otherLayer
      }
    };

    // @ts-ignore - plugin augments L with control.groupedLayers
    this.layerControl = (L as any).control
      .groupedLayers(baseMaps, groupedOverlays, {
        collapsed: false,
        exclusiveGroups: ["Boundaries"],
        groupCheckboxes: false
      })
      .addTo(this.map) as L.Control.Layers;

    this.map.on("overlayadd overlayremove baselayerchange", () => {
      requestAnimationFrame(() => this.refreshLayerControlCollapsibles());
    });

    // -------- Auto refresh AFTER layers exist --------
    this.startAutoRefresh(6 * 60 * 60 * 1000); // every 6 hrs
    document.addEventListener("visibilitychange", this.onVisibilityRefresh);
  }

  private makeGroupedLayerCollapsible(control: L.Control, groupName: string) {
    const container =
      (control as any).getContainer?.() as HTMLElement ??
      (control as any)._container as HTMLElement;
    if (!container) return;

    const nameSpan = Array.from(container.querySelectorAll(".leaflet-control-layers-group-name"))
      .find(el => (el.textContent ?? "").trim().toLowerCase() === groupName.trim().toLowerCase()) as HTMLElement | undefined;
    if (!nameSpan) return;

    const headerLabel = (nameSpan.closest(".leaflet-control-layers-group-label") as HTMLElement) || nameSpan;
    const groupRoot = (nameSpan.closest(".leaflet-control-layers-group") as HTMLElement) || headerLabel.parentElement!;
    if (!groupRoot) return;

    const hdr = headerLabel as HTMLElement & { _gliWired?: boolean };
    if (hdr._gliWired) return; hdr._gliWired = true;

    const items = Array.from(groupRoot.children).filter(el => el !== headerLabel) as HTMLElement[];
    if (!headerLabel.querySelector(".gli-caret")) {
      const caret = document.createElement("span");
      caret.className = "gli-caret";
      caret.setAttribute("role", "button");
      caret.setAttribute("aria-label", "Toggle group");
      caret.tabIndex = 0;
      headerLabel.insertBefore(caret, nameSpan);
    }
    const setCollapsed = (collapsed: boolean) => {
      items.forEach(el => { el.style.display = collapsed ? "none" : ""; });
      headerLabel.classList.toggle("is-collapsed", collapsed);
    };
    const toggle = () => setCollapsed(!(items[0]?.style.display === "none"));

    headerLabel.style.cursor = "pointer";
    headerLabel.addEventListener("click", (e: MouseEvent) => { e.preventDefault(); e.stopPropagation(); toggle(); });
    headerLabel.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); }
    });

    L.DomEvent.disableClickPropagation(headerLabel);
    L.DomEvent.disableScrollPropagation(headerLabel);

    setCollapsed(true);
  }

  private refreshLayerControlCollapsibles() {
    if (!this.layerControl) return;
    requestAnimationFrame(() => {
      this.makeGroupedLayerCollapsible(this.layerControl!, "Boundaries");
      this.makeGroupedLayerCollapsible(this.layerControl!, "Work Zones");
    });
  }

  // --------------- refresh helpers ---------------
  private onVisibilityRefresh = () => {
    if (document.visibilityState === "visible") this.refreshAllLayers();
  };

  private startAutoRefresh(ms: number) {
    if (this.autoRefreshTimer) clearInterval(this.autoRefreshTimer);
    if (ms > 0) this.autoRefreshTimer = window.setInterval(() => this.refreshAllLayers(), ms);
  }

  private refreshAllLayers() {
    const layers = [
      this.truckClosureLayer, this.constructionLayer, this.nightConstructionLayer,
      this.maintenanceLayer, this.nightMaintenanceLayer, this.emergencyLayer,
      this.obstructionLayer, this.weatherLayer, this.specialLayer, this.otherLayer
    ];
    for (const l of layers) (l as any)?.refresh?.();
  }

}