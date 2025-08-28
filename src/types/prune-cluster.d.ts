// src/types/prune-cluster.d.ts

import { Layer, Marker as LeafletMarker } from "leaflet";

declare module "@workingfamilies/prune-cluster" {
  export class Marker {
    constructor(lat: number, lon: number);
    data: any;
  }

  export class PruneClusterForLeaflet extends Layer {
    constructor();
    RemoveMarkers(): void;
    RegisterMarker(marker: Marker): void;
    ProcessView(): void;
    PrepareLeafletMarker?: (leafletMarker: LeafletMarker, data: any) => void;
  }

  const _default: {
    Marker: typeof Marker;
    PruneClusterForLeaflet: typeof PruneClusterForLeaflet;
  };
  export default _default;
}