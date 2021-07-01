/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT License.
 *
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is regenerated.
 */

import {
  TilesetID,
  RenderV2GetMapTilePreviewOptionalParams,
  RenderV2GetMapTilePreviewResponse
} from "../models";

/** Interface representing a RenderV2. */
export interface RenderV2 {
  /**
   * **Applies to**: S0 and S1 pricing tiers.
   *
   * The Get Map Tiles API allows users to request map tiles in vector or raster formats typically to be
   * integrated  into a map control or SDK. Some example tiles that can be requested are Azure Maps road
   * tiles, real-time  Weather Radar tiles or the map tiles created using [Azure Maps
   * Creator](https://aka.ms/amcreator). By default,  Azure Maps uses vector tiles for its web map
   * control (Web SDK) and Android SDK.
   * @param tilesetId A tileset is a collection of raster or vector data broken up into a uniform grid of
   *                  square tiles at preset  zoom levels. Every tileset has a **tilesetId** to use when making requests.
   *                  The **tilesetId** for tilesets created using [Azure Maps Creator](https://aka.ms/amcreator) are
   *                  generated through the  [Tileset Create API](https://docs.microsoft.com/en-us/rest/api/maps/tileset).
   *                  The ready-to-use tilesets supplied  by Azure Maps are listed below. For example, microsoft.base.
   * @param zoom Zoom level for the desired tile. Please find TilesetID list below for more details on
   *             supported zoom level for each tilesetId.<br>
   *
   * Please see [Zoom Levels and Tile
   *             Grid](https://docs.microsoft.com/en-us/azure/location-based-services/zoom-levels-and-tile-grid) for
   *             details.
   * @param xTileIndex X coordinate of the tile on zoom grid. Value must be in the range [0,
   *                   2<sup>`zoom`</sup> -1].
   *
   * Please see [Zoom Levels and Tile
   *                   Grid](https://docs.microsoft.com/en-us/azure/location-based-services/zoom-levels-and-tile-grid) for
   *                   details.
   * @param yTileIndex Y coordinate of the tile on zoom grid. Value must be in the range [0,
   *                   2<sup>`zoom`</sup> -1].
   *
   * Please see [Zoom Levels and Tile
   *                   Grid](https://docs.microsoft.com/en-us/azure/location-based-services/zoom-levels-and-tile-grid) for
   *                   details.
   * @param options The options parameters.
   */
  getMapTilePreview(
    tilesetId: TilesetID,
    zoom: number,
    xTileIndex: number,
    yTileIndex: number,
    options?: RenderV2GetMapTilePreviewOptionalParams
  ): Promise<RenderV2GetMapTilePreviewResponse>;
}
