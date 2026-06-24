import { registry } from "@skalfa/skalfa-app-core";
import { ExportExcel } from "./ExportExcel.component.js";
import { ImportExcel } from "./ImportExcel.component.js";
import { DocumentViewerComponent, DocumentViewerIcon } from "./DocumentViewer.component.js";

export * from "./DocumentViewer.component.js";
export * from "./ExportExcel.component.js";
export * from "./ImportExcel.component.js";
export * from "./PrintTable.component.js";
export * from "./RenderPDF.component.js";

registry.register("ExportExcel", ExportExcel);
registry.register("ImportExcel", ImportExcel);
registry.register("DocumentViewerComponent", DocumentViewerComponent);
registry.register("DocumentViewerIcon", DocumentViewerIcon);
