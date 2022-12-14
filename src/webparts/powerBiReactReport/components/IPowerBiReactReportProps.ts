import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ServiceScope } from "@microsoft/sp-core-library";

export interface IPowerBiReactReportProps {
  webPartContext: WebPartContext;
  serviceScope: ServiceScope;
  defaultWorkspaceId: string;
  defaultReportId: string;
  defaultWidthToHeight: number;
  items:any;
}

export interface IPowerBiReactReportState {
  loading: boolean;
  workspaceId: string;
  reportId: string;
  widthToHeight: number;
}
