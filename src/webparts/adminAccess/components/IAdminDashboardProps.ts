import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminDashboardProps {
    description: string;
    userDisplayName: string;
    context: WebPartContext;
}
