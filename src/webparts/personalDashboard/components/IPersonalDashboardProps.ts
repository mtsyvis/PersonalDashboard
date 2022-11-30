import { ServiceScope } from "@microsoft/sp-core-library";

export interface IPersonalDashboardProps {
    hasTeamsContext: boolean;
    userDisplayName: string;
    serviceScope: ServiceScope;
}
