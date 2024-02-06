import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { SPFI } from "@pnp/sp";

export interface IDocsGroupByEntKeywordProps {
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    sp: SPFI;
    currentTheme: IReadonlyTheme | undefined;
    siteUrl: string;
    keywords: string;
    searchPageUrl: string;
}
