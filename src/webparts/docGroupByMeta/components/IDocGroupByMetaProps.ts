import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { SPFI } from "@pnp/sp";

export interface IDocGroupByMetaProps {
	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
	sp: SPFI;
	currentTheme: IReadonlyTheme | undefined;
	docLibraryName: string;
	metadataFieldName: string;
	searchPageUrl: string;
}
