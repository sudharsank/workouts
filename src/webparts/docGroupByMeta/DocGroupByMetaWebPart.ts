import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
	PropertyPaneTextField,
	type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPFI } from '@pnp/sp';
import * as strings from 'DocGroupByMetaWebPartStrings';
import DocGroupByMeta from './components/DocGroupByMeta';
import { IDocGroupByMetaProps } from './components/IDocGroupByMetaProps';
import { getSP } from '../../pnp.config';

export interface IDocGroupByMetaWebPartProps {
	docLibraryName: string;
	metadataFieldName: string;
	searchPageUrl: string;
}

export default class DocGroupByMetaWebPart extends BaseClientSideWebPart<IDocGroupByMetaWebPartProps> {

	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = '';
	private _sp: SPFI;
	private _currentTheme: IReadonlyTheme | undefined;

	public render(): void {
		const element: React.ReactElement<IDocGroupByMetaProps> = React.createElement(
			DocGroupByMeta,
			{
				sp: this._sp,
				isDarkTheme: this._isDarkTheme,
				environmentMessage: this._environmentMessage,
				hasTeamsContext: !!this.context.sdks.microsoftTeams,
				userDisplayName: this.context.pageContext.user.displayName,
				currentTheme: this._currentTheme,
				docLibraryName: this.properties.docLibraryName,
				metadataFieldName: this.properties.metadataFieldName,
				searchPageUrl: this.properties.searchPageUrl
			}
		);

		ReactDom.render(element, this.domElement);
	}

	protected onInit(): Promise<void> {
		this._sp = getSP(this.context);
		return this._getEnvironmentMessage().then(message => {
			this._environmentMessage = message;
		});
	}



	private _getEnvironmentMessage(): Promise<string> {
		if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
			return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
				.then(context => {
					let environmentMessage: string = '';
					switch (context.app.host.name) {
						case 'Office': // running in Office
							environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
							break;
						case 'Outlook': // running in Outlook
							environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
							break;
						case 'Teams': // running in Teams
						case 'TeamsModern':
							environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
							break;
						default:
							environmentMessage = strings.UnknownEnvironment;
					}

					return environmentMessage;
				});
		}

		return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
	}

	protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
		if (!currentTheme) {
			return;
		}

		this._isDarkTheme = !!currentTheme.isInverted;
		const {
			semanticColors
		} = currentTheme;
		this._currentTheme = currentTheme;

		if (semanticColors) {
			this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
			this.domElement.style.setProperty('--link', semanticColors.link || null);
			this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
		}

	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription
					},
					groups: [
						{
							groupFields:[
								PropertyPaneTextField('docLibraryName', {
									label: 'Document library name',
									multiline: false,
									value: this.properties.docLibraryName
								}),
								PropertyPaneTextField('metadataFieldName', {
									label: 'Metadata Field name',
									multiline: false,
									value: this.properties.metadataFieldName
								}),
								PropertyPaneTextField('searchPageUrl', {
									label: 'Search results URL',
									multiline: false,
									value: this.properties.searchPageUrl
								})
							]
						}
					]
				}
			]
		};
	}
}
