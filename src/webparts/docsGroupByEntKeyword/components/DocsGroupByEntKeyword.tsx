import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './DocsGroupByEntKeyword.module.scss';
import type { IDocsGroupByEntKeywordProps } from './IDocsGroupByEntKeywordProps';
import * as _ from 'lodash';
import {
    DetailsList, IColumn, IDetailsGroupRenderProps, IDetailsList, IGroup,
    IGroupDividerProps, Icon, Link, MessageBar, MessageBarType, SelectionMode, Spinner, SpinnerSize, Stack
} from '@fluentui/react';
import { GroupedListV2FC } from '@fluentui/react/lib/GroupedList';
import { ISearchQuery, SearchResults } from "@pnp/sp/search";

export interface IDocs {
    Name: string;
    FileUrl: string;
    Keyword: string;
}

const DocsGroupByEntKeyword: React.FC<IDocsGroupByEntKeywordProps> = (props) => {
    const {
        isDarkTheme,
        environmentMessage,
        hasTeamsContext,
        userDisplayName
    } = props;

    const [loading, setLoading] = useState<boolean>(true);
    const [items, setItems] = useState<any[]>([]);
    const [groups, setGroups] = React.useState<IGroup[]>([]);

    const loadDocuments = async () => {
        let finalDocs: IDocs[] = [];
        const enterpriseKeywordsManagedProperty = "ows_MetadataFacetInfo";
        const pathManagedProperty = "Path";
        const fileExtensionManagedProperty = "FileExtension";
        const keywordsToSearch = props.keywords?.split(',');
        const keywordConditions = keywordsToSearch.map(keyword => `${enterpriseKeywordsManagedProperty}:"${keyword}"`).join(" OR ");

        const siteUrl = props.siteUrl!; // Replace with the URL of the specific site
        const queryText = `(${keywordConditions}) AND ${fileExtensionManagedProperty}:<> AND ${pathManagedProperty}:"${siteUrl}/*"`;

        let searchQuery: ISearchQuery = {
            Querytext: queryText,
            RowLimit: 50,
            SelectProperties: ['Path', enterpriseKeywordsManagedProperty, 'owstaxIdTaxKeyword',
                "DefaultEncodingURL",
                "FileType",
                "Filename",
                "OriginalPath", // Folder path
                "Author", // Document author
                "LastModifiedTime", // Last modified timestamp
                "Created"]
        }
        const results: SearchResults = await props.sp.search(searchQuery);
        
        let searchResults: any[] = results.PrimarySearchResults;
        searchResults.map((result: any) => {
            props.keywords.split(',').map((key: string) => {
                if (result[enterpriseKeywordsManagedProperty].toLowerCase().indexOf(key.toLowerCase()) >= 0) {
                    finalDocs.push({
                        Keyword: key,
                        Name: result.Filename,
                        FileUrl: result.DefaultEncodingURL
                    });
                }
            })
        });
        finalDocs = _.sortBy(finalDocs, 'Keyword');
        let groupedDocs = _.groupBy(finalDocs, 'Keyword');
        let docGroups: IGroup[] = [];
        _.map(groupedDocs, (value, groupkey) => {
            docGroups.push({
                key: groupkey,
                name: groupkey,
                count: value.length,
                startIndex: _.indexOf(finalDocs, _.filter(finalDocs, (d: any) => d[`Keyword`] == groupkey)[0]),
                data: _.filter(finalDocs, (d: any) => d[`Keyword`] == groupkey),
                level: 0
            });
        });
        setItems(finalDocs);
        setGroups(docGroups);
        setLoading(false);
    };

    const _onNavigate = (targeturl: string) => {
        window.open(targeturl, '_blank');
    };
    const root = React.useRef<IDetailsList>(null);
    const [columns] = React.useState<IColumn[]>([
        {
            key: 'name', name: 'Name', fieldName: 'Name', minWidth: 100, maxWidth: 200, isResizable: true,
            onRender: (item?: IDocs, index?: number, column?: IColumn) => {
                return (
                    <Link onClick={() => _onNavigate(item.FileUrl)} style={{ marginTop: '3px' }}>
                        {item.Name}
                    </Link>
                );
            }
        },
    ]);

    const _openSearchPage = (keyword: string): void => {
        if (props.searchPageUrl) window.open(`${props.searchPageUrl}/files?q=${keyword}`, '_blank');
    };

    const _onToggleCollapse = (props: IGroupDividerProps) => {
        return () => props!.onToggleCollapse!(props!.group!);
    };

    const _onRenderGroupHeader: IDetailsGroupRenderProps['onRenderHeader'] = props => {
        if (props) {
            return (
                <Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='start' style={{ marginTop: '10px' }}>
                    <Stack.Item>
                        <Link onClick={_onToggleCollapse(props)} style={{ marginTop: '3px' }}>
                            {props.group!.isCollapsed ? <Icon iconName='CaretRight8' /> : <Icon iconName='CaretDown8' />}
                        </Link>
                    </Stack.Item>
                    <Stack.Item style={{ fontSize: 15 }}>
                        {props.group?.name} ({props.group?.count}) -
                    </Stack.Item>
                    <Stack.Item>
                        <Link onClick={() => _openSearchPage(props.group?.name)} style={{ marginTop: '2px' }}>Show More...</Link>
                    </Stack.Item>
                </Stack>
            )
        }
    };

    useEffect(() => {
        loadDocuments();
    }, []);

    return (
        <section className={`${styles.docsGroupByEntKeyword} ${hasTeamsContext ? styles.teams : ''}`}>
            {loading ? (
                <Spinner size={SpinnerSize.large} label='Please wait...' labelPosition='top' />
            ) : (
                <>
                    {props.siteUrl && props.keywords ? (
                        <div>
                            <DetailsList
                                componentRef={root}
                                items={items}
                                groups={groups}
                                columns={columns}
                                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                ariaLabelForSelectionColumn="Toggle selection"
                                checkButtonAriaLabel="select row"
                                checkButtonGroupAriaLabel="select section"
                                groupProps={{
                                    onRenderHeader: _onRenderGroupHeader,
                                    groupedListAs: GroupedListV2FC
                                }}
                                selectionMode={SelectionMode.none}
                            />
                        </div>
                    ) : (
                        <MessageBar messageBarType={MessageBarType.warning}>Webpart configuration missing...</MessageBar>
                    )}
                </>
            )}
        </section>
    );

};

export default DocsGroupByEntKeyword;
