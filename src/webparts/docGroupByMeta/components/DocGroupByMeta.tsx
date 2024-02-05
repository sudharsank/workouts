import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import styles from './DocGroupByMeta.module.scss';
import type { IDocGroupByMetaProps } from './IDocGroupByMetaProps';
import * as _ from 'lodash';
import {
	DetailsList, IColumn, IDetailsGroupRenderProps, IDetailsList, IGroup,
	IGroupDividerProps, Icon, Link, MessageBar, MessageBarType, SelectionMode, Stack
} from '@fluentui/react';
import { GroupedListV2FC } from '@fluentui/react/lib/GroupedList';

const DocGroupByMeta: FC<IDocGroupByMetaProps> = (props) => {
	const {
		isDarkTheme,
		environmentMessage,
		hasTeamsContext,
		userDisplayName
	} = props;

	const [items, setItems] = useState<any[]>([]);
	const [groups, setGroups] = React.useState<IGroup[]>([]);

	const loadDocuments = async () => {
		const query: string = `<View Scope="RecursiveAll"><Query></Query>
									<ViewFields>
										<FieldRef Name='${props.metadataFieldName}'/><FieldRef Name='FileRef'/>
										<FieldRef Name='FileLeafRef'/>
									</ViewFields>
								</View>`
		let docs = await props.sp.web.lists.getByTitle(props.docLibraryName).getItemsByCAMLQuery({ ViewXml: query }, 'FileRef,FileLeafRef');
		docs = _.sortBy(docs, `${props.metadataFieldName}.Label`);
		var groupedDocs = _.groupBy(docs, `${props.metadataFieldName}.Label`);
		let docGroups: IGroup[] = [];
		_.map(groupedDocs, (value, groupkey) => {
			docGroups.push({
				key: groupkey,
				name: groupkey,
				count: value.length,
				startIndex: _.indexOf(docs, _.filter(docs, (d: any) => d[`${props.metadataFieldName}.Label`] == groupkey)[0]),
				data: _.filter(docs, (d: any) => d[`${props.metadataFieldName}.Label`] == groupkey),
				level: 0
			});
		});
		setItems(docs);
		setGroups(docGroups);
	};

	const root = React.useRef<IDetailsList>(null);
	const [columns] = React.useState<IColumn[]>([
		{ key: 'name', name: 'Name', fieldName: 'FileLeafRef', minWidth: 100, maxWidth: 200, isResizable: true },
	]);

	const _openSearchPage = (keyword: string): void => {
		if (props.searchPageUrl) window.open(`${props.searchPageUrl}?q=${keyword}`, '_blank');
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
		(async () => {
			if (props.docLibraryName && props.metadataFieldName) await loadDocuments();
		})();
	}, []);

	return (
		<section className={`${styles.docGroupByMeta} ${hasTeamsContext ? styles.teams : ''}`}>
			{props.docLibraryName && props.metadataFieldName ? (
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

		</section>
	);

};

export default DocGroupByMeta;
