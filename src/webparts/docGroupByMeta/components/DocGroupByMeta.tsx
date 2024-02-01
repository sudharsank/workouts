import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import styles from './DocGroupByMeta.module.scss';
import type { IDocGroupByMetaProps } from './IDocGroupByMetaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as _ from 'lodash';
import { DetailsList, IColumn, IDetailsGroupRenderProps, IDetailsHeaderProps, IDetailsList, IGroup, IGroupDividerProps, IRenderFunction, Icon, Link, SelectionMode, Stack } from '@fluentui/react';
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

	const demo = async () => {
		// const tempQuery = `<Where>
		// 						<Contains>
		// 							<FieldRef Name="Categories0"/>
		// 							<Value Type="TaxonomyFieldType">Credit</Value>
		// 						</Contains>
		// 					</Where>`;
		const query: string = `<View Scope="RecursiveAll"><Query></Query>
									<ViewFields>
										<FieldRef Name='Categories0'/><FieldRef Name='FileRef'/>
										<FieldRef Name='FileLeafRef'/>
									</ViewFields>
								</View>`
		let docs = await props.sp.web.lists.getByTitle('Documents').getItemsByCAMLQuery({ ViewXml: query }, 'FileRef,FileLeafRef');
		docs = _.sortBy(docs, 'Categories0.Label');
		var groupedDocs = _.groupBy(docs, 'Categories0.Label');
		let docGroups: IGroup[] = [];
		_.map(groupedDocs, (value, groupkey) => {
			docGroups.push({
				key: groupkey,
				name: groupkey,
				count: value.length,
				startIndex: _.indexOf(docs, _.filter(docs, (d: any) => d.Categories0.Label == groupkey)[0]),
				data: _.filter(docs, (d: any) => d.Categories0.Label == groupkey),
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
		window.open(`${window.location.origin}/_layouts/15/search.aspx/?q=${keyword}`, '_blank');
	};

	const _onToggleCollapse = (props: IGroupDividerProps) => {
		return () => props!.onToggleCollapse!(props!.group!);
	};

	const _onRenderGroupHeader: IDetailsGroupRenderProps['onRenderHeader'] = props => {
		if (props) {
			return (
				<Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='start' style={{ marginTop: '10px' }}>
					<Stack.Item>
						<Link onClick={_onToggleCollapse(props)} style={{marginTop: '3px'}}>
							{props.group!.isCollapsed ? <Icon iconName='CaretRight8' /> : <Icon iconName='CaretDown8' />}
						</Link>
					</Stack.Item>
					<Stack.Item style={{ fontSize: 15 }}>
						{props.group?.name} ({props.group?.count}) -
					</Stack.Item>
					<Stack.Item>
						<Link onClick={() => _openSearchPage(props.group?.name)} style={{marginTop: '2px'}}>Show More...</Link>
					</Stack.Item>
				</Stack>
			)
		}
	};

	useEffect(() => {
		(async () => {
			await demo();
		})();
	}, []);

	return (
		<section className={`${styles.docGroupByMeta} ${hasTeamsContext ? styles.teams : ''}`}>
			<div className={styles.welcome}>
				<img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
				<h2>Well done, {escape(userDisplayName)}!</h2>
				<div>{environmentMessage}</div>
			</div>
			<div>
				<h3>Welcome to SharePoint Framework!</h3>
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
		</section>
	);

};

export default DocGroupByMeta;
