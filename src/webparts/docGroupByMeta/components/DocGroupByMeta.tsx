import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import styles from './DocGroupByMeta.module.scss';
import type { IDocGroupByMetaProps } from './IDocGroupByMetaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as _ from 'lodash';
import { DetailsList, IColumn, IDetailsList, IGroup } from '@fluentui/react';

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
		console.log(docs[0].Categories0.Label);
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
		console.log('Documents: ', docGroups);
	};

	const root = React.useRef<IDetailsList>(null);
	const [columns] = React.useState<IColumn[]>([
		{ key: 'name', name: 'Name', fieldName: 'FileLeafRef', minWidth: 100, maxWidth: 200, isResizable: true },
		{ key: 'color', name: 'Color', fieldName: 'Categories0.Label', minWidth: 100, maxWidth: 200 },
	]);

	// const [items, setItems] = React.useState<any[]>([
	// 	{ key: 'a', name: 'a', color: 'red' },
	// 	{ key: 'b', name: 'b', color: 'red' },
	// 	{ key: 'c', name: 'c', color: 'blue' },
	// 	{ key: 'd', name: 'd', color: 'blue' },
	// 	{ key: 'e', name: 'e', color: 'blue' },
	// ]);

	// const [groups, setGroups] = React.useState<IGroup[]>([
	// 	{ key: 'groupred0', name: 'Color: "red"', startIndex: 0, count: 2, level: 0 },
	// 	{ key: 'groupgreen2', name: 'Color: "green"', startIndex: 2, count: 0, level: 0 },
	// 	{ key: 'groupblue2', name: 'Color: "blue"', startIndex: 2, count: 3, level: 0 },
	// ]);

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
				// onRenderDetailsHeader={onRenderDetailsHeader}
				// groupProps={{
				// 	showEmptyGroups: true,
				// 	groupedListAs: GroupedListV2,
				// }}
				// onRenderItemColumn={onRenderColumn}
				// compact={isCompactMode}
				/>
			</div>
		</section>
	);

};

export default DocGroupByMeta;
