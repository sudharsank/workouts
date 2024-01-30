import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import styles from './DocGroupByMeta.module.scss';
import type { IDocGroupByMetaProps } from './IDocGroupByMetaProps';
import { escape } from '@microsoft/sp-lodash-subset';

const DocGroupByMeta: FC<IDocGroupByMetaProps> = (props) => {
	const {
		isDarkTheme,
		environmentMessage,
		hasTeamsContext,
		userDisplayName
	} = props;

	const demo = async () => {
		const query: string = `<View>
									<Query>
										<Where>
											<Contains>
												<FieldRef Name="Categories0"/>
												<Value Type="TaxonomyFieldType">Credit</Value>
											</Contains>
										</Where>
									</Query>
									<ViewFields>
										<FieldRef Name='Categories0'/><FieldRef Name='FileLeaf'/>
										<FieldRef Name='FileLeafRef'/>
									</ViewFields>
								</View>`
		let docs = await props.sp.web.lists.getByTitle('Documents').getItemsByCAMLQuery({ ViewXml: query });

		//.filter(`Categories0 eq 'Credit Documents'`)();
		console.log('Documents: ', docs);
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
			</div>
		</section>
	);

};

export default DocGroupByMeta;
