import * as React from 'react';
import { FC, useEffect } from 'react';
import styles from './TestMultitenant.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI } from '@pnp/sp';
import { PrimaryButton } from '@fluentui/react';
import { GraphFI, SPFx as graphSPFx, graphfi } from "@pnp/graph";
import { SPFx as spSPFx, spfi } from "@pnp/sp";
import { MSAL } from "@pnp/msaljsclient";
import { Configuration, AuthenticationParameters } from "msal";
import "@pnp/graph/users";
import "@pnp/sp/webs";
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITestMultitenantProps {
    sp: SPFI;
    wpContext: WebPartContext;
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
}

const TestMultitenant: FC<ITestMultitenantProps> = (props) => {
    let graph: GraphFI = null;
    let sp: SPFI = null;
    const {
        description,
        isDarkTheme,
        environmentMessage,
        hasTeamsContext,
        userDisplayName
    } = props;

    const configuration: Configuration = {
        auth: {
            authority: "https://login.microsoftonline.com/common",
            clientId: "aa19eb94-4970-4b86-a5ff-7e6dd86f897b"            
        }
    };

    const authParams: AuthenticationParameters = {
        scopes: ["https://graph.microsoft.com/.default"]
    };

    const _getAuth = async () => {
        graph = graphfi().using(graphSPFx(props.wpContext), MSAL(configuration, authParams));
        const tenantUsers = await graph.users();

        console.log(tenantUsers);
    };

    useEffect(() => {
        console.log('Webpart load.');
        
    }, []);

    return (
        <section className={`${styles.testMultitenant} ${hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.welcome}>
                <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
                <h2>Well done, {escape(userDisplayName)}!</h2>
                <div>{environmentMessage}</div>
                <div>Web part property value: <strong>{escape(description)}</strong></div>
            </div>
            <div>
                <PrimaryButton onClick={_getAuth}>Get Token</PrimaryButton>
            </div>
        </section>
    );
};

export default TestMultitenant;
