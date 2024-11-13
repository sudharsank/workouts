import * as React from 'react';
import styles from './BrandFontDemo.module.scss';
import type { IBrandFontDemoProps } from './IBrandFontDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class BrandFontDemo extends React.Component<IBrandFontDemoProps, {}> {
  public render(): React.ReactElement<IBrandFontDemoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.brandFontDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
        </div>
        <div>
            <h2>Brand font demo</h2>
            <div className={styles.bcTitle}>Title</div>
            <div className={styles.bcBody}>Body</div>
            <div className={styles.bcInteractive}>Interactive</div>
            <div className={styles.bcHeadline}>Headline</div>
        </div>
      </section>
    );
  }
}
