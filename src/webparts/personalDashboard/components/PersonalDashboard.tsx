import * as React from 'react';
import styles from './PersonalDashboard.module.scss';
import { IPersonalDashboardProps } from './IPersonalDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PersonalDashboard extends React.Component<IPersonalDashboardProps, {}> {
    public render(): React.ReactElement<IPersonalDashboardProps> {
        const {
            hasTeamsContext,
            userDisplayName,
            serviceScope
        } = this.props;

        return (
            <section className={`${styles.personalDashboard} ${hasTeamsContext ? styles.teams : ''}`}>
                <div className={styles.welcome}>
                    <h2>Hello {escape(userDisplayName)}!</h2>
                </div>



            </section>
        );
    }
}
