import * as React from 'react';
import styles from './PersonalDashboard.module.scss';
import { IPersonalDashboardProps } from './IPersonalDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { useDispatch } from 'react-redux';
import { fetchEmails } from '../store/slices/mainSlice';
import type {} from 'redux-thunk/extend-redux';

const PersonalDashboard: React.FC<IPersonalDashboardProps> = (props: IPersonalDashboardProps) => {

    const dispatch = useDispatch();

    React.useEffect(() => {
        dispatch(fetchEmails(props.serviceScope));
        debugger;
    }, [dispatch]);


    return (
        <section className={`${styles.personalDashboard} ${props.hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.welcome}>
                <h2>Hello {escape(props.userDisplayName)}!</h2>
            </div>


        </section>
    );
}

export default PersonalDashboard;


