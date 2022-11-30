import * as React from 'react';
import styles from './PersonalDashboard.module.scss';
import { IPersonalDashboardProps } from './IPersonalDashboardProps';
import { useDispatch, useSelector } from 'react-redux';
import { fetchEmails, getEmails, getEmailsLoaded, getError } from '../store/slices/mainSlice';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import type { } from 'redux-thunk/extend-redux';
import IEmail from '../models/IEmail';


const PersonalDashboard: React.FC<IPersonalDashboardProps> = (props: IPersonalDashboardProps) => {

    const dispatch = useDispatch();
    const loaded = useSelector(getEmailsLoaded);
    const emails = useSelector(getEmails);
    const error = useSelector(getError);

    React.useEffect(() => {
        dispatch(fetchEmails(props.serviceScope));
    }, [dispatch, props.serviceScope]);


    const onRenderCell = (item: IEmail, index: number | undefined): JSX.Element => {
        return <Link href={item.webLink} className={styles.message} target='_blank'>
            <div className={styles.from}>{item.from.emailAddress.name || item.from.emailAddress.address}</div>
            <div className={styles.subject}>{item.subject}</div>
            <div className={styles.date}>{(new Date(item.receivedDateTime).toLocaleDateString())}</div>
            <div className={styles.preview}>{item.bodyPreview}</div>
        </Link>;
    }

    return (
        <section className={`${styles.personalDashboard} ${props.hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.personalEmail}>
                {!loaded &&
                    <Spinner label={"Loading"} size={SpinnerSize.large} />
                }
                {emails && emails.length > 0 ? (
                    <div>
                        <List items={emails}
                            onRenderCell={onRenderCell} className={styles.list} />
                        <Link href='https://outlook.office.com/owa/' target='_blank' className={styles.viewAll}>ViewAll</Link>
                    </div>
                ) : (
                    loaded && (
                        error ?
                            <span className={styles.error}>{error}</span> :
                            <span className={styles.noMessages}>No Messages</span>
                    )
                )
                }
            </div>
        </section>
    );
}

export default PersonalDashboard;


