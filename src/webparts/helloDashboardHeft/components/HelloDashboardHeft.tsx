import * as React from 'react';
import { DetailsList, IColumn, SelectionMode } from '@fluentui/react';
import styles from './HelloDashboardHeft.module.scss';
import type { IHelloDashboardHeftProps } from './IHelloDashboardHeftProps';
import type { IHelloDashboardHeftState } from './IHelloDashboardHeftState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SharePointService } from '../services/SharePointService';
import { ITeamMember } from '../models/ITeamMember';

export default class HelloDashboardHeft extends React.Component<IHelloDashboardHeftProps, IHelloDashboardHeftState> {
  private sharePointService: SharePointService;

  constructor(props: IHelloDashboardHeftProps) {
    super(props);
    this.state = {
      teamMembers: [],
      loading: true,
      error: ""
    };
    this.sharePointService = new SharePointService(this.props.context.pageContext);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadTeamMembers();
  }

  private async loadTeamMembers(): Promise<void> {
    try {
      const members = await this.sharePointService.getTeamMembers();
      this.setState({
        teamMembers: members,
        loading: false,
        error: ""
      });
    } catch (err) {
      this.setState({
        teamMembers: [],
        loading: false,
        error: `Error loading team members: ${err instanceof Error ? err.message : 'Unknown error'}`
      });
    }
  }

  public render(): React.ReactElement<IHelloDashboardHeftProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const { teamMembers, loading, error } = this.state;

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'ID',
        fieldName: 'ID',
        minWidth: 50,
        maxWidth: 80,
        isResizable: true
      },
      {
        key: 'column2',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'column3',
        name: 'Member ID',
        fieldName: 'memberId',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'column4',
        name: 'Display Name',
        fieldName: 'displayName',
        minWidth: 120,
        maxWidth: 200,
        isResizable: true
      },
      {
        key: 'column5',
        name: 'Role',
        fieldName: 'role',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'column6',
        name: 'Email',
        fieldName: 'email',
        minWidth: 150,
        maxWidth: 250,
        isResizable: true
      },
      {
        key: 'column7',
        name: 'Active',
        fieldName: 'active',
        minWidth: 70,
        maxWidth: 100,
        isResizable: true,
        onRender: (item: ITeamMember) => (
          <span>{item.active ? 'Yes' : 'No'}</span>
        )
      }
    ];

    return (
      <section className={`${styles.helloDashboardHeft} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
          
          <div style={{ marginTop: '20px' }}>
            <h3>Team Members</h3>
            {loading && <p>Loading team members...</p>}
            {error && <p style={{ color: 'red' }}>{error}</p>}
            {!loading && !error && (
              <DetailsList
                items={teamMembers}
                columns={columns}
                selectionMode={SelectionMode.none}
                layoutMode={0}
              />
            )}
          </div>
        </div>
      </section>
    );
  }
}
