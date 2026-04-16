import * as React from 'react';
import styles from './HelloDashboardHeft.module.scss';
import type { IHelloDashboardHeftProps } from './IHelloDashboardHeftProps';
import TeamMemberManager from './teamMembers/TeamMemberManager';

export default class HelloDashboardHeft extends React.Component<IHelloDashboardHeftProps> {
  public render(): React.ReactElement<IHelloDashboardHeftProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.helloDashboardHeft} ${hasTeamsContext ? styles.teams : ''}`}>
        <TeamMemberManager {...this.props} />
      </section>
    );
  }
}
