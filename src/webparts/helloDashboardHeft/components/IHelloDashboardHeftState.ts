import type { ITeamMember } from '../models/ITeamMember';

export interface IHelloDashboardHeftState {
  teamMembers: ITeamMember[];
  loading: boolean;
  error: string;
}
