import type { ITeamMember, ITeamMemberFormData } from '../../models/ITeamMember';

export interface ITeamMemberManagerState {
  teamMembers: ITeamMember[];
  loading: boolean;
  saving: boolean;
  error: string;
  success: string;
  isPanelOpen: boolean;
  editingMember?: ITeamMember;
  formData: ITeamMemberFormData;
}
