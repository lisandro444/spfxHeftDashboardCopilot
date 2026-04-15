export interface ITeamMember {
  ID: number;
  Title: string;
  memberId: string;
  displayName: string;
  role: string;
  email: string;
  active: boolean;
}

export interface ITeamMemberFormData {
  displayName: string;
  role: string;
  email: string;
  active: boolean;
}
