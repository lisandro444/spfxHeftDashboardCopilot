import type { ITeamMember, ITeamMemberFormData } from '../../models/ITeamMember';
import { SharePointService } from '../../services/SharePointService';

type SharePointTeamMemberItem = {
  ID: number;
  Title?: string;
  memberId?: string;
  displayName?: string;
  role?: string;
  email?: string;
  active?: boolean;
};

export class TeamMemberService {
  private readonly sharePointService: SharePointService;
  private readonly listTitle: string = 'TeamMembers';
  private readonly selectFields: string[] = ['ID', 'Title', 'memberId', 'displayName', 'role', 'email', 'active'];

  constructor(sharePointService: SharePointService) {
    this.sharePointService = sharePointService;
  }

  public async getTeamMembers(): Promise<ITeamMember[]> {
    const items = await this.sharePointService.getListItems<SharePointTeamMemberItem>(this.listTitle, this.selectFields, 100);
    return items.map((item) => this.mapTeamMember(item));
  }

  public async addTeamMember(member: ITeamMemberFormData): Promise<ITeamMember> {
    const createdItem = await this.sharePointService.addListItem<SharePointTeamMemberItem>(
      this.listTitle,
      this.mapPayload(member),
      this.selectFields
    );

    return this.mapTeamMember(createdItem);
  }

  public async updateTeamMember(id: number, member: ITeamMemberFormData): Promise<void> {
    await this.sharePointService.updateListItem(this.listTitle, id, this.mapPayload(member));
  }

  public async deleteTeamMember(id: number): Promise<void> {
    await this.sharePointService.deleteListItem(this.listTitle, id);
  }

  private mapPayload(member: ITeamMemberFormData): Record<string, string | boolean> {
    return {
      Title: member.displayName.trim(),
      memberId: this.buildMemberId(member),
      displayName: member.displayName.trim(),
      role: member.role.trim(),
      email: member.email.trim().toLowerCase(),
      active: member.active
    };
  }

  private mapTeamMember(item: SharePointTeamMemberItem): ITeamMember {
    const fallbackTitle = item.Title || '';

    return {
      ID: item.ID,
      Title: fallbackTitle,
      memberId: item.memberId || '',
      displayName: item.displayName || fallbackTitle,
      role: item.role || '',
      email: item.email || '',
      active: item.active ?? false
    };
  }

  private buildMemberId(member: ITeamMemberFormData): string {
    const sourceValue = member.email.split('@')[0] || member.displayName;

    return sourceValue
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-');
  }
}
