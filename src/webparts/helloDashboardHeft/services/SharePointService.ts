import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PageContext } from '@microsoft/sp-page-context';
import type { ITeamMember, ITeamMemberFormData } from '../models/ITeamMember';

type SharePointTeamMemberItem = {
  ID: number;
  Title?: string;
  memberId?: string;
  displayName?: string;
  role?: string;
  email?: string;
  active?: boolean;
};

export class SharePointService {
  private readonly sp: SPFI;
  private readonly listTitle: string = 'TeamMembers';

  constructor(pageContext: PageContext) {
    this.sp = spfi().using(SPFx({ pageContext }));
  }

  public async getTeamMembers(): Promise<ITeamMember[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('ID', 'Title', 'memberId', 'displayName', 'role', 'email', 'active')
        .top(100)();

      return (items as SharePointTeamMemberItem[]).map((item) => this.mapTeamMember(item));
    } catch (err) {
      console.error('Error fetching team members:', err);
      throw err;
    }
  }

  public async addTeamMember(member: ITeamMemberFormData): Promise<ITeamMember> {
    try {
      const payload = this.mapPayload(member);
      const result = await this.sp.web.lists.getByTitle(this.listTitle).items.add(payload);
      const createdItem = await result.item.select('ID', 'Title', 'memberId', 'displayName', 'role', 'email', 'active')();

      return this.mapTeamMember(createdItem as SharePointTeamMemberItem);
    } catch (err) {
      console.error('Error creating team member:', err);
      throw err;
    }
  }

  public async updateTeamMember(id: number, member: ITeamMemberFormData): Promise<void> {
    try {
      const payload = this.mapPayload(member);
      await this.sp.web.lists.getByTitle(this.listTitle).items.getById(id).update(payload);
    } catch (err) {
      console.error('Error updating team member:', err);
      throw err;
    }
  }

  public async deleteTeamMember(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listTitle).items.getById(id).delete();
    } catch (err) {
      console.error('Error deleting team member:', err);
      throw err;
    }
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
