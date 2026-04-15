import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PageContext } from '@microsoft/sp-page-context';
import type { ITeamMember } from '../models/ITeamMember';

export class SharePointService {
  private sp: SPFI;

  constructor(pageContext: PageContext) {
    this.sp = spfi().using(SPFx({ pageContext }));
  }

  public async getTeamMembers(): Promise<ITeamMember[]> {
    try {
      // Log the current web URL for debugging
      console.log('Current web URL:', this.sp.web.toUrl());

      const items = await this.sp.web.lists
        .getByTitle('TeamMembers')
        .items.select('ID', 'Title', 'memberId', 'displayName', 'role', 'email', 'active')
        .top(100)();

      return items as ITeamMember[];
    } catch (err) {
      console.error('Error fetching team members:', err);
      throw err;
    }
  }
}
