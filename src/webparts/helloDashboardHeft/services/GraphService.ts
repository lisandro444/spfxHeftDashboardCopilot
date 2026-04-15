import { graphfi, SPFx as GraphSPFx, GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/groups';
import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { PageContext } from '@microsoft/sp-page-context';

export interface IUserProfile {
  displayName: string;
  mail: string;
  mobilePhone: string;
  officeLocation: string;
  [key: string]: unknown;
}

export interface IUserInfo {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  [key: string]: unknown;
}

export class GraphService {
  private graph: GraphFI;

  constructor(pageContext: PageContext) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    this.graph = graphfi().using(GraphSPFx({ pageContext } as any));
  }

  /**
   * Get current logged-in user information
   */
  public async getCurrentUser(): Promise<IUserInfo> {
    try {
      const user = await this.graph.me.select('id', 'displayName', 'mail', 'userPrincipalName')();
      return user as IUserInfo;
    } catch (err) {
      console.error('Error fetching current user:', err);
      throw err;
    }
  }

  /**
   * Get current user detailed profile
   */
  public async getUserProfile(): Promise<IUserProfile> {
    try {
      const profile = await this.graph.me.select(
        'displayName',
        'mail',
        'mobilePhone',
        'officeLocation',
        'userPrincipalName',
        'jobTitle',
        'department'
      )();
      return profile as IUserProfile;
    } catch (err) {
      console.error('Error fetching user profile:', err);
      throw err;
    }
  }

  /**
   * Get user by email
   */
  public async getUserByEmail(email: string): Promise<IUserInfo> {
    try {
      const users = await this.graph.users.select('id', 'displayName', 'mail', 'userPrincipalName')
        .filter(`mail eq '${email}'`)();
      
      if (users.length > 0) {
        return users[0] as IUserInfo;
      }
      throw new Error(`User ${email} not found`);
    } catch (err) {
      console.error(`Error fetching user ${email}:`, err);
      throw err;
    }
  }

  /**
   * Get all members from a group
   */
  public async getGroupMembers(groupId: string): Promise<IUserInfo[]> {
    try {
      const members = await this.graph.groups.getById(groupId)();
      return members as IUserInfo[];
    } catch (err) {
      console.error(`Error fetching group members for ${groupId}:`, err);
      throw err;
    }
  }

  /**
   * Get current user's groups
   */
  public async getCurrentUserGroups(): Promise<DirectoryObject[]> {
    try {
      const groups = await this.graph.me.memberOf();
      return groups;
    } catch (err) {
      console.error('Error fetching user groups:', err);
      throw err;
    }
  }

  /**
   * Search users by display name
   */
  public async searchUsers(searchTerm: string): Promise<IUserInfo[]> {
    try {
      const users = await this.graph.users
        .select('id', 'displayName', 'mail', 'userPrincipalName')
        .filter(`startswith(displayName, '${searchTerm}')`)
        .top(10)();
      
      return users as IUserInfo[];
    } catch (err) {
      console.error(`Error searching users with term "${searchTerm}":`, err);
      throw err;
    }
  }
}
