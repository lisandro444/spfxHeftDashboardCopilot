import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  type IColumn,
  IconButton,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  SelectionMode,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
  Toggle
} from '@fluentui/react';
import type { ITeamMember } from '../../models/ITeamMember';
import { SharePointService } from '../../services/SharePointService';
import styles from './TeamMemberManager.module.scss';
import type { ITeamMemberManagerProps } from './ITeamMemberManagerProps';
import type { ITeamMemberManagerState } from './ITeamMemberManagerState';
import { TeamMemberService } from './TeamMemberService';

export default class TeamMemberManager extends React.Component<ITeamMemberManagerProps, ITeamMemberManagerState> {
  private readonly teamMemberService: TeamMemberService;

  constructor(props: ITeamMemberManagerProps) {
    super(props);

    this.teamMemberService = new TeamMemberService(new SharePointService(props.context.pageContext));
    this.state = {
      teamMembers: [],
      loading: true,
      saving: false,
      error: '',
      success: '',
      isPanelOpen: false,
      editingMember: undefined,
      formData: {
        displayName: '',
        role: '',
        email: '',
        active: true
      }
    };
  }

  public componentDidMount(): void {
    this.loadMembers().catch((err) => {
      this.setState({
        loading: false,
        error: `Error loading team members: ${err instanceof Error ? err.message : 'Unknown error'}`
      });
    });
  }

  public render(): React.ReactElement<ITeamMemberManagerProps> {
    const { teamMembers, loading, error, success, isPanelOpen, editingMember, formData, saving } = this.state;

    const columns: IColumn[] = [
      {
        key: 'name',
        name: 'Name',
        minWidth: 180,
        isResizable: true,
        onRender: (item: ITeamMember) => <span className={styles.memberName}>{item.displayName || item.Title}</span>
      },
      {
        key: 'role',
        name: 'Role',
        fieldName: 'role',
        minWidth: 140,
        isResizable: true
      },
      {
        key: 'email',
        name: 'Email',
        fieldName: 'email',
        minWidth: 220,
        isResizable: true
      },
      {
        key: 'status',
        name: 'Status',
        minWidth: 110,
        isResizable: true,
        onRender: (item: ITeamMember) => (
          <span className={item.active ? styles.statusBadgeActive : styles.statusBadgeInactive}>
            {item.active ? 'Active' : 'Inactive'}
          </span>
        )
      },
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 110,
        maxWidth: 120,
        onRender: (item: ITeamMember) => (
          <div className={styles.actionButtons}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit"
              onClick={() => this.openEditPanel(item)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete"
              onClick={() => {
                this.handleDelete(item).catch((err) => {
                  console.error('Unexpected delete error:', err);
                });
              }}
            />
          </div>
        )
      }
    ];

    return (
      <div className={styles.teamMembersShell}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="end" className={styles.pageHeader}>
          <div>
            <Text variant="xxLarge" block className={styles.pageTitle}>
              Team Members
            </Text>
            <Text variant="medium" className={styles.pageSubtitle}>
              Manage your team who can post status updates
            </Text>
          </div>
          <PrimaryButton
            text="Add Team Member"
            iconProps={{ iconName: 'Add' }}
            onClick={this.openNewPanel}
            className={styles.addButton}
          />
        </Stack>

        {error && (
          <div className={styles.messageBar}>
            <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
          </div>
        )}

        {success && (
          <div className={styles.messageBar}>
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
              {success}
            </MessageBar>
          </div>
        )}

        <div className={styles.tableCard}>
          {loading ? (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Loading team members..." />
            </div>
          ) : teamMembers.length === 0 ? (
            <div className={styles.emptyState}>
              <Text>No team members found. Create the first one to get started.</Text>
            </div>
          ) : (
            <DetailsList
              items={teamMembers}
              columns={columns}
              selectionMode={SelectionMode.none}
              className={styles.membersList}
            />
          )}
        </div>

        <Panel
          isOpen={isPanelOpen}
          onDismiss={this.closePanel}
          type={PanelType.smallFixedFar}
          headerText={editingMember ? 'Edit Team Member' : 'Add Team Member'}
          closeButtonAriaLabel="Close"
        >
          <Stack tokens={{ childrenGap: 14 }}>
            <TextField
              label="Name"
              required
              value={formData.displayName}
              onChange={(_, value) => this.updateFormField('displayName', value || '')}
            />
            <TextField
              label="Role"
              required
              value={formData.role}
              onChange={(_, value) => this.updateFormField('role', value || '')}
            />
            <TextField
              label="Email"
              type="email"
              required
              value={formData.email}
              onChange={(_, value) => this.updateFormField('email', value || '')}
            />
            <Toggle
              label="Status"
              checked={formData.active}
              onText="Active"
              offText="Inactive"
              onChange={(_, checked) => this.updateFormField('active', !!checked)}
            />

            <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.panelFooter}>
              <PrimaryButton
                text={saving ? 'Saving...' : 'Save'}
                onClick={() => {
                  this.handleSave().catch((err) => {
                    console.error('Unexpected save error:', err);
                  });
                }}
                disabled={saving}
              />
              <DefaultButton text="Cancel" onClick={this.closePanel} disabled={saving} />
            </Stack>
          </Stack>
        </Panel>
      </div>
    );
  }

  private readonly loadMembers = async (): Promise<void> => {
    this.setState({ loading: true, error: '' });

    const items = await this.teamMemberService.getTeamMembers();
    const sortedItems = [...items].sort((a, b) => (a.displayName || a.Title).localeCompare(b.displayName || b.Title));

    this.setState({
      teamMembers: sortedItems,
      loading: false
    });
  };

  private readonly openNewPanel = (): void => {
    this.setState({
      editingMember: undefined,
      isPanelOpen: true,
      error: '',
      success: '',
      formData: {
        displayName: '',
        role: '',
        email: '',
        active: true
      }
    });
  };

  private openEditPanel(member: ITeamMember): void {
    this.setState({
      editingMember: member,
      isPanelOpen: true,
      error: '',
      success: '',
      formData: {
        displayName: member.displayName || member.Title,
        role: member.role,
        email: member.email,
        active: member.active
      }
    });
  }

  private readonly closePanel = (): void => {
    if (!this.state.saving) {
      this.setState({ isPanelOpen: false });
    }
  };

  private updateFormField(field: 'displayName' | 'role' | 'email' | 'active', value: string | boolean): void {
    this.setState((currentState) => ({
      formData: {
        ...currentState.formData,
        [field]: value
      }
    }));
  }

  private validateForm(): string | undefined {
    const { formData } = this.state;

    if (!formData.displayName.trim()) {
      return 'Name is required.';
    }

    if (!formData.role.trim()) {
      return 'Role is required.';
    }

    if (!formData.email.trim()) {
      return 'Email is required.';
    }

    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(formData.email)) {
      return 'Please enter a valid email address.';
    }

    return undefined;
  }

  private async handleSave(): Promise<void> {
    const { editingMember, formData } = this.state;
    const validationError = this.validateForm();

    if (validationError) {
      this.setState({ error: validationError });
      return;
    }

    try {
      this.setState({ saving: true, error: '', success: '' });

      if (editingMember) {
        await this.teamMemberService.updateTeamMember(editingMember.ID, formData);
        this.setState({ success: 'Team member updated successfully.' });
      } else {
        await this.teamMemberService.addTeamMember(formData);
        this.setState({ success: 'Team member created successfully.' });
      }

      this.setState({
        isPanelOpen: false,
        editingMember: undefined,
        formData: {
          displayName: '',
          role: '',
          email: '',
          active: true
        }
      });

      await this.loadMembers();
    } catch (err) {
      this.setState({
        error: `Error saving team member: ${err instanceof Error ? err.message : 'Unknown error'}`
      });
    } finally {
      this.setState({ saving: false });
    }
  }

  private async handleDelete(member: ITeamMember): Promise<void> {
    const memberName = member.displayName || member.Title;
    const confirmed = window.confirm(`Are you sure you want to delete ${memberName}?`);

    if (!confirmed) {
      return;
    }

    try {
      this.setState({ error: '', success: '' });
      await this.teamMemberService.deleteTeamMember(member.ID);
      this.setState({ success: 'Team member deleted successfully.' });
      await this.loadMembers();
    } catch (err) {
      this.setState({
        error: `Error deleting team member: ${err instanceof Error ? err.message : 'Unknown error'}`
      });
    }
  }
}
