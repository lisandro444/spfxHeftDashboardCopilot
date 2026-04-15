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
import styles from './HelloDashboardHeft.module.scss';
import type { IHelloDashboardHeftProps } from './IHelloDashboardHeftProps';
import type { ITeamMember, ITeamMemberFormData } from '../models/ITeamMember';
import { SharePointService } from '../services/SharePointService';

const emptyFormData: ITeamMemberFormData = {
  displayName: '',
  role: '',
  email: '',
  active: true
};

const TeamMemberManager: React.FC<IHelloDashboardHeftProps> = (props) => {
  const serviceRef = React.useRef(new SharePointService(props.context.pageContext));
  const [members, setMembers] = React.useState<ITeamMember[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [success, setSuccess] = React.useState<string>('');
  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  const [editingMember, setEditingMember] = React.useState<ITeamMember | undefined>(undefined);
  const [formData, setFormData] = React.useState<ITeamMemberFormData>(emptyFormData);

  const loadMembers = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError('');

      const items = await serviceRef.current.getTeamMembers();
      const sortedItems = [...items].sort((a, b) =>
        (a.displayName || a.Title).localeCompare(b.displayName || b.Title)
      );

      setMembers(sortedItems);
    } catch (err) {
      setError(`Error loading team members: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    loadMembers().catch((err) => {
      setError(`Error loading team members: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setLoading(false);
    });
  }, [loadMembers]);

  const resetForm = (): void => {
    setEditingMember(undefined);
    setFormData(emptyFormData);
  };

  const openNewPanel = (): void => {
    resetForm();
    setSuccess('');
    setError('');
    setIsPanelOpen(true);
  };

  const openEditPanel = (member: ITeamMember): void => {
    setEditingMember(member);
    setFormData({
      displayName: member.displayName || member.Title,
      role: member.role,
      email: member.email,
      active: member.active
    });
    setSuccess('');
    setError('');
    setIsPanelOpen(true);
  };

  const closePanel = (): void => {
    if (!saving) {
      setIsPanelOpen(false);
    }
  };

  const updateFormField = (field: keyof ITeamMemberFormData, value: string | boolean): void => {
    setFormData((current) => ({
      ...current,
      [field]: value
    }));
  };

  const validateForm = (): string | undefined => {
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
  };

  const handleSave = async (): Promise<void> => {
    const validationError = validateForm();
    if (validationError) {
      setError(validationError);
      return;
    }

    try {
      setSaving(true);
      setError('');
      setSuccess('');

      if (editingMember) {
        await serviceRef.current.updateTeamMember(editingMember.ID, formData);
        setSuccess('Team member updated successfully.');
      } else {
        await serviceRef.current.addTeamMember(formData);
        setSuccess('Team member created successfully.');
      }

      setIsPanelOpen(false);
      resetForm();
      await loadMembers();
    } catch (err) {
      setError(`Error saving team member: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (member: ITeamMember): Promise<void> => {
    const memberName = member.displayName || member.Title;
    const confirmed = window.confirm(`Are you sure you want to delete ${memberName}?`);

    if (!confirmed) {
      return;
    }

    try {
      setError('');
      setSuccess('');
      await serviceRef.current.deleteTeamMember(member.ID);
      setSuccess('Team member deleted successfully.');
      await loadMembers();
    } catch (err) {
      setError(`Error deleting team member: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

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
            onClick={() => openEditPanel(item)}
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Delete"
            ariaLabel="Delete"
            onClick={() => {
              handleDelete(item).catch((err) => {
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
          onClick={openNewPanel}
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
        ) : members.length === 0 ? (
          <div className={styles.emptyState}>
            <Text>No team members found. Create the first one to get started.</Text>
          </div>
        ) : (
          <DetailsList
            items={members}
            columns={columns}
            selectionMode={SelectionMode.none}
            className={styles.membersList}
          />
        )}
      </div>

      <Panel
        isOpen={isPanelOpen}
        onDismiss={closePanel}
        type={PanelType.smallFixedFar}
        headerText={editingMember ? 'Edit Team Member' : 'Add Team Member'}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 14 }}>
          <TextField
            label="Name"
            required
            value={formData.displayName}
            onChange={(_, value) => updateFormField('displayName', value || '')}
          />
          <TextField
            label="Role"
            required
            value={formData.role}
            onChange={(_, value) => updateFormField('role', value || '')}
          />
          <TextField
            label="Email"
            type="email"
            required
            value={formData.email}
            onChange={(_, value) => updateFormField('email', value || '')}
          />
          <Toggle
            label="Status"
            checked={formData.active}
            onText="Active"
            offText="Inactive"
            onChange={(_, checked) => updateFormField('active', !!checked)}
          />

          <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.panelFooter}>
            <PrimaryButton
              text={saving ? 'Saving...' : 'Save'}
              onClick={() => {
                handleSave().catch((err) => {
                  console.error('Unexpected save error:', err);
                });
              }}
              disabled={saving}
            />
            <DefaultButton text="Cancel" onClick={closePanel} disabled={saving} />
          </Stack>
        </Stack>
      </Panel>
    </div>
  );
};

export default TeamMemberManager;
