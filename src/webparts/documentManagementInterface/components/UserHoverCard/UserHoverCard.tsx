
import * as React from 'react';
import styles from './UserHoverCard.module.scss';
import { IUserHoverCardProps } from './IUserHoverCardProps';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { IconButton } from '@fluentui/react/lib/Button';
import { HoverCard, HoverCardType, IPlainCardProps } from '@fluentui/react/lib/HoverCard';
import { UsersService } from '../../../services/SharePointService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// Presentational component: mostra i dettagli utente
export const UserHoverCardDetails: React.FC<IUserHoverCardProps> = ({ user, loading }) => {
  if (loading) {
    return <div className={styles.loading}>Loading...</div>;
  }
  if (!user) {
    return <div className={styles.empty}>No user data</div>;
  }
  if (user.displayName === 'User not found') {
    return <div className={styles.empty}>User not found</div>;
  }
  let firstName = '';
  let lastName = '';
  if (user.displayName) {
    const parts = user.displayName.split(' ');
    if (parts.length > 1) {
      firstName = parts[0];
      lastName = parts.slice(1).join(' ');
    } else {
      firstName = user.displayName;
    }
  }
  return (
    <div className={styles.card}>
      <Persona
        text={user.displayName}
        imageUrl={user.pictureUrl}
        size={PersonaSize.size40}
        styles={{ root: { marginBottom: 12 } }}
      />
      <div className={styles.details}>
        <div className={styles.row}><span className={styles.label}>First name:</span><span className={styles.value}>{firstName}</span></div>
        <div className={styles.row}><span className={styles.label}>Last name:</span><span className={styles.value}>{lastName}</span></div>
        <div className={styles.row}><span className={styles.label}>Job title:</span><span className={styles.value}>{user.jobTitle || ''}</span></div>
        <div className={styles.row}><span className={styles.label}>Department:</span><span className={styles.value}>{user.department || ''}</span></div>
        <div className={styles.row}><span className={styles.label}>Office:</span><span className={styles.value}>{user.officeLocation || ''}</span></div>
        <div className={styles.row}><span className={styles.label}>Office phone:</span><span className={styles.value}>{user.businessPhones && user.businessPhones.length > 0 ? user.businessPhones[0] : ''}</span></div>
        <div className={styles.row}><span className={styles.label}>Mobile phone:</span><span className={styles.value}>{user.mobilePhone || ''}</span></div>
        <div className={styles.row}>
          <span className={styles.label}>Contatta:</span>
          <span className={styles.value + ' ' + styles.gapValue}>
            {user.mail && (
              <IconButton
                iconProps={{ iconName: 'Mail' }}
                title="Invia email"
                ariaLabel="Invia email"
                href={`mailto:${user.mail}`}
                target="_blank"
                styles={{ root: { padding: 0, height: 24, width: 24 } }}
              />
            )}
            {user.mail && (
              <IconButton
                iconProps={{ iconName: 'TeamsLogo' }}
                title="Chatta su Teams"
                ariaLabel="Chatta su Teams"
                href={`https://teams.microsoft.com/l/chat/0/0?users=${user.mail}`}
                target="_blank"
                styles={{ root: { padding: 0, height: 24, width: 24 } }}
              />
            )}
          </span>
        </div>
      </div>
    </div>
  );
};

// Wrapper base: mostra Persona e card statica (dati già disponibili)
export interface IUserHoverCardWrapperProps {
  user: {
    id: string;
    displayName: string;
    mail?: string;
    pictureUrl?: string;
    jobTitle?: string;
    department?: string;
    officeLocation?: string;
    businessPhones?: string[];
    mobilePhone?: string;
  };
}

export const UserHoverCardWrapper: React.FC<IUserHoverCardWrapperProps> = ({ user }) => {
  const plainCardProps: IPlainCardProps = {
    onRenderPlainCard: () => <UserHoverCardDetails user={user} />,
  };
  return (
    <HoverCard
      type={HoverCardType.plain}
      plainCardProps={plainCardProps}
      instantOpenOnClick={false}
    >
      <Persona
        text={user.displayName}
        imageUrl={user.pictureUrl}
        size={PersonaSize.size32}
      />
    </HoverCard>
  );
};

// Smart: fa fetch da Graph su hover
export interface IUserHoverCardSmartProps {
  email: string;
  displayName: string;
  pictureUrl?: string;
  context: WebPartContext;
}

export const UserHoverCardSmart: React.FC<IUserHoverCardSmartProps> = ({ email, displayName, pictureUrl, context }) => {
  const [user, setUser] = React.useState<any>(null);
  const [loading, setLoading] = React.useState(false);

  const fetchUser = React.useCallback(async () => {
    setLoading(true);
    try {
      const usersService = new UsersService(context);
      const userData = await usersService.getUserByEmail(email);
      let userPic = pictureUrl;
      try {
        const userProfileService = new (await import('../../../services/SharePointService')).UserProfileService(context);
        if (userData && userData.id) {
          userPic = await userProfileService.getUserProfilePicture(userData.id);
        } else {
          userPic = await userProfileService.getUserProfilePicture(email);
        }
      } catch {}
      setUser({
        id: userData?.id || email,
        displayName: userData?.displayName || displayName,
        mail: userData?.mail || userData?.userPrincipalName || email,
        jobTitle: userData?.jobTitle,
        department: userData?.department,
        officeLocation: userData?.officeLocation,
        businessPhones: userData?.businessPhones,
        mobilePhone: userData?.mobilePhone,
        pictureUrl: userPic
      });
    } catch {
      setUser({
        id: email,
        displayName,
        mail: email,
        pictureUrl
      });
    } finally {
      setLoading(false);
    }
  }, [email, displayName, pictureUrl, context]);

  const plainCardProps: IPlainCardProps = {
    onRenderPlainCard: () => <UserHoverCardDetails user={user} loading={loading} />,
  };

  return (
    <HoverCard
      type={HoverCardType.plain}
      plainCardProps={plainCardProps}
      onCardVisible={fetchUser}
      instantOpenOnClick={false}
    >
      <Persona
        text={displayName}
        imageUrl={pictureUrl}
        size={PersonaSize.size32}
      />
    </HoverCard>
  );
};
