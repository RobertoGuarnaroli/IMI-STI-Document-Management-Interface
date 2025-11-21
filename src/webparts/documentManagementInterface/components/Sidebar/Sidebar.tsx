import * as React from 'react';
import styles from './Sidebar.module.scss';
import type { ISidebarProps } from './ISidebarProps';
import projectConfig from '../../../../config/projectConfig';
import { Icon } from '@fluentui/react/lib/Icon';

export const Sidebar: React.FC<ISidebarProps> = ({ isVisible, selectedTab, onTabChange, onClose }) => {
    const sidebarClass = `${styles.sidebar} ${isVisible ? styles.visible : styles.hidden}`;

    return (
        <div className={sidebarClass}>
            {/* Close button per mobile */}
            {onClose && (
                <button
                    className={styles.closeButton}
                    onClick={onClose}
                    aria-label="Chiudi menu"
                >
                    <Icon iconName="Cancel" />
                </button>
            )}

            <ul className={styles.tabList}>
                {projectConfig.lists.map((list) => (
                    <li
                        key={list.id}
                        className={
                            selectedTab === list.id
                                ? `${styles.tabItem} ${styles.selected}`
                                : styles.tabItem
                        }
                        onClick={() => onTabChange(list.id)}
                    >
                        <Icon iconName={list.icon || 'Page'} className={styles.icon} />
                        {list.name}
                    </li>
                ))}
            </ul>
        </div>
    );
};