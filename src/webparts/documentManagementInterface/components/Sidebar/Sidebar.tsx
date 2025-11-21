import * as React from 'react';
import styles from './Sidebar.module.scss';
import type { ISidebarProps } from './ISidebarProps';
import projectConfig from '../../../../config/projectConfig';
import { Icon } from '@fluentui/react/lib/Icon';

export const Sidebar: React.FC<ISidebarProps> = ({ isVisible, selectedTab, onTabChange }) => {
    const sidebarClass = `${styles.sidebar} ${isVisible ? styles.visible : styles.hidden}`;
    return (
        <div className={sidebarClass}>
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
                        <Icon iconName={list.icon || 'Page'} className={styles.icon}/>
                        {list.name}
                    </li>
                ))}
            </ul>
        </div>
    );
};
