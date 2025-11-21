export interface ISidebarProps {
    isVisible: boolean;
    selectedTab: string;
    onTabChange: (tabId: string) => void;
}