export interface ISidebarProps {
    isVisible: boolean;
    selectedTab: string;
    onTabChange: (tabId: string) => void;
    onClose?: () => void; // Opzionale per chiudere la sidebar su mobile
}