export interface IModalContainerProps {
    isOpen: boolean;
    title: string;
    children: React.ReactNode;
    onSave?: () => void;
    onCancel?: () => void;
    saving?: boolean;
    saveText?: string;
    cancelText?: string;
    width?: string;
}
