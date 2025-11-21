import { IButtonProps } from "@fluentui/react";

export interface IButtonsRibbonProps{
    onNewClick: () => void;
    buttonText?: string;
}

export interface IButtonsRibbonButton {
    key: string;
    text: string;
    iconName?: string;
    onClick: () => void;
    disabled?: boolean;
    color?: string; // colore di sfondo opzionale
    style?: React.CSSProperties;
    buttonProps?: Partial<IButtonProps>;
    visible?: boolean;
}