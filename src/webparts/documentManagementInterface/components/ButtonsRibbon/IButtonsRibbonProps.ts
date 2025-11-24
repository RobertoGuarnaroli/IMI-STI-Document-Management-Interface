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
    color?: string;
    border?: string;
    buttonProps?: Partial<IButtonProps>;
    visible?: boolean;
}