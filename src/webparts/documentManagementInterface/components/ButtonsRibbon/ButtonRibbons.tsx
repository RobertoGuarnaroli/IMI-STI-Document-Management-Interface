import * as React from 'react';
import { IButtonProps, PrimaryButton } from '@fluentui/react/lib/Button';
import styles from './ButtonsRibbon.module.scss';

export interface IButtonsRibbonButton {
    key: string;
    text: string;
    iconName?: string;
    onClick: () => void;
    disabled?: boolean;
    color?: string; // colore di sfondo opzionale
    style?: React.CSSProperties;
    buttonProps?: Partial<IButtonProps>;
}

export interface IButtonsRibbonProps {
    buttons: IButtonsRibbonButton[];
}

export const ButtonsRibbon: React.FC<IButtonsRibbonProps> = ({ buttons }) => (
    <div className={styles.buttonsRibbon}>
        {buttons.map(btn => (
            <PrimaryButton
                key={btn.key}
                text={btn.text}
                iconProps={btn.iconName ? { iconName: btn.iconName } : undefined}
                onClick={btn.onClick}
                disabled={btn.disabled}
                style={{
                    backgroundColor: btn.color || undefined,
                    ...btn.style
                }}
                {...btn.buttonProps}
            />
        ))}
    </div>
);
