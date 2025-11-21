import * as React from 'react';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import styles from './ButtonsRibbon.module.scss';
import { IButtonsRibbonButton } from './IButtonsRibbonProps';


export interface IButtonsRibbonProps {
    buttons: IButtonsRibbonButton[];
}

export const ButtonsRibbon: React.FC<IButtonsRibbonProps> = ({ buttons }) => (
    <div className={styles.buttonsRibbon}>
        {buttons.map(btn =>
            btn.visible !== false && (
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
            )
        )}
    </div>
);
