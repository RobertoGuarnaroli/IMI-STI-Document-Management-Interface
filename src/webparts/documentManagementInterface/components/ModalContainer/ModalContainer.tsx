import * as React from 'react';
import styles from './ModalContainer.module.scss';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import type { IModalContainerProps } from './IModalContainerProps';

export const ModalContainer: React.FC<IModalContainerProps> = ({
    isOpen,
    title,
    children,
    onSave,
    onCancel,
    saving = false,
    saveText = 'Salva',
    cancelText = 'Annulla',
    width = '480px',
}) => {
    if (!isOpen) return null;
    return (
        <>
            <div className={styles.modalOverlay} onClick={onCancel} />
            <div className={styles.modalWrapper}>
                <div className={styles.modalContainer} style={{ width }}>
                    <h2>{title}</h2>
                    {children}
                    <div className={styles.modalActions}>
                        {onSave && (
                            <PrimaryButton
                                text={saving ? 'Salvataggio...' : saveText}
                                onClick={onSave}
                                disabled={saving}
                                styles={{
                                    root: {
                                        backgroundColor: '#441f53',
                                        border: 'none',
                                        color: '#ffffff',
                                        fontWeight: 600,
                                        padding: '6px 12px',
                                        borderRadius: 4,
                                        fontSize: 14,
                                        marginRight: 12,
                                    },
                                    rootHovered: {
                                        backgroundColor: '#5a2a6b',
                                    },
                                }}
                            />
                        )}
                        <DefaultButton text={cancelText} onClick={onCancel} />
                    </div>
                </div>
            </div>
        </>
    );
};
