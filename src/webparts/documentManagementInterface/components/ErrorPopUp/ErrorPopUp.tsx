import * as React from 'react';
import styles from './ErrorPopUp.module.scss';

export interface ErrorPopUpProps {
    message: string;
    onClose?: () => void;
    duration?: number; // ms
}

export const ErrorPopUp: React.FC<ErrorPopUpProps> = ({ message, onClose, duration = 4000 }) => {
    React.useEffect(() => {
        if (onClose && duration > 0) {
            const timer = setTimeout(onClose, duration);
            return () => clearTimeout(timer);
        }
    }, [onClose, duration]);

    if (!message) return null;
    return (
        <div className={styles.errorPopUp}>
            <span>{message}</span>
            {onClose && (
                <button className={styles.closeBtn} onClick={onClose} title="Chiudi">×</button>
            )}
        </div>
    );
};
