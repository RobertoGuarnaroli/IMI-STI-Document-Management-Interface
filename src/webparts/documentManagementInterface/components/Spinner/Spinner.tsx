import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './Spinner.module.scss';

export const LoadingSpinner: React.FC = () => {
    return (
        <div className={styles.spinnerContainer}>
            <Spinner label="Loading..." size={SpinnerSize.large} />
        </div>
    );
}

