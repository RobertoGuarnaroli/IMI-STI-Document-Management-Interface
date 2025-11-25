import * as React from 'react';
import { DefaultButton, Stack } from '@fluentui/react';
import styles from './FileUpload.module.scss';

export interface FileUploadProps {
  onUpload: (file: File) => void;
  onUploadMultiple?: (files: File[]) => void;
  disabled?: boolean;
}

export const FileUpload: React.FC<FileUploadProps> = ({ onUpload, onUploadMultiple, disabled }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = React.useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      if (onUploadMultiple && e.target.files.length > 1) {
        onUploadMultiple(Array.from(e.target.files));
      } else {
        onUpload(e.target.files[0]);
      }
      e.target.value = '';
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (disabled) return;
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      if (onUploadMultiple && e.dataTransfer.files.length > 1) {
        onUploadMultiple(Array.from(e.dataTransfer.files));
      } else {
        onUpload(e.dataTransfer.files[0]);
      }
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) setDragActive(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ marginBottom: 16 }}>
      <div
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        className={dragActive ? `${styles.dragDropArea} ${styles.dragDropAreaActive}` : styles.dragDropArea}
        style={{ cursor: disabled ? 'not-allowed' : 'pointer' }}
      >
        Trascina qui uno o più file oppure
        <label htmlFor="file-upload-input" style={{ display: 'none' }}>Seleziona file da caricare</label>
        <input
          id="file-upload-input"
          type="file"
          ref={fileInputRef}
          style={{ display: 'none' }}
          onChange={handleFileChange}
          disabled={disabled}
          multiple
          aria-label="Seleziona file da caricare"
        />
        <DefaultButton
          id="file-upload-btn"
          text="Scegli file..."
          onClick={() => fileInputRef.current?.click()}
          disabled={disabled}
          style={{ marginLeft: 8, marginTop: 8 }}
        />
      </div>
    </Stack>
  );
};
