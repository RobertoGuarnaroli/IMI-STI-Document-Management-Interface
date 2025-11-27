// ============================================
// ExcelUpload.tsx
// Componente per upload e parsing Excel
// ============================================

import * as React from 'react';
import { DefaultButton, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
import * as XLSX from 'xlsx';
import styles from './ExcelUpload.module.scss';

export interface IExcelData {
  ProjectCode?: string;
  Title?: string;
  Customer?: string;
  ProjectManager?: string;
  Status?: string;
  StartDate?: string | Date;
  EndDate?: string | Date;
  Notes?: string;
}

export interface IExcelUploadProps {
  onDataExtracted: (data: IExcelData) => void;
  onClear?: () => void;
  disabled?: boolean;
}

export const ExcelUpload: React.FC<IExcelUploadProps> = ({ onDataExtracted, onClear, disabled }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [success, setSuccess] = React.useState(false);

  // Mapping tra possibili nomi colonne Excel e i nostri field names
  const columnMappings: Record<string, string[]> = {
    ProjectCode: [
      'PROJECT CODE',
      'PROJECT_CODE',
      'Project Code',
      'ProjectCode',
      'Codice Progetto'
    ],
    Title: [
      'TITLE',
      'TITLE',
      'Project Title',
      'Titolo',
      'Title'
    ],
    Customer: [
      'CUSTOMER',
      'CUSTOMER',
      'Customer Name',
      'Cliente',
      'Customer'
    ],
    ProjectManager: [
      'PROJECT MANAGER',
      'PROJECT_MANAGER',
      'Project Manager',
      'Responsabile',
      'Manager',
      'Project Manager Email',
      'ProjectManagerEmail',
      'Manager Email',
      'Responsabile Email',
      'Email Project Manager',
      'Email Responsabile'
    ],
    Status: [
      'STATUS',
      'STATUS',
      'Stato',
      'Status'
    ],
    StartDate: [
      'START DATE',
      'START_DATE',
      'Data Inizio',
      'StartDate'
    ],
    EndDate: [
      'END DATE',
      'END_DATE',
      'Data Fine',
      'EndDate'
    ],
    Notes: [
      'NOTES',
      'NOTES',
      'Note',
      'Notes'
    ]
  };

  const findColumnValue = (row: any, fieldName: string): any => {
    const possibleNames = columnMappings[fieldName] || [];
    
    for (const colName of possibleNames) {
      // Case-insensitive search
      const key = Object.keys(row).find(
        k => k.toLowerCase().trim() === colName.toLowerCase().trim()
      );
      
      if (key && row[key] !== undefined && row[key] !== null && row[key] !== '') {
        return row[key];
      }
    }
    
    return undefined;
  };

  const parseExcelFile = async (file: File): Promise<void> => {
    setLoading(true);
    setError(null);
    setSuccess(false);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });

      // Prendi il primo sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Converti in JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (!jsonData || jsonData.length === 0) {
        throw new Error('The Excel file is empty or does not contain valid data');
      }

      // Take the first data row (after header)
      const firstRow = jsonData[0] as any;

      // Extract data using mappings
      const extractedData: IExcelData = {
        ProjectCode: findColumnValue(firstRow, 'ProjectCode'),
        Title: findColumnValue(firstRow, 'Title'),
        Customer: findColumnValue(firstRow, 'Customer'),
        ProjectManager: findColumnValue(firstRow, 'ProjectManager'),
        Status: findColumnValue(firstRow, 'Status'),
        StartDate: findColumnValue(firstRow, 'StartDate'),
        EndDate: findColumnValue(firstRow, 'EndDate'),
        Notes: findColumnValue(firstRow, 'Notes')
      };

      // Format dates if present: restituisci oggetti Date
      if (extractedData.StartDate) {
        const dateValue = extractedData.StartDate;
        if (typeof dateValue === 'number') {
          const excelDate = XLSX.SSF.parse_date_code(dateValue);
          extractedData.StartDate = new Date(
            excelDate.y,
            excelDate.m - 1,
            excelDate.d
          );
        } else if (typeof dateValue === 'string') {
          const parsed = new Date(dateValue);
          if (!isNaN(parsed.getTime())) {
            extractedData.StartDate = parsed;
          }
        }
      }
      if (extractedData.EndDate) {
        const dateValue = extractedData.EndDate;
        if (typeof dateValue === 'number') {
          const excelDate = XLSX.SSF.parse_date_code(dateValue);
          extractedData.EndDate = new Date(
            excelDate.y,
            excelDate.m - 1,
            excelDate.d
          );
        } else if (typeof dateValue === 'string') {
          const parsed = new Date(dateValue);
          if (!isNaN(parsed.getTime())) {
            extractedData.EndDate = parsed;
          }
        }
      }

      console.log('Extracted data from Excel:', extractedData);

      // Check that at least one field is populated
      const hasData = Object.values(extractedData).some(v => v !== undefined && v !== '');
      
      if (!hasData) {
        throw new Error('No valid data found in the Excel file. Check the column names.');
      }

      onDataExtracted(extractedData);
      setSuccess(true);
      
      // Reset del file input
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }

    } catch (err: any) {
      console.error('Excel parsing error:', err);
      setError(err.message || 'Error reading the Excel file');
    } finally {
      setLoading(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      // Check extension
      const fileName = file.name.toLowerCase();
      if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        setError('Unsupported file format. Use .xlsx or .xls files');
        return;
      }

      void parseExcelFile(file);
    }
  };

  return (
    <div className={styles.excelUploadContainer}>
      <div className={styles.uploadSection}>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          className={styles.hiddenInput}
          disabled={disabled || loading}
          aria-label="Upload Excel file"
        />
        
        <PrimaryButton
          text={loading ? 'Loading...' : 'Upload Excel'}
          iconProps={{ iconName: 'ExcelDocument' }}
          onClick={() => fileInputRef.current?.click()}
          disabled={disabled || loading}
          className={styles.uploadButton}
        />

        <DefaultButton
          text="Clear"
          iconProps={{ iconName: 'Cancel' }}
          onClick={() => {
            setError(null);
            setSuccess(false);
            if (fileInputRef.current) {
              fileInputRef.current.value = '';
            }
            if (typeof onClear === 'function') {
              onClear();
            }
          }}
          disabled={disabled || loading}
          className={styles.cancelButton}
        />
      </div>

      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel="Close"
          className={styles.messageBar}
        >
          {error}
        </MessageBar>
      )}

      {success && (
        <MessageBar
          messageBarType={MessageBarType.success}
          isMultiline={false}
          onDismiss={() => setSuccess(false)}
          dismissButtonAriaLabel="Close"
          className={styles.messageBar}
        >
          Data extracted successfully! Check the fields in the form.
        </MessageBar>
      )}

      <div className={styles.helpText}>
        <strong>Required Excel format:</strong>
        <ul>
          <li>First row: column headers</li>
          <li>Second row: document data</li>
          <li>Supported columns: Project Code, Title, Customer, Status etc.</li>
        </ul>
      </div>
    </div>
  );
};