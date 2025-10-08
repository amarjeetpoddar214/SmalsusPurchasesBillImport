
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { BankTransaction } from './types';
import Modal from './ui/Modal';
import Button from './ui/Button';

interface ImportModalProps {
  onSave: (transactions: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>[]) => void;
  onClose: () => void;
}

type ParsedTransaction = Omit<BankTransaction, 'id' | 'matchedPurchaseId'>;

const ImportModal: React.FC<ImportModalProps> = ({ onSave, onClose }) => {
  const [parsedData, setParsedData] = useState<ParsedTransaction[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const handleDownloadTemplate = () => {
    const csvContent = "data:text/csv;charset=utf-8,"
      + "Description,Date,Amount,Type\n"
      + '"Zomato Order","2024-08-01",450.00,"debit"\n'
      + '"Salary credit","2024-07-31",75000.00,"credit"\n';

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "transaction_template.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const processFile = (selectedFile: File) => {
    if (!selectedFile) return;

    setFile(selectedFile);
    setError(null);
    setParsedData([]);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook: any = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        if (json.length < 2) {
          throw new Error("The file is empty or missing headers.");
        }

        const headers = (json[0] as string[]).map(h => h.toLowerCase().trim());
        const requiredHeaders = ['description', 'date', 'amount', 'type'];
        if (!requiredHeaders.every(rh => headers.includes(rh))) {
          throw new Error(`Missing required columns. Please ensure your file has these headers: ${requiredHeaders.join(', ')}.`);
        }

        const validatedData: ParsedTransaction[] = [];
        const headerIndexMap = headers.reduce((acc, h, i) => ({ ...acc, [h]: i }), {} as Record<string, number>);

        for (let i = 1; i < json.length; i++) {
          const row = json[i] as any[];
          if (row.length === 0) continue; // Skip empty rows

          const description = row[headerIndexMap['description']];
          const date = row[headerIndexMap['date']];
          const amount = row[headerIndexMap['amount']];
          const type = (row[headerIndexMap['type']] || '').toLowerCase().trim();

          if (!description || !date || amount === undefined || !type) {
            throw new Error(`Row ${i + 1}: Contains missing data for a required column.`);
          }
          if (typeof amount !== 'string') {
            throw new Error(`Row ${i + 1}: Invalid amount "${amount}". Amount must be a number.`);
          }



          // --- Date Validation and Conversion ---
          let parsedDate: Date | null = null;

          // If Excel stored it as a number (date code)
          if (typeof date === 'number') {
            const d = XLSX.SSF.parse_date_code(date);
            if (d) {
              parsedDate = new Date(d.y, d.m - 1, d.d); // Use local date directly
            }
          }

          // If Excel already returned a JS Date
          else if (date instanceof Date) {
            parsedDate = date;
          }
          else if (typeof date === 'string') {
            // Remove spaces and replace "/" with "-"
            const cleanDate = date.trim().replace(/\//g, '-');

            // Try YYYY-MM-DD first
            let [y, m, d] = cleanDate.split('-').map(Number);

            // If any part is NaN, try DD-MM-YYYY
            if (isNaN(y) || isNaN(m) || isNaN(d)) {
              const parts = cleanDate.split('-').map(Number);
              if (parts.length === 3) {
                [d, m, y] = parts; // assume DD-MM-YYYY
              }
            }

            parsedDate = new Date(y, m - 1, d);
          }


          // Final check
          if (!parsedDate) {
            throw new Error(`Row ${i + 1}: Invalid date "${date}". Please use a valid date format (YYYY-MM-DD or DD/MM/YYYY).`);
          }

          if (type !== 'debit' && type !== 'credit') {
            throw new Error(`Row ${i + 1}: Invalid type "${type}". Type must be 'debit' or 'credit'.`);
          }

          validatedData.push({
            description: String(description),
            amount: Number(amount),
            date: `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(2, '0')}-${String(parsedDate.getDate()).padStart(2, '0')}`,
            type: type as 'debit' | 'credit'
          });
        }
        setParsedData(validatedData);
      } catch (err: any) {
        setError(err.message || 'An unexpected error occurred while parsing the file.');
        setParsedData([]);
      }
    };
    reader.onerror = () => {
      setError("Failed to read the file.");
    };
    reader.readAsArrayBuffer(selectedFile);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) processFile(selectedFile);
  };

  const handleDragEvents = (e: React.DragEvent<HTMLDivElement>, isEntering: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(isEntering);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    handleDragEvents(e, false);
    const droppedFile = e.dataTransfer.files?.[0];
    if (droppedFile) processFile(droppedFile);
  };

  const handleImport = () => {
    if (parsedData.length > 0) {
      onSave(parsedData);
    }
  };

  return (
    <Modal title="Import Transactions from File" onClose={onClose} size="xl">
      <div className="space-y-4">
        <div className="bg-slate-700/50 p-4 rounded-md border border-slate-600 text-sm text-slate-300 space-y-2">
          <p>Upload a file (.xlsx, .xls, .csv) with the following columns:</p>
          <ul className="list-disc list-inside text-xs text-slate-400 pl-2">
            <li><code className="bg-slate-800 px-1 py-0.5 rounded">Description</code>: The transaction description (e.g., "Coffee Shop").</li>
            <li><code className="bg-slate-800 px-1 py-0.5 rounded">Date</code>: The transaction date (e.g., "2024-08-15").</li>
            <li><code className="bg-slate-800 px-1 py-0.5 rounded">Amount</code>: A numeric value for the transaction amount.</li>
            <li><code className="bg-slate-800 px-1 py-0.5 rounded">Type</code>: Either "debit" or "credit".</li>
          </ul>
          <Button onClick={handleDownloadTemplate} variant="secondary" size="normal" className="mt-2 text-xs py-1 px-2">
            Download Template
          </Button>
        </div>

        <div
          onDragEnter={(e) => handleDragEvents(e, true)}
          onDragLeave={(e) => handleDragEvents(e, false)}
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
          className={`relative p-6 border-2 border-dashed rounded-md text-center transition-colors ${isDragging ? 'border-indigo-500 bg-indigo-500/10' : 'border-slate-600 hover:border-slate-500'}`}
        >
          <input
            type="file"
            id="file-upload"
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            onChange={handleFileChange}
            accept=".xlsx, .xls, .csv"
          />
          <label htmlFor="file-upload" className="cursor-pointer text-slate-400">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10 mx-auto mb-2 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1}><path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M12 15l-4-4m0 0l4-4m-4 4h12" /></svg>
            <span className="font-semibold text-indigo-400">Click to upload</span> or drag and drop.
            {file && <p className="text-sm mt-2 text-slate-300">Selected file: {file.name}</p>}
          </label>
        </div>

        {error && <div className="bg-red-500/20 text-red-300 p-3 rounded-md text-sm">{error}</div>}

        {parsedData.length > 0 && (
          <div>
            <h3 className="font-semibold text-slate-200 mb-2">Preview ({parsedData.length} transactions found)</h3>
            <div className="max-h-60 overflow-y-auto border border-slate-700 rounded-md">
              <table className="w-full text-sm text-left">
                <thead className="sticky top-0 bg-slate-800">
                  <tr>
                    <th className="p-2">Description</th>
                    <th className="p-2">Date</th>
                    <th className="p-2 text-right">Amount</th>
                    <th className="p-2">Type</th>
                  </tr>
                </thead>
                <tbody className="bg-slate-900/50">
                  {parsedData.slice(0, 10).map((t, i) => ( // Preview first 10 rows
                    <tr key={i} className="border-t border-slate-700">
                      <td className="p-2">{t.description}</td>
                      <td className="p-2">{t.date}</td>
                      <td className="p-2 text-right">₹{t.amount.toFixed(2)}</td>
                      <td className="p-2">
                        <span className={`px-2 py-0.5 rounded-full text-xs ${t.type === 'debit' ? 'bg-red-500/20 text-red-300' : 'bg-green-500/20 text-green-300'}`}>
                          {t.type}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {parsedData.length > 10 && <p className="text-center text-xs text-slate-500 p-2 bg-slate-800">...and {parsedData.length - 10} more.</p>}
            </div>
          </div>
        )}

        <div className="flex justify-end gap-3 pt-4">
          <Button type="button" onClick={onClose} variant="secondary">Cancel</Button>
          <Button type="button" onClick={handleImport} variant="primary" disabled={parsedData.length === 0}>
            Import {parsedData.length > 0 ? parsedData.length : ''} Transactions
          </Button>
        </div>
      </div>
    </Modal>
  );
};

export default ImportModal;