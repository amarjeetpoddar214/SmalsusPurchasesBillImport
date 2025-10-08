
import React, { useState, useEffect, useRef } from 'react';
import { Purchase, Category } from './types';
import { suggestCategory } from '../../../services/geminiService';
import Modal from './ui/Modal';
import Input from './ui/Input';
import Select from './ui/Select';
import Button from './ui/Button';
import { Web } from 'sp-pnp-js';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface PurchaseFormProps {
  purchase: Purchase | null;
  onSave: (purchase: Omit<Purchase, 'id'> & { billFile?: File }) => void; // ✅ allow billFile
  onClose: () => void;
  context: WebPartContext;
}



const PurchaseForm: React.FC<PurchaseFormProps> = ({ purchase, onSave, onClose, context }) => {
  const [name, setName] = useState('');
  const [amount, setAmount] = useState('');
  const [date, setDate] = useState('');
  const [category, setCategory] = useState<Category>(Category.Other);
  const [billImage, setBillImage] = useState<string | null>(null);
  const [isSuggesting, setIsSuggesting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [transactions, setTransactions] = useState<{ id: string, title: string }[]>([]);
  const [selectedTransaction, setSelectedTransaction] = useState<string>('');
  const [billFile, setBillFile] = useState<File | null>(null); // store file for upload


  const webURL = context.pageContext.web.absoluteUrl;
  const transactionsListId = '38dbeddd-4e03-4d26-83a7-90305de6046e';


  useEffect(() => {
    if (purchase) {
      setName(purchase.name);
      setAmount(String(purchase.amount));
      setDate(purchase.date);
      setCategory(purchase.category);
      setBillImage(purchase.billImage || null); // existing URL
      setBillFile(null); // ✅ important: no new file yet
      setSelectedTransaction(purchase.matchedBankTransactionId || '');
    } else {
      setName('');
      setAmount('');
      setDate(new Date().toISOString().split('T')[0]);
      setCategory(Category.Other);
      setBillImage(null);
      setBillFile(null);
      setSelectedTransaction('');
    }
  }, [purchase]);



  useEffect(() => {
    const loadTransactions = async () => {
      try {
        const web = new Web(webURL);
        const items = await web.lists
          .getById(transactionsListId)
          .items.select("Id,Title")
          .get();

        setTransactions(items.map((item: any) => ({
          id: item.Id.toString(),
          title: item.Title
        })));
      } catch (error) {
        console.error("Error loading transactions:", error);
      }
    };

    loadTransactions();
  }, []);



  const handleSuggestCategory = async () => {
    if (!name) return;
    setIsSuggesting(true);
    const suggested = await suggestCategory(name);
    if (suggested) {
      setCategory(suggested);
    }
    setIsSuggesting(false);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setBillFile(file); // save the file
      const reader = new FileReader();
      reader.onloadend = () => {
        setBillImage(reader.result as string); // preview only
      };
      reader.readAsDataURL(file);
    }
  };



  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const parsedAmount = parseFloat(amount);

    if (name && !isNaN(parsedAmount) && date) {
      onSave({
        name,
        amount: parsedAmount,
        date,
        category,
        billFile: billFile || undefined, // ✅ use state instead of input ref
        matchedBankTransactionId: selectedTransaction || null
      });
    }
  };




  return (
    <Modal title={purchase ? 'Edit Purchase' : 'Add Purchase'} onClose={onClose}>
      <form onSubmit={handleSubmit} className="space-y-4">
        <div>
          <label htmlFor="name" className="block text-sm font-medium text-slate-300 mb-1">Item Name</label>
          <Input id="name" type="text" value={name} onChange={(e) => setName(e.target.value)} required />
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label htmlFor="amount" className="block text-sm font-medium text-slate-300 mb-1">Amount (₹)</label>
            <Input id="amount" type="number" step="0.01" value={amount} onChange={(e) => setAmount(e.target.value)} required />
          </div>
          <div>
            <label htmlFor="date" className="block text-sm font-medium text-slate-300 mb-1">Date</label>
            <Input id="date" type="date" value={date} onChange={(e) => setDate(e.target.value)} required />
          </div>
        </div>

        <div>
          <label htmlFor="category" className="block text-sm font-medium text-slate-300 mb-1">Category</label>
          <div className="flex items-center gap-2">
            <Select id="category" value={category} onChange={(e) => setCategory(e.target.value as Category)} className="flex-grow">
              {Object.values(Category).map(cat => (
                <option key={cat} value={cat}>{cat}</option>
              ))}
            </Select>
            <Button type="button" onClick={handleSuggestCategory} disabled={isSuggesting || !name} variant="secondary">
              {isSuggesting ? '...' : 'Suggest'}
            </Button>
          </div>
        </div>

        <div>
          <label htmlFor="lookup" className="block text-sm font-medium text-slate-300 mb-1">Match Transaction</label>
          <Select
            id="lookup"
            value={selectedTransaction}
            onChange={(e) => setSelectedTransaction(e.target.value)}
            className="flex-grow"
          >
            <option value="">-- Select Transaction --</option>
            {transactions.map(tx => (
              <option key={tx.id} value={tx.id}>{tx.title}</option>
            ))}
          </Select>
        </div>


        <div>
          <label className="block text-sm font-medium text-slate-300 mb-1">Bill / Receipt</label>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            className="hidden"
            accept="image/png, image/jpeg, image/webp, application/pdf"
          />
          {billImage ? (
            <div className="mt-2 bg-slate-900 rounded-md p-3 flex flex-col items-center justify-center text-center">
              {billImage.toLowerCase().endsWith('.pdf') || billImage.startsWith('data:application/pdf') ? (
                <div className="flex flex-col items-center space-y-2">
                  <img
                    src="https://upload.wikimedia.org/wikipedia/commons/8/87/PDF_file_icon.svg"
                    alt="PDF Icon"
                    className="w-12 h-12"
                  />
                  <a
                    href={billImage}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-blue-400 underline"
                  >
                    Open PDF Bill
                  </a>
                  <Button
                    type="button"
                    onClick={() => setBillImage(null)}
                    variant="secondary"
                    size="normal"
                  >
                    Remove File
                  </Button>
                </div>
              ) : (
                <div className="relative group w-full">
                  <img
                    src={billImage}
                    alt="Bill preview"
                    className="rounded-md max-h-48 w-full object-contain"
                  />
                  <div className="absolute inset-0 bg-black/60 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity rounded-md">
                    <Button
                      type="button"
                      onClick={() => setBillImage(null)}
                      variant="secondary"
                      size="normal"
                    >
                      Remove Image
                    </Button>
                  </div>
                </div>
              )}
            </div>
          ) : (
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="mt-1 w-full flex flex-col justify-center items-center px-4 py-6 border-2 border-slate-600 border-dashed rounded-md text-sm text-slate-400 hover:bg-slate-700/50 hover:border-slate-500 transition"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8 mb-2 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
              </svg>
              <span>Click to upload an image or PDF</span>
            </button>
          )}


        </div>

        <div className="flex justify-end gap-3 pt-4">
          <Button type="button" onClick={onClose} variant="secondary">Cancel</Button>
          <Button type="submit" variant="primary">Save Purchase</Button>
        </div>
      </form>
    </Modal>
  );
};

export default PurchaseForm;