import React, { useState } from 'react';
import { BankTransaction } from './types'
import Modal from './ui/Modal';
import Input from './ui/Input';
import Select from './ui/Select';
import Button from './ui/Button';

interface BankTransactionFormProps {
  onSave: (transaction: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>) => void;
  onClose: () => void;
}

const BankTransactionForm: React.FC<BankTransactionFormProps> = ({ onSave, onClose }) => {
  const [description, setDescription] = useState('');
  const [amount, setAmount] = useState('');
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [type, setType] = useState<'debit' | 'credit'>('debit');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const parsedAmount = parseFloat(amount);
    if (description && !isNaN(parsedAmount) && date) {
      onSave({ description, amount: parsedAmount, date, type });
    }
  };

  return (
    <Modal title="Add Bank Transaction" onClose={onClose}>
      <form onSubmit={handleSubmit} className="space-y-4">
        <div>
          <label htmlFor="description" className="block text-sm font-medium text-slate-300 mb-1">Description</label>
          <Input id="description" type="text" value={description} onChange={(e) => setDescription(e.target.value)} required />
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
          <label htmlFor="type" className="block text-sm font-medium text-slate-300 mb-1">Type</label>
          <Select id="type" value={type} onChange={(e) => setType(e.target.value as 'debit' | 'credit')}>
            <option value="debit">debit</option>
            <option value="credit">credit</option>
          </Select>
        </div>

        <div className="flex justify-end gap-3 pt-4">
          <Button type="button" onClick={onClose} variant="secondary">Cancel</Button>
          <Button type="submit" variant="primary">Save Transaction</Button>
        </div>
      </form>
    </Modal>
  );
};

export default BankTransactionForm;
