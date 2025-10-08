import React from 'react';
import { BankTransaction, Purchase } from './types';
import Modal from './ui/Modal';
import Button from './ui/Button';

interface MatchModalProps {
  transaction: BankTransaction;
  unmatchedPurchases: Purchase[];
  onMatch: (transactionId: string, purchaseId: string) => void;
  onClose: () => void;
}

const MatchModal: React.FC<MatchModalProps> = ({ transaction, unmatchedPurchases, onMatch, onClose }) => {
  return (
    <Modal title={`Match transaction: "${transaction.description}"`} onClose={onClose} size="lg">
      <div className="space-y-3 max-h-[60vh] overflow-y-auto pr-2">
        <p className="text-slate-400 text-sm mb-4">Select a purchase to match with this transaction of <span className="font-bold text-slate-200">₹{transaction.amount.toFixed(2)}</span>.</p>
        {unmatchedPurchases.length > 0 ? (
          unmatchedPurchases.map(purchase => (
            <button
              key={purchase.id}
              onClick={() => onMatch(transaction.id, purchase.id)}
              className="w-full text-left bg-slate-800/50 rounded-lg p-3 flex items-center justify-between gap-4 transition-all hover:bg-slate-700 border border-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <div>
                <p className="font-semibold text-white">{purchase.name}</p>
                <p className="text-sm text-slate-400">{new Date(purchase.date + 'T00:00:00').toLocaleDateString()}</p>
              </div>
              <p className="text-lg font-bold text-indigo-400">
                ₹{purchase.amount.toFixed(2)}
              </p>
            </button>
          ))
        ) : (
          <div className="text-center py-10">
            <p className="text-slate-400">No unmatched purchases found.</p>
            <p className="text-slate-500 text-sm">You can add new purchases from the main screen.</p>
          </div>
        )}
      </div>
      <div className="flex justify-end pt-6">
        <Button type="button" onClick={onClose} variant="secondary">Cancel</Button>
      </div>
    </Modal>
  );
};

export default MatchModal;
