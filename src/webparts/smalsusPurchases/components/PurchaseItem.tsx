import React, { useState } from 'react';
import { Purchase, Category } from './types';
import Button from './ui/Button';
import Modal from './ui/Modal';

interface PurchaseItemProps {
  purchase: Purchase;
  onEdit: (purchase: Purchase) => void;
  onDelete: (id: string) => void;
  isMatched: boolean;
}

const categoryColors: Record<Category, string> = {
    [Category.Groceries]: 'bg-green-500/20 text-green-300 border-green-500/30',
    [Category.Utilities]: 'bg-blue-500/20 text-blue-300 border-blue-500/30',
    [Category.Entertainment]: 'bg-purple-500/20 text-purple-300 border-purple-500/30',
    [Category.Transport]: 'bg-yellow-500/20 text-yellow-300 border-yellow-500/30',
    [Category.Health]: 'bg-red-500/20 text-red-300 border-red-500/30',
    [Category.Shopping]: 'bg-pink-500/20 text-pink-300 border-pink-500/30',
    [Category.FoodAndDining]: 'bg-orange-500/20 text-orange-300 border-orange-500/30',
    [Category.Other]: 'bg-slate-500/20 text-slate-300 border-slate-500/30',
};

const PurchaseItem: React.FC<PurchaseItemProps> = ({ purchase, onEdit, onDelete, isMatched }) => {
  const [isBillVisible, setIsBillVisible] = useState(false);

  const formattedDate = new Date(purchase.date + 'T00:00:00').toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });

  return (
    <>
      <div className="bg-slate-800/50 rounded-lg p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 transition-all hover:bg-slate-800 border border-slate-700">
        <div className="flex-1">
          <div className="flex items-center gap-2">
            <p className="font-semibold text-white">{purchase.name}</p>
            {isMatched && (
                <span title="Matched with bank transaction" className="text-green-400">
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor">
                        <path fillRule="evenodd" d="M12.586 4.586a2 2 0 112.828 2.828l-3 3a2 2 0 01-2.828 0l-1.5-1.5a2 2 0 112.828-2.828l1.5 1.5 3-3z" clipRule="evenodd" />
                        <path fillRule="evenodd" d="M6.586 12.586a2 2 0 10-2.828 2.828l3 3a2 2 0 002.828 0l1.5-1.5a2 2 0 10-2.828-2.828l-1.5 1.5-3-3z" clipRule="evenodd" />
                    </svg>
                </span>
            )}
            {purchase.billImage && (
              <button 
                onClick={() => setIsBillVisible(true)} 
                className="text-slate-400 hover:text-indigo-400 transition-colors" 
                aria-label="View bill"
                title="View bill"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M4 4a2 2 0 012-2h8a2 2 0 012 2v12a1 1 0 110 2h-3a1 1 0 01-1-1v-2a1 1 0 00-1-1H9a1 1 0 00-1 1v2a1 1 0 01-1 1H4a1 1 0 110-2V4zm3 1a1 1 0 000 2h8a1 1 0 100-2H7zM7 9a1 1 0 000 2h2a1 1 0 100-2H7z" clipRule="evenodd" />
                </svg>
              </button>
            )}
          </div>
          <div className="flex items-center gap-x-3 text-sm text-slate-400 mt-1">
            <span>{formattedDate}</span>
            <span className={`px-2 py-0.5 rounded-full text-xs font-medium border ${categoryColors[purchase.category]}`}>
              {purchase.category}
            </span>
          </div>
        </div>
        <div className="flex items-center gap-x-4 w-full sm:w-auto">
          <p className="text-lg font-bold text-indigo-400 flex-1 sm:flex-none">
            ₹{purchase.amount.toFixed(2)}
          </p>
          <div className="flex items-center gap-x-2">
              <Button onClick={() => onEdit(purchase)} variant="ghost" size="icon">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.5L15.232 5.232z" /></svg>
              </Button>
              <Button onClick={() => onDelete(purchase.id)} variant="ghost" size="icon" className="text-red-400 hover:bg-red-500/10 hover:text-red-300">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
              </Button>
          </div>
        </div>
      </div>
      {isBillVisible && purchase.billImage && (
        <Modal title={`Bill for ${purchase.name}`} onClose={() => setIsBillVisible(false)} size="lg">
            <img src={purchase.billImage} alt={`Bill for ${purchase.name}`} className="w-full h-auto rounded-md object-contain max-h-[80vh]" />
        </Modal>
      )}
    </>
  );
};

export default PurchaseItem;
