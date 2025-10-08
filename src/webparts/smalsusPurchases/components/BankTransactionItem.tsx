import React from 'react';
import { BankTransaction, Purchase } from './types';
import Button from './ui/Button';

interface BankTransactionItemProps {
  transaction: BankTransaction;
  purchases: Purchase[];
  onMatch: (transaction: BankTransaction) => void;
  onUnmatch: (transactionId: string, purchaseId: string) => void;
}

const BankTransactionItem: React.FC<BankTransactionItemProps> = ({ transaction, purchases, onMatch, onUnmatch }) => {
  const formattedDate = new Date(transaction.date + 'T00:00:00').toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });

  const matchedPurchase = transaction.matchedPurchaseId
    ? purchases.find(p => p.id === transaction.matchedPurchaseId)
    : null;

  const amountColor = transaction.type === 'debit' ? 'text-red-400' : 'text-green-400';
  const amountPrefix = transaction.type === 'debit' ? '-' : '+';

  return (
    <div className="bg-slate-800/50 rounded-lg p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 transition-all hover:bg-slate-800 border border-slate-700">
      <div className="flex-1">
        <p className="font-semibold text-white">{transaction.description}</p>
        <div className="flex items-center gap-x-3 text-sm text-slate-400 mt-1">
          <span>{formattedDate}</span>
          {matchedPurchase && (
            <>
              <span className="text-green-400 flex items-center gap-1">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M12.586 4.586a2 2 0 112.828 2.828l-3 3a2 2 0 01-2.828 0l-1.5-1.5a2 2 0 112.828-2.828l1.5 1.5 3-3z" clipRule="evenodd" />
                  <path fillRule="evenodd" d="M6.586 12.586a2 2 0 10-2.828 2.828l3 3a2 2 0 002.828 0l1.5-1.5a2 2 0 10-2.828-2.828l-1.5 1.5-3-3z" clipRule="evenodd" />
                </svg>
                Matched: {matchedPurchase.name}
              </span>
            </>
          )}
        </div>
      </div>
      <div className="flex items-center gap-x-4 w-full sm:w-auto">
        <p className={`text-lg font-bold ${amountColor} flex-1 sm:flex-none`}>
          {amountPrefix}₹{transaction.amount.toFixed(2)}
        </p>
        <div className="flex items-center">
          {matchedPurchase ? (
            <Button onClick={() => onUnmatch(transaction.id, matchedPurchase.id)} variant="secondary" size="normal">Unmatch</Button>
          ) : (
            <Button onClick={() => onMatch(transaction)} variant="primary" size="normal">Match</Button>
          )}
        </div>
      </div>
    </div>
  );
};

export default BankTransactionItem;