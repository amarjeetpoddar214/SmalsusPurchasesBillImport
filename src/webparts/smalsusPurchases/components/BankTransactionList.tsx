import React from 'react';
import { BankTransaction, Purchase } from './types';
import BankTransactionItem from './BankTransactionItem';
import Card from './ui/Card';

interface BankTransactionListProps {
  transactions: BankTransaction[];
  purchases: Purchase[];
  onMatch: (transaction: BankTransaction) => void;
  onUnmatch: (transactionId: string, purchaseId: string) => void;
}

const BankTransactionList: React.FC<BankTransactionListProps> = ({ transactions, purchases, onMatch, onUnmatch }) => {
  return (
    <Card>
      <h2 className="text-xl font-bold text-white mb-4">Bank Transactions</h2>
      <div className="space-y-3">
        {transactions.length > 0 ? (
          transactions.map(transaction => (
            <BankTransactionItem
              key={transaction.id}
              transaction={transaction}
              purchases={purchases}
              onMatch={onMatch}
              onUnmatch={onUnmatch}
            />
          ))
        ) : (
          <div className="text-center py-10">
            <p className="text-slate-400">No bank transactions yet.</p>
            <p className="text-slate-500 text-sm">Click "Add Transaction" to get started.</p>
          </div>
        )}
      </div>
    </Card>
  );
};

export default BankTransactionList;
