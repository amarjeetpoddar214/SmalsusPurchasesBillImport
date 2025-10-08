import React from 'react';
import { Purchase } from './types';
import PurchaseItem from './PurchaseItem';
import Card from './ui/Card';

interface PurchaseListProps {
  purchases: Purchase[];
  onEdit: (purchase: Purchase) => void;
  onDelete: (id: string) => void;
}

const PurchaseList: React.FC<PurchaseListProps> = ({ purchases, onEdit, onDelete }) => {
  return (
    <Card>
      <h2 className="text-xl font-bold text-white mb-4">Recent Purchases</h2>
      <div className="space-y-3">
        {purchases.length > 0 ? (
          purchases.map(purchase => (
            <PurchaseItem
              key={purchase.id}
              purchase={purchase}
              onEdit={onEdit}
              onDelete={onDelete}
              isMatched={!!purchase.matchedBankTransactionId}
            />
          ))
        ) : (
          <div className="text-center py-10">
            <p className="text-slate-400">No purchases yet.</p>
            <p className="text-slate-500 text-sm">Click "Add Purchase" to get started.</p>
          </div>
        )}
      </div>
    </Card>
  );
};

export default PurchaseList;
