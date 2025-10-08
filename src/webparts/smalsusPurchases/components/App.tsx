import * as React from 'react';
import { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Purchase, BankTransaction } from './types';
import { useAppData, MonthlyData } from '../../../hooks/useAppData';
import Header from '../components/Header';
import PurchaseList from '../components/PurchaseList';
import PurchaseForm from '../components/PurchaseForm';
import CategoryChart from '../components/CategoryChart';
import Insights from '../components/Insights';
import Card from '../components/ui/Card';
import BankTransactionList from '../components/BankTransactionList';
import BankTransactionForm from '../components/BankTransactionForm';
import MatchModal from '../components/MatchModal';
import ImportModal from '../components/ImportModal';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface AppProps {
  context: WebPartContext;
}

const getCurrentMonthYYYYMM = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month: any = today.getMonth() + 1; // getMonth is 0-indexed
  return `${year}-${String(month).padStart(2, '0')}`;
};

const App: React.FC<AppProps> = ({ context }) => {
  const { appData, isLoaded, addPurchase, updatePurchase, deletePurchase, addTransaction, updateTransaction, addMultipleTransactions } = useAppData(context);

  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthYYYYMM());

  const [isPurchaseModalOpen, setIsPurchaseModalOpen] = useState(false);
  const [isTransactionModalOpen, setIsTransactionModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [editingPurchase, setEditingPurchase] = useState<Purchase | null>(null);

  const [isMatchingModalOpen, setIsMatchingModalOpen] = useState(false);
  const [transactionToMatch, setTransactionToMatch] = useState<BankTransaction | null>(null);

  const handlePreviousMonth = () => {
    setSelectedMonth(prev => {
      let [year, month] = prev.split('-').map(Number);
      if (month === 1) {
        month = 12;
        year -= 1;
      } else {
        month -= 1;
      }
      return `${year}-${String(month).padStart(2, '0')}`;
    });
  };

  const handleNextMonth = () => {
    setSelectedMonth(prev => {
      let [year, month] = prev.split('-').map(Number);
      if (month === 12) {
        month = 1;
        year += 1;
      } else {
        month += 1;
      }
      return `${year}-${String(month).padStart(2, '0')}`;
    });
  };

  const { purchases, transactions } = useMemo<MonthlyData>(() => {
    return appData[selectedMonth] || { purchases: [], transactions: [] };
  }, [appData, selectedMonth]);

  const allPurchases = useMemo(() => Object.values(appData).flatMap((d: MonthlyData) => d.purchases), [appData]);
  const allTransactions = useMemo(() => Object.values(appData).flatMap((d: MonthlyData) => d.transactions), [appData]);

  const handleAddPurchaseClick = () => {
    setEditingPurchase(null);
    setIsPurchaseModalOpen(true);
  };

  const handleAddTransactionClick = () => {
    setIsTransactionModalOpen(true);
  };

  const handleImportClick = () => {
    setIsImportModalOpen(true);
  };

  const handleExportData = () => {
    const purchasesForExport = purchases.map(p => ({
      Name: p.name,
      Amount: p.amount,
      Date: p.date,
      Category: p.category,
      'Bill Attached': p.billImage ? 'Yes' : 'No',
      'Matched': p.matchedBankTransactionId ? 'Yes' : 'No',
    }));

    const transactionsForExport = transactions.map(t => ({
      Description: t.description,
      Amount: t.amount,
      Date: t.date,
      Type: t.type,
      'Matched': t.matchedPurchaseId ? 'Yes' : 'No',
    }));

    const purchasesSheet = XLSX.utils.json_to_sheet(purchasesForExport);
    const transactionsSheet = XLSX.utils.json_to_sheet(transactionsForExport);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, purchasesSheet, 'Purchases');
    XLSX.utils.book_append_sheet(workbook, transactionsSheet, 'Bank Transactions');

    XLSX.writeFile(workbook, `Purchase_Report_${selectedMonth}.xlsx`);
  };

  const handleEditPurchaseClick = (purchase: Purchase) => {
    setEditingPurchase(purchase);
    setIsPurchaseModalOpen(true);
  };

  const handleSavePurchase = (purchaseData: Omit<Purchase, 'id'>) => {
    if (editingPurchase) {
      // ✅ Preserve the existing billImage if user didn’t upload a new file
      const updatedPurchase = {
        ...editingPurchase,
        ...purchaseData,
        id: editingPurchase.id,
        billImage: (purchaseData as any).billImage ? undefined : editingPurchase.billImage // ✅ fix property name
      };
      updatePurchase(updatedPurchase);
    } else {
      addPurchase(purchaseData);
    }

    setIsPurchaseModalOpen(false);
    setEditingPurchase(null);
  };



  const handleUnmatch = useCallback((transactionId: string, purchaseId: string) => {
    const transaction = allTransactions.find(t => t.id === transactionId);
    const purchase = allPurchases.find(p => p.id === purchaseId);
    if (transaction && purchase) {
      updateTransaction({ ...transaction, matchedPurchaseId: null });
      updatePurchase({ ...purchase, matchedBankTransactionId: null });
    }
  }, [allTransactions, allPurchases, updateTransaction, updatePurchase]);

  const handleDeletePurchase = (id: string) => {
    if (window.confirm('Are you sure you want to delete this purchase?')) {
      const purchaseToDelete = allPurchases.find(p => p.id === id);
      if (purchaseToDelete) {
        if (purchaseToDelete.matchedBankTransactionId) {
          handleUnmatch(purchaseToDelete.matchedBankTransactionId, id);
        }
        deletePurchase(purchaseToDelete);
      }
    }
  };

  const handleSaveTransaction = (transactionData: Omit<BankTransaction, 'id'>) => {
    addTransaction(transactionData);
    setIsTransactionModalOpen(false);
  }

  const handleImportSave = (newTransactions: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>[]) => {
    addMultipleTransactions(newTransactions);
    setIsImportModalOpen(false);
  };

  const handleMatchClick = (transaction: BankTransaction) => {
    setTransactionToMatch(transaction);
    setIsMatchingModalOpen(true);
  };

  const handleMatch = useCallback((transactionId: string, purchaseId: string) => {
    const transaction = allTransactions.find(t => t.id === transactionId);
    const purchase = allPurchases.find(p => p.id === purchaseId);

    if (transaction && purchase) {
      updateTransaction({ ...transaction, matchedPurchaseId: purchaseId });
      updatePurchase({ ...purchase, matchedBankTransactionId: transactionId });
    }
    setIsMatchingModalOpen(false);
    setTransactionToMatch(null);
  }, [allTransactions, allPurchases, updateTransaction, updatePurchase]);



  const totalExpenses = useMemo(() => {
    return purchases.reduce((total, p) => total + p.amount, 0).toFixed(2);
  }, [purchases]);

  const unmatchedPurchases = useMemo(() => allPurchases.filter(p => !p.matchedBankTransactionId), [allPurchases]);

  if (!isLoaded) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center">
        <p className="text-slate-300 text-lg">Loading your financial data...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 font-sans">
      <Header
        context={context}
        selectedMonth={selectedMonth}
        onPreviousMonth={handlePreviousMonth}
        onNextMonth={handleNextMonth}
        onAddPurchase={handleAddPurchaseClick}
        onAddTransaction={handleAddTransactionClick}
        onImportTransactions={handleImportClick}
        onExportData={handleExportData}
      />
      <main className="p-4 sm:p-6 lg:p-8  mx-auto">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <div className="lg:col-span-2 space-y-8">
            <PurchaseList
              purchases={purchases}
              onEdit={handleEditPurchaseClick}
              onDelete={handleDeletePurchase}
            />
            <BankTransactionList
              transactions={transactions}
              purchases={allPurchases}
              onMatch={handleMatchClick}
              onUnmatch={handleUnmatch}
            />
          </div>
          <div className="space-y-8">
            <Card>
              <h2 className="text-xl font-bold text-white mb-2">Total Expenses ({new Date(selectedMonth + '-02').toLocaleString('default', { month: 'long', year: 'numeric' })})</h2>
              <p className="text-4xl font-extrabold text-indigo-400">₹{totalExpenses}</p>
            </Card>
            <CategoryChart purchases={purchases} />
            <Insights purchases={purchases} />
          </div>
        </div>
      </main>

      {isPurchaseModalOpen && (
        <PurchaseForm
          purchase={editingPurchase}
          onSave={handleSavePurchase}
          onClose={() => setIsPurchaseModalOpen(false)}
          context={context}
        />
      )}

      {isTransactionModalOpen && (
        <BankTransactionForm
          onSave={handleSaveTransaction}
          onClose={() => setIsTransactionModalOpen(false)}
        />
      )}

      {isImportModalOpen && (
        <ImportModal
          onSave={handleImportSave}
          onClose={() => setIsImportModalOpen(false)}
        />
      )}

      {isMatchingModalOpen && transactionToMatch && (
        <MatchModal
          transaction={transactionToMatch}
          unmatchedPurchases={unmatchedPurchases}
          onMatch={handleMatch}
          onClose={() => setIsMatchingModalOpen(false)}
        />
      )}
    </div>
  );
};

export default App;