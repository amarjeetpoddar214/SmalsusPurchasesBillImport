// src/webparts/.../components/App.tsx
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
  const month: number = today.getMonth() + 1;
  return `${year}-${String(month).padStart(2, '0')}`;
};

const App: React.FC<AppProps> = ({ context }) => {
  const { appData, isLoaded, addPurchase, updatePurchase, deletePurchase, addTransaction, updateTransaction, addMultipleTransactions } = useAppData(context);

  // ---------- State ----------
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthYYYYMM());

  // viewMode: monthly | quarterly | half-yearly | yearly
  const [viewMode, setViewMode] = useState<'monthly' | 'quarterly' | 'half-yearly' | 'yearly'>('monthly');

  const [isPurchaseModalOpen, setIsPurchaseModalOpen] = useState(false);
  const [isTransactionModalOpen, setIsTransactionModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [editingPurchase, setEditingPurchase] = useState<Purchase | null>(null);

  const [isMatchingModalOpen, setIsMatchingModalOpen] = useState(false);
  const [transactionToMatch, setTransactionToMatch] = useState<BankTransaction | null>(null);

  // ---------- Navigation helpers ----------
  const handlePreviousMonth = () => {
    setSelectedMonth(prev => {
      let [year, month] = prev.split('-').map(Number);
      if (month === 1) { month = 12; year -= 1; } else { month -= 1; }
      return `${year}-${String(month).padStart(2, '0')}`;
    });
  };

  const handleNextMonth = () => {
    setSelectedMonth(prev => {
      let [year, month] = prev.split('-').map(Number);
      if (month === 12) { month = 1; year += 1; } else { month += 1; }
      return `${year}-${String(month).padStart(2, '0')}`;
    });
  };

  // ---------- Helper: build months included in a view ----------
  // returns array of month keys like ["2025-01","2025-02",...]
  const getFilteredMonths = (selectedMonthKey: string): string[] => {
    const [year, month] = selectedMonthKey.split('-').map(Number);

    if (viewMode === 'monthly') {
      return [selectedMonthKey];
    }

    if (viewMode === 'quarterly') {
      const quarterStart = Math.floor((month - 1) / 3) * 3 + 1;
      return [0, 1, 2].map(i => `${year}-${String(quarterStart + i).padStart(2, '0')}`);
    }

    if (viewMode === 'half-yearly') {
      const halfStart = month <= 6 ? 1 : 7;
      return [0, 1, 2, 3, 4, 5].map(i => `${year}-${String(halfStart + i).padStart(2, '0')}`);
    }

    // yearly
    return Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, '0')}`);
  };

  // ---------- Optional: Friendly period label ----------
  const getPeriodLabel = (): string => {
    if (!selectedMonth) return '';
    const [y, m] = selectedMonth.split('-').map(Number);
    if (viewMode === 'monthly') {
      return new Date(selectedMonth + '-02').toLocaleString('default', { month: 'long', year: 'numeric' });
    }
    if (viewMode === 'quarterly') {
      const quarter = Math.floor((m - 1) / 3) + 1;
      const labels: Record<number, string> = { 1: 'Jan - Mar', 2: 'Apr - Jun', 3: 'Jul - Sep', 4: 'Oct - Dec' };
      return `Q${quarter} ${y} (${labels[quarter]})`;
    }
    if (viewMode === 'half-yearly') {
      const half = m <= 6 ? 1 : 2;
      return `H${half} ${y} (${half === 1 ? 'Jan - Jun' : 'Jul - Dec'})`;
    }
    return String(y);
  };

  // ---------- Filter purchases & transactions based on view ----------
  const { purchases, transactions } = useMemo(() => {
    const months = getFilteredMonths(selectedMonth);
    const filteredPurchases = months.flatMap(m => appData[m]?.purchases || []);
    const filteredTransactions = months.flatMap(m => appData[m]?.transactions || []);
    return { purchases: filteredPurchases, transactions: filteredTransactions };
  }, [appData, selectedMonth, viewMode]);

  // ---------- All months flat lists (for cross-month operations) ----------
  const allPurchases = useMemo(() => Object.values(appData).flatMap((d: MonthlyData) => d.purchases), [appData]);
  const allTransactions = useMemo(() => Object.values(appData).flatMap((d: MonthlyData) => d.transactions), [appData]);

  // ---------- UI action handlers (unchanged) ----------
  const handleAddPurchaseClick = () => { setEditingPurchase(null); setIsPurchaseModalOpen(true); };
  const handleAddTransactionClick = () => { setIsTransactionModalOpen(true); };
  const handleImportClick = () => { setIsImportModalOpen(true); };

  const handleExportData = () => {
    const periodLabelSafe = getPeriodLabel().replace(/[^a-zA-Z0-9]/g, '_') || selectedMonth.replace('-', '_');
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
    XLSX.writeFile(workbook, `Purchase_Report_${periodLabelSafe}.xlsx`);
  };

  const handleEditPurchaseClick = (purchase: Purchase) => { setEditingPurchase(purchase); setIsPurchaseModalOpen(true); };

  const handleSavePurchase = (purchaseData: Omit<Purchase, 'id'>) => {
    if (editingPurchase) {
      const updatedPurchase = {
        ...editingPurchase,
        ...purchaseData,
        id: editingPurchase.id,
        billImage: (purchaseData as any).billImage ? undefined : editingPurchase.billImage
      };
      updatePurchase(updatedPurchase);
    } else {
      addPurchase(purchaseData);
    }
    setIsPurchaseModalOpen(false);
    setEditingPurchase(null);
  };

  const handleUnmatch = useCallback(async (transactionId: string, purchaseId: string) => {
    const transaction = allTransactions.find(t => t.id === transactionId);
    const purchase = allPurchases.find(p => p.id === purchaseId);
    if (transaction && purchase) {
      await updateTransaction({ ...transaction, matchedPurchaseId: null });
      await updatePurchase({ ...purchase, matchedBankTransactionId: null, billImage: purchase.billImage });
    }
  }, [allTransactions, allPurchases, updateTransaction, updatePurchase]);

  const handleDeletePurchase = async (id: string) => {
    if (!window.confirm('Are you sure you want to delete this purchase?')) return;
    const purchaseToDelete = allPurchases.find(p => p.id === id);
    if (!purchaseToDelete) return;
    try {
      if (purchaseToDelete.matchedBankTransactionId) {
        await handleUnmatch(purchaseToDelete.matchedBankTransactionId, id);
      }
      await deletePurchase(purchaseToDelete);
    } catch (error) {
      console.error('Error while unmatching/deleting purchase:', error);
      alert('Failed to delete purchase. Please try again.');
    }
  };

  const handleSaveTransaction = (transactionData: Omit<BankTransaction, 'id'>) => {
    addTransaction(transactionData);
    setIsTransactionModalOpen(false);
  };

  const handleImportSave = (newTransactions: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>[]) => {
    addMultipleTransactions(newTransactions);
    setIsImportModalOpen(false);
  };

  const handleMatchClick = (transaction: BankTransaction) => { setTransactionToMatch(transaction); setIsMatchingModalOpen(true); };

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

  const totalExpenses = useMemo(() => purchases.reduce((total, p) => total + p.amount, 0).toFixed(2), [purchases]);

  const unmatchedPurchases = useMemo(() => {
    if (!transactionToMatch) return [];
    const monthKey = transactionToMatch.date.slice(0, 7); // YYYY-MM
    return appData[monthKey]?.purchases.filter(p => !p.matchedBankTransactionId) || [];
  }, [appData, transactionToMatch]);

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
        viewMode={viewMode}
        onViewModeChange={setViewMode}
      />

      <main className="w-full max-w-full mx-auto p-4 sm:p-6 lg:p-8">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <div className="lg:col-span-2 space-y-8">
            <PurchaseList purchases={purchases} onEdit={handleEditPurchaseClick} onDelete={handleDeletePurchase} />
            <BankTransactionList
              transactions={transactions}
              purchases={allPurchases}
              onMatch={handleMatchClick}   // <-- pass the click handler that receives a BankTransaction
              onUnmatch={handleUnmatch}    // <-- keep unmatch as-is (matches your signature)
            />
          </div>

          <div className="space-y-8">
            <Card>
              <h2 className="text-xl font-bold text-white mb-2">Total Expenses ({getPeriodLabel()})</h2>
              <p className="text-4xl font-extrabold text-indigo-400">₹{totalExpenses}</p>
            </Card>
            <CategoryChart purchases={purchases} />
            <Insights purchases={purchases} />
          </div>
        </div>
      </main>

      {isPurchaseModalOpen && (
        <PurchaseForm purchase={editingPurchase} onSave={handleSavePurchase} onClose={() => setIsPurchaseModalOpen(false)} context={context} />
      )}

      {isTransactionModalOpen && (
        <BankTransactionForm onSave={handleSaveTransaction} onClose={() => setIsTransactionModalOpen(false)} />
      )}

      {isImportModalOpen && (
        <ImportModal onSave={handleImportSave} onClose={() => setIsImportModalOpen(false)} />
      )}

      {isMatchingModalOpen && transactionToMatch && (
        <MatchModal transaction={transactionToMatch} unmatchedPurchases={unmatchedPurchases} onMatch={handleMatch} onClose={() => setIsMatchingModalOpen(false)} />
      )}
    </div>
  );
};

export default App;
