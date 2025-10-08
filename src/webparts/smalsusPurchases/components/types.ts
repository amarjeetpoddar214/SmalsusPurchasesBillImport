export enum Category {
  Groceries = 'Groceries',
  Utilities = 'Utilities',
  Entertainment = 'Entertainment',
  Transport = 'Transport',
  Health = 'Health',
  Shopping = 'Shopping',
  FoodAndDining = 'Food & Dining',
  Other = 'Other',
}

export interface Purchase {
  id: string;
  name: string;
  amount: number;
  date: string; // Storing as YYYY-MM-DD for input compatibility
  category: Category;
  billImage?: string; // Base64 encoded image string
  matchedBankTransactionId?: string | null;
  matchedBankTransactionTitle?: string;
}

export interface BankTransaction {
  id: string;
  description: string;
  amount: number;
  date: string;
  type: 'debit' | 'credit';
  matchedPurchaseId?: string | null;
}
