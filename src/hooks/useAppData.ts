import { useState, useEffect, useCallback } from 'react';
import { Purchase, BankTransaction } from '../webparts/smalsusPurchases/components/types';
import { Web } from 'sp-pnp-js';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface MonthlyData {
    purchases: Purchase[];
    transactions: BankTransaction[];
}

export interface AppData {
    [key: string]: MonthlyData; // YYYY-MM
}

const uuidv4 = () => {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

export const useAppData = (context: WebPartContext) => {
    console.log('Context Data', context);


    if (!context) {
        console.warn('useAppData: WebPartContext not provided. SharePoint calls will fail until context is passed.');
    }

    const webURL = context.pageContext.web.absoluteUrl;
    const purchasesListId = '47460239-25cf-4928-ba07-b8b75237c714';
    const transactionsListId = '38dbeddd-4e03-4d26-83a7-90305de6046e';

    const [appData, setAppData] = useState<AppData>({});
    const [isLoaded, setIsLoaded] = useState(false);



    const fetchAppDataFromSP = async () => {
        try {
            const web = new Web(webURL);

            // Fetch Purchases
            const purchasesRes = await web.lists.getById(purchasesListId)
                .items
                .select('Id', 'Title', 'amount', 'date', 'category', 'matchTransaction/Id', 'matchTransaction/Title', 'BillImageUrl')
                .expand('matchTransaction')
                .getAll();
            console.log('PurchaseList', purchasesRes);

            const mappedPurchases: Purchase[] = purchasesRes.map(item => ({
                id: `p-${item.Id}`,
                name: item.Title,
                amount: Number(item.amount),
                date: new Date(item.date).toISOString().slice(0, 10),
                category: item.category || 'Other',
                // store the transaction id in the same format we use for transactions: 't-<id>'
                matchedBankTransactionId: item.matchTransaction?.Id ? `t-${item.matchTransaction.Id}` : null,
                matchedBankTransactionTitle: item.matchTransaction?.Title || '',
                billImage: item.BillImageUrl?.Url || undefined
            }));


            // Fetch Transactions
            const transactionsRes = await web.lists.getById(transactionsListId).items.getAll();
            console.log('Transaction List', transactionsRes);

            const mappedTransactions: BankTransaction[] = transactionsRes.map(item => ({
                id: `t-${item.Id}`,
                description: item.Title,
                amount: Number(item.amount),
                date: new Date(item.date).toISOString().slice(0, 10),
                type: (item.typeChoice || 'debit').toLowerCase() === 'credit' ? 'credit' : 'debit',
                matchedPurchaseId: null // fill in from purchases below
            }));

            // Link transactions -> purchases so UI knows a transaction's matched purchase
            mappedTransactions.forEach(t => {
                const matchedPurchase = mappedPurchases.find(p => p.matchedBankTransactionId === t.id);
                if (matchedPurchase) {
                    t.matchedPurchaseId = matchedPurchase.id;
                }
            });


            // Convert to AppData grouped by month
            const newAppData: AppData = {};

            mappedPurchases.forEach(p => {
                const monthKey = p.date.slice(0, 7);
                if (!newAppData[monthKey]) newAppData[monthKey] = { purchases: [], transactions: [] };
                newAppData[monthKey].purchases.push(p);
            });

            mappedTransactions.forEach(t => {
                const monthKey = t.date.slice(0, 7);
                if (!newAppData[monthKey]) newAppData[monthKey] = { purchases: [], transactions: [] };
                newAppData[monthKey].transactions.push(t);
            });

            // Sort each month's purchases and transactions by date
            Object.keys(newAppData).forEach(month => {
                newAppData[month].purchases.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
                newAppData[month].transactions.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
            });

            // Set state
            setAppData(newAppData);
        } catch (error) {
            console.error('Error fetching data from SharePoint', error);
        }
    };


    useEffect(() => {
        const loadData = async () => {
            try {
                await fetchAppDataFromSP(); // primary: fetch fresh data from SharePoint
            } catch (err) {
                console.error(err);
            } finally {
                setIsLoaded(true); // mark data as loaded in any case
            }
        };

        loadData();
    }, []);


    const getMonthKey = (date: string) => date.slice(0, 7); // YYYY-MM

    const addPurchase = useCallback(
        async (purchaseData: Omit<Purchase, 'id'> & { billFile?: File }) => {
            setIsLoaded(false);
            try {
                const web = new Web(webURL);
                let uploadedFileUrl: string | undefined = undefined;

                // compute a reliable server-relative path for the library folder
                const siteServerRelativeUrl = new URL(webURL).pathname.replace(/\/$/, ''); // e.g. '/sites/Smalsus'
                const folderServerRelativeUrl = `${siteServerRelativeUrl}/BillImages`; // update 'BillImages' if your library is named differently

                // 1️⃣ Upload file if exists (use raw File - works for PDFs & images)
                if (purchaseData.billFile) {
                    console.log('Uploading file to folder:', folderServerRelativeUrl, purchaseData.billFile.name);
                    const uploadedFile = await web
                        .getFolderByServerRelativeUrl(folderServerRelativeUrl)
                        .files.add(purchaseData.billFile.name, purchaseData.billFile, true);

                    console.log('Upload response:', uploadedFile);
                    uploadedFileUrl = uploadedFile.data.ServerRelativeUrl; // server-relative URL, e.g. /sites/Smalsus/BillImages/receipt.pdf
                    console.log('Uploaded file serverRelativeUrl:', uploadedFileUrl);
                }

                // 2️⃣ Save item to SharePoint (use SP.FieldUrlValue for Hyperlink/Picture column)
                const spItem = await web.lists.getById(purchasesListId).items.add({
                    Title: purchaseData.name,
                    amount: purchaseData.amount,
                    date: purchaseData.date,
                    category: purchaseData.category,
                    matchTransactionId: purchaseData.matchedBankTransactionId
                        ? Number(String(purchaseData.matchedBankTransactionId).replace(/^t-/, ''))
                        : null,
                    BillImageUrl: uploadedFileUrl
                        ? {
                            __metadata: { type: 'SP.FieldUrlValue' },
                            Url: uploadedFileUrl,
                            Description: 'Bill/Receipt'
                        }
                        : null
                });

                // 3️⃣ Map SP item to local Purchase and update state
                const newPurchase: Purchase = {
                    id: `p-${spItem.data.Id}`,
                    name: purchaseData.name,
                    amount: purchaseData.amount,
                    date: purchaseData.date,
                    category: purchaseData.category,
                    matchedBankTransactionId: purchaseData.matchedBankTransactionId || null,
                    matchedBankTransactionTitle: purchaseData.matchedBankTransactionTitle || '',
                    billImage: uploadedFileUrl // store server-relative URL for display
                };

                // 4️⃣ Update state grouped by month
                const monthKey = getMonthKey(newPurchase.date);
                setAppData(prev => {
                    const newAppData = { ...prev };
                    const monthData = newAppData[monthKey] || { purchases: [], transactions: [] };
                    monthData.purchases = [...monthData.purchases, newPurchase].sort(
                        (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()
                    );
                    newAppData[monthKey] = monthData;
                    return newAppData;
                });
            } catch (error) {
                console.error("Error adding purchase:", error);
                alert("Failed to add purchase. Please try again.");
            } finally {
                setIsLoaded(true);
            }
        },
        [webURL, purchasesListId] // add dependencies
    );





    const updatePurchase = useCallback(
        async (updatedPurchase: Purchase & { billFile?: File }) => {
            setIsLoaded(false);
            try {
                const web = new Web(webURL);
                const siteServerRelativeUrl = new URL(webURL).pathname.replace(/\/$/, '');
                const folderServerRelativeUrl = `${siteServerRelativeUrl}/BillImages`;

                let uploadedFileUrl: string | undefined = undefined;

                // Upload new bill file if provided
                if (updatedPurchase.billFile) {
                    const uploadedFile = await web
                        .getFolderByServerRelativeUrl(folderServerRelativeUrl)
                        .files.add(updatedPurchase.billFile.name, updatedPurchase.billFile, true);
                    uploadedFileUrl = uploadedFile.data.ServerRelativeUrl;
                }

                // Prepare SharePoint payload
                const updatePayload: any = {
                    Title: updatedPurchase.name,
                    amount: updatedPurchase.amount,
                    date: updatedPurchase.date,
                    category: updatedPurchase.category,
                    matchTransactionId: updatedPurchase.matchedBankTransactionId
                        ? Number(String(updatedPurchase.matchedBankTransactionId).replace(/^t-/, ''))
                        : null
                };

                if (uploadedFileUrl) {
                    updatePayload.BillImageUrl = {
                        __metadata: { type: 'SP.FieldUrlValue' },
                        Url: uploadedFileUrl,
                        Description: 'Bill/Receipt'
                    };
                }

                // Update SharePoint item
                await web.lists
                    .getById(purchasesListId)
                    .items.getById(Number(updatedPurchase.id.replace('p-', '')))
                    .update(updatePayload);

                // Use new file URL if uploaded, else keep existing
                const finalBillUrl = uploadedFileUrl ?? updatedPurchase.billImage;

                // Update local state
                setAppData(prev => {
                    const newAppData = { ...prev };

                    // Remove old purchase from all months
                    for (const month in newAppData) {
                        if (Object.prototype.hasOwnProperty.call(newAppData, month)) {
                            newAppData[month].purchases = newAppData[month].purchases.filter(
                                p => p.id !== updatedPurchase.id
                            );
                        }
                    }

                    // Add updated purchase to correct month
                    const monthKey = updatedPurchase.date.slice(0, 7);
                    const monthData = newAppData[monthKey] || { purchases: [], transactions: [] };
                    monthData.purchases = [
                        ...monthData.purchases,
                        { ...updatedPurchase, billImage: finalBillUrl } // <-- use finalBillUrl
                    ].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
                    newAppData[monthKey] = monthData;

                    return newAppData;
                });
            } catch (error) {
                console.error("Error updating purchase:", error);
                alert("Failed to update purchase. Please try again.");
            } finally {
                setIsLoaded(true);
            }
        },
        [webURL, purchasesListId]
    );






    const deletePurchase = useCallback(
        async (purchaseToDelete: Purchase) => {
            setIsLoaded(false); // show loading while deleting

            try {
                const web = new Web(webURL);

                // 1. Delete from SharePoint
                await web.lists
                    .getById(purchasesListId) // Purchases list ID
                    .items.getById(Number(purchaseToDelete.id.replace('p-', '')))
                    .delete();

                // 2. Update local state only after successful deletion
                setAppData(prev => {
                    const newAppData = { ...prev };
                    const monthKey = getMonthKey(purchaseToDelete.date);
                    const monthData = newAppData[monthKey];
                    if (monthData) {
                        monthData.purchases = monthData.purchases.filter(
                            p => p.id !== purchaseToDelete.id
                        );
                        newAppData[monthKey] = { ...monthData };
                    }
                    return newAppData;
                });

            } catch (error) {
                console.error('Error deleting purchase:', error);
                alert('Failed to delete purchase. Please try again.');
            } finally {
                setIsLoaded(true); // stop loading
            }
        },
        []
    );


    const addTransaction = useCallback(
        async (transactionData: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>) => {
            setIsLoaded(false); // show loading until transaction is saved

            try {
                const newTransaction: BankTransaction = {
                    ...transactionData,
                    id: uuidv4(), // temporary ID until SharePoint gives one
                    matchedPurchaseId: null
                };

                const web = new Web(webURL);
                const spItem = await web.lists.getById(transactionsListId).items.add({
                    Title: newTransaction.description,
                    amount: newTransaction.amount,
                    date: newTransaction.date,
                    typeChoice: newTransaction.type
                });

                newTransaction.id = `t-${spItem.data.Id}`; // standard 't-<SPId>' id

                const monthKey = getMonthKey(newTransaction.date);
                setAppData(prev => {
                    const newAppData = { ...prev };
                    const monthData = newAppData[monthKey] || { purchases: [], transactions: [] };
                    monthData.transactions = [...monthData.transactions, newTransaction].sort(
                        (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()
                    );
                    newAppData[monthKey] = monthData;
                    return newAppData;
                });
            } catch (error) {
                console.error("Error adding transaction:", error);
                alert("Failed to add transaction. Please try again.");
            } finally {
                setIsLoaded(true); // stop loading
            }
        },
        []
    );



    const addMultipleTransactions = useCallback(
        async (transactionsData: Omit<BankTransaction, 'id' | 'matchedPurchaseId'>[]) => {
            setIsLoaded(false);

            try {
                const web = new Web(webURL);

                const newTransactions: BankTransaction[] = [];

                // 1️⃣ Add each transaction to SharePoint
                for (const transactionData of transactionsData) {
                    const spItem = await web.lists
                        .getById(transactionsListId)
                        .items.add({
                            Title: transactionData.description,
                            amount: transactionData.amount,
                            date: transactionData.date,
                            typeChoice: transactionData.type
                        });

                    const newTransaction: BankTransaction = {
                        ...transactionData,
                        id: `t-${spItem.data.Id}`, // Use SharePoint ID
                        matchedPurchaseId: null
                    };

                    newTransactions.push(newTransaction);
                }

                // 2️⃣ Add all new transactions into state grouped by month
                setAppData(prev => {
                    const newAppData = { ...prev };

                    newTransactions.forEach(t => {
                        const monthKey = getMonthKey(t.date);
                        const monthData = newAppData[monthKey] || { purchases: [], transactions: [] };
                        monthData.transactions.push(t);
                        // sort newest first
                        monthData.transactions.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
                        newAppData[monthKey] = monthData;
                    });

                    return newAppData;
                });
            } catch (error) {
                console.error('Error adding multiple transactions:', error);
                alert('Failed to add multiple transactions. Please try again.');
            } finally {
                setIsLoaded(true);
            }
        },
        []
    );


    const updateTransaction = useCallback(
        async (updatedTransaction: BankTransaction) => {
            setIsLoaded(false); // show loading

            try {
                const web = new Web(webURL);

                // 1️⃣ Update SharePoint
                await web.lists
                    .getById(transactionsListId) // Transactions list ID
                    .items.getById(Number(updatedTransaction.id.replace('t-', '')))
                    .update({
                        Title: updatedTransaction.description,
                        amount: updatedTransaction.amount,
                        date: updatedTransaction.date,
                        typeChoice: updatedTransaction.type
                    });

                // 2️⃣ Update local state
                setAppData(prev => {
                    const newAppData = { ...prev };

                    // Remove transaction from all months
                    for (const month in newAppData) {
                        if (Object.prototype.hasOwnProperty.call(newAppData, month)) {
                            newAppData[month].transactions = newAppData[month].transactions.filter(
                                t => t.id !== updatedTransaction.id
                            );
                        }
                    }

                    // Add to correct month
                    const monthKey = getMonthKey(updatedTransaction.date);
                    const monthData = newAppData[monthKey] || { purchases: [], transactions: [] };
                    monthData.transactions = [...monthData.transactions, updatedTransaction].sort(
                        (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()
                    );
                    newAppData[monthKey] = monthData;

                    return newAppData;
                });
            } catch (error) {
                console.error('Error updating transaction:', error);
                alert('Failed to update transaction. Please try again.');
            } finally {
                setIsLoaded(true); // stop loading
            }
        },
        []
    );



    return { appData, isLoaded, addPurchase, updatePurchase, deletePurchase, addTransaction, updateTransaction, addMultipleTransactions };
};