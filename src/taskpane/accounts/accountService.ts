// src/taskpane/account/accountService.ts

import { AuthService } from '../auth/auth';

interface Account {
    name: string;
    balance: number;
}

interface Transaction {
    account: string;
    amount: number;
    date: Date;
    vendorId: string;
    vendorName: string;
    type: 'payment' | 'deposit' | 'adjustment';
}

export const AccountService = {
    // Initialize accounts
    accounts: {
        'Account 1': { name: 'Account 1', balance: 200000 },
        'Account 2': { name: 'Account 2', balance: 200000 }
    },

    // Transaction history
    transactions: [] as Transaction[],

    // Initialize accounts from storage
    initialize(): void {
        const storedAccounts = localStorage.getItem('accounts');
        if (storedAccounts) {
            this.accounts = JSON.parse(storedAccounts);
        }

        const storedTransactions = localStorage.getItem('transactions');
        if (storedTransactions) {
            this.transactions = JSON.parse(storedTransactions);
        }
    },

    // Get account balance
    getBalance(accountName: string): number {
        return this.accounts[accountName]?.balance || 0;
    },

    // Process payment
    processPayment(accountName: string, amount: number, vendorId: string): boolean {
        if (!this.accounts[accountName]) return false;
        if (this.accounts[accountName].balance < amount) return false;

        this.accounts[accountName].balance -= amount;
        
        const vendor = AuthService.getVendors().find(v => v.id === vendorId);
        this.transactions.push({
            account: accountName,
            amount,
            date: new Date(),
            vendorId,
            vendorName: vendor?.name || 'Unknown Vendor',
            type: 'payment'
        });

        this.saveToStorage();
        return true;
    },

    // Add funds to account
    deposit(accountName: string, amount: number): boolean {
        if (!this.accounts[accountName]) return false;
        
        this.accounts[accountName].balance += amount;
        this.transactions.push({
            account: accountName,
            amount,
            date: new Date(),
            vendorId: '',
            vendorName: 'System',
            type: 'deposit'
        });

        this.saveToStorage();
        return true;
    },

    // Manual balance adjustment
    adjustBalance(accountName: string, newBalance: number): boolean {
        if (!this.accounts[accountName]) return false;
        
        const difference = newBalance - this.accounts[accountName].balance;
        this.accounts[accountName].balance = newBalance;
        
        this.transactions.push({
            account: accountName,
            amount: difference,
            date: new Date(),
            vendorId: '',
            vendorName: 'System',
            type: 'adjustment'
        });

        this.saveToStorage();
        return true;
    },

    // Get transaction history
    getTransactions(accountName?: string): Transaction[] {
        if (accountName) {
            return this.transactions.filter(t => t.account === accountName);
        }
        return this.transactions;
    },

    // Save to localStorage
    saveToStorage(): void {
        localStorage.setItem('accounts', JSON.stringify(this.accounts));
        localStorage.setItem('transactions', JSON.stringify(this.transactions));
    },

    // Update UI balances
    updateAccountBalancesUI(): void {
        const account1Element = document.querySelector('#account-balances div:nth-child(1) p');
        const account2Element = document.querySelector('#account-balances div:nth-child(2) p');
        
        if (account1Element) {
            account1Element.textContent = `$${this.accounts['Account 1'].balance.toLocaleString()}`;
        }
        if (account2Element) {
            account2Element.textContent = `$${this.accounts['Account 2'].balance.toLocaleString()}`;
        }
    },

    // Sync balances to Excel
    async syncBalancesToExcel(): Promise<void> {
        try {
            await Excel.run(async (context) => {
                // Get or create Accounts sheet
                let accountsSheet;
                try {
                    accountsSheet = context.workbook.worksheets.getItem("Accounts");
                } catch {
                    accountsSheet = context.workbook.worksheets.add("Accounts");
                    accountsSheet.getRange("A1:B1").values = [["Account Name", "Balance"]];
                    accountsSheet.getRange("A1:B1").format.font.bold = true;
                }

                // Update account balances
                accountsSheet.getRange("A2:B3").values = [
                    ["Account 1", this.accounts['Account 1'].balance],
                    ["Account 2", this.accounts['Account 2'].balance]
                ];

                // Auto-fit columns
                accountsSheet.getUsedRange().format.autofitColumns();
                await context.sync();
            });
        } catch (error) {
            console.error("Error syncing balances to Excel:", error);
        }
    }
};

// Initialize accounts when module loads
AccountService.initialize();