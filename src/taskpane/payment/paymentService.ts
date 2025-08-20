// src/taskpane/payment/paymentService.ts

import { AuthService } from '../auth/auth';

// Define the Vendor interface that matches AuthService's Vendor structure
interface Vendor {
    id: string;
    name: string;
    paymentType: 'Weekly' | 'Biweekly' | 'On-Demand';
    account: 'Account 1' | 'Account 2';
    lastPaymentDate?: Date;
    nextPaymentDate?: Date;
    pendingPayment?: boolean;
}

interface Payment {
    vendorId: string;
    amount: number;
    date: Date;
    account: 'Account 1' | 'Account 2';
    status: 'completed' | 'pending' | 'failed';
}

export const PaymentService = {
    // Account balances
    accountBalances: {
        'Account 1': 200000,
        'Account 2': 200000
    },

    // Payment rules
    paymentRules: {
        getBaseAmount(vendorId: string): number {
            const idNum = parseInt(vendorId);
            if (idNum >= 1 && idNum <= 5) return 100; // Weekly
            if (idNum >= 6 && idNum <= 10) return 200; // Biweekly (double amount)
            return 0; // On-demand (amount specified manually)
        },
        
        getPaymentSchedule(vendorId: string): 'weekly' | 'biweekly' | 'on-demand' {
            const idNum = parseInt(vendorId);
            if (idNum >= 1 && idNum <= 5) return 'weekly';
            if (idNum >= 6 && idNum <= 10) return 'biweekly';
            return 'on-demand';
        },

        getDefaultAccount(vendorId: string): 'Account 1' | 'Account 2' {
            const idNum = parseInt(vendorId);
            if (idNum >= 1 && idNum <= 10) return 'Account 1'; // Scheduled vendors
            return 'Account 2'; // On-demand vendors
        }
    },

    // Process scheduled payments
    async processScheduledPayments(forceProcess = false): Promise<{processed: number, skipped: number, pending: number}> {
    const today = new Date();
    console.log(`Running payment check on ${today.toDateString()}`);

    // Only check for Friday if not forced
    if (!forceProcess && today.getDay() !== 5) {
        console.log("Not Friday - skipping payment processing");
        return {processed: 0, skipped: 0, pending: 0};
    }

    const vendors = AuthService.getVendors() as Vendor[];
    const result = {
        processed: 0,
        skipped: 0,
        pending: this.getPendingPayments().length
    };

    const payments: Payment[] = [];

    vendors.forEach(vendor => {
        const schedule = this.paymentRules.getPaymentSchedule(vendor.id);
        
        if (schedule === 'on-demand') {
            result.skipped++;
            return;
        }

        // Process if payment is due or we're forcing processing
        if (forceProcess || (vendor.nextPaymentDate && new Date(vendor.nextPaymentDate).toDateString() === today.toDateString())) {
            const amount = this.paymentRules.getBaseAmount(vendor.id);
            payments.push(this.createPayment(vendor, amount));
        } else {
            result.skipped++;
        }
    });

    if (payments.length > 0) {
        await this.processPayments(payments);
        result.processed = payments.filter(p => p.status === 'completed').length;
        result.pending = this.getPendingPayments().length;
    }

    return result;
},

    // Process on-demand payment
    async processOnDemandPayment(vendorId: string, amount: number): Promise<boolean> {
        const vendor = AuthService.getVendors().find(v => v.id === vendorId) as Vendor | undefined;
        if (!vendor) return false;

        // Check if scheduled vendor is being paid on-demand
        const schedule = this.paymentRules.getPaymentSchedule(vendorId);
        if (schedule !== 'on-demand') {
            const skipScheduled = confirm(`${vendor.name} is scheduled for ${schedule} payments. Skip next scheduled payment?`);
            if (skipScheduled) {
                vendor.lastPaymentDate = new Date();
                vendor.nextPaymentDate = this.getNextPaymentDate(vendorId);
                AuthService.updateVendor(vendor);
            }
        }

        const payment = this.createPayment(vendor, amount);
        return this.processPayment(payment);
    },

    // Retry payment for a specific vendor
    async retryPaymentForVendor(vendorId: string): Promise<boolean> {
        const pendingPayment = this.pendingPayments.find(p => p.vendorId === vendorId);
        if (!pendingPayment) return false;

        const success = await this.processPayment(pendingPayment);
        if (success) {
            // Remove from pending payments if successful
            this.pendingPayments = this.pendingPayments.filter(p => p.vendorId !== vendorId);
            localStorage.setItem('pendingPayments', JSON.stringify(this.pendingPayments));
        }
        return success;
    },

    // Helper methods
    createPayment(vendor: Vendor, amount: number): Payment {
        return {
            vendorId: vendor.id,
            amount,
            date: new Date(),
            account: vendor.account,
            status: 'pending'
        };
    },

    async processPayments(payments: Payment[]): Promise<void> {
        for (const payment of payments) {
            await this.processPayment(payment);
        }
    },

    async processPayment(payment: Payment): Promise<boolean> {
        if (this.accountBalances[payment.account] >= payment.amount) {
            // Sufficient funds
            this.accountBalances[payment.account] -= payment.amount;
            payment.status = 'completed';
            this.updateVendorPaymentDate(payment.vendorId);
            
            // Update UI
            this.updateAccountBalancesUI();
            return true;
        } else {
            // Insufficient funds
            payment.status = 'pending';
            this.addToPendingPayments(payment);
            
            // Show notification
            const vendor = AuthService.getVendors().find(v => v.id === payment.vendorId);
            if (vendor) {
                alert(`Insufficient funds in ${payment.account} to pay ${vendor.name}. Payment moved to pending.`);
            }
            
            return false;
        }
    },

    updateVendorPaymentDate(vendorId: string): void {
        const vendor = AuthService.getVendors().find(v => v.id === vendorId) as Vendor | undefined;
        if (vendor) {
            vendor.lastPaymentDate = new Date();
            vendor.nextPaymentDate = this.getNextPaymentDate(vendorId);
            vendor.pendingPayment = false;
            AuthService.updateVendor(vendor);
        }
    },

    getNextPaymentDate(vendorId: string): Date {
        const date = new Date();
        const schedule = this.paymentRules.getPaymentSchedule(vendorId);
        
        if (schedule === 'weekly') {
            date.setDate(date.getDate() + 7);
        } else if (schedule === 'biweekly') {
            date.setDate(date.getDate() + 14);
        }
        
        return date;
    },

    getWeekNumber(date: Date): number {
        const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
        const pastDaysOfYear = (date.getTime() - firstDayOfYear.getTime()) / 86400000;
        return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
    },

    pendingPayments: [] as Payment[],

    addToPendingPayments(payment: Payment): void {
        this.pendingPayments.push(payment);
        localStorage.setItem('pendingPayments', JSON.stringify(this.pendingPayments));
    },

    getPendingPayments(): Payment[] {
        const stored = localStorage.getItem('pendingPayments');
        this.pendingPayments = stored ? JSON.parse(stored) : [];
        return this.pendingPayments;
    },

    async retryPendingPayments(): Promise<void> {
        const pending = this.getPendingPayments();
        for (const payment of pending) {
            if (this.accountBalances[payment.account] >= payment.amount) {
                await this.processPayment(payment);
            }
        }
        this.updateAccountBalancesUI();
    },

    updateAccountBalancesUI(): void {
        const account1Element = document.querySelector('#account-balances div:nth-child(1) p');
        const account2Element = document.querySelector('#account-balances div:nth-child(2) p');
        
        if (account1Element) {
            account1Element.textContent = `$${this.accountBalances['Account 1'].toLocaleString()}`;
        }
        if (account2Element) {
            account2Element.textContent = `$${this.accountBalances['Account 2'].toLocaleString()}`;
        }
    }
};