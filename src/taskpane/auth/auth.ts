// src/taskpane/auth/auth.ts
localStorage.removeItem('vendors');
interface Vendor {
    id: string;
    name: string;
    paymentType: 'Weekly' | 'Biweekly' | 'On-Demand';
    account: 'Account 1' | 'Account 2';
    lastPaymentDate?: Date;
    nextPaymentDate?: Date;
    pendingPayment?: boolean;
}

export const AuthService = {
    // Existing authentication methods
    login(username: string, password: string): boolean {
        if (username === "aniljanjal2000@gmail.com" && password === "aniljanjal2000@gmail.com") {
            localStorage.setItem("isAuthenticated", "true");
            return true;
        }
        return false;
    },

    logout(): void {
        localStorage.removeItem("isAuthenticated");
    },

    isAuthenticated(): boolean {
        return localStorage.getItem("isAuthenticated") === "true";
    },

    // Vendor management methods
    getVendors(): Vendor[] {
        const vendors = localStorage.getItem('vendors');
        return vendors ? JSON.parse(vendors) : [];
    },

    saveVendors(vendors: Vendor[]): void {
        localStorage.setItem('vendors', JSON.stringify(vendors));
    },

    addVendor(vendor: Omit<Vendor, 'id'>): Vendor {
        const vendors = this.getVendors();
        const newVendor = {
            ...vendor,
            id: Date.now().toString()
        };
        vendors.push(newVendor);
        this.saveVendors(vendors);
        return newVendor;
    },

    updateVendor(updatedVendor: Vendor): void {
        const vendors = this.getVendors();
        const index = vendors.findIndex(v => v.id === updatedVendor.id);
        if (index !== -1) {
            vendors[index] = updatedVendor;
            this.saveVendors(vendors);
        }
    },

    deleteVendor(id: string): void {
    const vendors = this.getVendors().filter(v => v.id !== id);
    this.saveVendors(vendors);
}
};