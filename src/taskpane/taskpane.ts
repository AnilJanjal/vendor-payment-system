// src/taskpane/taskpane.ts

import { AuthService } from './auth/auth';
import { PaymentService } from './payment/paymentService';

// Global elements
let loginButton: HTMLElement | null;
let logoutButton: HTMLElement | null;
let currentVendorIdToDelete: string | null = null;

// Declare Excel for TypeScript
declare const Excel: any;

// Global flag to track Excel availability
let excelEnabled = false;

// Main initialization
Office.onReady(() => {
    initializeAddIn();
});

async function initializeAddIn() {
    try {
        // Check if we're running in Excel
        if (Office.context.host === Office.HostType.Excel) {
            // Check Excel API support safely
            if (Office.context.requirements?.isSetSupported?.('ExcelApi', 1.1)) {
                excelEnabled = true;
                await initializeExcelTable();
            }
        }
        
        // Initialize payment service
        PaymentService.getPendingPayments();
        // Check for payments daily (86400000ms = 1 day)
        setInterval(() => PaymentService.processScheduledPayments(), 86400000);
        
        initializeUI();
    } catch (error) {
        console.error("Initialization failed:", error);
        initializeUI(); // Fallback to UI-only mode
    }
}

async function initializeExcelTable() {
    if (!excelEnabled) return;

    try {
        await Excel.run(async (context) => {
            // Get or create Vendors sheet
            let vendorsSheet;
            try {
                vendorsSheet = context.workbook.worksheets.getItem("Vendors");
            } catch {
                vendorsSheet = context.workbook.worksheets.add("Vendors");
                vendorsSheet.getRange("A1:D1").values = [["Vendor Name", "Payment Type", "Account", "Next Payment Date"]];
                vendorsSheet.getRange("A1:D1").format.font.bold = true;
            }

            // Clear and update existing data
            const vendors = AuthService.getVendors();
            if (vendors.length > 0) {
                vendorsSheet.getRange("A2:D" + (vendors.length + 1)).values = 
                    vendors.map(vendor => [
                        vendor.name,
                        vendor.paymentType,
                        vendor.account,
                        vendor.nextPaymentDate ? new Date(vendor.nextPaymentDate).toLocaleDateString() : "N/A"
                    ]);
            }

            await context.sync();
        });
    } catch (error) {
        console.warn("Excel table initialization failed:", error);
        excelEnabled = false;
    }
}


async function initializeUI() {
    loginButton?.removeEventListener('click', handleLogin);
    logoutButton?.removeEventListener('click', handleLogout);

    if (AuthService.isAuthenticated()) {
        showApp();
    } else {
        showLogin();
    }
}

async function syncVendorsToExcel() {
    if (!excelEnabled) return;

    try {
        await Excel.run(async (context) => {
            const vendorsSheet = context.workbook.worksheets.getItem("Vendors");
            const vendors = AuthService.getVendors();

            // Clear existing data (keep headers)
            if (vendors.length > 0) {
                vendorsSheet.getRange("A2:D" + (vendors.length + 1)).clear();
            } else {
                vendorsSheet.getRange("A2:D2").clear();
            }

            // Add new data
            vendors.forEach((vendor, index) => {
                const row = vendorsSheet.getRange(`A${index + 2}:D${index + 2}`);
                row.values = [
                    [
                        vendor.name,
                        vendor.paymentType,
                        vendor.account,
                        vendor.nextPaymentDate ? new Date(vendor.nextPaymentDate).toLocaleDateString() : "N/A"
                    ]
                ];
            });

            // Auto-fit columns
            vendorsSheet.getUsedRange().format.autofitColumns();
            await context.sync();
        });
    } catch (error) {
        console.warn("Excel sync failed:", error);
        excelEnabled = false;
    }
}


function showLogin() {
    const container = document.getElementById('container');
    if (!container) return;

    container.innerHTML = `
        <div id="login-section" class="banking-container">
            <div class="banking-header">
                <img src="../../assets/vendor.jpg" class="banking-logo" alt="Bank Logo">
                <h2 class="banking-title">Vendor Payment Portal</h2>
                <p class="banking-subtitle">Secure access to your payment dashboard</p>
            </div>
            
            <div class="banking-input">
                <label for="username">User ID</label>
                <input id="username" type="text" placeholder="Enter your user ID">
            </div>
            
            <div class="banking-input">
                <label for="password">Password</label>
                <input id="password" type="password" placeholder="Enter your password">
            </div>
            
            <button id="login-button" class="banking-button">Sign In</button>
            <p id="login-error" class="banking-error" style="display: none;"></p>
            
            <div class="banking-footer">
                <p>Need help? Contact your system administrator</p>
            </div>
        </div>
    `;

    loginButton = document.getElementById('login-button');
    loginButton?.addEventListener('click', handleLogin);
}

function handleLogin() {
    const username = (document.getElementById('username') as HTMLInputElement)?.value;
    const password = (document.getElementById('password') as HTMLInputElement)?.value;
    const errorElement = document.getElementById('login-error');

    if (!username || !password) {
        if (errorElement) {
            errorElement.textContent = 'Please enter both username and password';
            errorElement.style.display = 'block';
        }
        return;
    }

    if (AuthService.login(username, password)) {
        showApp();
    } else {
        if (errorElement) {
            errorElement.textContent = 'Invalid credentials';
            errorElement.style.display = 'block';
        }
    }
}

function handleLogout() {
    AuthService.logout();
    initializeUI();
}

function showApp() {
    const container = document.getElementById('container');
    if (!container) return;

    container.innerHTML = `
        <div class="banking-container" style="max-width: 800px;">
            <div class="banking-header">
                <img src="../../assets/vendor.jpg" class="banking-logo" alt="Vendor Logo">
                <h2 class="banking-title">Vendor Management System</h2>
                <button id="logout-button" class="banking-button" 
                    style="position: absolute; top: 20px; right: 20px; width: auto; padding: 8px 16px;">
                    Sign Out
                </button>
            </div>
            
            <div style="margin-top: 20px;">
                <div style="display: flex; gap: 10px; margin-bottom: 20px;">
                    <button id="add-vendor-button" class="banking-button">
                        + Add New Vendor
                    </button>
                    <button id="process-payments" class="banking-button">
                        Process Payments
                    </button>
                    <button id="retry-pending" class="banking-button">
                        Retry Pending
                    </button>
                </div>
                
                <div id="account-balances" style="display: flex; gap: 20px; margin-bottom: 20px;">
                    <div class="account-balance">
                        <h3>Account 1</h3>
                        <p>$${PaymentService.accountBalances['Account 1'].toLocaleString()}</p>
                    </div>
                    <div class="account-balance">
                        <h3>Account 2</h3>
                        <p>$${PaymentService.accountBalances['Account 2'].toLocaleString()}</p>
                    </div>
                </div>
                
                <div id="vendors-list" style="max-height: 400px; overflow-y: auto;">
                    ${getVendorsHtml()}
                </div>
                
                <div id="pending-payments" style="margin-top: 30px;">
                    <h3>Pending Payments</h3>
                    <div id="pending-list">
                        ${getPendingPaymentsHtml()}
                    </div>
                </div>
            </div>
        </div>
    `;

    setupVendorControls();
    logoutButton = document.getElementById('logout-button');
    logoutButton?.addEventListener('click', handleLogout);
    
    // Add payment event listeners
    // Update this in the showApp() function
// Update in showApp() function
document.getElementById('process-payments')?.addEventListener('click', async () => {
    try {
        const result = await PaymentService.processScheduledPayments(true);
        
        // Show results in UI
        const resultsHtml = `
            <div class="payment-results" style="margin: 15px 0; padding: 10px; background: #f3f3f3; border-radius: 4px;">
                <h4 style="margin: 0 0 10px 0;">Payment Results</h4>
                <p>Processed: ${result.processed} payments</p>
                <p>Skipped: ${result.skipped} vendors</p>
                <p>Pending: ${result.pending} payments</p>
            </div>
        `;
        
        const container = document.getElementById('pending-payments');
        if (container) {
            container.insertAdjacentHTML('afterbegin', resultsHtml);
            setTimeout(() => {
                const results = document.querySelector('.payment-results');
                results?.remove();
            }, 5000);
        }

        // Update all data
        updateAccountBalances();
        updateVendorsList();
        showPendingPayments();
        
        // Sync to all Excel sheets
        await syncVendorsToExcel();
        
    } catch (error) {
        console.error("Error processing payments:", error);
        alert("Error processing payments. Check console for details.");
    }
});
    
    document.getElementById('retry-pending')?.addEventListener('click', async () => {
        await PaymentService.retryPendingPayments();
        updateVendorsList();
        updateAccountBalances();
        showPendingPayments();
        syncVendorsToExcel();
    });
}

function getVendorsHtml(): string {
    const vendors = AuthService.getVendors();
    
    if (vendors.length === 0) {
        return `
            <div style="text-align: center; padding: 20px; color: #666;">
                No vendors found. Click "Add New Vendor" to create one.
            </div>
        `;
    }

    return vendors.map(vendor => `
        <div class="vendor-card" data-vendor-id="${vendor.id}">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <h3 style="margin: 0; color: #0078d4;">${vendor.name}</h3>
                    <p style="margin: 5px 0; color: #666;">
                        Payment: ${vendor.paymentType} | Account: ${vendor.account}
                    </p>
                    ${vendor.nextPaymentDate ? `
                    <p style="margin: 5px 0; color: #666; font-size: 0.9em;">
                        Next payment: ${new Date(vendor.nextPaymentDate).toLocaleDateString()}
                    </p>
                    ` : ''}
                    ${vendor.pendingPayment ? `
                    <p style="margin: 5px 0; color: #d13438; font-size: 0.9em;">
                        Payment pending (insufficient funds)
                    </p>
                    ` : ''}
                </div>
                <div style="display: flex; align-items: center;">
                    <button class="pay-now" data-id="${vendor.id}" 
                        style="background: none; border: none; color: #0078d4; cursor: pointer; margin-right: 10px;">
                        Pay Now
                    </button>
                    <button class="edit-vendor" data-id="${vendor.id}" 
                        style="background: none; border: none; color: #0078d4; cursor: pointer; margin-right: 10px;">
                        Edit
                    </button>
                    <button class="delete-vendor" data-id="${vendor.id}" 
                        style="background: none; border: none; color: #d13438; cursor: pointer;">
                        Delete
                    </button>
                </div>
            </div>
        </div>
    `).join('');
}

function getPendingPaymentsHtml(): string {
    const payments = PaymentService.getPendingPayments();
    const vendors = AuthService.getVendors();
    
    if (payments.length === 0) {
        return '<p style="color: #666;">No pending payments</p>';
    }

    return payments.map(payment => {
        const vendor = vendors.find(v => v.id === payment.vendorId);
        return `
            <div class="pending-payment" style="padding: 10px; border-bottom: 1px solid #eee;">
                <div style="display: flex; justify-content: space-between;">
                    <div>
                        <strong>${vendor?.name || 'Unknown Vendor'}</strong>
                        <p>Amount: $${payment.amount} | Account: ${payment.account}</p>
                        <p>Date: ${new Date(payment.date).toLocaleDateString()}</p>
                    </div>
                    <button class="retry-single" data-id="${payment.vendorId}" 
                        style="background: none; border: none; color: #0078d4; cursor: pointer;">
                        Retry
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

function setupVendorControls() {
    // Add Vendor Button
    document.getElementById('add-vendor-button')?.addEventListener('click', () => showVendorForm());
    
    // Vendor List Event Delegation
    document.getElementById('vendors-list')?.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        const vendorId = target.getAttribute('data-id');
        
        if (target.classList.contains('delete-vendor') && vendorId) {
            currentVendorIdToDelete = vendorId;
            showDeleteConfirmation();
        } else if (target.classList.contains('edit-vendor') && vendorId) {
            const vendor = AuthService.getVendors().find(v => v.id === vendorId);
            if (vendor) showVendorForm(vendor);
        } else if (target.classList.contains('pay-now') && vendorId) {
            handlePayNow(vendorId);
        }
    });
    
    // Pending Payments Event Delegation
    document.getElementById('pending-list')?.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        const vendorId = target.getAttribute('data-id');
        
        if (target.classList.contains('retry-single') && vendorId) {
            PaymentService.retryPaymentForVendor(vendorId).then(() => {
                updateAccountBalances();
                updateVendorsList();
                showPendingPayments();
                syncVendorsToExcel();
            });
        }
    });
    
    // Confirm Delete Button
    document.getElementById('confirm-delete')?.addEventListener('click', () => {
        if (currentVendorIdToDelete) {
            deleteVendor(currentVendorIdToDelete);
            hideDeleteConfirmation();
        }
    });
    
    // Cancel Delete Button
    document.getElementById('cancel-delete')?.addEventListener('click', hideDeleteConfirmation);
}

async function handlePayNow(vendorId: string) {
    const vendor = AuthService.getVendors().find(v => v.id === vendorId);
    if (!vendor) return;

    let amount = PaymentService.paymentRules.getBaseAmount(vendorId);
    
    // For on-demand vendors, prompt for amount
    if (PaymentService.paymentRules.getPaymentSchedule(vendorId) === 'on-demand') {
        const input = prompt(`Enter payment amount for ${vendor.name}:`, amount.toString());
        if (input === null) return; // User cancelled
        amount = parseFloat(input) || 0;
        if (amount <= 0) {
            alert('Please enter a valid positive amount');
            return;
        }
    }

    const result = await PaymentService.processOnDemandPayment(vendorId, amount);
    
    if (result) {
        updateAccountBalances();
        updateVendorsList();
        showPendingPayments();
        syncVendorsToExcel();
    }
}

function updateAccountBalances() {
    const account1 = document.querySelector('#account-balances div:nth-child(1) p');
    const account2 = document.querySelector('#account-balances div:nth-child(2) p');
    
    if (account1) account1.textContent = `$${PaymentService.accountBalances['Account 1'].toLocaleString()}`;
    if (account2) account2.textContent = `$${PaymentService.accountBalances['Account 2'].toLocaleString()}`;
}

function updateVendorsList() {
    const vendorsList = document.getElementById('vendors-list');
    if (vendorsList) {
        vendorsList.innerHTML = getVendorsHtml();
    }
}

function showPendingPayments() {
    const pendingList = document.getElementById('pending-list');
    if (pendingList) {
        pendingList.innerHTML = getPendingPaymentsHtml();
    }
}

function showDeleteConfirmation() {
    const dialog = document.getElementById('confirm-dialog');
    if (dialog) dialog.style.display = 'flex';
}

function hideDeleteConfirmation() {
    const dialog = document.getElementById('confirm-dialog');
    if (dialog) dialog.style.display = 'none';
    currentVendorIdToDelete = null;
}

function deleteVendor(vendorId: string) {
    try {
        AuthService.deleteVendor(vendorId);
        syncVendorsToExcel(); // Ensure Excel is updated
        
        const vendorCard = document.querySelector(`.vendor-card[data-vendor-id="${vendorId}"]`);
        vendorCard?.remove();
        
        if (AuthService.getVendors().length === 0) {
            const vendorsList = document.getElementById('vendors-list');
            if (vendorsList) {
                vendorsList.innerHTML = `
                    <div style="text-align: center; padding: 20px; color: #666;">
                        No vendors found. Click "Add New Vendor" to create one.
                    </div>
                `;
            }
        }
    } catch (error) {
        console.error('Error deleting vendor:', error);
        alert('Error deleting vendor. Please try again.');
    }
}

function showVendorForm(vendor?: Vendor) {
    const container = document.getElementById('container');
    if (!container) return;

    container.innerHTML = `
        <div class="banking-container" style="max-width: 600px;">
            <div class="banking-header">
                <h2 class="banking-title">${vendor ? 'Edit Vendor' : 'Add New Vendor'}</h2>
            </div>
            
            <div style="margin-top: 20px;">
                <div class="banking-input">
                    <label>Vendor Name</label>
                    <input id="vendor-name" type="text" value="${vendor?.name || ''}" 
                        placeholder="Enter vendor name" style="width: 100%;">
                </div>
                
                <div class="banking-input">
                    <label>Payment Type</label>
                    <select id="payment-type" style="width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px;">
                        <option value="Weekly" ${vendor?.paymentType === 'Weekly' ? 'selected' : ''}>Weekly</option>
                        <option value="Biweekly" ${vendor?.paymentType === 'Biweekly' ? 'selected' : ''}>Biweekly</option>
                        <option value="On-Demand" ${vendor?.paymentType === 'On-Demand' ? 'selected' : ''}>On-Demand</option>
                    </select>
                </div>
                
                <div class="banking-input">
                    <label>Assigned Account</label>
                    <select id="assigned-account" style="width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 4px;">
                        <option value="Account 1" ${vendor?.account === 'Account 1' ? 'selected' : ''}>Account 1</option>
                        <option value="Account 2" ${vendor?.account === 'Account 2' ? 'selected' : ''}>Account 2</option>
                    </select>
                </div>
                
                <div style="display: flex; gap: 10px; margin-top: 30px;">
                    <button id="save-vendor" class="banking-button" style="flex: 1;">
                        ${vendor ? 'Update' : 'Save'} Vendor
                    </button>
                    <button id="cancel-form" class="banking-button" 
                        style="flex: 1; background: #f3f3f3; color: #333; border: 1px solid #ddd;">
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    `;

    document.getElementById('save-vendor')?.addEventListener('click', async () => {
        const nameInput = document.getElementById('vendor-name') as HTMLInputElement;
        const paymentTypeSelect = document.getElementById('payment-type') as HTMLSelectElement;
        const accountSelect = document.getElementById('assigned-account') as HTMLSelectElement;

        if (!nameInput.value.trim()) {
            alert('Please enter a vendor name');
            return;
        }

        if (vendor) {
            // Update existing vendor
            const updatedVendor = {
                ...vendor,
                name: nameInput.value.trim(),
                paymentType: paymentTypeSelect.value as 'Weekly' | 'Biweekly' | 'On-Demand',
                account: accountSelect.value as 'Account 1' | 'Account 2'
            };
            AuthService.updateVendor(updatedVendor);
        } else {
            // Create new vendor
            AuthService.addVendor({
                name: nameInput.value.trim(),
                paymentType: paymentTypeSelect.value as 'Weekly' | 'Biweekly' | 'On-Demand',
                account: accountSelect.value as 'Account 1' | 'Account 2'
            });
        }
        
        await syncVendorsToExcel(); // Ensure Excel is updated
        showApp();
    });

    document.getElementById('cancel-form')?.addEventListener('click', showApp);
}


interface Vendor {
    id: string;
    name: string;
    paymentType: 'Weekly' | 'Biweekly' | 'On-Demand';
    account: 'Account 1' | 'Account 2';
    lastPaymentDate?: Date;
    nextPaymentDate?: Date;
    pendingPayment?: boolean;
}