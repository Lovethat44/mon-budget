document.addEventListener('DOMContentLoaded', async () => {

    // =========================================================================
    // CONFIGURATION MICROSOFT ONEDRIVE (MSAL)
    // =========================================================================
    const msalConfig = {
        auth: {
            clientId: "TON_ID_CLIENT_ICI", // <--- REMPLACE PAR TON ID CLIENT AZURE
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "https://lovethat44.github.io/mon-budget/",
        },
        cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    let accessToken = null;
    const ONEDRIVE_FILE_NAME = "ma_sauvegarde_budget.json";
    const ONEDRIVE_PATH = `https://graph.microsoft.com/v1.0/me/drive/root:/Apps/MonBudget/${ONEDRIVE_FILE_NAME}:/content`;

    // =========================================================================
    // 1. RÉFÉRENCES DOM
    // =========================================================================
    const form = document.getElementById('transaction-form');
    const dateInput = document.getElementById('transaction-date');
    const descInput = document.getElementById('transaction-description');
    const amountInput = document.getElementById('transaction-amount');
    const typeInput = document.getElementById('transaction-type');
    const categorySelect = document.getElementById('transaction-category');
    const subcategorySelect = document.getElementById('transaction-subcategory');
    const newCategoryInput = document.getElementById('new-category-name');
    const newSubcategoryInput = document.getElementById('new-subcategory-name');
    const isRecurringCheckbox = document.getElementById('is-recurring');
    const recurrenceTypeSelect = document.getElementById('recurrence-type');
    const list = document.getElementById('transaction-list');
    const balanceDisplay = document.getElementById('current-balance');
    const realBalanceDisplay = document.getElementById('real-balance');
    const initialBalanceDisplay = document.getElementById('initial-balance');
    const monthYearDisplay = document.getElementById('month-year-display');
    const prevMonthBtn = document.getElementById('prev-month-btn');
    const nextMonthBtn = document.getElementById('next-month-btn');
    const categoryManagementList = document.getElementById('category-management-list');
    const filterAllBtn = document.getElementById('filter-all');
    const filterPendingBtn = document.getElementById('filter-pending');
    const loginBtn = document.getElementById('login-btn');

    // =========================================================================
    // VARIABLES D'ÉTAT
    // =========================================================================
    let transactions = [];
    let categories = {
        'Salaire': ['N/A'],
        'Loyer': ['N/A'],
        'Courses': ['Alimentation', 'Ménage'],
        'Loisirs': ['Sorties', 'Sports'],
        'N/A': ['N/A']
    };
    let currentDate = new Date();
    let filter = 'all';
    let editingTransactionId = null;
    let fileHandle = null; 

    // =========================================================================
    // 2. FONCTIONS ONEDRIVE (API)
    // =========================================================================

    const updateLoginButtonUI = (isConnected) => {
        if (!loginBtn) return;
        if (isConnected) {
            loginBtn.innerHTML = '<i class="fas fa-cloud"></i> Synchronisé';
            loginBtn.style.backgroundColor = "#28a745";
        } else {
            loginBtn.innerHTML = '<i class="fab fa-microsoft"></i> Connexion OneDrive';
            loginBtn.style.backgroundColor = "#0078d4";
        }
    };

    const signIn = async () => {
        try {
            const loginRequest = { scopes: ["Files.ReadWrite", "User.Read"] };
            const loginResponse = await msalInstance.loginPopup(loginRequest);
            accessToken = loginResponse.accessToken;
            console.log("Connecté à Microsoft Graph.");
            updateLoginButtonUI(true);
            await downloadFromOneDrive(); // Tenter de charger les données après connexion
            return true;
        } catch (err) {
            console.error("Erreur de connexion :", err);
            return false;
        }
    };

    const uploadToOneDrive = async (data) => {
        if (!accessToken) return;
        try {
            await fetch(ONEDRIVE_PATH, {
                method: 'PUT',
                headers: { 
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify(data)
            });
            console.log("Cloud synchronisé !");
        } catch (err) { console.error("Erreur Upload OneDrive :", err); }
    };

    const downloadFromOneDrive = async () => {
        if (!accessToken) return;
        try {
            const response = await fetch(ONEDRIVE_PATH, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
            if (response.ok) {
                const cloudData = await response.json();
                if (cloudData.transactions && cloudData.categories) {
                    transactions = cloudData.transactions;
                    categories = cloudData.categories;
                    saveAndRender();
                    console.log("Données récupérées depuis OneDrive.");
                }
            }
        } catch (err) { console.log("Aucun fichier trouvé sur le Cloud, prêt pour premier usage."); }
    };

    // =========================================================================
    // 3. FONCTIONS UTILITAIRES & SAUVEGARDE
    // =========================================================================

    const saveState = async () => {
        const dataExport = { transactions, categories };
        
        // 1. Sauvegarde LocalStorage
        localStorage.setItem('transactions', JSON.stringify(transactions));
        localStorage.setItem('categories', JSON.stringify(categories));
        localStorage.setItem('currentDate', currentDate.toISOString());

        // 2. Sauvegarde Cloud OneDrive
        if (accessToken) {
            await uploadToOneDrive(dataExport);
        }

        // 3. Sauvegarde sur fichier local (ton ancien système)
        try {
            if (window.showSaveFilePicker && fileHandle) {
                const writable = await fileHandle.createWritable();
                await writable.write(JSON.stringify(dataExport, null, 2));
                await writable.close();
                console.log("Fichier local synchronisé.");
            }
        } catch (err) { console.error("Fichier local inaccessible :", err); }
    };

    const loadState = () => {
        const storedTransactions = localStorage.getItem('transactions');
        const storedCategories = localStorage.getItem('categories');
        const storedDate = localStorage.getItem('currentDate');

        if (storedTransactions) transactions = JSON.parse(storedTransactions);
        if (storedCategories) categories = JSON.parse(storedCategories);
        if (storedDate) {
            currentDate = new Date(storedDate);
        } else {
            currentDate.setDate(1);
            currentDate.setHours(0, 0, 0, 0);
        }
    };

    const saveAndRender = () => {
        saveState();
        updateMonthYearDisplay(); 
        updateBalances();
        renderList();
        renderCategoryManagement(); 
    };

    const formatDateForDisplay = (dateString) => {
        return new Date(dateString).toLocaleDateString('fr-FR', {
            year: 'numeric', month: '2-digit', day: '2-digit'
        });
    };

    const formatDateForInput = (dateString) => {
        const date = new Date(dateString);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    const setDefaultDate = () => {
        if (dateInput) dateInput.value = formatDateForInput(new Date().toISOString());
    };

    const updateMonthYearDisplay = () => {
        if (!monthYearDisplay) return;
        const monthYearString = currentDate.toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });
        monthYearDisplay.textContent = monthYearString.charAt(0).toUpperCase() + monthYearString.slice(1);
    };

    const changeMonth = (offset) => {
        currentDate.setMonth(currentDate.getMonth() + offset);
        currentDate.setDate(1);
        currentDate.setHours(0, 0, 0, 0);
        saveAndRender(); 
    };

    // =========================================================================
    // 4. GESTION DES CATÉGORIES
    // =========================================================================

    const renderCategories = (selectCat = null) => {
        if (!categorySelect) return;
        const sortedCategories = Object.keys(categories).sort((a, b) => a.localeCompare(b));
        categorySelect.innerHTML = sortedCategories.map(cat =>
            `<option value="${cat}" ${selectCat === cat ? 'selected' : ''}>${cat}</option>`
        ).join('');
        categorySelect.innerHTML += `<option value="create_new">-- Créer Nouvelle Catégorie --</option>`;
        if (selectCat === 'create_new') categorySelect.value = 'create_new';
        renderSubcategories(categorySelect.value);
    };
    
    const renderSubcategories = (cat, selectSub = null) => {
        if (!subcategorySelect) return;
        subcategorySelect.innerHTML = '';
        if (cat && categories[cat] && cat !== 'create_new') {
            const sortedSubcategories = categories[cat].sort((a, b) => a.localeCompare(b));
            subcategorySelect.innerHTML = sortedSubcategories.map(sub =>
                `<option value="${sub}" ${selectSub === sub ? 'selected' : ''}>${sub}</option>`
            ).join('');
            subcategorySelect.disabled = false;
        } else {
            subcategorySelect.innerHTML = `<option value="" disabled selected>-- Choisir Sous-catégorie --</option>`;
            subcategorySelect.disabled = true;
        }
        subcategorySelect.innerHTML += `<option value="create_new_sub">-- Créer Nouvelle Sous-catégorie --</option>`;
        if (selectSub === 'create_new_sub') subcategorySelect.value = 'create_new_sub';
    };

    const renderCategoryManagement = () => {
        if (!categoryManagementList) return;
        const totals = {};
        Object.keys(categories).forEach(cat => {
            totals[cat] = { total: 0, subcategories: {} };
            categories[cat].forEach(sub => totals[cat].subcategories[sub] = 0);
        });
        const currentMonthTransactions = transactions.filter(t => {
            const tDate = new Date(t.date);
            return tDate.getMonth() === currentDate.getMonth() && tDate.getFullYear() === currentDate.getFullYear();
        });
        currentMonthTransactions.forEach(t => {
            const cat = t.category;
            const sub = t.subcategory;
            if (totals[cat]) {
                const val = t.type === 'credit' ? t.amount : -t.amount;
                totals[cat].total += val;
                if (totals[cat].subcategories[sub] !== undefined) totals[cat].subcategories[sub] += val;
            }
        });
        const sortedCategories = Object.keys(categories).sort((a, b) => a.localeCompare(b));
        categoryManagementList.innerHTML = '';
        sortedCategories.forEach(catName => {
            const subcategories = categories[catName].sort((a, b) => a.localeCompare(b));
            const isProtected = catName === 'N/A';
            const catAmount = totals[catName].total;
            const catClass = catAmount >= 0 ? 'credit' : 'debit';
            const catAmountDisplay = Math.abs(catAmount) > 0 ? `${catAmount.toFixed(2)} €` : '0.00 €';
            const subListItems = subcategories.map(subName => {
                const isSubProtected = subName === 'N/A';
                const subAmount = totals[catName].subcategories[subName] || 0;
                const subClass = subAmount >= 0 ? 'credit' : 'debit';
                const subAmountDisplay = Math.abs(subAmount) > 0 ? `${subAmount.toFixed(2)} €` : '';
                const deleteBtnHtml = isSubProtected ? `<button class="btn-delete-sub btn-disabled" disabled><i class="fas fa-trash"></i></button>` : `<button class="btn-delete-sub" data-cat="${catName}" data-sub="${subName}"><i class="fas fa-trash"></i></button>`;
                const editBtnHtml = isSubProtected ? `<button class="btn-edit-sub btn-disabled" disabled><i class="fas fa-edit"></i></button>` : `<button class="btn-edit-sub" data-cat="${catName}" data-sub="${subName}"><i class="fas fa-edit"></i></button>`;
                return `<li class="subcategory-item" data-cat="${catName}" data-sub="${subName}"><div style="display:flex; align-items:center; gap:10px;"><span class="sub-name">${subName}</span><span class="${subClass}" style="font-size:0.9em; font-weight:bold;">${subAmountDisplay}</span></div><div class="sub-actions">${editBtnHtml}${deleteBtnHtml}</div></li>`;
            }).join('');
            const deleteCatBtnHtml = isProtected ? `<button class="btn-delete-cat btn-disabled" disabled><i class="fas fa-trash"></i></button>` : `<button class="btn-delete-cat" data-cat="${catName}"><i class="fas fa-trash"></i></button>`;
            const editCatBtnHtml = isProtected ? `<button class="btn-edit-cat btn-disabled" disabled><i class="fas fa-edit"></i></button>` : `<button class="btn-edit-cat" data-cat="${catName}"><i class="fas fa-edit"></i></button>`;
            const listItem = document.createElement('li');
            listItem.className = 'category-item';
            listItem.dataset.cat = catName;
            listItem.innerHTML = `<div class="category-header"><div style="display:flex; align-items:center; gap:10px;"><span class="cat-name" style="font-weight:bold;">${catName}</span><span class="${catClass}" style="font-weight:bold;">${catAmountDisplay}</span></div><div class="cat-actions">${editCatBtnHtml}${deleteCatBtnHtml}<i class="fas fa-chevron-down toggle-icon"></i></div></div><ul class="subcategory-list" style="display: none;">${subListItems}</ul>`;
            categoryManagementList.appendChild(listItem);
        });
    };

    const startEdit = (e, type, oldValue) => {
        const item = e.target.closest(type === 'cat' ? '.category-item' : '.subcategory-item');
        const nameSpan = item.querySelector(type === 'cat' ? '.cat-name' : '.sub-name');
        if (nameSpan.querySelector('input')) return;
        const input = document.createElement('input');
        input.type = 'text'; input.value = oldValue; input.className = 'edit-input';
        nameSpan.innerHTML = ''; nameSpan.appendChild(input); input.focus();
        const finishEdit = () => {
            const newValue = input.value.trim();
            if (newValue === '' || newValue === oldValue) { nameSpan.innerHTML = oldValue; return; }
            if (newValue === 'N/A') { alert(`Réservé.`); nameSpan.innerHTML = oldValue; return; }
            type === 'cat' ? saveCategoryEdit(oldValue, newValue, nameSpan) : saveSubcategoryEdit(item.dataset.cat, oldValue, newValue, nameSpan);
        };
        input.addEventListener('blur', finishEdit);
        input.addEventListener('keypress', (event) => { if (event.key === 'Enter') { event.preventDefault(); finishEdit(); } });
    };

    const saveCategoryEdit = (oldName, newName, nameSpan) => {
        if (categories[newName]) { alert(`Existe déjà.`); nameSpan.innerHTML = oldName; return; }
        categories[newName] = categories[oldName]; delete categories[oldName];
        transactions.forEach(t => { if (t.category === oldName) t.category = newName; });
        saveAndRender(); nameSpan.innerHTML = newName;
    };

    const saveSubcategoryEdit = (catName, oldName, newName, nameSpan) => {
        if (!categories[catName] || categories[catName].includes(newName)) { alert(`Existe déjà.`); nameSpan.innerHTML = oldName; return; }
        const index = categories[catName].indexOf(oldName);
        if (index > -1) categories[catName][index] = newName;
        transactions.forEach(t => { if (t.category === catName && t.subcategory === oldName) t.subcategory = newName; });
        saveAndRender(); nameSpan.innerHTML = newName;
    };

    const deleteCategory = (catName) => {
        if (catName === 'N/A') return;
        if (confirm(`Supprimer "${catName}" ?`)) {
            transactions.forEach(t => { if (t.category === catName) { t.category = 'N/A'; t.subcategory = 'N/A'; } });
            delete categories[catName]; saveAndRender();
        }
    };

    const deleteSubcategory = (catName, subName) => {
        if (subName === 'N/A') return;
        if (confirm(`Supprimer "${subName}" ?`)) {
            transactions.forEach(t => { if (t.category === catName && t.subcategory === subName) t.subcategory = 'N/A'; });
            categories[catName] = categories[catName].filter(sub => sub !== subName);
            if (categories[catName].length === 0) categories[catName].push('N/A');
            saveAndRender();
        }
    };

    // =========================================================================
    // 5. CALCULS ET SOLDES
    // =========================================================================

    const updateBalances = () => {
        if (!balanceDisplay || !initialBalanceDisplay || !realBalanceDisplay) return;
        const startOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
        const endOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0, 23, 59, 59);
        let previousBalance = 0;
        transactions.forEach(t => {
            const tDate = new Date(t.date);
            if (t.validated && tDate < startOfMonth) { t.type === 'credit' ? previousBalance += t.amount : previousBalance -= t.amount; }
        });
        const currentMonthTransactions = transactions.filter(t => {
            const tDate = new Date(t.date);
            return tDate >= startOfMonth && tDate <= endOfMonth;
        });
        let currentMonthTotalValidated = 0;
        let pendingTotal = 0;
        let realBalance = previousBalance;
        currentMonthTransactions.forEach(t => {
            if (t.validated) {
                t.type === 'credit' ? realBalance += t.amount : realBalance -= t.amount;
                t.type === 'credit' ? currentMonthTotalValidated += t.amount : currentMonthTotalValidated -= t.amount;
            } else {
                t.type === 'credit' ? pendingTotal += t.amount : pendingTotal -= t.amount;
            }
        });
        const finalBalance = realBalance + pendingTotal;
        realBalanceDisplay.innerHTML = `<span class="label">Solde RÉEL Actuel</span><span class="amount ${realBalance >= 0 ? 'credit' : 'debit'}" style="font-size: 1.5em;">${realBalance.toFixed(2)} €</span>`;
        initialBalanceDisplay.innerHTML = `<span class="label">Solde de Départ</span><span class="amount ${previousBalance >= 0 ? 'credit' : 'debit'}">${previousBalance.toFixed(2)} €</span>`;
        balanceDisplay.innerHTML = `<span class="label">Solde ESTIMÉ</span><span class="amount ${finalBalance >= 0 ? 'credit' : 'debit'}" style="font-size: 1.5em;">${finalBalance.toFixed(2)} €</span><div style="display:flex; justify-content: space-around; width: 100%; margin-top: 5px;"><span class="validated-total">Validé : <span class="${currentMonthTotalValidated >= 0 ? 'credit' : 'debit'}">${currentMonthTotalValidated.toFixed(2)} €</span></span><span class="pending-total">Attente : <span class="${pendingTotal >= 0 ? 'credit' : 'debit'}">${pendingTotal.toFixed(2)} €</span></span></div>`;
        const vSpan = document.getElementById('current-validated-total');
        const pSpan = document.getElementById('current-pending-total');
        if (vSpan) vSpan.textContent = `Validé : ${currentMonthTotalValidated.toFixed(2)} €`;
        if (pSpan) pSpan.textContent = `En Attente : ${pendingTotal.toFixed(2)} €`;
    };

    // =========================================================================
    // 6. FONCTIONS DE TRANSACTION
    // =========================================================================

    const getCategoryOptionsHTML = (selectedCategory, selectedSubcategory) => {
        const sortedCategories = Object.keys(categories).sort((a, b) => a.localeCompare(b));
        const categoriesOptions = sortedCategories.map(cat => `<option value="${cat}" ${selectedCategory === cat ? 'selected' : ''}>${cat}</option>`).join('');
        const currentSubcategories = categories[selectedCategory] || ['N/A'];
        const sortedSubcategories = currentSubcategories.sort((a, b) => a.localeCompare(b));
        const subcategoriesOptions = sortedSubcategories.map(sub => `<option value="${sub}" ${selectedSubcategory === sub ? 'selected' : ''}>${sub}</option>`).join('');
        return { categoriesOptions, subcategoriesOptions };
    };

    const renderList = () => {
        if (!list) return;
        list.innerHTML = '';
        const firstDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
        const lastDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0, 23, 59, 59);
        let filteredTransactions = transactions.filter(t => {
            const tDate = new Date(t.date);
            const isCurrentMonth = tDate >= firstDayOfMonth && tDate <= lastDayOfMonth;
            return filter === 'pending' ? (isCurrentMonth && !t.validated) : isCurrentMonth;
        });
        filteredTransactions.sort((a, b) => new Date(a.date) - new Date(b.date));
        filteredTransactions.forEach(t => {
            const dateString = formatDateForDisplay(t.date);
            const amountText = (t.type === 'debit' ? '-' : '+') + parseFloat(t.amount).toFixed(2) + ' €';
            const isEditing = editingTransactionId === t.id;
            let recurrenceText = t.isRecurring ? `<span class="t-recurrence"> (${t.recurrenceType === 'monthly' ? 'Mensuel' : 'Annuel'})</span>` : '';
            const li = document.createElement('li');
            li.className = `transaction-item ${t.type} ${t.validated ? 'validated' : 'pending'}`;
            li.dataset.id = t.id;
            if (isEditing) {
                const { categoriesOptions, subcategoriesOptions } = getCategoryOptionsHTML(t.category, t.subcategory);
                li.classList.add('editing');
                li.innerHTML = `<div class="t-edit-form"><input type="date" value="${formatDateForInput(t.date)}" data-field="date"><input type="text" value="${t.description}" data-field="description"><select data-field="category">${categoriesOptions}</select><select data-field="subcategory">${subcategoriesOptions}</select><input type="number" value="${parseFloat(t.amount).toFixed(2)}" data-field="amount" step="0.01"><select data-field="type"><option value="debit" ${t.type === 'debit' ? 'selected' : ''}>Dépense</option><option value="credit" ${t.type === 'credit' ? 'selected' : ''}>Revenu</option></select><div class="t-actions"><button class="btn-save-edit"><i class="fas fa-save"></i></button><button class="btn-cancel-edit"><i class="fas fa-times"></i></button></div></div>`;
            } else {
                li.innerHTML = `<div class="t-info"><span class="t-desc">${t.description}</span><span class="t-date">${dateString} - ${t.category} (${t.subcategory})${recurrenceText}</span></div><div class="t-actions"><span class="t-amount ${t.type}">${amountText}</span><button class="btn-duplicate"><i class="fas fa-copy"></i></button><button class="btn-edit"><i class="fas fa-edit"></i></button><button class="btn-validate"><i class="far ${t.validated ? 'fa-check-square' : 'fa-square'}"></i></button><button class="btn-delete"><i class="fas fa-trash"></i></button></div>`;
            }
            list.appendChild(li);
        });
    };

    const generateRecurringInstances = (baseTransaction) => {
        if (!baseTransaction.isRecurring || !baseTransaction.recurrenceType) return [];
        const baseDate = new Date(baseTransaction.date);
        let numInstances = baseTransaction.recurrenceType === 'monthly' ? 12 : 5;
        let interval = baseTransaction.recurrenceType === 'monthly' ? 1 : 12;
        const recurringInstances = [];
        const baseDay = baseDate.getDate();
        for (let i = 1; i <= numInstances; i++) {
            let nextDate = new Date(baseDate.getFullYear(), baseDate.getMonth() + i * interval, baseDay);
            if (nextDate.getDate() !== baseDay && interval === 1) nextDate = new Date(baseDate.getFullYear(), baseDate.getMonth() + i * interval, 0);
            if (nextDate.getTime() > baseDate.getTime()) {
                recurringInstances.push({ ...baseTransaction, id: Date.now().toString() + '-' + i + '-' + Math.random().toString(36).substring(2, 9), date: nextDate.toISOString(), validated: false, baseId: baseTransaction.baseId });
            }
        }
        return recurringInstances;
    };

    const startTransactionEdit = (id) => { editingTransactionId = id; renderList(); };
    const cancelTransactionEdit = () => { editingTransactionId = null; renderList(); };

    const saveTransactionEdit = (id, li) => {
        const index = transactions.findIndex(t => t.id === id);
        if (index === -1) return;
        const original = transactions[index];
        const dateVal = li.querySelector('[data-field="date"]').value;
        const newDate = new Date(dateVal);
        const amount = parseFloat(li.querySelector('[data-field="amount"]').value);
        const changes = { date: newDate.toISOString(), description: li.querySelector('[data-field="description"]').value, category: li.querySelector('[data-field="category"]').value, subcategory: li.querySelector('[data-field="subcategory"]').value, amount, type: li.querySelector('[data-field="type"]').value };
        if (!dateVal || changes.description === '' || isNaN(amount) || amount <= 0) return alert('Champs invalides.');

        if (original.isRecurring && original.baseId) {
            if (confirm(`Modifier toute la série ?`)) {
                const baseId = original.baseId;
                const startFrom = new Date(original.date);
                transactions.forEach((t, i) => {
                    if (t.baseId === baseId && new Date(t.date) >= startFrom) {
                        let d = t.date;
                        if (new Date(original.date).getTime() !== newDate.getTime()) {
                           let instanceDate = new Date(t.date);
                           let nd = new Date(instanceDate.getFullYear(), instanceDate.getMonth(), newDate.getDate());
                           d = nd.toISOString();
                        }
                        transactions[i] = { ...t, ...changes, date: d };
                    }
                });
            } else { transactions[index] = { ...original, ...changes }; }
        } else { transactions[index] = { ...original, ...changes }; }
        editingTransactionId = null; saveAndRender();
    };

    const duplicateTransaction = (id) => {
        const original = transactions.find(t => t.id === id);
        if (!original) return;
        const newT = { ...original, id: Date.now().toString() + '-' + Math.random().toString(36).substring(2, 9), baseId: null, validated: false, isRecurring: false, date: new Date().toISOString() };
        transactions.push(newT); saveState(); updateBalances(); startTransactionEdit(newT.id);
    };

    // =========================================================================
    // 7. ÉVÉNEMENTS
    // =========================================================================

    if (loginBtn) {
        loginBtn.addEventListener('click', async () => {
            await signIn();
        });
    }

    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const dateToKeep = dateInput.value;
            let cat = categorySelect.value;
            let subcat = subcategorySelect.value;
            const isRec = isRecurringCheckbox?.checked;
            const recType = isRec ? recurrenceTypeSelect.value : null;

            if (cat === 'create_new') {
                const val = newCategoryInput.value.trim();
                if (val && val !== 'N/A') { categories[val] = ['N/A']; cat = val; } else return alert('Invalide');
            }
            if (subcat === 'create_new_sub') {
                const val = newSubcategoryInput.value.trim();
                if (val && val !== 'N/A') { 
                    if (categories[cat][0] === 'N/A') categories[cat] = [];
                    categories[cat].push(val); subcat = val; 
                } else return alert('Invalide');
            }

            const amount = parseFloat(amountInput.value);
            if (!cat || !subcat || descInput.value.trim() === '' || isNaN(amount) || amount <= 0 || !dateInput.value) return alert('Champs manquants.');

            const newT = { id: Date.now().toString() + '-' + Math.random().toString(36).substring(2, 9), description: descInput.value, amount, type: typeInput.value, category: cat, subcategory: subcat, date: new Date(dateInput.value).toISOString(), validated: false, isRecurring: isRec, recurrenceType: recType, baseId: null };
            newT.baseId = newT.id;
            transactions.push(newT);
            if (isRec) transactions.push(...generateRecurringInstances(newT));
            
            saveAndRender();
            form.reset();
            dateInput.value = dateToKeep;
            renderCategories(cat);
        });
    }

    if (list) {
        list.addEventListener('click', (e) => {
            const li = e.target.closest('.transaction-item');
            if (!li) return;
            const id = li.dataset.id;
            const idx = transactions.findIndex(t => t.id === id);
            if (idx === -1) return;

            if (e.target.closest('.btn-validate')) { transactions[idx].validated = !transactions[idx].validated; saveAndRender(); }
            else if (e.target.closest('.btn-delete')) {
                const t = transactions[idx];
                if (t.isRecurring && t.baseId) {
                    if (confirm("Supprimer la série future ?")) {
                        transactions = transactions.filter(tr => !(tr.baseId === t.baseId && new Date(tr.date) >= new Date(t.date)));
                    } else { transactions.splice(idx, 1); }
                } else { if (confirm("Supprimer ?")) transactions.splice(idx, 1); }
                saveAndRender();
            }
            else if (e.target.closest('.btn-duplicate')) duplicateTransaction(id);
            else if (e.target.closest('.btn-edit')) startTransactionEdit(id);
            else if (e.target.closest('.btn-cancel-edit')) cancelTransactionEdit();
            else if (e.target.closest('.btn-save-edit')) saveTransactionEdit(id, li);
        });

        list.addEventListener('change', (e) => {
            if (e.target.dataset.field === 'category') {
                const cat = e.target.value;
                const subSelect = e.target.closest('.t-edit-form').querySelector('[data-field="subcategory"]');
                const subs = categories[cat] || ['N/A'];
                subSelect.innerHTML = subs.map(s => `<option value="${s}">${s}</option>`).join('');
            }
        });
    }

    if (categoryManagementList) {
        categoryManagementList.addEventListener('click', (e) => {
            const target = e.target.closest('button');
            const header = e.target.closest('.category-header');
            const item = e.target.closest('.category-item');
            if (header && item && !target) {
                const subList = item.querySelector('.subcategory-list');
                const isHidden = subList.style.display === 'none';
                subList.style.display = isHidden ? 'block' : 'none';
                return;
            }
            if (!target || target.disabled) return;
            if (target.classList.contains('btn-delete-cat')) deleteCategory(target.dataset.cat);
            else if (target.classList.contains('btn-delete-sub')) deleteSubcategory(target.dataset.cat, target.dataset.sub);
            else if (target.classList.contains('btn-edit-cat')) startEdit(e, 'cat', target.dataset.cat);
            else if (target.classList.contains('btn-edit-sub')) startEdit(e, 'sub', target.dataset.sub);
        });
    }

    if (filterAllBtn) filterAllBtn.addEventListener('click', () => { filter = 'all'; filterAllBtn.classList.add('active'); filterPendingBtn.classList.remove('active'); renderList(); });
    if (filterPendingBtn) filterPendingBtn.addEventListener('click', () => { filter = 'pending'; filterPendingBtn.classList.add('active'); filterAllBtn.classList.remove('active'); renderList(); });
    if (prevMonthBtn) prevMonthBtn.addEventListener('click', () => changeMonth(-1));
    if (nextMonthBtn) nextMonthBtn.addEventListener('click', () => changeMonth(1));

    if (isRecurringCheckbox && recurrenceTypeSelect) {
        isRecurringCheckbox.addEventListener('change', () => {
            recurrenceTypeSelect.disabled = !isRecurringCheckbox.checked;
            recurrenceTypeSelect.value = isRecurringCheckbox.checked ? 'monthly' : '';
        });
    }

    if (categorySelect) {
        categorySelect.addEventListener('change', () => {
            if (categorySelect.value === 'create_new') {
                newCategoryInput.style.display = 'block'; newCategoryInput.focus();
                renderSubcategories(null);
            } else {
                newCategoryInput.style.display = 'none';
                renderSubcategories(categorySelect.value);
            }
        });
    }

    if (newCategoryInput) {
        newCategoryInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const val = newCategoryInput.value.trim();
                if (val && val !== 'N/A' && !categories[val]) {
                    categories[val] = ['N/A']; renderCategories(val);
                    newCategoryInput.style.display = 'none'; newCategoryInput.value = '';
                }
            }
        });
    }

    if (subcategorySelect) {
        subcategorySelect.addEventListener('change', () => {
            newSubcategoryInput.style.display = subcategorySelect.value === 'create_new_sub' ? 'block' : 'none';
            if(newSubcategoryInput.style.display === 'block') newSubcategoryInput.focus();
        });
    }

    if (newSubcategoryInput) {
        newSubcategoryInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const val = newSubcategoryInput.value.trim();
                const cat = categorySelect.value;
                if (val && val !== 'N/A' && !categories[cat].includes(val)) {
                    if (categories[cat][0] === 'N/A') categories[cat] = [];
                    categories[cat].push(val); renderSubcategories(cat, val);
                    newSubcategoryInput.style.display = 'none'; newSubcategoryInput.value = '';
                }
            }
        });
    }

    // =========================================================================
    // 8. EXPORT / IMPORT JSON
    // =========================================================================
    const exportBtn = document.getElementById('export-data-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', () => {
            const dataExport = { transactions, categories };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(dataExport));
            const downloadAnchorNode = document.createElement('a');
            downloadAnchorNode.setAttribute("href", dataStr);
            downloadAnchorNode.setAttribute("download", "ma_sauvegarde_budget.json");
            document.body.appendChild(downloadAnchorNode);
            downloadAnchorNode.click();
            downloadAnchorNode.remove();
        });
    }

    const importInput = document.getElementById('import-data-input');
    const importBtn = document.getElementById('import-data-btn');
    if (importBtn && importInput) {
        importBtn.addEventListener('click', () => importInput.click());
        importInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const importedData = JSON.parse(event.target.result);
                    if (importedData.transactions && importedData.categories) {
                        transactions = importedData.transactions;
                        categories = importedData.categories;
                        saveAndRender();
                        alert("Importation réussie !");
                    }
                } catch (err) { alert("Erreur lors de la lecture du fichier."); }
            };
            reader.readAsText(file);
        });
    }

    const clearDataBtn = document.getElementById('clear-data-btn');
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', () => {
            if (confirm("Voulez-vous vraiment TOUT supprimer ? Cette action est irréversible (sauf si vous avez un export JSON).")) {
                transactions = [];
                categories = { 'Salaire': ['N/A'], 'Loyer': ['N/A'], 'Courses': ['Alimentation'], 'Loisirs': ['Sorties'], 'N/A': ['N/A'] };
                saveAndRender();
            }
        });
    }

    // =========================================================================
    // 9. INITIALISATION
    // =========================================================================
    
    // Tentative de reconnexion silencieuse MSAL
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const response = await msalInstance.acquireTokenSilent({
                scopes: ["Files.ReadWrite", "User.Read"],
                account: accounts[0]
            });
            accessToken = response.accessToken;
            updateLoginButtonUI(true);
            await downloadFromOneDrive();
        } catch (e) { console.log("Session expirée, reconnexion manuelle nécessaire."); }
    }

    loadState();
    renderCategories();
    setDefaultDate();
    saveAndRender();

    if (recurrenceTypeSelect) recurrenceTypeSelect.disabled = true;
    if (newCategoryInput) newCategoryInput.style.display = 'none';
    if (newSubcategoryInput) newSubcategoryInput.style.display = 'none';
    if (subcategorySelect) subcategorySelect.disabled = true;

});