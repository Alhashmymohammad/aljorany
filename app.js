// Aljorany Pro - Universal Excel Reader (Fixed)
// يدعم جميع صيغ Excel مع إظهار كامل السطر

class AljoranyPro {
    constructor() {
        this.filesData = new Map();
        this.currentResults = [];
        this.isProcessing = false;
        this.currentTheme = 'dark';
        this.searchFilter = 'all';
        this.showFullRow = true;
        
        this.supportedExtensions = [
            'xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xlt',
            'xlam', 'xla', 'xml', 'csv', 'txt', 'prn', 'dif', 
            'slk', 'dbf', 'ods', 'fods', 'uos', 'html', 'htm', 'numbers'
        ];
        
        this.initElements();
        this.initEventListeners();
        this.initTheme();
    }
    
    initElements() {
        this.uploadZone = document.getElementById('uploadZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileList = document.getElementById('fileList');
        this.progressContainer = document.getElementById('progressContainer');
        this.progressFill = document.getElementById('progressFill');
        this.progressPercent = document.getElementById('progressPercent');
        this.searchContainer = document.getElementById('searchContainer');
        this.searchInput = document.getElementById('searchInput');
        this.searchBtn = document.getElementById('searchBtn');
        this.filterChips = document.querySelectorAll('.filter-chip');
        this.resultsContainer = document.getElementById('resultsContainer');
        this.resultsList = document.getElementById('resultsList');
        this.resultsCount = document.getElementById('resultsCount');
        this.copyAllBtn = document.getElementById('copyAllBtn');
        this.exportBtn = document.getElementById('exportBtn');
        this.clearResultsBtn = document.getElementById('clearResultsBtn');
        this.statsBar = document.getElementById('statsBar');
        this.statFiles = document.getElementById('statFiles');
        this.statRows = document.getElementById('statRows');
        this.statCols = document.getElementById('statCols');
        this.statResults = document.getElementById('statResults');
        this.emptyState = document.getElementById('emptyState');
        this.themeToggle = document.getElementById('themeToggle');
        this.toastContainer = document.getElementById('toastContainer');
    }
    
    initEventListeners() {
        this.uploadZone.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFiles(e.target.files));
        
        this.uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadZone.classList.add('dragover');
        });
        
        this.uploadZone.addEventListener('dragleave', () => {
            this.uploadZone.classList.remove('dragover');
        });
        
        this.uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadZone.classList.remove('dragover');
            this.handleFiles(e.dataTransfer.files);
        });
        
        this.searchBtn.addEventListener('click', () => this.performSearch());
        this.searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.performSearch();
        });
        
        this.filterChips.forEach(chip => {
            chip.addEventListener('click', () => {
                this.filterChips.forEach(c => c.classList.remove('active'));
                chip.classList.add('active');
                this.searchFilter = chip.dataset.filter;
            });
        });
        
        this.copyAllBtn.addEventListener('click', () => this.copyAllResults());
        this.exportBtn.addEventListener('click', () => this.exportResults());
        this.clearResultsBtn.addEventListener('click', () => this.clearResults());
        this.themeToggle.addEventListener('click', () => this.toggleTheme());
        
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey || e.metaKey) {
                if (e.key === 'f') {
                    e.preventDefault();
                    this.searchInput.focus();
                } else if (e.key === 'o') {
                    e.preventDefault();
                    this.fileInput.click();
                }
            }
        });
    }
    
    initTheme() {
        const savedTheme = localStorage.getItem('aljorany-theme') || 'dark';
        this.setTheme(savedTheme);
    }
    
    setTheme(theme) {
        this.currentTheme = theme;
        document.documentElement.setAttribute('data-theme', theme);
        this.themeToggle.textContent = theme === 'dark' ? '☀️' : '🌙';
        localStorage.setItem('aljorany-theme', theme);
    }
    
    toggleTheme() {
        const newTheme = this.currentTheme === 'dark' ? 'light' : 'dark';
        this.setTheme(newTheme);
    }
    
    isValidExtension(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        return this.supportedExtensions.includes(ext);
    }
    
    getFileExtension(filename) {
        return filename.split('.').pop().toLowerCase();
    }
    
    formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }
    
    async handleFiles(files) {
        if (this.isProcessing) return;
        
        const validFiles = Array.from(files).filter(file => {
            const isValid = this.isValidExtension(file.name);
            if (!isValid) {
                this.showToast('⚠️ ' + file.name + ': الصيغة غير مدعومة', 'warning');
            }
            return isValid;
        });
        
        if (validFiles.length === 0) {
            this.showToast('❌ لا توجد ملفات Excel صالحة', 'error');
            return;
        }
        
        this.isProcessing = true;
        this.progressContainer.classList.add('show');
        
        let processedCount = 0;
        let totalRows = 0;
        let totalCols = 0;
        
        for (let i = 0; i < validFiles.length; i++) {
            const file = validFiles[i];
            
            try {
                const progress = ((i + 0.5) / validFiles.length) * 100;
                this.updateProgress(progress);
                
                const data = await this.readFile(file);
                
                this.updateProgress(((i + 1) / validFiles.length) * 100);
                
                if (data && data.length > 0) {
                    const columns = Object.keys(data[0]);
                    
                    this.filesData.set(file.name, {
                        data: data,
                        rowCount: data.length,
                        colCount: columns.length,
                        columns: columns,
                        size: file.size,
                        lastModified: file.lastModified,
                        extension: this.getFileExtension(file.name)
                    });
                    
                    totalRows += data.length;
                    totalCols = Math.max(totalCols, columns.length);
                    processedCount++;
                    
                    this.addFileToList(file, data.length);
                } else {
                    this.showToast('⚠️ ' + file.name + ': الملف فارغ', 'warning');
                }
            } catch (error) {
                console.error('Error:', error);
                this.showToast('❌ خطأ في ' + file.name + ': ' + error.message, 'error');
            }
        }
        
        this.updateProgress(100);
        
        setTimeout(() => {
            this.progressContainer.classList.remove('show');
            this.updateProgress(0);
        }, 500);
        
        this.isProcessing = false;
        
        if (processedCount > 0) {
            this.showToast('✅ تم استيراد ' + processedCount + ' ملف', 'success');
            this.statsBar.style.display = 'flex';
            this.searchContainer.classList.add('show');
            this.emptyState.style.display = 'none';
            this.updateStats();
            this.searchInput.focus();
        }
    }
    
    readFile(file) {
        return new Promise((resolve, reject) => {
            const ext = this.getFileExtension(file.name);
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    let workbook;
                    let data = e.target.result;
                    
                    if (ext === 'csv' || ext === 'txt' || ext === 'prn') {
                        workbook = XLSX.read(data, { 
                            type: 'string',
                            raw: true,
                            cellDates: true
                        });
                    } else if (ext === 'html' || ext === 'htm') {
                        workbook = XLSX.read(data, { 
                            type: 'string',
                            cellDates: true
                        });
                    } else if (ext === 'dbf') {
                        const dbfData = new Uint8Array(data);
                        workbook = XLSX.read(dbfData, { 
                            type: 'array',
                            cellDates: true
                        });
                    } else {
                        const binaryData = new Uint8Array(data);
                        workbook = XLSX.read(binaryData, { 
                            type: 'array',
                            cellFormula: false,
                            cellHTML: false,
                            cellText: true,
                            cellDates: true,
                            raw: true
                        });
                    }
                    
                    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                        reject(new Error('الملف لا يحتوي على بيانات'));
                        return;
                    }
                    
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                        defval: '',
                        blankrows: false,
                        raw: false,
                        dateNF: 'yyyy-mm-dd'
                    });
                    
                    const cleanedData = this.cleanData(jsonData);
                    
                    resolve(cleanedData);
                    
                } catch (error) {
                    reject(new Error('فشل قراءة الملف: ' + error.message));
                }
            };
            
            reader.onerror = () => reject(new Error('فشل قراءة الملف'));
            
            if (['csv', 'txt', 'prn', 'html', 'htm'].includes(ext)) {
                reader.readAsText(file);
            } else {
                reader.readAsArrayBuffer(file);
            }
        });
    }
    
    cleanData(jsonData) {
        return jsonData
            .filter(row => {
                const values = Object.values(row);
                return values.some(val => {
                    if (val === null || val === undefined) return false;
                    const str = String(val).trim();
                    return str !== '' && str !== 'null' && str !== 'undefined';
                });
            })
            .map(row => {
                const cleanRow = {};
                Object.entries(row).forEach(([key, value]) => {
                    if (value !== null && value !== undefined) {
                        let cleanValue = String(value).trim();
                        cleanValue = cleanValue.replace(/\s+/g, ' ');
                        if (cleanValue !== '' && cleanValue !== 'null' && cleanValue !== 'undefined') {
                            cleanRow[key] = cleanValue;
                        }
                    }
                });
                return cleanRow;
            })
            .filter(row => Object.keys(row).length > 0);
    }
    
    addFileToList(file, rowCount) {
        const fileId = 'file-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        const ext = this.getFileExtension(file.name).toUpperCase();
        
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.id = fileId;
        fileItem.innerHTML = `
            <div class="file-icon">${this.getFileIcon(ext)}</div>
            <div class="file-info">
                <div class="file-name" title="${file.name}">${file.name}</div>
                <div class="file-meta">${this.formatFileSize(file.size)} • ${rowCount.toLocaleString()} صف</div>
            </div>
            <button class="file-remove" title="حذف الملف">✕</button>
        `;
        
        fileItem.querySelector('.file-remove').addEventListener('click', () => {
            this.removeFile(fileId, file.name);
        });
        
        this.fileList.appendChild(fileItem);
        this.fileList.classList.add('show');
    }
    
    getFileIcon(ext) {
        const icons = {
            'XLSX': '📗', 'XLS': '📗', 'XLSM': '📗', 'XLSB': '📗',
            'XLTX': '📘', 'XLTM': '📘', 'XLT': '📘',
            'XLAM': '📙', 'XLA': '📙',
            'CSV': '📄', 'TXT': '📄', 'PRN': '📄',
            'ODS': '📕', 'FODS': '📕', 'UOS': '📕',
            'HTML': '🌐', 'HTM': '🌐',
            'DBF': '🗄️', 'DIF': '📊', 'SLK': '📊',
            'XML': '📋', 'NUMBERS': '🍎'
        };
        return icons[ext] || '📄';
    }
    
    removeFile(fileId, fileName) {
        const element = document.getElementById(fileId);
        if (element) element.remove();
        
        this.filesData.delete(fileName);
        
        if (this.filesData.size === 0) {
            this.fileList.classList.remove('show');
            this.searchContainer.classList.remove('show');
            this.resultsContainer.classList.remove('show');
            this.emptyState.style.display = 'block';
            this.statsBar.style.display = 'none';
            this.currentResults = [];
        }
        
        this.updateStats();
        this.showToast('🗑️ تم حذف الملف', 'success');
    }
    
    updateProgress(percent) {
        this.progressFill.style.width = percent + '%';
        this.progressPercent.textContent = Math.round(percent) + '%';
    }
    
    updateStats() {
        let totalRows = 0;
        let maxCols = 0;
        
        this.filesData.forEach(file => {
            totalRows += file.rowCount;
            maxCols = Math.max(maxCols, file.colCount);
        });
        
        this.statFiles.textContent = this.filesData.size;
        this.statRows.textContent = totalRows.toLocaleString();
        this.statCols.textContent = maxCols;
        this.statResults.textContent = this.currentResults.length;
    }
    
    performSearch() {
        const searchTerm = this.searchInput.value.trim();
        
        if (!searchTerm) {
            this.showToast('⚠️ أدخل كلمة للبحث', 'error');
            return;
        }
        
        if (this.filesData.size === 0) {
            this.showToast('⚠️ لا توجد ملفات مستوردة', 'error');
            return;
        }
        
        const searchTerms = searchTerm.toLowerCase()
            .split(/\s+/)
            .filter(t => t.length > 0);
        
        this.currentResults = [];
        let resultId = 0;
        
        this.filesData.forEach((fileData, fileName) => {
            fileData.data.forEach((row, rowIndex) => {
                const rowValues = Object.values(row);
                const rowText = rowValues.join(' ').toLowerCase();
                
                let matches = false;
                let matchType = '';
                
                if (this.searchFilter === 'exact') {
                    matches = rowText.includes(searchTerm.toLowerCase());
                    matchType = 'تطابق تام';
                } else if (this.searchFilter === 'partial') {
                    matches = searchTerms.some(term => rowText.includes(term));
                    matchType = 'تطابق جزئي';
                } else {
                    matches = searchTerms.every(term => rowText.includes(term));
                    matchType = searchTerms.length > 1 ? 'تطابق كلي' : 'تطابق';
                }
                
                if (matches) {
                    const matchedFields = [];
                    const otherFields = [];
                    
                    Object.entries(row).forEach(([key, value]) => {
                        const valueLower = value.toLowerCase();
                        const isMatch = searchTerms.some(term => 
                            valueLower.includes(term)
                        );
                        
                        if (isMatch) {
                            matchedFields.push({ key, value, highlight: true });
                        } else {
                            otherFields.push({ key, value, highlight: false });
                        }
                    });
                    
                    this.currentResults.push({
                        id: resultId++,
                        fileName,
                        rowNumber: rowIndex + 2,
                        matchedFields,
                        otherFields,
                        allFields: row,
                        matchType,
                        fullRowText: rowValues.join(' | ')
                    });
                }
            });
        });
        
        this.displayResults();
        this.updateStats();
        
        if (this.currentResults.length === 0) {
            this.showToast('🔍 لم يتم العثور على نتائج', 'error');
        } else {
            this.showToast('✅ تم العثور على ' + this.currentResults.length + ' نتيجة', 'success');
        }
    }
    
    displayResults() {
        this.resultsCount.textContent = '(' + this.currentResults.length + ')';
        this.resultsList.innerHTML = '';
        
        if (this.currentResults.length === 0) {
            this.resultsContainer.classList.remove('show');
            return;
        }
        
        this.resultsContainer.classList.add('show');
        
        this.currentResults.forEach(result => {
            const card = this.createResultCard(result);
            this.resultsList.appendChild(card);
        });
        
        setTimeout(() => {
            this.resultsContainer.scrollIntoView({ 
                behavior: 'smooth', 
                block: 'start' 
            });
        }, 100);
    }
    
    createResultCard(result) {
        const card = document.createElement('div');
        card.className = 'result-card';
        
        const titleFields = result.matchedFields.slice(0, 3);
        const titleText = titleFields.map(f => f.value).join(' - ') || 'نتيجة بحث';
        
        // عرض كامل السطر
        const allFields = [...result.matchedFields, ...result.otherFields]
            .filter(f => f.value && f.value.trim() !== '');
        
        const fieldsHtml = allFields.map(field => `
            <div class="field-item ${field.highlight ? 'highlighted' : ''}">
                <div class="field-label">${field.key}</div>
                <div class="field-value">${this.escapeHtml(field.value)}</div>
            </div>
        `).join('');
        
        // preview للسطر الكامل
        const fullRowPreview = result.fullRowText.length > 150 
            ? result.fullRowText.substring(0, 150) + '...' 
            : result.fullRowText;
        
        card.innerHTML = `
            <div class="result-header">
                <div class="result-title">
                    ${this.escapeHtml(titleText)}
                    <span class="result-match">${result.matchType}</span>
                </div>
            </div>
            <div class="result-meta">
                <span>📄 ${result.fileName}</span>
                <span>📊 صف ${result.rowNumber}</span>
                <span>✓ ${result.matchedFields.length} تطابق</span>
            </div>
            <div class="full-row-preview" title="السطر الكامل - انقر للنسخ">
                <strong>السطر الكامل:</strong><br>
                ${this.escapeHtml(fullRowPreview)}
            </div>
            <div class="result-fields">
                ${fieldsHtml}
            </div>
            <div class="copy-indicator">انقر للنسخ 📋</div>
        `;
        
        card.addEventListener('click', () => this.copyResult(result, card));
        
        let pressTimer;
        card.addEventListener('touchstart', (e) => {
            pressTimer = setTimeout(() => {
                e.preventDefault();
                this.copyResult(result, card);
            }, 500);
        });
        card.addEventListener('touchend', () => clearTimeout(pressTimer));
        
        return card;
    }
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    copyResult(result, cardElement) {
        const textToCopy = result.fullRowText;
        
        navigator.clipboard.writeText(textToCopy).then(() => {
            cardElement.classList.add('copied');
            setTimeout(() => cardElement.classList.remove('copied'), 1000);
            this.showToast('✅ تم نسخ السطر الكامل', 'success');
        }).catch(() => {
            const textArea = document.createElement('textarea');
            textArea.value = textToCopy;
            textArea.style.position = 'fixed';
            textArea.style.left = '-9999px';
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            
            cardElement.classList.add('copied');
            setTimeout(() => cardElement.classList.remove('copied'), 1000);
            this.showToast('✅ تم نسخ السطر الكامل', 'success');
        });
    }
    
    copyAllResults() {
        if (this.currentResults.length === 0) return;
        
        const allText = this.currentResults.map(r => r.fullRowText).join('\n');
        
        navigator.clipboard.writeText(allText).then(() => {
            this.showToast('✅ تم نسخ ' + this.currentResults.length + ' سطر', 'success');
        }).catch(() => {
            this.showToast('❌ فشل النسخ', 'error');
        });
    }
    
    exportResults() {
        if (this.currentResults.length === 0) {
            this.showToast('⚠️ لا توجد نتائج للتصدير', 'error');
            return;
        }
        
        const exportData = this.currentResults.map(r => ({
            'اسم الملف': r.fileName,
            'رقم الصف': r.rowNumber,
            'نوع التطابق': r.matchType,
            'السطر الكامل': r.fullRowText,
            ...r.allFields
        }));
        
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'نتائج البحث');
        
        const timestamp = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, 'aljorany-results-' + timestamp + '.xlsx');
        
        this.showToast('📥 تم تصدير النتائج', 'success');
    }
    
    clearResults() {
        this.currentResults = [];
        this.resultsList.innerHTML = '';
        this.resultsContainer.classList.remove('show');
        this.searchInput.value = '';
        this.updateStats();
        this.showToast('🗑️ تم مسح النتائج', 'success');
    }
    
    showToast(message, type) {
        const toast = document.createElement('div');
        toast.className = 'toast ' + type;
        toast.textContent = message;
        
        this.toastContainer.appendChild(toast);
        
        requestAnimationFrame(() => {
            toast.classList.add('show');
        });
        
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 400);
        }, 3000);
    }
}

const app = new AljoranyPro();

if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js')
        .then(reg => console.log('Service Worker registered'))
        .catch(err => console.log('Service Worker error:', err));
}
