<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Advanced Invoice Reconciliation System</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.5/babel.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;500;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #6366f1;
            --primary-dark: #4f46e5;
            --success: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --gray-50: #f9fafb;
            --gray-100: #f3f4f6;
            --gray-200: #e5e7eb;
            --gray-600: #4b5563;
            --gray-800: #1f2937;
            --red-50: #fef2f2;
            --red-100: #fee2e2;
            --red-600: #dc2626;
            --green-50: #f0fdf4;
            --green-600: #16a34a;
            --yellow-50: #fefce8;
            --yellow-600: #ca8a04;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 24px;
        }

        .app {
            max-width: 1920px;
            margin: 0 auto;
        }

        .header {
            background: white;
            border-radius: 20px;
            padding: 40px;
            margin-bottom: 24px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }

        .title {
            font-size: 36px;
            font-weight: 800;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 8px;
        }

        .subtitle {
            color: var(--gray-600);
            font-size: 15px;
        }

        .upload-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }

        .upload-card {
            background: white;
            border-radius: 16px;
            padding: 28px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        .upload-header {
            font-size: 18px;
            font-weight: 700;
            color: var(--gray-800);
            margin-bottom: 6px;
        }

        .upload-desc {
            font-size: 13px;
            color: var(--gray-600);
            margin-bottom: 18px;
        }

        .upload-btn {
            width: 100%;
            padding: 14px;
            background: linear-gradient(135deg, var(--primary), var(--primary-dark));
            color: white;
            border: none;
            border-radius: 10px;
            font-weight: 600;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.3s;
        }

        .upload-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 16px rgba(99, 102, 241, 0.4);
        }

        .file-name {
            margin-top: 12px;
            padding: 10px 14px;
            background: var(--green-50);
            color: var(--green-600);
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
        }

        .file-input {
            display: none;
        }

        .controls {
            background: white;
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        .controls-row {
            display: flex;
            gap: 12px;
            margin-bottom: 16px;
        }

        .control-group {
            flex: 1;
        }

        .control-label {
            font-size: 12px;
            font-weight: 600;
            color: var(--gray-600);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
            display: block;
        }

        .select-input {
            width: 100%;
            padding: 12px;
            background: var(--gray-50);
            border: 1.5px solid var(--gray-200);
            border-radius: 8px;
            font-size: 14px;
            font-family: inherit;
            color: var(--gray-800);
            transition: all 0.2s;
        }

        .select-input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1);
        }

        .btn-group {
            display: flex;
            gap: 12px;
        }

        .action-btn {
            flex: 1;
            padding: 16px;
            border: none;
            border-radius: 10px;
            font-weight: 700;
            font-size: 15px;
            cursor: pointer;
            transition: all 0.3s;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .compare-btn {
            background: linear-gradient(135deg, var(--success), #059669);
            color: white;
        }

        .compare-btn:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
        }

        .export-btn {
            background: linear-gradient(135deg, var(--primary), var(--primary-dark));
            color: white;
        }

        .export-btn:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(99, 102, 241, 0.4);
        }

        .action-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .summary {
            background: white;
            border-radius: 16px;
            padding: 32px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        .summary-title {
            font-size: 20px;
            font-weight: 700;
            margin-bottom: 24px;
            color: var(--gray-800);
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
        }

        .stat-card {
            background: var(--gray-50);
            padding: 20px;
            border-radius: 12px;
            border-left: 4px solid var(--primary);
        }

        .stat-card.success {
            border-left-color: var(--success);
        }

        .stat-card.warning {
            border-left-color: var(--warning);
        }

        .stat-card.danger {
            border-left-color: var(--danger);
        }

        .stat-value {
            font-size: 32px;
            font-weight: 800;
            margin-bottom: 6px;
        }

        .stat-value.success {
            color: var(--success);
        }

        .stat-value.warning {
            color: var(--warning);
        }

        .stat-value.danger {
            color: var(--danger);
        }

        .stat-label {
            font-size: 12px;
            color: var(--gray-600);
            font-weight: 600;
        }

        .results {
            background: white;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            overflow: hidden;
        }

        .tabs {
            display: flex;
            background: var(--gray-50);
            padding: 8px;
            gap: 4px;
        }

        .tab {
            flex: 1;
            padding: 12px 20px;
            background: transparent;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            font-size: 13px;
            color: var(--gray-600);
            cursor: pointer;
            transition: all 0.2s;
        }

        .tab.active {
            background: white;
            color: var(--primary);
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }

        .tab:hover:not(.active) {
            color: var(--primary);
        }

        .table-container {
            overflow-x: auto;
            max-height: 600px;
            overflow-y: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        th {
            background: var(--gray-50);
            padding: 14px 12px;
            text-align: left;
            font-size: 11px;
            font-weight: 700;
            color: var(--gray-600);
            text-transform: uppercase;
            letter-spacing: 0.8px;
            position: sticky;
            top: 0;
            z-index: 10;
            border-bottom: 2px solid var(--gray-200);
        }

        td {
            padding: 12px;
            border-bottom: 1px solid var(--gray-100);
            font-size: 13px;
        }

        tr:hover {
            background: var(--gray-50);
        }

        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
        }

        .badge.match {
            background: var(--green-50);
            color: var(--green-600);
        }

        .badge.partial {
            background: var(--yellow-50);
            color: var(--yellow-600);
        }

        .badge.missing {
            background: var(--red-50);
            color: var(--red-600);
        }

        .highlight-red {
            background: var(--red-100) !important;
            color: var(--red-600) !important;
            font-weight: 600;
        }

        .highlight-green {
            background: var(--green-50) !important;
            color: var(--green-600) !important;
            font-weight: 600;
        }

        .side-by-side {
            display: grid;
            grid-template-columns: 1fr 1px 1fr;
            gap: 8px;
            align-items: start;
        }

        .side-label {
            font-size: 10px;
            font-weight: 700;
            color: var(--gray-600);
            text-transform: uppercase;
            margin-bottom: 4px;
        }

        .side-value {
            padding: 6px 8px;
            border-radius: 4px;
            background: var(--gray-50);
            font-size: 12px;
        }

        .divider {
            width: 1px;
            background: var(--gray-200);
            height: 100%;
        }

        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: var(--gray-600);
        }

        @media (max-width: 768px) {
            .upload-grid {
                grid-template-columns: 1fr;
            }
            .btn-group {
                flex-direction: column;
            }
        }
    </style>
</head>
<body>
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useMemo } = React;

        function InvoiceReconciliation() {
            const [workbook1, setWorkbook1] = useState(null);
            const [workbook2, setWorkbook2] = useState(null);
            const [data1, setData1] = useState(null);
            const [data2, setData2] = useState(null);
            const [results, setResults] = useState(null);
            const [activeTab, setActiveTab] = useState('matched');
            const [sortBy, setSortBy] = useState('none');
            const [filterHQ, setFilterHQ] = useState('all');

            const extractInvoiceNumber = (invoiceStr) => {
                if (!invoiceStr) return '';
                const str = String(invoiceStr).trim();
                const dashMatch = str.match(/--(\d+)/);
                if (dashMatch) return dashMatch[1];
                const numMatch = str.match(/(\d+)$/);
                if (numMatch) return numMatch[1];
                return str;
            };

            const normalizeCompanyName = (name) => {
                if (!name) return '';
                return String(name)
                    .toUpperCase()
                    .replace(/SDN\.?\s*BHD\.?/gi, '')
                    .replace(/\(SSD\)/gi, '')
                    .replace(/BERHAD/gi, '')
                    .replace(/\s+/g, ' ')
                    .trim();
            };

            const companyNamesMatch = (name1, name2) => {
                const norm1 = normalizeCompanyName(name1);
                const norm2 = normalizeCompanyName(name2);
                if (!norm1 || !norm2) return false;
                return norm1.includes(norm2) || norm2.includes(norm1) || norm1 === norm2;
            };

            const excelDateToJSDate = (serial) => {
                // Excel stores dates as days since 1900-01-01
                // Excel incorrectly treats 1900 as a leap year, so we need to account for that
                const utc_days = Math.floor(serial - 25569);
                const utc_value = utc_days * 86400;
                const date_info = new Date(utc_value * 1000);
                return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
            };

            const formatDate = (dateValue) => {
                if (!dateValue) return '';
                
                // Check if it's an Excel serial number (numeric value > 1000)
                if (typeof dateValue === 'number' && dateValue > 1000) {
                    const date = excelDateToJSDate(dateValue);
                    const day = String(date.getDate()).padStart(2, '0');
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const year = date.getFullYear();
                    return `${day}/${month}/${year}`;
                }
                
                // If it's already a Date object
                if (dateValue instanceof Date) {
                    const day = String(dateValue.getDate()).padStart(2, '0');
                    const month = String(dateValue.getMonth() + 1).padStart(2, '0');
                    const year = dateValue.getFullYear();
                    return `${day}/${month}/${year}`;
                }
                
                // If it's a string, return as is
                return String(dateValue);
            };

            const parseDate = (dateStr) => {
                if (!dateStr) return null;
                
                // Handle Excel serial numbers
                if (typeof dateStr === 'number' && dateStr > 1000) {
                    return excelDateToJSDate(dateStr);
                }
                
                if (dateStr instanceof Date) return dateStr;
                
                const str = String(dateStr).trim();
                const parts = str.split('/');
                if (parts.length === 3) {
                    return new Date(parts[2], parts[1] - 1, parts[0]);
                }
                return new Date(dateStr);
            };

            const datesMatch = (date1, date2) => {
                const d1 = parseDate(date1);
                const d2 = parseDate(date2);
                if (!d1 || !d2 || isNaN(d1) || isNaN(d2)) return false;
                return d1.toDateString() === d2.toDateString();
            };

            const amountsMatch = (amt1, amt2) => {
                const a1 = parseFloat(amt1) || 0;
                const a2 = parseFloat(amt2) || 0;
                return Math.abs(a1 - a2) < 0.01;
            };

            const handleFileUpload = (e, fileNum) => {
                const file = e.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = (event) => {
                    const workbook = XLSX.read(event.target.result, { type: 'binary' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    if (fileNum === 1) {
                        // Workbook 1 - header on row 4 (index 3)
                        const json = XLSX.utils.sheet_to_json(sheet, { 
                            header: 1, 
                            defval: '',
                            range: 3 
                        });
                        const headers = json[0];
                        const data = json.slice(1).map(row => {
                            const obj = {};
                            headers.forEach((header, idx) => {
                                obj[header] = row[idx];
                            });
                            return obj;
                        }).filter(row => row['No Invois']);
                        
                        setWorkbook1(file);
                        setData1(data);
                    } else {
                        // Account file - header on row 3 (index 2)
                        const json = XLSX.utils.sheet_to_json(sheet, { 
                            header: 1,
                            defval: '',
                            range: 2 
                        });
                        const headers = json[0];
                        const data = json.slice(1).map(row => {
                            const obj = {};
                            headers.forEach((header, idx) => {
                                obj[header] = row[idx];
                            });
                            return obj;
                        }).filter(row => row['Invoice No.']);
                        
                        setWorkbook2(file);
                        setData2(data);
                    }
                };
                reader.readAsBinaryString(file);
            };

            const compareData = () => {
                if (!data1 || !data2) return;

                const matched = [];
                const partiallySame = [];
                const missingInAccount = [];
                const missingInWorkbook1 = [];

                const wb1Map = new Map();
                data1.forEach(row => {
                    const invoiceNum = extractInvoiceNumber(row['No Invois']);
                    if (invoiceNum) {
                        wb1Map.set(invoiceNum, row);
                    }
                });

                const matchedWb1Invoices = new Set();

                // Process Account file (Workbook 2)
                data2.forEach(accountRow => {
                    const invoiceNo = accountRow['Invoice No.'];
                    const extractedNum = extractInvoiceNumber(invoiceNo);
                    
                    if (wb1Map.has(extractedNum)) {
                        const wb1Row = wb1Map.get(extractedNum);
                        matchedWb1Invoices.add(extractedNum);
                        
                        // Compare all fields
                        const dateMatch = datesMatch(wb1Row['Tarikh'], accountRow['Date']);
                        const amountMatch = amountsMatch(wb1Row['Total'], accountRow['Net Total']);
                        const outstandingMatch = amountsMatch(wb1Row['Outstanding'], accountRow['Outstanding']);
                        
                        const pembaik = wb1Row['Pembaik'];
                        const pemilik = wb1Row['Pemilik'];
                        const companyName = accountRow['Company Name'];
                        
                        const companyMatch = companyNamesMatch(companyName, pembaik) || 
                                           companyNamesMatch(companyName, pemilik);
                        
                        const allMatch = dateMatch && amountMatch && outstandingMatch && companyMatch;
                        const hasAnyMismatch = !dateMatch || !amountMatch || !outstandingMatch || !companyMatch;
                        
                        const record = {
                            // Account file fields (as primary)
                            invoiceNo: accountRow['Invoice No.'],
                            date_account: formatDate(accountRow['Date']),
                            companyName: accountRow['Company Name'],
                            netTotal_account: accountRow['Net Total'],
                            outstanding_account: accountRow['Outstanding'],
                            hqCawangan: accountRow['HQ / Cawangan'],
                            
                            // Workbook 1 fields
                            invoiceNo_wb1: wb1Row['No Invois'],
                            date_wb1: formatDate(wb1Row['Tarikh']),
                            pembaik: pembaik,
                            pemilik: pemilik,
                            total_wb1: wb1Row['Total'],
                            outstanding_wb1: wb1Row['Outstanding'],
                            
                            // Match status
                            dateMatch,
                            amountMatch,
                            outstandingMatch,
                            companyMatch,
                            allMatch,
                            status: allMatch ? 'SAME' : 'PARTIALLY SAME'
                        };
                        
                        if (allMatch) {
                            matched.push(record);
                        } else {
                            partiallySame.push(record);
                        }
                    } else {
                        missingInWorkbook1.push({
                            invoiceNo: invoiceNo,
                            date: formatDate(accountRow['Date']),
                            companyName: accountRow['Company Name'],
                            netTotal: accountRow['Net Total'],
                            outstanding: accountRow['Outstanding'],
                            hqCawangan: accountRow['HQ / Cawangan']
                        });
                    }
                });

                // Find missing in Account
                data1.forEach(row => {
                    const invoiceNum = extractInvoiceNumber(row['No Invois']);
                    if (invoiceNum && !matchedWb1Invoices.has(invoiceNum)) {
                        missingInAccount.push({
                            invoiceNo: row['No Invois'],
                            date: formatDate(row['Tarikh']),
                            pembaik: row['Pembaik'],
                            pemilik: row['Pemilik'],
                            total: row['Total'],
                            outstanding: row['Outstanding']
                        });
                    }
                });

                setResults({
                    matched,
                    partiallySame,
                    missingInAccount,
                    missingInWorkbook1,
                    summary: {
                        total_wb1: data1.length,
                        total_account: data2.length,
                        same: matched.length,
                        partiallySame: partiallySame.length,
                        missingInAccount: missingInAccount.length,
                        missingInWorkbook1: missingInWorkbook1.length
                    }
                });
            };

            const sortedData = useMemo(() => {
                if (!results) return null;

                let data = null;
                if (activeTab === 'matched') data = [...results.matched];
                else if (activeTab === 'partial') data = [...results.partiallySame];
                else if (activeTab === 'missing-account') data = results.missingInAccount;
                else if (activeTab === 'missing-wb1') data = results.missingInWorkbook1;

                if (!data) return null;

                if (sortBy === 'hq' && (activeTab === 'matched' || activeTab === 'partial')) {
                    return data.sort((a, b) => (a.hqCawangan || '').localeCompare(b.hqCawangan || ''));
                } else if (sortBy === 'pembaik' && (activeTab === 'matched' || activeTab === 'partial')) {
                    return data.sort((a, b) => (a.pembaik || '').localeCompare(b.pembaik || ''));
                } else if (sortBy === 'pemilik' && (activeTab === 'matched' || activeTab === 'partial')) {
                    return data.sort((a, b) => (a.pemilik || '').localeCompare(b.pemilik || ''));
                }

                return data;
            }, [results, activeTab, sortBy]);

            const filteredData = useMemo(() => {
                if (!sortedData) return null;
                if (filterHQ === 'all' || activeTab === 'missing-account' || activeTab === 'missing-wb1') {
                    return sortedData;
                }
                return sortedData.filter(row => row.hqCawangan === filterHQ);
            }, [sortedData, filterHQ]);

            const hqOptions = useMemo(() => {
                if (!results) return [];
                const all = [...results.matched, ...results.partiallySame];
                const unique = [...new Set(all.map(r => r.hqCawangan).filter(Boolean))];
                return unique.sort();
            }, [results]);

            const exportToExcel = () => {
                if (!results) return;

                const wb = XLSX.utils.book_new();

                // Summary sheet
                const summaryData = [
                    ['Invoice Reconciliation Report'],
                    [''],
                    ['Summary'],
                    ['Total in Workbook 1', results.summary.total_wb1],
                    ['Total in Account', results.summary.total_account],
                    ['SAME (All Fields Match)', results.summary.same],
                    ['PARTIALLY SAME (Some Fields Mismatch)', results.summary.partiallySame],
                    ['Missing in Account', results.summary.missingInAccount],
                    ['Missing in Workbook 1', results.summary.missingInWorkbook1]
                ];
                const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
                XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

                // Same data sheet
                if (results.matched.length > 0) {
                    const sameData = [
                        ['Invoice No.', 'Date', 'Company Name', 'Net Total', 'Outstanding', 'HQ / Cawangan',
                         'WB1 Invoice', 'WB1 Date', 'Pembaik', 'Pemilik', 'WB1 Total', 'WB1 Outstanding'],
                        ...results.matched.map(r => [
                            r.invoiceNo, r.date_account, r.companyName, r.netTotal_account, r.outstanding_account, r.hqCawangan,
                            r.invoiceNo_wb1, r.date_wb1, r.pembaik, r.pemilik, r.total_wb1, r.outstanding_wb1
                        ])
                    ];
                    const sameSheet = XLSX.utils.aoa_to_sheet(sameData);
                    XLSX.utils.book_append_sheet(wb, sameSheet, 'SAME');
                }

                // Partially same sheet
                if (results.partiallySame.length > 0) {
                    const partialData = [
                        ['Invoice No.', 'Date (Account)', 'Date (WB1)', 'Date Match', 
                         'Company Name', 'Pembaik', 'Pemilik', 'Company Match',
                         'Net Total (Account)', 'Total (WB1)', 'Amount Match',
                         'Outstanding (Account)', 'Outstanding (WB1)', 'Outstanding Match',
                         'HQ / Cawangan'],
                        ...results.partiallySame.map(r => [
                            r.invoiceNo, r.date_account, r.date_wb1, r.dateMatch ? 'YES' : 'NO',
                            r.companyName, r.pembaik, r.pemilik, r.companyMatch ? 'YES' : 'NO',
                            r.netTotal_account, r.total_wb1, r.amountMatch ? 'YES' : 'NO',
                            r.outstanding_account, r.outstanding_wb1, r.outstandingMatch ? 'YES' : 'NO',
                            r.hqCawangan
                        ])
                    ];
                    const partialSheet = XLSX.utils.aoa_to_sheet(partialData);
                    XLSX.utils.book_append_sheet(wb, partialSheet, 'PARTIALLY SAME');
                }

                // Missing sheets
                if (results.missingInAccount.length > 0) {
                    const missingAccData = [
                        ['Invoice No.', 'Date', 'Pembaik', 'Pemilik', 'Total', 'Outstanding'],
                        ...results.missingInAccount.map(r => [
                            r.invoiceNo, r.date, r.pembaik, r.pemilik, r.total, r.outstanding
                        ])
                    ];
                    const missingAccSheet = XLSX.utils.aoa_to_sheet(missingAccData);
                    XLSX.utils.book_append_sheet(wb, missingAccSheet, 'Missing in Account');
                }

                if (results.missingInWorkbook1.length > 0) {
                    const missingWb1Data = [
                        ['Invoice No.', 'Date', 'Company Name', 'Net Total', 'Outstanding', 'HQ / Cawangan'],
                        ...results.missingInWorkbook1.map(r => [
                            r.invoiceNo, r.date, r.companyName, r.netTotal, r.outstanding, r.hqCawangan
                        ])
                    ];
                    const missingWb1Sheet = XLSX.utils.aoa_to_sheet(missingWb1Data);
                    XLSX.utils.book_append_sheet(wb, missingWb1Sheet, 'Missing in Workbook1');
                }

                XLSX.writeFile(wb, 'invoice-reconciliation-detailed-report.xlsx');
            };

            return (
                <div className="app">
                    <div className="header">
                        <h1 className="title">Advanced Invoice Reconciliation</h1>
                        <p className="subtitle">Compare Workbook 1 (Internal) vs Account File with detailed matching and highlighting</p>
                    </div>

                    <div className="upload-grid">
                        <div className="upload-card">
                            <div className="upload-header">Workbook 1 (Internal Records)</div>
                            <div className="upload-desc">Header Row 4 • No Invois, Tarikh, Pembaik, Pemilik, Total, Outstanding</div>
                            <input type="file" accept=".xlsx,.xls" className="file-input" id="file1"
                                onChange={(e) => handleFileUpload(e, 1)} />
                            <button className="upload-btn" onClick={() => document.getElementById('file1').click()}>
                                Upload Workbook 1
                            </button>
                            {workbook1 && <div className="file-name">✓ {workbook1.name}</div>}
                        </div>

                        <div className="upload-card">
                            <div className="upload-header">Account File</div>
                            <div className="upload-desc">Header Row 3 • Invoice No., Date, Company Name, Net Total, Outstanding, HQ/Cawangan</div>
                            <input type="file" accept=".xlsx,.xls" className="file-input" id="file2"
                                onChange={(e) => handleFileUpload(e, 2)} />
                            <button className="upload-btn" onClick={() => document.getElementById('file2').click()}>
                                Upload Account File
                            </button>
                            {workbook2 && <div className="file-name">✓ {workbook2.name}</div>}
                        </div>
                    </div>

                    {data1 && data2 && (
                        <div className="controls">
                            <div className="controls-row">
                                <div className="control-group">
                                    <label className="control-label">Sort By</label>
                                    <select className="select-input" value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
                                        <option value="none">Default Order</option>
                                        <option value="hq">HQ / Cawangan (from Account)</option>
                                        <option value="pembaik">Pembaik (from Workbook 1)</option>
                                        <option value="pemilik">Pemilik (from Workbook 1)</option>
                                    </select>
                                </div>

                                <div className="control-group">
                                    <label className="control-label">Filter by HQ/Cawangan</label>
                                    <select className="select-input" value={filterHQ} onChange={(e) => setFilterHQ(e.target.value)}>
                                        <option value="all">All Locations</option>
                                        {hqOptions.map((hq, idx) => (
                                            <option key={idx} value={hq}>{hq}</option>
                                        ))}
                                    </select>
                                </div>
                            </div>

                            <div className="btn-group">
                                <button className="action-btn compare-btn" onClick={compareData}>
                                    🔍 Compare & Analyze
                                </button>
                                <button className="action-btn export-btn" onClick={exportToExcel} disabled={!results}>
                                    📊 Export Report (3 Sheets)
                                </button>
                            </div>
                        </div>
                    )}

                    {results && (
                        <>
                            <div className="summary">
                                <div className="summary-title">Reconciliation Summary</div>
                                <div className="stats-grid">
                                    <div className="stat-card success">
                                        <div className="stat-value success">{results.summary.same}</div>
                                        <div className="stat-label">SAME (Perfect Match)</div>
                                    </div>
                                    <div className="stat-card warning">
                                        <div className="stat-value warning">{results.summary.partiallySame}</div>
                                        <div className="stat-label">PARTIALLY SAME</div>
                                    </div>
                                    <div className="stat-card danger">
                                        <div className="stat-value danger">{results.summary.missingInAccount}</div>
                                        <div className="stat-label">Missing in Account</div>
                                    </div>
                                    <div className="stat-card danger">
                                        <div className="stat-value danger">{results.summary.missingInWorkbook1}</div>
                                        <div className="stat-label">Missing in Workbook 1</div>
                                    </div>
                                </div>
                            </div>

                            <div className="results">
                                <div className="tabs">
                                    <button className={`tab ${activeTab === 'matched' ? 'active' : ''}`}
                                        onClick={() => setActiveTab('matched')}>
                                        SAME ({results.matched.length})
                                    </button>
                                    <button className={`tab ${activeTab === 'partial' ? 'active' : ''}`}
                                        onClick={() => setActiveTab('partial')}>
                                        PARTIALLY SAME ({results.partiallySame.length})
                                    </button>
                                    <button className={`tab ${activeTab === 'missing-account' ? 'active' : ''}`}
                                        onClick={() => setActiveTab('missing-account')}>
                                        Missing in Account ({results.missingInAccount.length})
                                    </button>
                                    <button className={`tab ${activeTab === 'missing-wb1' ? 'active' : ''}`}
                                        onClick={() => setActiveTab('missing-wb1')}>
                                        Missing in WB1 ({results.missingInWorkbook1.length})
                                    </button>
                                </div>

                                <div className="table-container">
                                    {(activeTab === 'matched' || activeTab === 'partial') && filteredData && (
                                        <table>
                                            <thead>
                                                <tr>
                                                    <th>Invoice No.</th>
                                                    <th>Date</th>
                                                    <th>Company Name</th>
                                                    <th>Pembaik / Pemilik</th>
                                                    <th>Net Total</th>
                                                    <th>Outstanding</th>
                                                    <th>HQ / Cawangan</th>
                                                    <th>Status</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {filteredData.map((row, idx) => (
                                                    <tr key={idx}>
                                                        <td>{row.invoiceNo}</td>
                                                        <td className={!row.dateMatch ? 'highlight-red' : ''}>
                                                            {row.date_account}
                                                            {!row.dateMatch && <div style={{fontSize: '10px', marginTop: '4px'}}>WB1: {row.date_wb1}</div>}
                                                        </td>
                                                        <td className={!row.companyMatch ? 'highlight-red' : ''}>
                                                            {row.companyName}
                                                            {!row.companyMatch && (
                                                                <div style={{fontSize: '10px', marginTop: '4px'}}>
                                                                    Pembaik: {row.pembaik}<br/>Pemilik: {row.pemilik}
                                                                </div>
                                                            )}
                                                        </td>
                                                        <td>{row.pembaik} / {row.pemilik}</td>
                                                        <td className={!row.amountMatch ? 'highlight-red' : ''}>
                                                            {row.netTotal_account}
                                                            {!row.amountMatch && <div style={{fontSize: '10px', marginTop: '4px'}}>WB1: {row.total_wb1}</div>}
                                                        </td>
                                                        <td className={!row.outstandingMatch ? 'highlight-red' : ''}>
                                                            {row.outstanding_account}
                                                            {!row.outstandingMatch && <div style={{fontSize: '10px', marginTop: '4px'}}>WB1: {row.outstanding_wb1}</div>}
                                                        </td>
                                                        <td>{row.hqCawangan}</td>
                                                        <td>
                                                            <span className={`badge ${row.allMatch ? 'match' : 'partial'}`}>
                                                                {row.status}
                                                            </span>
                                                        </td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}

                                    {activeTab === 'missing-account' && (
                                        <table>
                                            <thead>
                                                <tr>
                                                    <th>Invoice No.</th>
                                                    <th>Date</th>
                                                    <th>Pembaik</th>
                                                    <th>Pemilik</th>
                                                    <th>Total</th>
                                                    <th>Outstanding</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {results.missingInAccount.map((row, idx) => (
                                                    <tr key={idx}>
                                                        <td>{row.invoiceNo}</td>
                                                        <td>{row.date}</td>
                                                        <td>{row.pembaik}</td>
                                                        <td>{row.pemilik}</td>
                                                        <td>{row.total}</td>
                                                        <td>{row.outstanding}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}

                                    {activeTab === 'missing-wb1' && (
                                        <table>
                                            <thead>
                                                <tr>
                                                    <th>Invoice No.</th>
                                                    <th>Date</th>
                                                    <th>Company Name</th>
                                                    <th>Net Total</th>
                                                    <th>Outstanding</th>
                                                    <th>HQ / Cawangan</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {results.missingInWorkbook1.map((row, idx) => (
                                                    <tr key={idx}>
                                                        <td>{row.invoiceNo}</td>
                                                        <td>{row.date}</td>
                                                        <td>{row.companyName}</td>
                                                        <td>{row.netTotal}</td>
                                                        <td>{row.outstanding}</td>
                                                        <td>{row.hqCawangan}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}
                                </div>
                            </div>
                        </>
                    )}

                    {!results && (
                        <div className="empty-state">
                            <div style={{fontSize: '48px', marginBottom: '16px'}}>📊</div>
                            <p>Upload both files and click "Compare & Analyze" to start reconciliation</p>
                        </div>
                    )}
                </div>
            );
        }

        ReactDOM.render(<InvoiceReconciliation />, document.getElementById('root'));
    </script>
</body>
</html>
