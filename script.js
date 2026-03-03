document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('verification-form');
    const input = document.getElementById('certificate-input');
    const verifyBtn = document.getElementById('verify-btn');
    const errorMessage = document.getElementById('error-message');
    const resultSection = document.getElementById('result-section');

    // Result elements
    const resCertNo = document.getElementById('res-cert-no');
    const resName = document.getElementById('res-name');
    const resProgramme = document.getElementById('res-programme');
    const resDate = document.getElementById('res-date');
    const resWorkplace = document.getElementById('res-workplace');
    const resSerial = document.getElementById('res-serial');
    const resPrintDate = document.getElementById('res-print-date');

    const SPREADSHEET_ID = '2PACX-1vS-E6XqaS3hZ9FqMUtZgWKs6UvkNiKc8ytlUyHh6FJZed7Lo7Z3T4R2mfEPA8-TcqpF6u7mLfou4BLU';
    const PUBHTML_URL = `https://docs.google.com/spreadsheets/d/e/${SPREADSHEET_ID}/pubhtml`;

    // Helper: CORS proxy list (tried in order, cycling per attempt)
    const PROXIES = [
        (url) => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
        (url) => `https://corsproxy.io/?${encodeURIComponent(url)}`,
        (url) => `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(url)}`,
    ];
    const getProxyUrl = (url, proxyIndex = 0) => PROXIES[proxyIndex % PROXIES.length](url);

    // Helper: retry fetch cycling through proxies on each failure
    const fetchWithRetry = async (url, maxRetries = 6, delayMs = 1000) => {
        for (let attempt = 0; attempt < maxRetries; attempt++) {
            const proxyUrl = getProxyUrl(url, attempt);
            try {
                const response = await fetch(proxyUrl, { cache: 'no-store' });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                return response;
            } catch (err) {
                if (attempt === maxRetries - 1) throw err;
                const proxyName = ['allorigins', 'corsproxy.io', 'codetabs'][attempt % PROXIES.length];
                console.warn(`Fetch attempt ${attempt + 1} via ${proxyName} failed, trying next…`);
                await new Promise(res => setTimeout(res, delayMs));
            }
        }
    };

    // Cache
    let sheetIdMap = null;
    let cachedDataByYear = {};
    let isMapLoaded = false;
    let mapLoadPromise = null;  // prevent parallel duplicate fetches

    // Year dropdown
    const yearSelectElement = document.getElementById('year-select');

    // Load sheet map (with retry) and populate dropdown
    const loadAvailableYears = () => {
        if (mapLoadPromise) return mapLoadPromise;   // already in-flight or done

        mapLoadPromise = (async () => {
            if (isMapLoaded) return;

            sheetIdMap = {};
            try {
                const htmlResponse = await fetchWithRetry(getProxyUrl(PUBHTML_URL));
                const htmlText = await htmlResponse.text();

                // Extract GIDs from Google's inline JS
                const regex = /\{name:\s*"([^"]+)"[^\}]+gid:\s*"(\d+)"/g;
                let match;
                while ((match = regex.exec(htmlText)) !== null) {
                    sheetIdMap[match[1]] = match[2];
                }

                isMapLoaded = true;

                // Populate year dropdown with only real sheet years
                if (yearSelectElement) {
                    yearSelectElement.innerHTML = '<option value="">Auto-Detect Year</option>';
                    const years = Object.keys(sheetIdMap)
                        .map(n => { const m = n.match(/^(\d{4})/); return m ? parseInt(m[1], 10) : null; })
                        .filter(Boolean)
                        .sort((a, b) => b - a);

                    years.forEach(year => {
                        const opt = document.createElement('option');
                        opt.value = year;
                        opt.textContent = year;
                        yearSelectElement.appendChild(opt);
                    });
                }
            } catch (error) {
                console.error('Failed to load sheet map:', error);
                mapLoadPromise = null;   // allow retry on next verify click
                throw error;
            }
        })();

        return mapLoadPromise;
    };

    // Warm up the connection immediately on page load
    loadAvailableYears().catch(() => {/* silent – will retry on verify click */ });

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const certNumber = input.value.trim();
        if (!certNumber) return;

        hideError();
        hideResult();
        setLoading(true, 'Connecting to database…');

        try {
            // 1. Parse year from certificate number
            const yearPrefixStr = certNumber.substring(0, 2);
            if (!/^\d{2}$/.test(yearPrefixStr)) throw new Error('invalid_format');

            const fullYear = 2000 + parseInt(yearPrefixStr, 10);
            const selectedYear = (yearSelectElement && yearSelectElement.value)
                ? parseInt(yearSelectElement.value, 10)
                : fullYear;

            if (selectedYear < 2026) {
                showError('Certificates issued before 2026 are not available to verify online. Please contact NILS for manual verification.');
                setLoading(false);
                return;
            }

            const targetSheetName = `${selectedYear} Certificate Registry`;

            // 2. Ensure sheet map is loaded (waits & retries if needed)
            if (!isMapLoaded) {
                setLoading(true, 'Loading registry index…');
                await loadAvailableYears();
            }

            const targetGid = sheetIdMap?.[targetSheetName];
            if (!targetGid) throw new Error('not_found');

            // 3. Fetch CSV for the target year (with retry via Papa + proxy)
            if (!cachedDataByYear[selectedYear]) {
                setLoading(true, 'Fetching certificate data…');
                const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SPREADSHEET_ID}/pub?gid=${targetGid}&single=true&output=csv`;

                // Retry CSV fetch cycling through all proxy fallbacks (6 attempts)
                let lastError;
                const MAX_CSV_ATTEMPTS = 6;
                for (let attempt = 0; attempt < MAX_CSV_ATTEMPTS; attempt++) {
                    const proxiedCsvUrl = getProxyUrl(csvUrl, attempt);
                    try {
                        cachedDataByYear[selectedYear] = await new Promise((resolve, reject) => {
                            Papa.parse(proxiedCsvUrl, {
                                download: true,
                                header: true,
                                skipEmptyLines: true,
                                complete: (results) => resolve(results.data),
                                error: (err) => reject(err)
                            });
                        });
                        break; // success
                    } catch (err) {
                        lastError = err;
                        if (attempt < MAX_CSV_ATTEMPTS - 1) {
                            const proxyName = ['allorigins', 'corsproxy.io', 'codetabs'][attempt % PROXIES.length];
                            console.warn(`CSV attempt ${attempt + 1} via ${proxyName} failed, trying next…`);
                            await new Promise(res => setTimeout(res, 1000));
                        }
                    }
                }
                if (!cachedDataByYear[selectedYear]) throw lastError;
            }

            // 4. Search for the certificate
            setLoading(true, 'Searching…');
            const sheetData = cachedDataByYear[selectedYear];
            const foundRecord = sheetData.find(row =>
                row['Certificate No'] && String(row['Certificate No']).trim() === certNumber
            );

            if (foundRecord) {
                resCertNo.textContent = foundRecord['Certificate No'] || 'N/A';
                resName.textContent = foundRecord['Name with initial'] || 'N/A';
                resProgramme.textContent = foundRecord['Programme Name'] || 'N/A';
                resDate.textContent = foundRecord['Date'] || 'N/A';
                resWorkplace.textContent = foundRecord['Working Place'] || 'N/A';
                resSerial.textContent = foundRecord['Serial No'] || 'N/A';
                resPrintDate.textContent = foundRecord['Date of Printing'] || 'N/A';

                // Populate print document fields
                const now = new Date();
                const refDate = now.toISOString().replace('T', ' ').substring(0, 19);
                const refSerial = (foundRecord['Serial No'] || certNumber).replace(/\s/g, '');
                document.getElementById('print-reference').textContent =
                    `${refDate.replace(/-/g, '').replace(/:/g, '').replace(' ', '-')}-${refSerial}`;
                document.getElementById('print-name').textContent = foundRecord['Name with initial'] || 'N/A';
                document.getElementById('print-programme').textContent = foundRecord['Programme Name'] || 'N/A';
                document.getElementById('print-award-date').textContent = foundRecord['Date'] || 'N/A';
                document.getElementById('print-cert-no').textContent = foundRecord['Certificate No'] || 'N/A';
                document.getElementById('print-serial').textContent = foundRecord['Serial No'] || 'N/A';
                document.getElementById('print-workplace').textContent = foundRecord['Working Place'] || 'N/A';
                document.getElementById('print-generated-date').textContent =
                    now.toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' }) +
                    ' ' + now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

                showResult();
            } else {
                throw new Error('not_found_in_sheet');
            }

        } catch (error) {
            console.error('Verification error:', error);
            if (['invalid_format', 'not_found', 'not_found_in_sheet'].includes(error.message)) {
                showError('Certificate not found. Please review the number and try again. For more information please contact NILS.');
            } else {
                showError('Unable to connect to the verification database. Please try again later.');
            }
        } finally {
            setLoading(false);
        }
    });

    // ---- UI helpers ----
    const setLoading = (isLoading, statusText = 'Verifying…') => {
        const btnText = verifyBtn.querySelector('.btn-text');
        if (isLoading) {
            verifyBtn.classList.add('loading');
            if (btnText) btnText.textContent = statusText;
            input.disabled = true;
            if (yearSelectElement) yearSelectElement.disabled = true;
        } else {
            verifyBtn.classList.remove('loading');
            if (btnText) btnText.textContent = 'Verify';
            input.disabled = false;
            if (yearSelectElement) yearSelectElement.disabled = false;
        }
    };

    const showError = (msg) => {
        document.getElementById('error-text').textContent = msg;
        errorMessage.classList.remove('hidden');
    };

    const hideError = () => errorMessage.classList.add('hidden');
    const showResult = () => {
        resultSection.classList.remove('hidden');
        setTimeout(() => resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' }), 100);
    };
    const hideResult = () => resultSection.classList.add('hidden');
});

// ---- Print / Download as PDF ----
function printVerificationDoc() {
    window.print();
}
