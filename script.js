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

    // Helper: CORS proxy
    const getProxyUrl = (url) => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`;

    // Helper: retry fetch up to `maxRetries` times with increasing delay
    const fetchWithRetry = async (url, maxRetries = 3, delayMs = 1500) => {
        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                const response = await fetch(url);
                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                return response;
            } catch (err) {
                if (attempt === maxRetries) throw err;
                console.warn(`Fetch attempt ${attempt} failed, retrying in ${delayMs * attempt}ms…`);
                await new Promise(res => setTimeout(res, delayMs * attempt));
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

                // Retry CSV fetch up to 3 times
                let lastError;
                for (let attempt = 1; attempt <= 3; attempt++) {
                    try {
                        cachedDataByYear[selectedYear] = await new Promise((resolve, reject) => {
                            Papa.parse(getProxyUrl(csvUrl), {
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
                        if (attempt < 3) {
                            console.warn(`CSV fetch attempt ${attempt} failed, retrying…`);
                            await new Promise(res => setTimeout(res, 1500 * attempt));
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


