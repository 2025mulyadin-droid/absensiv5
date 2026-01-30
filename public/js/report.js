let currentData = null;

// Initialize dates
document.addEventListener('DOMContentLoaded', () => {
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);

    // Adjust to local date string YYYY-MM-DD
    const formatDate = (d) => d.toLocaleDateString('en-CA', { timeZone: 'Asia/Jakarta' });

    document.getElementById('startDate').value = formatDate(firstDay);
    document.getElementById('endDate').value = formatDate(today);
});

async function loadReport() {
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;

    if (!start || !end) {
        alert("Pilih tanggal terlebih dahulu");
        return;
    }

    const btn = document.querySelector('button[onclick="loadReport()"]');
    const originalText = btn.textContent;
    btn.textContent = "Memuat...";
    btn.disabled = true;

    try {
        const res = await fetch(`/api/report?start=${start}&end=${end}`);
        const data = await res.json();
        currentData = data;

        renderSummary(data.summary, data.list);
        renderDetail(data.list);

        document.getElementById('reportResult').style.display = 'block';
    } catch (e) {
        console.error(e);
        alert("Gagal memuat laporan");
    } finally {
        btn.textContent = originalText;
        btn.disabled = false;
    }
}

// Helper: Generate array of dates between start and end (inclusive)
function getDatesInRange(startDate, endDate) {
    const dates = [];
    let curr = new Date(startDate);
    const last = new Date(endDate);

    while (curr <= last) {
        dates.push(new Date(curr).toLocaleDateString('en-CA', { timeZone: 'Asia/Jakarta' }));
        curr.setDate(curr.getDate() + 1);
    }
    return dates;
}

// Helper: Check if date string YYYY-MM-DD is Weekend (Sat/Sun)
function isWeekend(dateStr) {
    const d = new Date(dateStr);
    const day = d.getUTCDay(); // 0=Sun, 6=Sat (Verified for YYYY-MM-DD format which parses as UTC)
    return day === 0 || day === 6;
}

function renderSummary(summary, list) {
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;
    const dateRange = getDatesInRange(startDate, endDate);

    // Create a map for quick lookup: user_id + date -> status
    const attendanceMap = new Map();
    list.forEach(item => {
        attendanceMap.set(`${item.name}_${item.date}`, item.status);
    });

    const thead = document.querySelector('#summaryTable thead tr');
    const tbody = document.querySelector('#summaryTable tbody');

    // Build Header
    // Order: Nama | JML HARI KERJA | H | S | I | TK | Date Columns (DD)
    let headerHtml = '<th>Nama</th>';
    headerHtml += '<th class="text-center" style="background:#f1f5f9; white-space:normal; width:80px; font-size:0.8rem;">JML HARI KERJA</th>';
    headerHtml += '<th class="text-center" style="color:var(--status-hadir)">H</th>';
    headerHtml += '<th class="text-center" style="color:var(--status-sakit)">S</th>';
    headerHtml += '<th class="text-center" style="color:var(--status-izin)">I</th>';
    headerHtml += '<th class="text-center" style="color:var(--status-alpha)">TK</th>';

    dateRange.forEach(date => {
        const dayLabel = date.split('-')[2]; // Get DD
        headerHtml += `<th class="text-center" style="min-width:40px;">${dayLabel}</th>`;
    });
    thead.innerHTML = headerHtml;

    // Build Body
    tbody.innerHTML = '';

    summary.forEach(row => {
        const tr = document.createElement('tr');

        // Calculate stats
        let tkCount = 0;
        let cellsHtml = '';

        dateRange.forEach(date => {
            const status = attendanceMap.get(`${row.name}_${date}`);
            let cellContent = '-';
            let cellClass = 'status-cell status-a';

            if (status === 'Hadir') {
                cellContent = 'H';
                cellClass = 'status-cell status-h';
            } else if (status === 'Sakit') {
                cellContent = 'S';
                cellClass = 'status-cell status-s';
            } else if (status === 'Izin') {
                cellContent = 'I';
                cellClass = 'status-cell status-i';
            } else {
                // No status
                if (!isWeekend(date)) {
                    tkCount++;
                    cellContent = 'TK';
                    cellClass = 'status-cell status-tk';
                }
            }
            cellsHtml += `<td class="${cellClass}">${cellContent}</td>`;
        });

        // Calculate JML HARI KERJA (Total H + S + I + TK)
        const totalHariKerja = (row.total_hadir || 0) + (row.total_sakit || 0) + (row.total_izin || 0) + tkCount;

        // Construct Row
        let rowHtml = `<td>${row.name}</td>`;
        rowHtml += `<td class="text-center" style="font-weight:bold; background:#f8fafc;">${totalHariKerja}</td>`;
        rowHtml += `<td class="text-center status-h">${row.total_hadir || 0}</td>`;
        rowHtml += `<td class="text-center status-s">${row.total_sakit || 0}</td>`;
        rowHtml += `<td class="text-center status-i">${row.total_izin || 0}</td>`;
        rowHtml += `<td class="text-center status-a" style="font-weight:bold; color:var(--status-alpha);">${tkCount}</td>`;

        rowHtml += cellsHtml;

        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });
}

function renderDetail(list) {
    const tbody = document.querySelector('#detailTable tbody');
    tbody.innerHTML = '';

    list.forEach(row => {
        const tr = document.createElement('tr');
        const badgeClass = row.status.toLowerCase();
        tr.innerHTML = `
            <td>${row.date}</td>
            <td>${row.name}</td>
            <td><span class="badge ${badgeClass}">${row.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function exportToExcel() {
    if (!currentData) return;

    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;
    const dateRange = getDatesInRange(startDate, endDate);

    // Prepare Data Array (Array of Arrays)
    const data = [];

    // Row 1: Title
    data.push([`Laporan Absensi GTK ISLAMADINA Periode ${startDate} - ${endDate}`]);
    data.push([]); // Empty row for spacing

    // Row 3: Headers
    const headers = ["Nama", "JML HARI KERJA", "H", "S", "I", "TK"];
    dateRange.forEach(d => headers.push(d.split('-')[2])); // Add DD
    data.push(headers);

    // Row 4+: Data
    const attendanceMap = new Map();
    currentData.list.forEach(item => {
        attendanceMap.set(`${item.name}_${item.date}`, item.status);
    });

    currentData.summary.forEach(row => {
        const rowData = [row.name];

        // Calculate TK
        let tkCount = 0;
        const dateCells = [];

        dateRange.forEach(date => {
            const status = attendanceMap.get(`${row.name}_${date}`);
            if (status === 'Hadir') dateCells.push('H');
            else if (status === 'Sakit') dateCells.push('S');
            else if (status === 'Izin') dateCells.push('I');
            else {
                if (!isWeekend(date)) {
                    tkCount++;
                    dateCells.push('TK');
                } else {
                    dateCells.push('-');
                }
            }
        });

        const totalHariKerja = (row.total_hadir || 0) + (row.total_sakit || 0) + (row.total_izin || 0) + tkCount;

        // Add Summaries
        rowData.push(totalHariKerja);
        rowData.push(row.total_hadir || 0);
        rowData.push(row.total_sakit || 0);
        rowData.push(row.total_izin || 0);
        rowData.push(tkCount);

        // Add Date Cells
        rowData.push(...dateCells);

        data.push(rowData);
    });

    // Create Worksheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Merge Title Cell
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } });

    // Create Workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekapitulasi");

    // Download
    XLSX.writeFile(wb, `Laporan_Absensi_${startDate}_sd_${endDate}.xlsx`);
}
