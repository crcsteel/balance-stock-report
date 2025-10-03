const SHEET_ID = '1N_qX_beldGyLspfU6lgeskWd-aYu6ugAhRFTrHgO_Sw';
const SHEET_NAME = 'data_group';
const API_URL_INVENTORY = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${SHEET_NAME}&tqx=out:json`;

// ---- Helper ดึงและแปลงข้อมูล ----
async function fetchSheetData(url) {
    const res = await fetch(url);
    const text = await res.text();
    const json = JSON.parse(text.substring(47).slice(0, -2)); // clean gviz response
    const table = json.table;

    return table.rows.map(row => ({
        group: row.c[0]?.v || "",
        weightSurin: parseFloat(row.c[1]?.v) || 0,
        weightNangrong: parseFloat(row.c[2]?.v) || 0,
        weightDetUdom: parseFloat(row.c[3]?.v) || 0,
        costSurin: parseFloat(row.c[4]?.v) || 0,
        costNangrong: parseFloat(row.c[5]?.v) || 0,
        costDetUdom: parseFloat(row.c[6]?.v) || 0
    }));
}

// ---- Format ตัวเลขไทย ----
function formatNumber(num) {
    return Number(num).toLocaleString('th-TH');
}

// ---- Summary ----
function updateSummary(data) {
    let weightSurin = 0, weightNangrong = 0, weightDetUdom = 0;
    let costSurin = 0, costNangrong = 0, costDetUdom = 0;

    data.forEach(item => {
        weightSurin += item.weightSurin;
        weightNangrong += item.weightNangrong;
        weightDetUdom += item.weightDetUdom;
        costSurin += item.costSurin;
        costNangrong += item.costNangrong;
        costDetUdom += item.costDetUdom;
    });

    document.getElementById('weightSurin').textContent = formatNumber(weightSurin.toFixed(2));
    document.getElementById('weightNangrong').textContent = formatNumber(weightNangrong.toFixed(2));
    document.getElementById('weightDetUdom').textContent = formatNumber(weightDetUdom.toFixed(2));
    document.getElementById('weightTotal').textContent = formatNumber((weightSurin + weightNangrong + weightDetUdom).toFixed(2));

    document.getElementById('costSurin').textContent = formatNumber(costSurin);
    document.getElementById('costNangrong').textContent = formatNumber(costNangrong);
    document.getElementById('costDetUdom').textContent = formatNumber(costDetUdom);
    document.getElementById('costTotal').textContent = formatNumber(costSurin + costNangrong + costDetUdom);
}


// ---- Main ----
$(document).ready(async function() {
    const inventoryData = await fetchSheetData(API_URL_INVENTORY);

    // Init DataTable
    const table = $('#inventoryTable').DataTable({
        data: inventoryData,
        columns: [
            { data: 'group' },
            { data: 'weightSurin', className: 'text-right', render: d => formatNumber(d.toFixed(2)) },
            { data: 'weightNangrong', className: 'text-right', render: d => formatNumber(d.toFixed(2)) },
            { data: 'weightDetUdom', className: 'text-right', render: d => formatNumber(d.toFixed(2)) },
            { data: 'costSurin', className: 'text-right', render: d => formatNumber(d) },
            { data: 'costNangrong', className: 'text-right', render: d => formatNumber(d) },
            { data: 'costDetUdom', className: 'text-right', render: d => formatNumber(d) }
        ],
        language: {
            search: "ค้นหา:",
            lengthMenu: "แสดง _MENU_ รายการ",
            info: "แสดง _START_ ถึง _END_ จาก _TOTAL_ รายการ",
            infoEmpty: "แสดง 0 ถึง 0 จาก 0 รายการ",
            infoFiltered: "(กรองจาก _MAX_ รายการทั้งหมด)",
            paginate: {
                first: "แรก",
                last: "สุดท้าย",
                next: "ถัดไป",
                previous: "ก่อนหน้า"
            },
            emptyTable: "ไม่มีข้อมูลในตาราง"
        },
        pageLength: 10,
        responsive: true,
        order: [[0, 'asc']]
    });

    // Summary
    updateSummary(inventoryData);

// Export
document.getElementById('exportBtn').addEventListener('click', function() {
    const exportData = inventoryData.map(item => ({
        'Group': item.group,
        'Weight Surin (ตัน)': item.weightSurin,
        'Weight Nangrong (ตัน)': item.weightNangrong,
        'Weight Det Udom (ตัน)': item.weightDetUdom,
        'Cost Surin (บาท)': item.costSurin,
        'Cost Nangrong (บาท)': item.costNangrong,
        'Cost Det Udom (บาท)': item.costDetUdom
    }));

    // ====== เพิ่ม Sum Row ======
    const totalWeightSurin = inventoryData.reduce((sum, i) => sum + i.weightSurin, 0);
    const totalWeightNangrong = inventoryData.reduce((sum, i) => sum + i.weightNangrong, 0);
    const totalWeightDetUdom = inventoryData.reduce((sum, i) => sum + i.weightDetUdom, 0);
    const totalCostSurin = inventoryData.reduce((sum, i) => sum + i.costSurin, 0);
    const totalCostNangrong = inventoryData.reduce((sum, i) => sum + i.costNangrong, 0);
    const totalCostDetUdom = inventoryData.reduce((sum, i) => sum + i.costDetUdom, 0);

    exportData.push({}); // แทรกแถวว่างก่อนรวม
    exportData.push({
        'Group': 'รวมทั้งหมด',
        'Weight Surin (ตัน)': totalWeightSurin.toFixed(2),
        'Weight Nangrong (ตัน)': totalWeightNangrong.toFixed(2),
        'Weight Det Udom (ตัน)': totalWeightDetUdom.toFixed(2),
        'Cost Surin (บาท)': totalCostSurin,
        'Cost Nangrong (บาท)': totalCostNangrong,
        'Cost Det Udom (บาท)': totalCostDetUdom
    });
    // ============================

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Steel Inventory");

    const today = new Date();
    const dateStr = today.getFullYear() + '-' + 
                  String(today.getMonth() + 1).padStart(2, '0') + '-' + 
                  String(today.getDate()).padStart(2, '0');
    const filename = `CRC_Steel_Inventory_${dateStr}.xlsx`;

    XLSX.writeFile(wb, filename);
});

});
