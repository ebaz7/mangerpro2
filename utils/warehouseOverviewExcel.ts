import * as XLSX from 'xlsx';

export interface WarehouseExcelDataset {
    summary: {
        reportDate?: string;
        report1Label?: string;
        report1Jalali?: string;
        report2Label?: string;
        report2Jalali?: string;
        signature?: string;
        ceoSignature?: string;
        lastYearYarnsWeight?: number;
        currentYarnsWeight?: number;
        yarnsDiffWeight?: number;
        yarnsRatio?: number;
        lastYearRawWeight?: number;
        currentRawWeight?: number;
        rawDiffWeight?: number;
        rawRatio?: number;
        lastYearTotalWeight?: number;
        currentTotalWeight?: number;
        totalDiffWeight?: number;
        totalRatio?: number;
        totalLastYearContainers?: number;
        containersTotal?: number | string;
        diffContainers?: number;
        totalLastYearDollars?: number;
        dollarsTotal?: number | string;
        diffDollars?: number;
    };
    yarnItems: any[];
    rawItems: any[];
    purchasingGoods?: any[];
    domesticPurchases?: any[];
    goodsInCustoms?: any[];
    commercialGoods?: any[];
    logisticsItems?: any[];
    growthItems?: any[];
    negativeItems?: any[];
    allComparedItems?: any[];
    getGroupChildItems?: (groupCode: string) => Array<{ itemName: string; itemCode: string }>;
}

export type WarehouseExcelScope = 'all' | 'summary' | 'yarns' | 'raw' | 'logistics' | 'variance';

/**
 * Formats numbers safely for Excel cells
 */
const n = (val: any): number => {
    if (val === undefined || val === null || val === '') return 0;
    const num = parseFloat(String(val).replace(/,/g, ''));
    return isNaN(num) ? 0 : Math.round(num * 100) / 100;
};

const formatPct = (pct: number): string => {
    if (isNaN(pct) || !isFinite(pct)) return '-';
    const sign = pct > 0 ? '+' : '';
    return `${sign}${pct.toFixed(1)}%`;
};

/**
 * Builds the exact 1:1 Excel replica matching the official print / PDF report layout.
 */
export const exportWarehouseOverviewToExcel = (
    dataset: WarehouseExcelDataset,
    scope: WarehouseExcelScope = 'all'
): void => {
    const {
        summary = {},
        yarnItems = [],
        rawItems = [],
        purchasingGoods = [],
        domesticPurchases = [],
        goodsInCustoms = [],
        commercialGoods = [],
        logisticsItems = []
    } = dataset;

    const reportDate = summary.reportDate || '۱۴۰۵/۰۵/۳۱';
    const r1Label = summary.report1Label || 'منتهی به سال ۱۴۰۴';
    const r1Date = summary.report1Jalali || '۱۴۰۴/۱۲/۲۹';
    const r2Label = summary.report2Label || 'وضعیت فعلی سال ۱۴۰۵';
    const r2Date = summary.report2Jalali || '۱۴۰۵/۰۵/۳۱';
    const manager = summary.signature || 'محمد ابراهیم حیدری';
    const ceo = summary.ceoSignature || 'جناب آقای محمد امین فتوت احمدی';

    // Summary calculations
    const yLast = n(summary.lastYearYarnsWeight) || yarnItems.reduce((s, i) => s + n(i.lastYearWeight), 0);
    const yCurr = n(summary.currentYarnsWeight) || yarnItems.reduce((s, i) => s + n(i.currentWeight), 0);
    const yDiff = yCurr - yLast;
    const yPct = yLast ? (yDiff / yLast) * 100 : 0;

    const rLast = n(summary.lastYearRawWeight) || rawItems.reduce((s, i) => s + n(i.lastYearWeight), 0);
    const rCurr = n(summary.currentRawWeight) || rawItems.reduce((s, i) => s + n(i.currentWeight), 0);
    const rDiff = rCurr - rLast;
    const rPct = rLast ? (rDiff / rLast) * 100 : 0;

    const tLast = n(summary.lastYearTotalWeight) || (yLast + rLast);
    const tCurr = n(summary.currentTotalWeight) || (yCurr + rCurr);
    const tDiff = tCurr - tLast;
    const tPct = tLast ? (tDiff / tLast) * 100 : 0;

    const containersCount = summary.containersTotal || '۰';
    const dollarsAmount = typeof summary.dollarsTotal === 'number' ? summary.dollarsTotal.toLocaleString('en-US') : (summary.dollarsTotal || '۰');

    // Combine logistics items
    const allLogistics = logisticsItems.length > 0 ? logisticsItems : [
        ...goodsInCustoms.map((r, i) => ({
            name: r.cargoType || 'بار گمرکی',
            proforma: r.proforma || '-',
            registrationNumber: r.registrationNumber || '-',
            status: r.statusBadge || 'در گمرک',
            containers: r.container || 0,
            value: r.weight ? `${n(r.weight).toLocaleString('fa-IR')} kg` : `${n(r.dollars).toLocaleString('fa-IR')} $`,
            currency: r.dollars > 0 ? 'USD' : 'IRR'
        })),
        ...purchasingGoods.map((r, i) => ({
            name: r.cargoType || 'بار در راه / خرید خارجی',
            proforma: r.proforma || '-',
            registrationNumber: r.registrationNumber || '-',
            status: r.statusBadge || 'در حال خرید / در راه',
            containers: r.container || 0,
            value: r.weight ? `${n(r.weight).toLocaleString('fa-IR')} kg` : `${n(r.dollars).toLocaleString('fa-IR')} $`,
            currency: r.dollars > 0 ? 'USD' : 'IRR'
        })),
        ...domesticPurchases.map((r, i) => ({
            name: `${r.cargoType || 'خرید داخلی'}${r.petrochemicalName ? ` (${r.petrochemicalName})` : ''}`,
            proforma: r.proforma || '-',
            registrationNumber: r.registrationNumber || '-',
            status: r.statusBadge || 'خرید پتروشیمی',
            containers: r.container || 0,
            value: r.weight ? `${n(r.weight).toLocaleString('fa-IR')} kg` : `${n(r.rialAmount).toLocaleString('fa-IR')} ریال`,
            currency: 'IRR'
        })),
        ...commercialGoods.map((r, i) => ({
            name: r.itemName || 'کالای تجاری',
            proforma: r.proforma || '-',
            registrationNumber: r.registrationNumber || '-',
            status: 'انبار تجاری',
            containers: r.container || 0,
            value: r.weight ? `${n(r.weight).toLocaleString('fa-IR')} kg` : `${n(r.dollars).toLocaleString('fa-IR')} $`,
            currency: r.dollars > 0 ? 'USD' : 'IRR'
        }))
    ];

    // 1. Create a new Workbook
    const wb = XLSX.utils.book_new();

    // Helper to apply RTL and column widths
    const finalizeSheet = (ws: XLSX.WorkSheet, colWidths: number[], merges?: XLSX.Range[]) => {
        ws['!views'] = [{ rightToLeft: true, RTL: true }];
        ws['!cols'] = colWidths.map(w => ({ wch: w }));
        if (merges && merges.length > 0) {
            ws['!merges'] = merges;
        }
    };

    // ==========================================
    // PRIMARY MASTER SHEET: Exact 1:1 Report Replica
    // ==========================================
    const masterRows: any[][] = [];
    const masterMerges: XLSX.Range[] = [];

    // Helper to push row and optional merge
    const addRow = (row: any[]) => {
        masterRows.push(row);
        return masterRows.length - 1;
    };

    const addMergedRow = (row: any[], startCol = 0, endCol = 5) => {
        const rIdx = addRow(row);
        masterMerges.push({ s: { r: rIdx, c: startCol }, e: { r: rIdx, c: endCol } });
        return rIdx;
    };

    // --- Header Section ---
    addMergedRow(['گزارش مدیریتی و تراز وزنی وضعیت انبارها و زنجیره تامین']);
    addMergedRow([`دوره مبنا: ${r1Label} (${r1Date}) | دوره جاری: ${r2Label} (${r2Date}) | تاریخ استعلام: ${reportDate}`]);
    addMergedRow(['سامانه یکپارچه مانیتورینگ کارخانجات، واردات و انبارهای سایان ERP']);
    addMergedRow(['نوع سند: گزارش جامع تراز و استانداردی (صفحه ۱)']);
    addRow([]); // Blank line

    // --- Audience Box ---
    addMergedRow([`مدیریت محترم لپان بافت جناب آقای محمد امین فتوت احمدی`]);
    addMergedRow([`با سلام، احتراماً گزارش موجودی منتهی به سال ۱۴۰۴ (مورخ ${r1Date}) و مقایسه آن با وضعیت فعلی سال ۱۴۰۵ (مورخ ${r2Date}) مستخرج از سامانه یکپارچه سایان به همراه جزئیات بارهای در راه و گمرک به شرح ذیل تقدیم حضور می‌گردد:`]);
    addRow([]); // Blank line

    // --- KPI Cards (4 summary blocks) ---
    addMergedRow(['📊 شاخص‌های کلان تراز وزنی و ارزش زنجیره تامین']);
    addRow([
        '🧵 نخ‌های تولیدی کارخانه',
        `kg ${yCurr.toLocaleString('fa-IR')}`,
        `تغییر: ${yDiff >= 0 ? '+' : ''}${yDiff.toLocaleString('fa-IR')} kg (${formatPct(yPct)})`,
        '📦 مواد اولیه و وارداتی',
        `kg ${rCurr.toLocaleString('fa-IR')}`,
        `تغییر: ${rDiff >= 0 ? '+' : ''}${rDiff.toLocaleString('fa-IR')} kg (${formatPct(rPct)})`
    ]);
    addRow([
        '👑 سرجمع کل زنجیره تامین',
        `kg ${tCurr.toLocaleString('fa-IR')}`,
        `تراز رشد: ${tDiff >= 0 ? '+' : ''}${tDiff.toLocaleString('fa-IR')} kg (${formatPct(tPct)})`,
        '🚢 کانتینرها و ارزش دلاری',
        `${containersCount} کانتینر`,
        `${dollarsAmount} $ ارزش در راه`
    ]);
    addRow([]); // Blank line

    // --- TABLE 1: نخ‌های تولیدی کارخانه (تولید داخلی) ---
    addMergedRow(['۱. جدول نخ‌های تولیدی کارخانه (تولید داخلی)']);
    addRow([
        '#',
        'نام گروه کالا / شرح تولید',
        'وزن سال قبل (kg)',
        'وزن سال جاری (kg)',
        'اختلاف وزنی (kg)',
        'درصد تغییر'
    ]);

    yarnItems.forEach((item, idx) => {
        const last = n(item.lastYearWeight);
        const curr = n(item.currentWeight);
        const diff = curr - last;
        const pct = last ? (diff / last) * 100 : 0;
        const displayName = `${item.name || item.groupName || '-'}${item.code ? ` (${item.code})` : ''}`;

        addRow([
            idx + 1,
            displayName,
            last,
            curr,
            diff,
            formatPct(pct)
        ]);
    });

    // Total for Table 1
    addRow([
        '∑',
        'جمع کل نخ‌های تولیدی کارخانه',
        yLast,
        yCurr,
        yDiff,
        formatPct(yPct)
    ]);
    addRow([]); // Blank line

    // --- TABLE 2: مواد اولیه، اقلام وارداتی و انبار سایان ---
    addMergedRow(['۲. جدول مواد اولیه، اقلام وارداتی و انبار سایان']);
    addRow([
        '#',
        'نام گروه کالا / مواد اولیه وارداتی',
        'وزن سال قبل (kg)',
        'وزن سال جاری (kg)',
        'اختلاف وزنی (kg)',
        'درصد تغییر'
    ]);

    rawItems.forEach((item, idx) => {
        const last = n(item.lastYearWeight);
        const curr = n(item.currentWeight);
        const diff = curr - last;
        const pct = last ? (diff / last) * 100 : 0;
        const displayName = `${item.name || item.groupName || '-'}${item.code ? ` (${item.code})` : ''}`;

        addRow([
            idx + 1,
            displayName,
            last,
            curr,
            diff,
            formatPct(pct)
        ]);
    });

    // Total for Table 2
    addRow([
        '∑',
        'جمع کل مواد اولیه و وارداتی',
        rLast,
        rCurr,
        rDiff,
        formatPct(rPct)
    ]);
    addRow([]); // Blank line

    // --- TABLE 3: بارهای در راه، گمرک و خریدهای در جریان ---
    if (allLogistics.length > 0) {
        addMergedRow(['۳. بارهای در راه، گمرک و خریدهای در جریان (کانتینری و ارزی)'], 0, 7);
        addRow([
            '#',
            'شرح محموله / کالا',
            'پروفرما',
            'ثبت سفارش',
            'وضعیت سند',
            'تعداد کانتینر',
            'مقدار / ارزش',
            'نوع ارز'
        ]);

        allLogistics.forEach((item, idx) => {
            addRow([
                idx + 1,
                item.name || '-',
                item.proforma || '-',
                item.registrationNumber || '-',
                item.status || 'در جریان',
                item.containers ? `${item.containers} کانتینر` : '-',
                item.value || item.currentValue || '-',
                item.currency || 'USD'
            ]);
        });
        addRow([]); // Blank line
    }

    // --- GRAND TOTAL BANNER: سرجمع کل موجودی زنجیره تامین و کارخانجات ---
    addMergedRow(['★ سرجمع کل موجودی زنجیره تامین و کارخانجات']);
    addRow([
        'عنوان تراز کلی',
        'موجودی کل سال قبل (kg)',
        'موجودی کل سال جاری (kg)',
        'اختلاف خالص وزنی (kg)',
        'درصد رشد تراز کل',
        'وضعیت کل زنجیره'
    ]);
    addRow([
        'مجموع کل زنجیره تامین و کارخانجات',
        tLast,
        tCurr,
        tDiff,
        formatPct(tPct),
        tDiff >= 0 ? 'رشد تراز وزنی ✅' : 'کسری موجودی ⚠️'
    ]);
    addRow([]); // Blank line

    // --- Signatures Section ---
    addMergedRow(['✍️ امضا و تایید رسمی']);
    addRow([
        'تنظیم‌کننده و مدیر بازرگانی:',
        manager,
        '',
        'تاییدکننده / مدیریت محترم عامل:',
        ceo,
        ''
    ]);

    const wsMaster = XLSX.utils.aoa_to_sheet(masterRows);
    finalizeSheet(wsMaster, [6, 38, 22, 22, 22, 20, 18, 14], masterMerges);
    XLSX.utils.book_append_sheet(wb, wsMaster, 'گزارش تراز و وضعیت انبارها');

    // ----------------------------------------------------
    // Trigger Clean Download in Browser
    // ----------------------------------------------------
    const safeDate = reportDate.replace(/[\/\\]/g, '-');
    const fileName = `گزارش_رسمی_موجودی_و_تراز_انبار_${safeDate}.xlsx`;

    XLSX.writeFile(wb, fileName, {
        bookType: 'xlsx',
        type: 'binary'
    });
};
