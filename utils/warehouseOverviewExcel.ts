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
        containersTotal?: number;
        diffContainers?: number;
        totalLastYearDollars?: number;
        dollarsTotal?: number;
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
    if (isNaN(pct) || !isFinite(pct)) return '۰٪';
    const sign = pct > 0 ? '+' : '';
    return `${sign}${pct.toFixed(1)}%`;
};

/**
 * Builds the enterprise-grade multi-sheet Excel file for Sayan Warehouse Balance
 * Matching the exact visual layout of the executive management report on screen.
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
        growthItems = [],
        negativeItems = [],
        allComparedItems = [],
        getGroupChildItems
    } = dataset;

    const reportDate = summary.reportDate || '۱۴۰۵/۰۵/۳۱';
    const r1Label = summary.report1Label || 'منتهی به سال ۱۴۰۴';
    const r1Date = summary.report1Jalali || '۱۴۰۴/۱۲/۲۹';
    const r2Label = summary.report2Label || 'وضعیت فعلی سال ۱۴۰۵';
    const r2Date = summary.report2Jalali || '۱۴۰۵/۰۵/۳۱';
    const manager = summary.signature || 'محمد ابراهیم حیدری';
    const ceo = summary.ceoSignature || 'مدیریت محترم لپان بافت جناب آقای محمد امین فتوت احمدی';

    // 1. Create a new Workbook
    const wb = XLSX.utils.book_new();

    // Helper to apply RTL and column widths to a worksheet
    const finalizeSheet = (ws: XLSX.WorkSheet, colWidths: number[], merges?: XLSX.Range[]) => {
        ws['!views'] = [{ rightToLeft: true, RTL: true }];
        ws['!cols'] = colWidths.map(w => ({ wch: w }));
        if (merges && merges.length > 0) {
            ws['!merges'] = merges;
        }
    };

    // =========================================================================
    // PRIMARY MASTER SHEET: گزارش جامع مدیریتی مقایسه‌ای وضعیت انبارها (Exact visual layout)
    // =========================================================================
    if (scope === 'all' || scope === 'summary') {
        const masterRows: any[][] = [];
        const merges: XLSX.Range[] = [];

        // Row 0: Title Header
        masterRows.push(['گزارش مدیریتی مقایسه‌ای وضعیت انبارها']);
        merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 14 } });

        // Row 1: Subtitle / Audience
        masterRows.push([`مخاطب: ${ceo}`]);
        merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 14 } });

        // Row 2: Metadata row
        masterRows.push([
            'تاریخ گزارش:', reportDate,
            '', 'دوره اول:', `${r1Label} (${r1Date})`,
            '', 'دوره دوم:', `${r2Label} (${r2Date})`,
            '', 'تنظیم‌کننده:', manager,
            '', 'سیستم انبار:', 'سامانه یکپارچه سایان (Live ERP)'
        ]);

        // Row 3: Letter introduction
        masterRows.push([
            `با سلام، احتراماً گزارش موجودی ${r1Label} (مورخ ${r1Date}) و مقایسه آن با ${r2Label} (مورخ ${r2Date}) مستخرج از سامانه یکپارچه سایان به همراه جزئیات بارهای در راه و گمرک به شرح ذیل تقدیم حضور می‌گردد:`
        ]);
        merges.push({ s: { r: 3, c: 0 }, e: { r: 3, c: 14 } });

        // Row 4: Empty space
        masterRows.push(['']);

        // Row 5: Table Top Header (Merged categories)
        const headerTopRowIdx = masterRows.length;
        masterRows.push([
            'ردیف',
            'کد کالا / نخ',
            'نوع کالا / نخ',
            'پروفرم',
            `موجودی انبارها (${r1Label} - مورخ ${r1Date})`, '', '', '',
            `موجودی انبارها (${r2Label} - مورخ ${r2Date})`, '', '', '',
            'تحلیل مغایرت و موازنه', '', ''
        ]);
        merges.push({ s: { r: headerTopRowIdx, c: 0 }, e: { r: headerTopRowIdx + 1, c: 0 } });
        merges.push({ s: { r: headerTopRowIdx, c: 1 }, e: { r: headerTopRowIdx + 1, c: 1 } });
        merges.push({ s: { r: headerTopRowIdx, c: 2 }, e: { r: headerTopRowIdx + 1, c: 2 } });
        merges.push({ s: { r: headerTopRowIdx, c: 3 }, e: { r: headerTopRowIdx + 1, c: 3 } });
        merges.push({ s: { r: headerTopRowIdx, c: 4 }, e: { r: headerTopRowIdx, c: 7 } });
        merges.push({ s: { r: headerTopRowIdx, c: 8 }, e: { r: headerTopRowIdx, c: 11 } });
        merges.push({ s: { r: headerTopRowIdx, c: 12 }, e: { r: headerTopRowIdx, c: 14 } });

        // Row 6: Sub Header Columns
        masterRows.push([
            'ردیف',
            'کد کالا',
            'نوع کالا / نخ',
            'پروفرم',
            'کارتن',
            'وزن (KG)',
            'کانتینر',
            'ارزش دلاری ($)',
            'کارتن',
            'وزن (KG)',
            'کانتینر',
            'ارزش دلاری ($)',
            'اختلاف وزن (KG)',
            'درصد تغییرات (%)',
            'وضعیت'
        ]);

        let rowCounter = 1;

        // --- SECTION 1: کالاهای تولیدی کارخانه (نخ‌ها) ---
        const s1BannerIdx = masterRows.length;
        masterRows.push(['🧵 ۱. کالاهای تولیدی کارخانه (نخ‌های بافته شده و انبار محصول - گروه ۰۴)']);
        merges.push({ s: { r: s1BannerIdx, c: 0 }, e: { r: s1BannerIdx, c: 14 } });

        let ySumLastW = 0, ySumCurrW = 0, ySumLastC = 0, ySumCurrC = 0, ySumLastCont = 0, ySumCurrCont = 0, ySumLastD = 0, ySumCurrD = 0;
        yarnItems.forEach((item) => {
            const lW = n(item.lastYearWeight);
            const cW = n(item.currentWeight);
            const lC = n(item.lastYearCartons);
            const cC = n(item.currentCartons);
            const lCont = n(item.lastYearContainers);
            const cCont = n(item.currentContainers);
            const lD = n(item.lastYearDollars);
            const cD = n(item.currentDollars);
            const diffW = cW - lW;
            const pct = lW > 0 ? (diffW / lW) * 100 : (cW > 0 ? 100 : 0);

            ySumLastW += lW; ySumCurrW += cW;
            ySumLastC += lC; ySumCurrC += cC;
            ySumLastCont += lCont; ySumCurrCont += cCont;
            ySumLastD += lD; ySumCurrD += cD;

            masterRows.push([
                rowCounter++,
                item.code || '-',
                item.name || '-',
                item.proforma || '-',
                lC || '-',
                lW,
                lCont || '-',
                lD || '-',
                cC || '-',
                cW,
                cCont || '-',
                cD || '-',
                diffW,
                formatPct(pct),
                diffW >= 0 ? 'افزایش ذخیره' : 'کاهش / مصرف'
            ]);

            if (getGroupChildItems && item.code) {
                const childItems = getGroupChildItems(item.code);
                if (childItems && childItems.length > 1) {
                    childItems.forEach(ch => {
                        masterRows.push([
                            '',
                            `  ↳ ${ch.itemCode}`,
                            `     └ ${ch.itemName}`,
                            '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', 'زیرمجموعه'
                        ]);
                    });
                }
            }
        });

        const yDiffTotal = ySumCurrW - ySumLastW;
        const yPctTotal = ySumLastW > 0 ? (yDiffTotal / ySumLastW) * 100 : 0;
        masterRows.push([
            '★',
            'جمع گروه ۰۴',
            'سرجمع کالاهای تولیدی کارخانه (نخ‌ها)',
            '-',
            ySumLastC,
            ySumLastW,
            ySumLastCont,
            ySumLastD,
            ySumCurrC,
            ySumCurrW,
            ySumCurrCont,
            ySumCurrD,
            yDiffTotal,
            formatPct(yPctTotal),
            yDiffTotal >= 0 ? 'رشد موجودی' : 'افت موجودی'
        ]);

        // --- SECTION 2: مواد اولیه، الیاف و کارخانه (گروه ۰۱) ---
        const s2BannerIdx = masterRows.length;
        masterRows.push(['🏭 ۲. مواد اولیه، الیاف و اقلام انبار کارخانه (چیپس، لاکرا، پلی‌استر و...)']);
        merges.push({ s: { r: s2BannerIdx, c: 0 }, e: { r: s2BannerIdx, c: 14 } });

        let rSumLastW = 0, rSumCurrW = 0, rSumLastC = 0, rSumCurrC = 0, rSumLastCont = 0, rSumCurrCont = 0, rSumLastD = 0, rSumCurrD = 0;
        rawItems.forEach((item) => {
            const lW = n(item.lastYearWeight);
            const cW = n(item.currentWeight);
            const lC = n(item.lastYearCartons);
            const cC = n(item.currentCartons);
            const lCont = n(item.lastYearContainers);
            const cCont = n(item.currentContainers);
            const lD = n(item.lastYearDollars);
            const cD = n(item.currentDollars);
            const diffW = cW - lW;
            const pct = lW > 0 ? (diffW / lW) * 100 : (cW > 0 ? 100 : 0);

            rSumLastW += lW; rSumCurrW += cW;
            rSumLastC += lC; rSumCurrC += cC;
            rSumLastCont += lCont; rSumCurrCont += cCont;
            rSumLastD += lD; rSumCurrD += cD;

            masterRows.push([
                rowCounter++,
                item.code || '-',
                item.name || '-',
                item.proforma || '-',
                lC || '-',
                lW,
                lCont || '-',
                lD || '-',
                cC || '-',
                cW,
                cCont || '-',
                cD || '-',
                diffW,
                formatPct(pct),
                diffW >= 0 ? 'افزایش ذخیره' : 'مصرف از انبار'
            ]);
        });

        const rDiffTotal = rSumCurrW - rSumLastW;
        const rPctTotal = rSumLastW > 0 ? (rDiffTotal / rSumLastW) * 100 : 0;
        masterRows.push([
            '★',
            'جمع گروه ۰۱',
            'سرجمع مواد اولیه و کارخانه',
            '-',
            rSumLastC,
            rSumLastW,
            rSumLastCont,
            rSumLastD,
            rSumCurrC,
            rSumCurrW,
            rSumCurrCont,
            rSumCurrD,
            rDiffTotal,
            formatPct(rPctTotal),
            rDiffTotal >= 0 ? 'رشد ذخیره' : 'مصرف از انبار'
        ]);

        // --- SECTION 3: بارهای در حال خرید خارجی و در راه ---
        if (purchasingGoods && purchasingGoods.length > 0) {
            const s3BannerIdx = masterRows.length;
            masterRows.push(['🚢🛒 ۳. بارهای در حال خرید خارجی و ترانزیت در راه']);
            merges.push({ s: { r: s3BannerIdx, c: 0 }, e: { r: s3BannerIdx, c: 14 } });

            let pSumW = 0, pSumC = 0, pSumCont = 0, pSumD = 0;
            purchasingGoods.forEach((item) => {
                const w = n(item.weight);
                const c = n(item.cartons);
                const cont = n(item.container);
                const d = n(item.dollars);
                pSumW += w; pSumC += c; pSumCont += cont; pSumD += d;

                masterRows.push([
                    rowCounter++,
                    item.registrationNumber ? `ثبت: ${item.registrationNumber}` : 'خرید خارجی',
                    item.cargoType || 'محموله در راه',
                    item.proforma || '-',
                    '-', 0, '-', '-',
                    c || '-',
                    w,
                    cont || '-',
                    d || '-',
                    w,
                    '+۱۰۰٪',
                    item.statusBadge || 'در راه ترانزیت'
                ]);
            });

            masterRows.push([
                '★',
                'جمع بارهای در راه',
                'سرجمع بارهای در حال خرید و ترانزیت',
                '-',
                '-', 0, '-', '-',
                pSumC,
                pSumW,
                pSumCont,
                pSumD,
                pSumW,
                '+۱۰۰٪',
                'ورودی‌های فعال'
            ]);
        }

        // --- SECTION 4: بارهای موجود در اماکن گمرکی ---
        if (goodsInCustoms && goodsInCustoms.length > 0) {
            const s4BannerIdx = masterRows.length;
            masterRows.push(['🏢 ۴. بارهای موجود در اماکن گمرکی (در حال ترخیص)']);
            merges.push({ s: { r: s4BannerIdx, c: 0 }, e: { r: s4BannerIdx, c: 14 } });

            let cSumW = 0, cSumC = 0, cSumCont = 0, cSumD = 0;
            goodsInCustoms.forEach((item) => {
                const w = n(item.weight);
                const c = n(item.cartons);
                const cont = n(item.container);
                const d = n(item.dollars);
                cSumW += w; cSumC += c; cSumCont += cont; cSumD += d;

                masterRows.push([
                    rowCounter++,
                    item.registrationNumber ? `ثبت: ${item.registrationNumber}` : 'گمرک',
                    item.cargoType || 'محموله گمرکی',
                    item.proforma || '-',
                    '-', 0, '-', '-',
                    c || '-',
                    w,
                    cont || '-',
                    d || '-',
                    w,
                    '+۱۰۰٪',
                    item.statusBadge || 'در حال ترخیص'
                ]);
            });

            masterRows.push([
                '★',
                'جمع گمرک',
                'سرجمع بارهای موجود در اماکن گمرکی',
                '-',
                '-', 0, '-', '-',
                cSumC,
                cSumW,
                cSumCont,
                cSumD,
                cSumW,
                '+۱۰۰٪',
                'آماده ترخیص'
            ]);
        }

        // --- SECTION 5: کالاهای تجاری و متفرقه ---
        if (commercialGoods && commercialGoods.length > 0) {
            const s5BannerIdx = masterRows.length;
            masterRows.push(['🏬 ۵. کالاهای تجاری و متفرقه']);
            merges.push({ s: { r: s5BannerIdx, c: 0 }, e: { r: s5BannerIdx, c: 14 } });

            let comSumW = 0, comSumC = 0, comSumCont = 0, comSumD = 0;
            commercialGoods.forEach((item) => {
                const w = n(item.weight);
                const c = n(item.cartons);
                const cont = n(item.container);
                const d = n(item.dollars);
                comSumW += w; comSumC += c; comSumCont += cont; comSumD += d;

                masterRows.push([
                    rowCounter++,
                    item.category || 'کالای تجاری',
                    item.itemName || 'کالای تجاری',
                    item.proforma || '-',
                    '-', 0, '-', '-',
                    c || '-',
                    w,
                    cont || '-',
                    d || '-',
                    w,
                    '+۱۰۰٪',
                    'انبار تجاری'
                ]);
            });

            masterRows.push([
                '★',
                'جمع تجاری',
                'سرجمع کالاهای تجاری و متفرقه',
                '-',
                '-', 0, '-', '-',
                comSumC,
                comSumW,
                comSumCont,
                comSumD,
                comSumW,
                '+۱۰۰٪',
                'موجودی تجاری'
            ]);
        }

        // --- SECTION 6: خریدهای داخلی پتروشیمی ---
        if (domesticPurchases && domesticPurchases.length > 0) {
            const s6BannerIdx = masterRows.length;
            masterRows.push(['📦 ۶. خریدهای داخلی پتروشیمی و بورس کالا']);
            merges.push({ s: { r: s6BannerIdx, c: 0 }, e: { r: s6BannerIdx, c: 14 } });

            let dSumW = 0, dSumC = 0, dSumCont = 0, dSumD = 0;
            domesticPurchases.forEach((item) => {
                const w = n(item.weight);
                const c = n(item.cartons);
                const cont = n(item.container);
                const d = n(item.dollars);
                dSumW += w; dSumC += c; dSumCont += cont; dSumD += d;

                masterRows.push([
                    rowCounter++,
                    item.petrochemicalName || 'بورس کالا',
                    item.cargoType || 'پتروشیمی',
                    item.proforma || '-',
                    '-', 0, '-', '-',
                    c || '-',
                    w,
                    cont || '-',
                    d || '-',
                    w,
                    '+۱۰۰٪',
                    item.statusBadge || 'خرید قطعی'
                ]);
            });

            masterRows.push([
                '★',
                'جمع پتروشیمی',
                'سرجمع خریدهای داخلی و بورس',
                '-',
                '-', 0, '-', '-',
                dSumC,
                dSumW,
                dSumCont,
                dSumD,
                dSumW,
                '+۱۰۰٪',
                'خرید داخلی'
            ]);
        }

        // --- SECTION 7: GRAND TOTAL ROW (سرجمع کل زنجیره تامین) ---
        const transitWeight = (purchasingGoods || []).reduce((s, x) => s + n(x.weight), 0);
        const customsWeight = (goodsInCustoms || []).reduce((s, x) => s + n(x.weight), 0);
        const commercialWeight = (commercialGoods || []).reduce((s, x) => s + n(x.weight), 0);
        const domesticWeight = (domesticPurchases || []).reduce((s, x) => s + n(x.weight), 0);

        const grandLastW = n(summary.lastYearTotalWeight) || (ySumLastW + rSumLastW);
        const grandCurrW = n(summary.currentTotalWeight) || (ySumCurrW + rSumCurrW + transitWeight + customsWeight + commercialWeight + domesticWeight);
        const grandDiffW = grandCurrW - grandLastW;
        const grandPct = grandLastW > 0 ? (grandDiffW / grandLastW) * 100 : 0;

        const grandLastC = ySumLastC + rSumLastC;
        const grandCurrC = ySumCurrC + rSumCurrC + (purchasingGoods || []).reduce((s, x) => s + n(x.cartons), 0) + (goodsInCustoms || []).reduce((s, x) => s + n(x.cartons), 0) + (commercialGoods || []).reduce((s, x) => s + n(x.cartons), 0);

        const grandLastCont = n(summary.totalLastYearContainers) || (ySumLastCont + rSumLastCont);
        const grandCurrCont = n(summary.containersTotal) || (ySumCurrCont + rSumCurrCont + (purchasingGoods || []).reduce((s, x) => s + n(x.container), 0) + (goodsInCustoms || []).reduce((s, x) => s + n(x.container), 0) + (commercialGoods || []).reduce((s, x) => s + n(x.container), 0));

        const grandLastD = n(summary.totalLastYearDollars) || (ySumLastD + rSumLastD);
        const grandCurrD = n(summary.dollarsTotal) || (ySumCurrD + rSumCurrD + (purchasingGoods || []).reduce((s, x) => s + n(x.dollars), 0) + (goodsInCustoms || []).reduce((s, x) => s + n(x.dollars), 0) + (commercialGoods || []).reduce((s, x) => s + n(x.dollars), 0));

        masterRows.push(['']);
        masterRows.push([
            '👑',
            'سرجمع کل',
            'سرجمع کل موازنه وزنی، کانتینری و ارزی زنجیره تامین انبار',
            '-',
            grandLastC,
            grandLastW,
            grandLastCont,
            grandLastD,
            grandCurrC,
            grandCurrW,
            grandCurrCont,
            grandCurrD,
            grandDiffW,
            formatPct(grandPct),
            grandDiffW >= 0 ? 'تراز مثبت و پایدار' : 'تراز نزولی'
        ]);

        // --- KPI SUMMARY TABLES BELOW MASTER TABLE ---
        masterRows.push(['']);
        masterRows.push(['جدول خلاصه مدیریتی و موازنه شاخص‌های کلان']);
        merges.push({ s: { r: masterRows.length - 1, c: 0 }, e: { r: masterRows.length - 1, c: 6 } });

        masterRows.push([
            'ردیف',
            'سرفصل زنجیره تامین',
            `وزن ${r1Label} (KG)`,
            `وزن ${r2Label} (KG)`,
            'اختلاف وزنی (KG)',
            'درصد تغییر (%)',
            'تحلیل وضعیت'
        ]);

        masterRows.push([
            1, 'کالاهای تولیدی کارخانه (نخ‌ها)', ySumLastW, ySumCurrW, yDiffTotal, formatPct(yPctTotal),
            yDiffTotal >= 0 ? 'افزایش موجودی' : 'کاهش / مصرف'
        ]);
        masterRows.push([
            2, 'مواد اولیه و کارخانه (چیپس، لاکرا و...)', rSumLastW, rSumCurrW, rDiffTotal, formatPct(rPctTotal),
            rDiffTotal >= 0 ? 'افزایش ذخیره' : 'مصرف از انبار'
        ]);
        masterRows.push([
            3, 'بارهای در حال خرید خارجی و در راه', 0, transitWeight, transitWeight, '+۱۰۰٪', 'محموله فعال ترانزیت'
        ]);
        masterRows.push([
            4, 'بارهای موجود در اماکن گمرکی', 0, customsWeight, customsWeight, customsWeight > 0 ? '+۱۰۰٪' : '۰٪', 'در انتظار ترخیص'
        ]);
        masterRows.push([
            5, 'خریدهای داخلی پتروشیمی و بورس', 0, domesticWeight, domesticWeight, domesticWeight > 0 ? '+۱۰۰٪' : '۰٪', 'سفارشات قطعی'
        ]);
        masterRows.push([
            6, 'کالاهای تجاری و متفرقه', 0, commercialWeight, commercialWeight, commercialWeight > 0 ? '+۱۰۰٪' : '۰٪', 'موجودی تجاری'
        ]);
        masterRows.push([
            '★', 'سرجمع کل زنجیره تامین و انبارها', grandLastW, grandCurrW, grandDiffW, formatPct(grandPct),
            grandDiffW >= 0 ? 'تراز مثبت و پایدار' : 'تراز نزولی'
        ]);

        // Signatures Block
        masterRows.push(['']);
        masterRows.push(['']);
        const sigRowIdx = masterRows.length;
        masterRows.push([
            '', 'امضای تنظیم‌کننده / مدیر بازرگانی:', manager, '', '', '', '',
            '', 'امضای تاییدکننده / مدیریت عامل:', ceo, '', '', '', '', ''
        ]);
        merges.push({ s: { r: sigRowIdx, c: 1 }, e: { r: sigRowIdx, c: 3 } });
        merges.push({ s: { r: sigRowIdx, c: 8 }, e: { r: sigRowIdx, c: 11 } });

        const wsMaster = XLSX.utils.aoa_to_sheet(masterRows);
        finalizeSheet(wsMaster, [6, 16, 34, 16, 14, 20, 14, 16, 14, 20, 14, 16, 20, 16, 22], merges);
        XLSX.utils.book_append_sheet(wb, wsMaster, 'گزارش جامع وضعیت انبارها');
    }

    // =========================================================================
    // SHEET 2: کالاهای تولیدی (نخ‌ها) - تفکیک گروه 04
    // =========================================================================
    if (scope === 'all' || scope === 'yarns') {
        const yarnRows: any[][] = [
            ['کالاهای تولیدی کارخانه (نخ‌ها و محصولات نهایی) - مقایسه تفصیلی سالانه'],
            [`تاریخ گزارش: ${reportDate} | تنظیم‌کننده: ${manager} | تایید: ${ceo}`],
            [''],
            [
                'ردیف',
                'کد گروه / کالا',
                'نوع کالا / نخ تولیدی',
                'پروفرما',
                `کارتن ${r1Label}`,
                `وزن ${r1Label} (KG)`,
                `کانتینر ${r1Label}`,
                `دلار ${r1Label} ($)`,
                `کارتن ${r2Label}`,
                `وزن ${r2Label} (KG)`,
                `کانتینر ${r2Label}`,
                `دلار ${r2Label} ($)`,
                'تغییرات وزنی (KG)',
                'درصد تغییر (%)',
                'وضعیت روند'
            ]
        ];

        let ySumLastWeight = 0;
        let ySumCurrWeight = 0;
        let ySumLastCartons = 0;
        let ySumCurrCartons = 0;
        let ySumLastContainers = 0;
        let ySumCurrContainers = 0;
        let ySumLastDollars = 0;
        let ySumCurrDollars = 0;

        yarnItems.forEach((item, idx) => {
            const lastW = n(item.lastYearWeight);
            const currW = n(item.currentWeight);
            const diffW = currW - lastW;
            const pct = lastW > 0 ? (diffW / lastW) * 100 : (currW > 0 ? 100 : 0);

            const lastC = n(item.lastYearCartons);
            const currC = n(item.currentCartons);
            const lastCont = n(item.lastYearContainers);
            const currCont = n(item.currentContainers);
            const lastD = n(item.lastYearDollars);
            const currD = n(item.currentDollars);

            ySumLastWeight += lastW;
            ySumCurrWeight += currW;
            ySumLastCartons += lastC;
            ySumCurrCartons += currC;
            ySumLastContainers += lastCont;
            ySumCurrContainers += currCont;
            ySumLastDollars += lastD;
            ySumCurrDollars += currD;

            yarnRows.push([
                idx + 1,
                item.code || '-',
                item.name || '-',
                item.proforma || '-',
                lastC || '-',
                lastW,
                lastCont || '-',
                lastD || '-',
                currC || '-',
                currW,
                currCont || '-',
                currD || '-',
                diffW,
                formatPct(pct),
                diffW >= 0 ? 'افزایش ذخیره' : 'کاهش / مصرف'
            ]);

            if (getGroupChildItems && item.code) {
                const childItems = getGroupChildItems(item.code);
                if (childItems && childItems.length > 1) {
                    childItems.forEach(ch => {
                        yarnRows.push([
                            '',
                            `  ↳ ${ch.itemCode}`,
                            `     └ ${ch.itemName}`,
                            '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', 'زیرمجموعه'
                        ]);
                    });
                }
            }
        });

        const yTotalDiff = ySumCurrWeight - ySumLastWeight;
        const yTotalPct = ySumLastWeight > 0 ? (yTotalDiff / ySumLastWeight) * 100 : 0;

        yarnRows.push([
            '★',
            'جمع کل',
            'سرجمع کالاهای تولیدی کارخانه (گروه ۰۴)',
            '-',
            ySumLastCartons,
            ySumLastWeight,
            ySumLastContainers,
            ySumLastDollars,
            ySumCurrCartons,
            ySumCurrWeight,
            ySumCurrContainers,
            ySumCurrDollars,
            yTotalDiff,
            formatPct(yTotalPct),
            yTotalDiff >= 0 ? 'رشد کلی' : 'افت کلی'
        ]);

        const wsYarns = XLSX.utils.aoa_to_sheet(yarnRows);
        finalizeSheet(wsYarns, [6, 16, 34, 14, 14, 20, 14, 16, 14, 20, 14, 16, 20, 16, 18]);
        XLSX.utils.book_append_sheet(wb, wsYarns, 'کالاهای تولیدی (نخ‌ها)');
    }

    // =========================================================================
    // SHEET 3: مواد اولیه و کارخانه - گروه 01
    // =========================================================================
    if (scope === 'all' || scope === 'raw') {
        const rawRows: any[][] = [
            ['مواد اولیه، چیپس و اقلام مصرفی انبار کارخانه - مقایسه تفصیلی'],
            [`تاریخ گزارش: ${reportDate} | تنظیم‌کننده: ${manager} | تایید: ${ceo}`],
            [''],
            [
                'ردیف',
                'کد ماده اولیه',
                'شرح ماده اولیه / کالا',
                'پروفرما',
                `کارتن / بسته ${r1Label}`,
                `وزن ${r1Label} (KG)`,
                `کانتینر ${r1Label}`,
                `دلار ${r1Label} ($)`,
                `کارتن / بسته ${r2Label}`,
                `وزن ${r2Label} (KG)`,
                `کانتینر ${r2Label}`,
                `دلار ${r2Label} ($)`,
                'تغییرات وزنی (KG)',
                'درصد تغییر (%)',
                'وضعیت روند'
            ]
        ];

        let rSumLastWeight = 0;
        let rSumCurrWeight = 0;
        let rSumLastCartons = 0;
        let rSumCurrCartons = 0;
        let rSumLastContainers = 0;
        let rSumCurrContainers = 0;
        let rSumLastDollars = 0;
        let rSumCurrDollars = 0;

        rawItems.forEach((item, idx) => {
            const lastW = n(item.lastYearWeight);
            const currW = n(item.currentWeight);
            const diffW = currW - lastW;
            const pct = lastW > 0 ? (diffW / lastW) * 100 : (currW > 0 ? 100 : 0);

            const lastC = n(item.lastYearCartons);
            const currC = n(item.currentCartons);
            const lastCont = n(item.lastYearContainers);
            const currCont = n(item.currentContainers);
            const lastD = n(item.lastYearDollars);
            const currD = n(item.currentDollars);

            rSumLastWeight += lastW;
            rSumCurrWeight += currW;
            rSumLastCartons += lastC;
            rSumCurrCartons += currC;
            rSumLastContainers += lastCont;
            rSumCurrContainers += currCont;
            rSumLastDollars += lastD;
            rSumCurrDollars += currD;

            rawRows.push([
                idx + 1,
                item.code || '-',
                item.name || '-',
                item.proforma || '-',
                lastC || '-',
                lastW,
                lastCont || '-',
                lastD || '-',
                currC || '-',
                currW,
                currCont || '-',
                currD || '-',
                diffW,
                formatPct(pct),
                diffW >= 0 ? 'افزایش ذخیره' : 'مصرف از انبار'
            ]);
        });

        const rTotalDiff = rSumCurrWeight - rSumLastWeight;
        const rTotalPct = rSumLastWeight > 0 ? (rTotalDiff / rSumLastWeight) * 100 : 0;

        rawRows.push([
            '★',
            'جمع کل',
            'سرجمع مواد اولیه و کارخانه (گروه ۰۱)',
            '-',
            rSumLastCartons,
            rSumLastWeight,
            rSumLastContainers,
            rSumLastDollars,
            rSumCurrCartons,
            rSumCurrWeight,
            rSumCurrContainers,
            rSumCurrDollars,
            rTotalDiff,
            formatPct(rTotalPct),
            rTotalDiff >= 0 ? 'رشد کلی' : 'افت کلی'
        ]);

        const wsRaw = XLSX.utils.aoa_to_sheet(rawRows);
        finalizeSheet(wsRaw, [6, 16, 34, 14, 16, 20, 14, 16, 16, 20, 14, 16, 20, 16, 18]);
        XLSX.utils.book_append_sheet(wb, wsRaw, 'مواد اولیه کارخانه');
    }

    // =========================================================================
    // SHEET 4: بارهای در راه و خرید خارجی
    // =========================================================================
    if (scope === 'all' || scope === 'logistics') {
        const transitRows: any[][] = [
            ['بارهای در حال خرید خارجی و در راه ترانزیت'],
            [`تاریخ گزارش: ${reportDate} | مخاطب: ${ceo}`],
            [''],
            [
                'ردیف',
                'نوع محموله / کالا',
                'شماره پروفرما',
                'شماره ثبت سفارش',
                'وزن خالص (KG)',
                'تعداد کارتن / بسته',
                'تعداد کانتینر',
                'ارزش دلاری ($)',
                'مبلغ ریالی (ریال)',
                'تامین‌کننده / مبدا',
                'روش تسویه',
                'وضعیت محموله'
            ]
        ];

        let pSumWeight = 0;
        let pSumCartons = 0;
        let pSumContainers = 0;
        let pSumDollars = 0;
        let pSumRials = 0;

        purchasingGoods.forEach((item, idx) => {
            const w = n(item.weight);
            const c = n(item.cartons);
            const cont = n(item.container);
            const dol = n(item.dollars);
            const rial = n(item.rialAmount);

            pSumWeight += w;
            pSumCartons += c;
            pSumContainers += cont;
            pSumDollars += dol;
            pSumRials += rial;

            transitRows.push([
                idx + 1,
                item.cargoType || 'محموله خارجی',
                item.proforma || '-',
                item.registrationNumber || '-',
                w,
                c || '-',
                cont || '-',
                dol || '-',
                rial || '-',
                item.petrochemicalName || 'تامین خارجی',
                item.paymentMethod || 'ارزی / حواله',
                item.statusBadge || 'در راه / خرید'
            ]);
        });

        transitRows.push([
            '★',
            'جمع کل بارهای در حال خرید و در راه',
            '-',
            '-',
            pSumWeight,
            pSumCartons,
            pSumContainers,
            pSumDollars,
            pSumRials,
            '-',
            '-',
            'کل محموله‌های در راه'
        ]);

        const wsTransit = XLSX.utils.aoa_to_sheet(transitRows);
        finalizeSheet(wsTransit, [6, 32, 18, 20, 18, 16, 14, 18, 22, 22, 18, 22]);
        XLSX.utils.book_append_sheet(wb, wsTransit, 'بارهای در راه و خرید');
    }

    // =========================================================================
    // SHEET 5: بارهای موجود در گمرک و ترخیص
    // =========================================================================
    if (scope === 'all' || scope === 'logistics') {
        const customsRows: any[][] = [
            ['بارهای موجود در اماکن گمرکی و در حال انجام تشریفات ترخیص'],
            [`تاریخ گزارش: ${reportDate} | مخاطب: ${ceo}`],
            [''],
            [
                'ردیف',
                'شرح کالا / محموله گمرکی',
                'شماره پروفرما',
                'شماره ثبت سفارش',
                'وزن محموله (KG)',
                'تعداد کارتن / بسته',
                'تعداد کانتینر',
                'ارزش دلاری ($)',
                'وضعیت ترخیص / اظهارنامه'
            ]
        ];

        let cSumWeight = 0;
        let cSumCartons = 0;
        let cSumContainers = 0;
        let cSumDollars = 0;

        goodsInCustoms.forEach((item, idx) => {
            const w = n(item.weight);
            const c = n(item.cartons);
            const cont = n(item.container);
            const dol = n(item.dollars);

            cSumWeight += w;
            cSumCartons += c;
            cSumContainers += cont;
            cSumDollars += dol;

            customsRows.push([
                idx + 1,
                item.cargoType || 'محموله گمرکی',
                item.proforma || '-',
                item.registrationNumber || '-',
                w,
                c || '-',
                cont || '-',
                dol || '-',
                item.statusBadge || 'در حال ترخیص گمرکی'
            ]);
        });

        customsRows.push([
            '★',
            'جمع کل بارهای گمرک',
            '-',
            '-',
            cSumWeight,
            cSumCartons,
            cSumContainers,
            cSumDollars,
            'کل موجودی گمرک'
        ]);

        const wsCustoms = XLSX.utils.aoa_to_sheet(customsRows);
        finalizeSheet(wsCustoms, [6, 32, 18, 20, 18, 16, 14, 18, 24]);
        XLSX.utils.book_append_sheet(wb, wsCustoms, 'بارهای گمرک و ترخیص');
    }

    // =========================================================================
    // SHEET 6: کالاهای تجاری
    // =========================================================================
    if (scope === 'all' || scope === 'logistics') {
        const commRows: any[][] = [
            ['کالاهای تجاری و اقلام انبارهای متفرقه'],
            [`تاریخ گزارش: ${reportDate} | مخاطب: ${ceo}`],
            [''],
            [
                'ردیف',
                'شرح کالا',
                'دسته‌بندی تجاری',
                'شماره پروفرما',
                'شماره ثبت سفارش',
                'تعداد کارتن / بسته',
                'وزن خالص (KG)',
                'تعداد کانتینر',
                'ارزش دلاری ($)'
            ]
        ];

        let comSumWeight = 0;
        let comSumCartons = 0;
        let comSumContainers = 0;
        let comSumDollars = 0;

        commercialGoods.forEach((item, idx) => {
            const w = n(item.weight);
            const c = n(item.cartons);
            const cont = n(item.container);
            const dol = n(item.dollars);

            comSumWeight += w;
            comSumCartons += c;
            comSumContainers += cont;
            comSumDollars += dol;

            commRows.push([
                idx + 1,
                item.itemName || 'کالای تجاری',
                item.category || 'عمومی',
                item.proforma || '-',
                item.registrationNumber || '-',
                c || '-',
                w,
                cont || '-',
                dol || '-'
            ]);
        });

        commRows.push([
            '★',
            'جمع کل کالاهای تجاری',
            '-',
            '-',
            '-',
            comSumCartons,
            comSumWeight,
            comSumContainers,
            comSumDollars
        ]);

        const wsCommercial = XLSX.utils.aoa_to_sheet(commRows);
        finalizeSheet(wsCommercial, [6, 30, 20, 18, 20, 16, 18, 14, 18]);
        XLSX.utils.book_append_sheet(wb, wsCommercial, 'کالای تجاری');
    }

    // =========================================================================
    // SHEET 7: خریدهای داخلی پتروشیمی
    // =========================================================================
    if (scope === 'all' || scope === 'logistics') {
        const domesticRows: any[][] = [
            ['خریدهای داخلی پتروشیمی و بورس کالا'],
            [`تاریخ گزارش: ${reportDate} | مخاطب: ${ceo}`],
            [''],
            [
                'ردیف',
                'نوع محموله / کالا',
                'پتروشیمی / کارخانه',
                'شماره قرارداد / پروفرما',
                'وزن خالص (KG)',
                'تعداد بسته / کارتن',
                'تعداد کانتینر / تریلی',
                'ارزش دلاری معادل ($)',
                'مبلغ ریالی (ریال)',
                'روش پرداخت',
                'وضعیت محموله'
            ]
        ];

        let dSumWeight = 0;
        let dSumCartons = 0;
        let dSumContainers = 0;
        let dSumDollars = 0;
        let dSumRials = 0;

        domesticPurchases.forEach((item, idx) => {
            const w = n(item.weight);
            const c = n(item.cartons);
            const cont = n(item.container);
            const dol = n(item.dollars);
            const rial = n(item.rialAmount);

            dSumWeight += w;
            dSumCartons += c;
            dSumContainers += cont;
            dSumDollars += dol;
            dSumRials += rial;

            domesticRows.push([
                idx + 1,
                item.cargoType || 'پتروشیمی',
                item.petrochemicalName || 'بورس کالا',
                item.proforma || item.registrationNumber || '-',
                w,
                c || '-',
                cont || '-',
                dol || '-',
                rial || '-',
                item.paymentMethod || 'نقدی / LC داخلی',
                item.statusBadge || 'خرید قطعی'
            ]);
        });

        domesticRows.push([
            '★',
            'جمع کل خریدهای داخلی پتروشیمی',
            '-',
            '-',
            dSumWeight,
            dSumCartons,
            dSumContainers,
            dSumDollars,
            dSumRials,
            '-',
            'کل خریدهای داخلی'
        ]);

        const wsDomestic = XLSX.utils.aoa_to_sheet(domesticRows);
        finalizeSheet(wsDomestic, [6, 28, 22, 22, 18, 16, 18, 18, 22, 20, 20]);
        XLSX.utils.book_append_sheet(wb, wsDomestic, 'خریدهای داخلی پتروشیمی');
    }

    // =========================================================================
    // SHEET 8: ماتریس نظارتی اقلام دارای افت ذخیره و کسری منفی
    // =========================================================================
    if (scope === 'all' || scope === 'variance') {
        const varRows: any[][] = [
            ['ماتریس نظارتی اقلام دارای افت موجودی و هشدار کسری انبار'],
            [`تاریخ گزارش: ${reportDate} | اولویت پایش: بالا | مدیر مربوطه: ${manager}`],
            [''],
            [
                'ردیف',
                'کد کالا / گروه',
                'شرح کالا / گروه انبار',
                'دسته‌بندی سازمانی',
                `وزن ${r1Label} (KG)`,
                `وزن ${r2Label} (KG)`,
                'میزان افت / کسری (KG)',
                'درصد کاهش (%)',
                'درجه اولویت هشدار',
                'اقدام پیشگیرانه پیشنهادی'
            ]
        ];

        negativeItems.forEach((item, idx) => {
            const lastW = n(item.lastYearWeight);
            const currW = n(item.currentWeight);
            const diffW = currW - lastW;
            const pct = lastW > 0 ? (diffW / lastW) * 100 : (currW < 0 ? -100 : 0);
            const isSevere = Math.abs(diffW) > 10000 || pct < -40 || currW < 0;

            varRows.push([
                idx + 1,
                item.code || '-',
                item.name || item.groupName || '-',
                item.categoryLabel || item.category || 'انبار',
                lastW,
                currW,
                diffW,
                formatPct(pct),
                isSevere ? '🔴 بحرانی (اقدام فوری)' : '🟡 هشدار (نیازمند پیگیری)',
                currW < 0
                    ? 'رفع مغایرت منفی در کاردکس انبار و ثبت فاکتورهای کسری'
                    : 'صدور پیش‌فاکتور خرید جدید و هماهنگی با بازرگانی'
            ]);
        });

        const wsVariance = XLSX.utils.aoa_to_sheet(varRows);
        finalizeSheet(wsVariance, [6, 18, 34, 20, 20, 20, 20, 16, 22, 40]);
        XLSX.utils.book_append_sheet(wb, wsVariance, 'هشدار کسری منفی');
    }

    // 2. Generate and trigger download
    const cleanDate = reportDate.replace(/[\/\\]/g, '-');
    const scopeSuffix = scope === 'all' ? 'جامع_مدیریتی' : (scope === 'summary' ? 'خلاصه_تراز' : scope);
    const fileName = `گزارش_جامع_تراز_انبار_${cleanDate}_${scopeSuffix}.xlsx`;

    XLSX.writeFile(wb, fileName);
};
