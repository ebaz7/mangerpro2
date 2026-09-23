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
    const finalizeSheet = (ws: XLSX.WorkSheet, colWidths: number[]) => {
        ws['!views'] = [{ rightToLeft: true, RTL: true }];
        ws['!cols'] = colWidths.map(w => ({ wch: w }));
    };

    // ==========================================
    // SHEET 1: خلاصه مدیریتی و تراز کل (Summary & KPIs)
    // ==========================================
    if (scope === 'all' || scope === 'summary') {
        const sumRows: any[][] = [
            ['سامانه نمای کلی موجودی و مغایرت سالانه انبار - گزارش مدیریتی جامع'],
            ['پایش همزمان موجودی‌های سامانه یکپارچه سایان، انبارهای تجاری، بارهای در راه و گمرک'],
            [''],
            ['تاریخ گزارش:', reportDate, '', 'مخاطب گزارش:', ceo],
            ['دوره مقایسه اول:', `${r1Label} (مورخ ${r1Date})`, '', 'تنظیم‌کننده:', manager],
            ['دوره مقایسه دوم:', `${r2Label} (مورخ ${r2Date})`, '', 'سیستم مالی/انبار:', 'سامانه یکپارچه سایان (Live ERP)'],
            [''],
            ['', '', '', '', '', '', ''],
            ['جدول ۱: شاخص‌های کلان تراز وزنی زنجیره تامین و انبار (کیلوگرم)'],
            [
                'ردیف',
                'سرفصل زنجیره تامین و موجودی انبارها',
                `وزن ${r1Label} (KG)`,
                `وزن ${r2Label} (KG)`,
                'اختلاف وزنی (KG)',
                'درصد تغییرات (%)',
                'تحلیل روند و وضعیت'
            ]
        ];

        const yLast = n(summary.lastYearYarnsWeight);
        const yCurr = n(summary.currentYarnsWeight);
        const yDiff = yCurr - yLast;
        const yPct = yLast > 0 ? (yDiff / yLast) * 100 : 0;

        const rLast = n(summary.lastYearRawWeight);
        const rCurr = n(summary.currentRawWeight);
        const rDiff = rCurr - rLast;
        const rPct = rLast > 0 ? (rDiff / rLast) * 100 : 0;

        const transitWeight = purchasingGoods.reduce((s, x) => s + n(x.weight), 0);
        const customsWeight = goodsInCustoms.reduce((s, x) => s + n(x.weight), 0);
        const domesticWeight = domesticPurchases.reduce((s, x) => s + n(x.weight), 0);
        const commercialWeight = commercialGoods.reduce((s, x) => s + n(x.weight), 0);

        const tLast = n(summary.lastYearTotalWeight) || (yLast + rLast);
        const tCurr = n(summary.currentTotalWeight) || (yCurr + rCurr);
        const tDiff = tCurr - tLast;
        const tPct = tLast > 0 ? (tDiff / tLast) * 100 : 0;

        sumRows.push([
            1,
            'کالاهای تولیدی کارخانه (نخ‌های بافته شده و انبار محصول)',
            yLast,
            yCurr,
            yDiff,
            formatPct(yPct),
            yDiff >= 0 ? 'افزایش موجودی انبار' : 'کاهش موجودی (خروج بیشتر از تولید)'
        ]);

        sumRows.push([
            2,
            'مواد اولیه، الیاف و اقلام انبار کارخانه (چیپس، لاکرا، پلی‌استر و...)',
            rLast,
            rCurr,
            rDiff,
            formatPct(rPct),
            rDiff >= 0 ? 'افزایش ذخیره مواد اولیه' : 'افت موجودی و مصرف از انبار'
        ]);

        sumRows.push([
            3,
            'بارهای در حال خرید خارجی و ترانزیت در راه',
            0,
            transitWeight,
            transitWeight,
            '+۱۰۰٪',
            'محموله فعال در زنجیره لجستیک'
        ]);

        sumRows.push([
            4,
            'بارهای موجود در اماکن گمرکی (در حال ترخیص)',
            0,
            customsWeight,
            customsWeight,
            customsWeight > 0 ? '+۱۰۰٪' : '۰٪',
            'در انتظار تکمیل تشریفات گمرکی'
        ]);

        sumRows.push([
            5,
            'خریدهای داخلی پتروشیمی و بورس کالا',
            0,
            domesticWeight,
            domesticWeight,
            domesticWeight > 0 ? '+۱۰۰٪' : '۰٪',
            'سفارشات قطعی داخلی'
        ]);

        sumRows.push([
            6,
            'کالاهای تجاری و متفرقه',
            0,
            commercialWeight,
            commercialWeight,
            commercialWeight > 0 ? '+۱۰۰٪' : '۰٪',
            'موجودی انبارهای تجاری'
        ]);

        sumRows.push([
            '★',
            'سرجمع کل موجودی و ورودی‌های زنجیره تامین (کل انبار)',
            tLast,
            tCurr,
            tDiff,
            formatPct(tPct),
            tDiff >= 0 ? 'تراز مثبت و پایدار' : 'تراز نزولی و مصرفی'
        ]);

        sumRows.push(['']);
        sumRows.push(['جدول ۲: مقایسه کانتینری و ارزش ارزی']);
        sumRows.push([
            'ردیف',
            'شاخص سنجش',
            `${r1Label}`,
            `${r2Label}`,
            'اختلاف',
            'درصد تغییر (%)',
            'توضیحات'
        ]);

        const cLast = n(summary.totalLastYearContainers);
        const cCurr = n(summary.containersTotal);
        const cDiff = cCurr - cLast;
        const cPct = cLast > 0 ? (cDiff / cLast) * 100 : 0;

        const dLast = n(summary.totalLastYearDollars);
        const dCurr = n(summary.dollarsTotal);
        const dDiff = dCurr - dLast;
        const dPct = dLast > 0 ? (dDiff / dLast) * 100 : 0;

        sumRows.push([
            1,
            'تعداد کانتینر معادل (Container)',
            cLast,
            cCurr,
            cDiff,
            formatPct(cPct),
            cDiff >= 0 ? 'رشد ظرفیت کانتینری' : 'کاهش ظرفیت کانتینری'
        ]);

        sumRows.push([
            2,
            'ارزش دلاری بارهای ورودی و تجاری ($)',
            dLast,
            dCurr,
            dDiff,
            formatPct(dPct),
            dDiff >= 0 ? 'افزایش ارزش ورودی‌ها' : 'کاهش ارزش ورودی‌ها'
        ]);

        sumRows.push(['']);
        sumRows.push(['جدول ۳: خلاصه اقلام دارای نوسان شدید و کسری']);
        sumRows.push([
            'ردیف',
            'وضعیت تحلیلی',
            'تعداد اقلام',
            'مجموع نوسان وزنی (KG)',
            'توصیه مدیریتی'
        ]);

        const negWeight = negativeItems.reduce((s, x) => s + (x.diffWeight || 0), 0);
        const groWeight = growthItems.reduce((s, x) => s + (x.diffWeight || 0), 0);

        sumRows.push([
            1,
            'اقلام دارای کسری منفی یا افت ذخیره',
            negativeItems.length,
            negWeight,
            'نیازمند برنامه‌ریزی فوری خرید و سفارش‌گذاری جایگزین'
        ]);

        sumRows.push([
            2,
            'اقلام دارای رشد موجودی و مازاد',
            growthItems.length,
            groWeight,
            'پایش نقطه سفارش جهت جلوگیری از خواب سرمایه'
        ]);

        const wsSummary = XLSX.utils.aoa_to_sheet(sumRows);
        finalizeSheet(wsSummary, [6, 42, 22, 22, 22, 18, 38]);
        XLSX.utils.book_append_sheet(wb, wsSummary, 'خلاصه تراز و شاخص‌ها');
    }

    // ==========================================
    // SHEET 2: کالاهای تولیدی (نخ‌ها) - گروه 04
    // ==========================================
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

            // If detail child items function is available, append sub-rows
            if (getGroupChildItems && item.code) {
                const childItems = getGroupChildItems(item.code);
                if (childItems && childItems.length > 1) {
                    childItems.forEach(ch => {
                        yarnRows.push([
                            '',
                            `  ↳ ${ch.itemCode}`,
                            `     └ ${ch.itemName}`,
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            '-',
                            'زیرمجموعه'
                        ]);
                    });
                }
            }
        });

        const yTotalDiff = ySumCurrWeight - ySumLastWeight;
        const yTotalPct = ySumLastWeight > 0 ? (yTotalDiff / ySumLastWeight) * 100 : 0;

        // Total Row
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
        finalizeSheet(wsYarns, [6, 16, 32, 14, 14, 20, 14, 16, 14, 20, 14, 16, 20, 16, 18]);
        XLSX.utils.book_append_sheet(wb, wsYarns, 'کالاهای تولیدی (نخ‌ها)');
    }

    // ==========================================
    // SHEET 3: مواد اولیه و کارخانه - گروه 01
    // ==========================================
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

        // Total Row
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
        finalizeSheet(wsRaw, [6, 16, 32, 14, 16, 20, 14, 16, 16, 20, 14, 16, 20, 16, 18]);
        XLSX.utils.book_append_sheet(wb, wsRaw, 'مواد اولیه کارخانه');
    }

    // ==========================================
    // SHEET 4: بارهای در راه، خرید خارجی و داخلی
    // ==========================================
    if (scope === 'all' || scope === 'logistics') {
        const transitRows: any[][] = [
            ['بارهای در حال خرید خارجی، در راه ترانزیت و خریدهای بورس و پتروشیمی'],
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
                'پتروشیمی / تامین‌کننده',
                'روش تسویه',
                'وضعیت محموله'
            ]
        ];

        let pSumWeight = 0;
        let pSumCartons = 0;
        let pSumContainers = 0;
        let pSumDollars = 0;
        let pSumRials = 0;

        // Foreign purchases & in-transit
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

        // Domestic & Petrochemical purchases
        domesticPurchases.forEach((item, idx) => {
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
                purchasingGoods.length + idx + 1,
                `${item.cargoType || 'خرید داخلی'} (پتروشیمی)`,
                item.proforma || '-',
                item.registrationNumber || '-',
                w,
                c || '-',
                cont || '-',
                dol || '-',
                rial || '-',
                item.petrochemicalName || 'بورس کالا',
                item.paymentMethod || 'نقدی / LC داخلی',
                item.statusBadge || 'خرید قطعی پتروشیمی'
            ]);
        });

        // Summary row
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
            'کل محموله‌ها'
        ]);

        const wsTransit = XLSX.utils.aoa_to_sheet(transitRows);
        finalizeSheet(wsTransit, [6, 32, 18, 20, 18, 16, 14, 18, 22, 22, 18, 22]);
        XLSX.utils.book_append_sheet(wb, wsTransit, 'بارهای در راه و خرید');
    }

    // ==========================================
    // SHEET 5: بارهای موجود در گمرک و ترخیص
    // ==========================================
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

        // Summary row
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

    // ==========================================
    // SHEET 6: کالاهای تجاری
    // ==========================================
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

        // Summary row
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

    // ==========================================
    // SHEET 7: ماتریس تحلیل، اقلام کسری و هشدار
    // ==========================================
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
        finalizeSheet(wsVariance, [6, 18, 32, 20, 20, 20, 20, 16, 22, 40]);
        XLSX.utils.book_append_sheet(wb, wsVariance, 'هشدار کسری منفی');
    }

    // ==========================================
    // SHEET 8: ماتریس جامع مقایسه‌ای (کل اقلام)
    // ==========================================
    if (scope === 'all' && allComparedItems && allComparedItems.length > 0) {
        const allRows: any[][] = [
            ['ماتریس جامع مقایسه‌ای تمام اقلام زنجیره تامین و انبار'],
            [`تاریخ استخراج داده: ${reportDate} | منبع: سامانه سایان ERP`],
            [''],
            [
                'ردیف',
                'کد سیستمی',
                'شرح کالا / محموله',
                'دسته‌بندی',
                `وزن دوره قبل (KG)`,
                `وزن دوره جاری (KG)`,
                'تغییرات خالص (KG)',
                'درصد نوسان (%)',
                'وضعیت نوسان'
            ]
        ];

        allComparedItems.forEach((item, idx) => {
            const lastW = n(item.lastYearWeight);
            const currW = n(item.currentWeight);
            const diffW = currW - lastW;
            const pct = item.ratio !== undefined ? item.ratio : (lastW > 0 ? (diffW / lastW) * 100 : 0);

            allRows.push([
                idx + 1,
                item.code || '-',
                item.name || '-',
                item.categoryLabel || item.category || '-',
                lastW,
                currW,
                diffW,
                formatPct(pct),
                diffW > 0 ? 'افزایش' : (diffW < 0 ? 'کاهش' : 'بدون تغییر')
            ]);
        });

        const wsAll = XLSX.utils.aoa_to_sheet(allRows);
        finalizeSheet(wsAll, [6, 18, 34, 20, 20, 20, 20, 16, 16]);
        XLSX.utils.book_append_sheet(wb, wsAll, 'تراز جامع مقایسه‌ای');
    }

    // 2. Generate and trigger download
    const cleanDate = reportDate.replace(/[\/\\]/g, '-');
    const scopeSuffix = scope === 'all' ? 'کامل_چندشیته' : (scope === 'summary' ? 'خلاصه_مدیریتی' : scope);
    const fileName = `گزارش_تراز_انبار_${cleanDate}_${scopeSuffix}.xlsx`;

    XLSX.writeFile(wb, fileName);
};
