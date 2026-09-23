import React, { useState } from 'react';
import { TradeRecord, PetrochemicalPurchaseData, User } from '../../types';
import { formatCurrency, formatNumberString, deformatNumberString, formatDate } from '../../constants';
import { Building2, FileText, Landmark, Truck, Warehouse, Calculator, Plus, Trash2, CheckCircle2, Clock, AlertCircle, ShieldCheck, Scale, FileSpreadsheet, Paperclip, ChevronRight, Save } from 'lucide-react';
import { TradeDatePicker } from '../TradeDatePicker';

interface DomesticPetrochemicalTabProps {
    record: TradeRecord;
    onUpdateRecord: (updatedRecord: TradeRecord) => Promise<void> | void;
    currentUser: User;
}

export const DomesticPetrochemicalTab: React.FC<DomesticPetrochemicalTabProps> = ({
    record,
    onUpdateRecord,
    currentUser
}) => {
    const [subTab, setSubTab] = useState<'contract' | 'settlement' | 'loading_shipping' | 'warehouse' | 'costing'>('contract');

    const petroData: PetrochemicalPurchaseData = record.petrochemicalData || {
        petrochemicalName: record.sellerName || 'پتروشیمی شهید تندگویان',
        brokerName: 'کارگزاری بورس کالا',
        contractNumber: record.registrationNumber || '',
        proformaNumber: record.proformaNumber || '',
        proformaDate: record.startDate ? record.startDate.split('T')[0] : '',
        paymentMethod: 'internal_lc',
        gradeName: record.goodsName || 'چیپس پلی استر نساجی TG642',
        quantityKg: (record.items || []).reduce((s, i) => s + (i.weight || 0), 0) || 50000,
        basePricePerKg: 0,
        totalGoodsPrice: 0,
        vatAmount: 0,
        brokerageFee: 0,
        totalInvoiceAmount: 0,
        internalLc: {
            lcNumber: '',
            issuingBank: 'بانک تجارت',
            branch: 'شعبه مرکزی',
            issueDate: '',
            dueDate: '',
            lcAmount: 0,
            prepaymentAmount: 0,
            collateralDesc: '',
            commissionFee: 0,
            status: 'draft'
        },
        draftBarat: {
            baratNumber: '',
            sepamCode: '',
            bankName: 'بانک ملت',
            issueDate: '',
            dueDate: '',
            amount: 0,
            drawerName: 'شرکت کارخانجات',
            draweeName: 'پتروشیمی',
            status: 'draft'
        },
        loadingNotice: {
            remittanceNumber: '',
            behenyabCode: '',
            driverName: '',
            driverPhone: '',
            driverNationalCode: '',
            truckPlate: '',
            waybillNumber: '',
            freightCostRial: 0,
            loadingDate: '',
            deliveryStatus: 'pending_loading'
        },
        warehouseReceipt: {
            receiptNumber: '',
            receivedWeightKg: 0,
            receiptDate: '',
            warehouseName: 'انبار مرکزی مواد اولیه کارخانه',
            isConfirmed: false
        }
    };

    const updatePetro = async (partial: Partial<PetrochemicalPurchaseData>) => {
        const next: PetrochemicalPurchaseData = {
            ...petroData,
            ...partial
        };

        // Recalculate amounts if quantity or basePrice changed
        const qty = next.quantityKg || 0;
        const base = next.basePricePerKg || 0;
        const totalGoods = qty * base;
        const vat = next.vatAmount !== undefined && next.vatAmount > 0 ? next.vatAmount : Math.round(totalGoods * 0.10);
        const brokerage = next.brokerageFee || Math.round(totalGoods * 0.0035); // standard bourse fee ~0.35%
        const totalInvoice = totalGoods + vat + brokerage;

        next.totalGoodsPrice = totalGoods;
        if (!partial.vatAmount) next.vatAmount = vat;
        if (!partial.brokerageFee) next.brokerageFee = brokerage;
        next.totalInvoiceAmount = totalInvoice;

        await onUpdateRecord({
            ...record,
            petrochemicalData: next,
            sellerName: next.petrochemicalName || record.sellerName,
            goodsName: next.gradeName || record.goodsName,
            proformaNumber: next.proformaNumber || record.proformaNumber,
            registrationNumber: next.contractNumber || record.registrationNumber
        });
    };

    // Calculate final costing metrics
    const totalFreight = petroData.loadingNotice?.freightCostRial || 0;
    const totalBankCommission = petroData.paymentMethod === 'internal_lc' 
        ? (petroData.internalLc?.commissionFee || 0) 
        : 0;
    const finalTotalCost = (petroData.totalInvoiceAmount || 0) + totalFreight + totalBankCommission;
    const finalReceivedWeight = petroData.warehouseReceipt?.receivedWeightKg || petroData.quantityKg || 1;
    const costPerKg = finalReceivedWeight > 0 ? Math.round(finalTotalCost / finalReceivedWeight) : 0;
    const weightVarianceKg = (petroData.warehouseReceipt?.receivedWeightKg || 0) > 0 
        ? (petroData.warehouseReceipt!.receivedWeightKg! - (petroData.quantityKg || 0)) 
        : 0;

    return (
        <div className="space-y-6 p-4 sm:p-6 max-w-6xl mx-auto" dir="rtl">
            {/* Header Banner */}
            <div className="bg-gradient-to-r from-emerald-800 via-teal-800 to-slate-900 text-white p-5 rounded-3xl shadow-xl flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div className="space-y-1">
                    <div className="inline-flex items-center gap-2 px-3 py-1 bg-emerald-500/20 text-emerald-300 rounded-full text-xs font-bold border border-emerald-400/30">
                        <Building2 size={14} />
                        <span>پرونده خرید داخلی / پتروشیمی و بورس کالا</span>
                    </div>
                    <h2 className="text-xl sm:text-2xl font-black text-white">
                        {petroData.petrochemicalName || 'پتروشیمی'} - {petroData.gradeName || 'کالای بورسی'}
                    </h2>
                    <p className="text-xs text-emerald-200">
                        قرارداد بورس: <span className="font-mono font-bold text-white">{petroData.contractNumber || 'ثبت نشده'}</span> | 
                        پیش‌فاکتور: <span className="font-mono font-bold text-white">{petroData.proformaNumber || 'ثبت نشده'}</span> | 
                        روش تسویه: <span className="font-bold text-amber-300">
                            {petroData.paymentMethod === 'internal_lc' ? 'اعتبار اسنادی داخلی (LC ریالی)' : petroData.paymentMethod === 'draft_barat' ? 'برات الکترونیک (سپام)' : 'نقدی بورس'}
                        </span>
                    </p>
                </div>

                {/* Quick Cost Card */}
                <div className="bg-white/10 backdrop-blur-md px-5 py-3 rounded-2xl border border-white/20 text-left shrink-0">
                    <span className="text-[11px] text-emerald-200 block">بهای تمام‌شده هر کیلوگرم</span>
                    <span className="font-mono text-xl sm:text-2xl font-black text-white">
                        {costPerKg.toLocaleString('fa-IR')} <span className="text-xs font-normal">ریال/kg</span>
                    </span>
                    <span className="text-[10px] text-emerald-300 block mt-0.5 font-mono">
                        ارزش کل: {(finalTotalCost || 0).toLocaleString('fa-IR')} ریال
                    </span>
                </div>
            </div>

            {/* Navigation Sub-Tabs */}
            <div className="flex gap-2 overflow-x-auto pb-1 custom-scrollbar border-b border-gray-200 dark:border-gray-800">
                <button
                    type="button"
                    onClick={() => setSubTab('contract')}
                    className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap ${
                        subTab === 'contract'
                            ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
                            : 'bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700'
                    }`}
                >
                    <FileText size={16} />
                    <span>۱. قرارداد بورس و پیش‌فاکتور</span>
                </button>
                <button
                    type="button"
                    onClick={() => setSubTab('settlement')}
                    className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap ${
                        subTab === 'settlement'
                            ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
                            : 'bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700'
                    }`}
                >
                    <Landmark size={16} />
                    <span>۲. تسویه (LC ریالی / برات / نقدی)</span>
                </button>
                <button
                    type="button"
                    onClick={() => setSubTab('loading_shipping')}
                    className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap ${
                        subTab === 'loading_shipping'
                            ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
                            : 'bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700'
                    }`}
                >
                    <Truck size={16} />
                    <span>۳. حواله پتروشیمی، بهین‌یاب و حمل</span>
                </button>
                <button
                    type="button"
                    onClick={() => setSubTab('warehouse')}
                    className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap ${
                        subTab === 'warehouse'
                            ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
                            : 'bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700'
                    }`}
                >
                    <Warehouse size={16} />
                    <span>۴. رسید انبار کارخانه و باسکول</span>
                </button>
                <button
                    type="button"
                    onClick={() => setSubTab('costing')}
                    className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap ${
                        subTab === 'costing'
                            ? 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
                            : 'bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700'
                    }`}
                >
                    <Calculator size={16} />
                    <span>۵. آنالیز بهای تمام‌شده نهایی</span>
                </button>
            </div>

            {/* SUB-TAB 1: Bourse Contract & Proforma */}
            {subTab === 'contract' && (
                <div className="bg-white dark:bg-gray-800 p-6 rounded-3xl border border-gray-200 dark:border-gray-700 space-y-6 shadow-sm">
                    <div className="flex items-center justify-between border-b border-gray-100 dark:border-gray-700 pb-4">
                        <div>
                            <h3 className="text-base font-bold text-gray-900 dark:text-gray-100">مشخصات معامله بورس کالا و پیش‌فاکتور پتروشیمی</h3>
                            <p className="text-xs text-gray-500 mt-0.5">اطلاعات پایه عرضه، کارگزاری، گرید و مبالغ معامله شده</p>
                        </div>
                        <span className="px-3 py-1 bg-emerald-50 text-emerald-700 dark:bg-emerald-950/40 dark:text-emerald-300 text-xs font-bold rounded-xl border border-emerald-200">
                            بورس کالای ایران
                        </span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">نام مجتمع پتروشیمی / فروشنده *</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.petrochemicalName}
                                onChange={e => updatePetro({ petrochemicalName: e.target.value })}
                                placeholder="مثال: پتروشیمی شهید تندگویان، اروند، شازند..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">کارگزاری بورس کالا</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.brokerName || ''}
                                onChange={e => updatePetro({ brokerName: e.target.value })}
                                placeholder="مثال: کاریزما، آگاه، مفید، صبا جهاد..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره قرارداد / اطلاعیه خرید بورس</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.contractNumber || ''}
                                onChange={e => updatePetro({ contractNumber: e.target.value })}
                                placeholder="شناسه معامله / قرارداد..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره پیش‌فاکتور داخلی پتروشیمی</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.proformaNumber || ''}
                                onChange={e => updatePetro({ proformaNumber: e.target.value })}
                                placeholder="شماره پیش‌فاکتور..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">تاریخ پیش‌فاکتور</label>
                            <TradeDatePicker
                                value={petroData.proformaDate || ''}
                                onChange={d => updatePetro({ proformaDate: d })}
                                placeholder="تاریخ صدور..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">گرید کالایی پتروشیمی *</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.gradeName || ''}
                                onChange={e => updatePetro({ gradeName: e.target.value })}
                                placeholder="مثال: چیپس نساجی TG642، پلی‌پروپیلن C30S..."
                            />
                        </div>
                    </div>

                    {/* Quantities & Price Breakdown Grid */}
                    <div className="p-5 bg-slate-50 dark:bg-gray-900/50 rounded-2xl border border-slate-200 dark:border-gray-700 space-y-4">
                        <h4 className="font-bold text-sm text-slate-800 dark:text-slate-200 flex items-center gap-2">
                            <Calculator size={18} className="text-emerald-600" />
                            <span>محاسبه مالی پیش‌فاکتور و معامله بورس کالا</span>
                        </h4>

                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-700 dark:text-gray-300">وزن معامله شده (کیلوگرم) *</label>
                                <input
                                    type="number"
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-base font-mono font-bold text-gray-900 dark:text-gray-100 text-center"
                                    value={petroData.quantityKg || ''}
                                    onChange={e => updatePetro({ quantityKg: parseFloat(e.target.value) || 0 })}
                                    placeholder="مثال: 50000"
                                />
                                <span className="text-[11px] text-gray-500 block text-left">
                                    معادل: {((petroData.quantityKg || 0) / 1000).toFixed(2)} تن
                                </span>
                            </div>

                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-700 dark:text-gray-300">قیمت معامله شده هر کیلو (ریال) *</label>
                                <input
                                    type="text"
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-base font-mono font-bold text-emerald-700 dark:text-emerald-400 text-center"
                                    value={formatNumberString(petroData.basePricePerKg || 0)}
                                    onChange={e => updatePetro({ basePricePerKg: deformatNumberString(e.target.value) })}
                                    placeholder="نرخ هر کیلوگرم..."
                                />
                                <span className="text-[11px] text-gray-500 block text-left">
                                    {(petroData.basePricePerKg || 0).toLocaleString('fa-IR')} ریال
                                </span>
                            </div>

                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-700 dark:text-gray-300">مبلغ خالص کالا (ریال)</label>
                                <div className="w-full p-3 bg-emerald-50/50 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-800/60 rounded-xl font-mono text-base font-black text-emerald-800 dark:text-emerald-300 text-center">
                                    {(petroData.totalGoodsPrice || 0).toLocaleString('fa-IR')}
                                </div>
                            </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs pt-2 border-t border-slate-200 dark:border-gray-800">
                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-700 dark:text-gray-300">مالیات بر ارزش افزوده (۱۰٪) (ریال)</label>
                                <input
                                    type="text"
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-gray-800 dark:text-gray-200 text-center"
                                    value={formatNumberString(petroData.vatAmount || 0)}
                                    onChange={e => updatePetro({ vatAmount: deformatNumberString(e.target.value) })}
                                />
                            </div>

                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-700 dark:text-gray-300">کارمزد بورس کالا و کارگزاری (ریال)</label>
                                <input
                                    type="text"
                                    className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-gray-800 dark:text-gray-200 text-center"
                                    value={formatNumberString(petroData.brokerageFee || 0)}
                                    onChange={e => updatePetro({ brokerageFee: deformatNumberString(e.target.value) })}
                                />
                            </div>

                            <div className="space-y-1.5">
                                <label className="font-bold text-gray-900 dark:text-gray-100">جمع کل پیش‌فاکتور و فاکتور (ریال)</label>
                                <div className="w-full p-3 bg-indigo-50 dark:bg-indigo-950/40 border border-indigo-200 dark:border-indigo-800 rounded-xl font-mono text-base font-black text-indigo-900 dark:text-indigo-200 text-center">
                                    {(petroData.totalInvoiceAmount || 0).toLocaleString('fa-IR')}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            )}

            {/* SUB-TAB 2: Settlement (LC / Draft / Cash) */}
            {subTab === 'settlement' && (
                <div className="bg-white dark:bg-gray-800 p-6 rounded-3xl border border-gray-200 dark:border-gray-700 space-y-6 shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-3 border-b border-gray-100 dark:border-gray-700 pb-4">
                        <div>
                            <h3 className="text-base font-bold text-gray-900 dark:text-gray-100">روش و ابزار تسویه مالی با پتروشیمی و بورس</h3>
                            <p className="text-xs text-gray-500 mt-0.5">مشخصات اعتبار اسنادی داخلی (LC ریالی)، برات الکترونیک یا حواله نقدی</p>
                        </div>

                        {/* Payment Method Selector */}
                        <div className="flex gap-2">
                            <button
                                type="button"
                                onClick={() => updatePetro({ paymentMethod: 'internal_lc' })}
                                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                                    petroData.paymentMethod === 'internal_lc'
                                        ? 'bg-indigo-600 text-white shadow-md'
                                        : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300'
                                }`}
                            >
                                اعتبار اسنادی داخلی (LC)
                            </button>
                            <button
                                type="button"
                                onClick={() => updatePetro({ paymentMethod: 'draft_barat' })}
                                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                                    petroData.paymentMethod === 'draft_barat'
                                        ? 'bg-indigo-600 text-white shadow-md'
                                        : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300'
                                }`}
                            >
                                برات الکترونیک (سپام)
                            </button>
                            <button
                                type="button"
                                onClick={() => updatePetro({ paymentMethod: 'cash' })}
                                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                                    petroData.paymentMethod === 'cash'
                                        ? 'bg-indigo-600 text-white shadow-md'
                                        : 'bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300'
                                }`}
                            >
                                تسویه نقدی بورس
                            </button>
                        </div>
                    </div>

                    {/* Internal LC Fields */}
                    {petroData.paymentMethod === 'internal_lc' && (
                        <div className="space-y-4">
                            <div className="p-4 bg-indigo-50/60 dark:bg-indigo-950/30 rounded-2xl border border-indigo-200 dark:border-indigo-800 flex items-center justify-between">
                                <div className="flex items-center gap-2">
                                    <Landmark className="text-indigo-600 dark:text-indigo-400" size={20} />
                                    <span className="font-bold text-sm text-indigo-950 dark:text-indigo-200">
                                        اعتبار اسنادی دیداری / مدت‌دار ریالی (Internal LC)
                                    </span>
                                </div>
                                <span className="text-xs px-2.5 py-1 bg-white dark:bg-gray-900 text-indigo-700 dark:text-indigo-300 font-bold rounded-lg border border-indigo-200">
                                    تسهیلات بانکی پتروشیمی
                                </span>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">شماره اعتبار اسنادی (LC Number)</label>
                                    <input
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                        value={petroData.internalLc?.lcNumber || ''}
                                        onChange={e => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), lcNumber: e.target.value }
                                        })}
                                        placeholder="شماره LC..."
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">بانک گشایش‌کننده و شعبه</label>
                                    <input
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                        value={petroData.internalLc?.issuingBank || ''}
                                        onChange={e => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), issuingBank: e.target.value }
                                        })}
                                        placeholder="مثال: بانک تجارت شعبه مرکزی"
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">مبلغ کل اعتبار LC (ریال)</label>
                                    <input
                                        type="text"
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-indigo-700 dark:text-indigo-300 text-center"
                                        value={formatNumberString(petroData.internalLc?.lcAmount || petroData.totalInvoiceAmount || 0)}
                                        onChange={e => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), lcAmount: deformatNumberString(e.target.value) }
                                        })}
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">تاریخ گشایش LC</label>
                                    <TradeDatePicker
                                        value={petroData.internalLc?.issueDate || ''}
                                        onChange={d => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), issueDate: d }
                                        })}
                                        placeholder="تاریخ گشایش..."
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">تاریخ سررسید LC</label>
                                    <TradeDatePicker
                                        value={petroData.internalLc?.dueDate || ''}
                                        onChange={d => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), dueDate: d }
                                        })}
                                        placeholder="تاریخ سررسید..."
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">کارمزد گشایش بانکی (ریال)</label>
                                    <input
                                        type="text"
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-gray-800 dark:text-gray-200 text-center"
                                        value={formatNumberString(petroData.internalLc?.commissionFee || 0)}
                                        onChange={e => updatePetro({
                                            internalLc: { ...(petroData.internalLc || { lcNumber: '', issuingBank: '', branch: '', issueDate: '', dueDate: '', lcAmount: 0, status: 'draft' }), commissionFee: deformatNumberString(e.target.value) }
                                        })}
                                        placeholder="کارمزد بانک..."
                                    />
                                </div>
                            </div>
                        </div>
                    )}

                    {/* Draft / Barat Fields */}
                    {petroData.paymentMethod === 'draft_barat' && (
                        <div className="space-y-4">
                            <div className="p-4 bg-teal-50/60 dark:bg-teal-950/30 rounded-2xl border border-teal-200 dark:border-teal-800 flex items-center justify-between">
                                <div className="flex items-center gap-2">
                                    <FileText className="text-teal-600 dark:text-teal-400" size={20} />
                                    <span className="font-bold text-sm text-teal-950 dark:text-teal-200">
                                        برات الکترونیک سامانه سپام بانک مرکزی
                                    </span>
                                </div>
                                <span className="text-xs px-2.5 py-1 bg-white dark:bg-gray-900 text-teal-700 dark:text-teal-300 font-bold rounded-lg border border-teal-200">
                                    تسویه براتی بورس
                                </span>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">شناسه سپام / برات الکترونیک</label>
                                    <input
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                        value={petroData.draftBarat?.sepamCode || ''}
                                        onChange={e => updatePetro({
                                            draftBarat: { ...(petroData.draftBarat || { baratNumber: '', sepamCode: '', bankName: '', issueDate: '', dueDate: '', amount: 0, drawerName: '', draweeName: '', status: 'draft' }), sepamCode: e.target.value }
                                        })}
                                        placeholder="کد ۱۶ رقمی سپام..."
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">بانک عامل برات</label>
                                    <input
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                        value={petroData.draftBarat?.bankName || ''}
                                        onChange={e => updatePetro({
                                            draftBarat: { ...(petroData.draftBarat || { baratNumber: '', sepamCode: '', bankName: '', issueDate: '', dueDate: '', amount: 0, drawerName: '', draweeName: '', status: 'draft' }), bankName: e.target.value }
                                        })}
                                        placeholder="بانک عامل..."
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="font-bold text-gray-700 dark:text-gray-300">مبلغ برات (ریال)</label>
                                    <input
                                        type="text"
                                        className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-teal-700 dark:text-teal-300 text-center"
                                        value={formatNumberString(petroData.draftBarat?.amount || petroData.totalInvoiceAmount || 0)}
                                        onChange={e => updatePetro({
                                            draftBarat: { ...(petroData.draftBarat || { baratNumber: '', sepamCode: '', bankName: '', issueDate: '', dueDate: '', amount: 0, drawerName: '', draweeName: '', status: 'draft' }), amount: deformatNumberString(e.target.value) }
                                        })}
                                    />
                                </div>
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* SUB-TAB 3: Loading, Remittance & Logistics */}
            {subTab === 'loading_shipping' && (
                <div className="bg-white dark:bg-gray-800 p-6 rounded-3xl border border-gray-200 dark:border-gray-700 space-y-6 shadow-sm">
                    <div className="flex items-center justify-between border-b border-gray-100 dark:border-gray-700 pb-4">
                        <div>
                            <h3 className="text-base font-bold text-gray-900 dark:text-gray-100">حواله پتروشیمی، بهین‌یاب و اطلاعات ناوگان حمل</h3>
                            <p className="text-xs text-gray-500 mt-0.5">ثبت حواله بارگیری، بارنامه، راننده، پلاک کامیون و کرایه حمل</p>
                        </div>
                        <span className="px-3 py-1 bg-amber-50 text-amber-800 dark:bg-amber-950/40 dark:text-amber-300 text-xs font-bold rounded-xl border border-amber-200">
                            ناوگان و باربری
                        </span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره حواله فروش پتروشیمی</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.loadingNotice?.remittanceNumber || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), remittanceNumber: e.target.value }
                                })}
                                placeholder="شماره حواله..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">کد سهمیه سامانه بهین‌یاب</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.loadingNotice?.behenyabCode || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), behenyabCode: e.target.value }
                                })}
                                placeholder="کد بهین‌یاب..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره بارنامه تمبردار</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.loadingNotice?.waybillNumber || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), waybillNumber: e.target.value }
                                })}
                                placeholder="شماره بارنامه..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">نام راننده</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.loadingNotice?.driverName || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), driverName: e.target.value }
                                })}
                                placeholder="نام و نام خانوادگی راننده..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره تماس راننده</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100 text-left dir-ltr"
                                value={petroData.loadingNotice?.driverPhone || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), driverPhone: e.target.value }
                                })}
                                placeholder="0912..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">پلاک کامیون / تریلی</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100 text-center"
                                value={petroData.loadingNotice?.truckPlate || ''}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), truckPlate: e.target.value }
                                })}
                                placeholder="مثال: 12 ع 345 ایران 68"
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">کرایه حمل و باربری (ریال)</label>
                            <input
                                type="text"
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono font-bold text-emerald-700 dark:text-emerald-400 text-center"
                                value={formatNumberString(petroData.loadingNotice?.freightCostRial || 0)}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), freightCostRial: deformatNumberString(e.target.value) }
                                })}
                                placeholder="مبلغ کرایه حمل..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">تاریخ بارگیری از پتروشیمی</label>
                            <TradeDatePicker
                                value={petroData.loadingNotice?.loadingDate || ''}
                                onChange={d => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || { deliveryStatus: 'pending_loading' }), loadingDate: d }
                                })}
                                placeholder="تاریخ بارگیری..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">وضعیت حمل و ارسال</label>
                            <select
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.loadingNotice?.deliveryStatus || 'pending_loading'}
                                onChange={e => updatePetro({
                                    loadingNotice: { ...(petroData.loadingNotice || {}), deliveryStatus: e.target.value as any }
                                })}
                            >
                                <option value="pending_loading">در انتظار صدور حواله و بارگیری</option>
                                <option value="loaded">بارگیری شده در پتروشیمی</option>
                                <option value="dispatched">در حال حمل به کارخانه (در راه)</option>
                                <option value="delivered_warehouse">تحویل و تخلیه در انبار کارخانه</option>
                            </select>
                        </div>
                    </div>
                </div>
            )}

            {/* SUB-TAB 4: Factory Warehouse Receipt */}
            {subTab === 'warehouse' && (
                <div className="bg-white dark:bg-gray-800 p-6 rounded-3xl border border-gray-200 dark:border-gray-700 space-y-6 shadow-sm">
                    <div className="flex items-center justify-between border-b border-gray-100 dark:border-gray-700 pb-4">
                        <div>
                            <h3 className="text-base font-bold text-gray-900 dark:text-gray-100">رسید انبار کارخانه، توزین باسکول و تحویل کالا</h3>
                            <p className="text-xs text-gray-500 mt-0.5">ثبت اطلاعات ورود بار به کارخانه و مقایسه وزن باسکول با پیش‌فاکتور</p>
                        </div>
                        <span className="px-3 py-1 bg-teal-50 text-teal-700 dark:bg-teal-950/40 dark:text-teal-300 text-xs font-bold rounded-xl border border-teal-200">
                            انبار کارخانه
                        </span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">شماره رسید انبار</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-mono text-gray-900 dark:text-gray-100"
                                value={petroData.warehouseReceipt?.receiptNumber || ''}
                                onChange={e => updatePetro({
                                    warehouseReceipt: { ...(petroData.warehouseReceipt || {}), receiptNumber: e.target.value }
                                })}
                                placeholder="شماره قبض انبار..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">وزن خالص باسکول کارخانه (kg)</label>
                            <input
                                type="number"
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-base font-mono font-bold text-emerald-700 dark:text-emerald-400 text-center"
                                value={petroData.warehouseReceipt?.receivedWeightKg || ''}
                                onChange={e => updatePetro({
                                    warehouseReceipt: { ...(petroData.warehouseReceipt || {}), receivedWeightKg: parseFloat(e.target.value) || 0 }
                                })}
                                placeholder="وزن باسکول..."
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="font-bold text-gray-700 dark:text-gray-300">تاریخ تحویل به انبار</label>
                            <TradeDatePicker
                                value={petroData.warehouseReceipt?.receiptDate || ''}
                                onChange={d => updatePetro({
                                    warehouseReceipt: { ...(petroData.warehouseReceipt || {}), receiptDate: d }
                                })}
                                placeholder="تاریخ تخلیه..."
                            />
                        </div>

                        <div className="space-y-1.5 md:col-span-2">
                            <label className="font-bold text-gray-700 dark:text-gray-300">انبار مقصد</label>
                            <input
                                className="w-full border border-gray-300 dark:border-gray-600 rounded-xl p-3 bg-white dark:bg-gray-900 text-sm font-bold text-gray-900 dark:text-gray-100"
                                value={petroData.warehouseReceipt?.warehouseName || ''}
                                onChange={e => updatePetro({
                                    warehouseReceipt: { ...(petroData.warehouseReceipt || {}), warehouseName: e.target.value }
                                })}
                                placeholder="نام انبار..."
                            />
                        </div>

                        <div className="flex items-center gap-3 pt-6">
                            <label className="flex items-center gap-2 cursor-pointer">
                                <input
                                    type="checkbox"
                                    checked={petroData.warehouseReceipt?.isConfirmed || false}
                                    onChange={e => updatePetro({
                                        warehouseReceipt: { ...(petroData.warehouseReceipt || {}), isConfirmed: e.target.checked }
                                    })}
                                    className="w-5 h-5 rounded text-emerald-600 focus:ring-emerald-500"
                                />
                                <span className="font-bold text-xs text-gray-800 dark:text-gray-200">
                                    تایید نهایی ورود بار به انبار کارخانه
                                </span>
                            </label>
                        </div>
                    </div>

                    {/* Weight Variance Notice */}
                    {weightVarianceKg !== 0 && (
                        <div className={`p-4 rounded-2xl border text-xs flex items-center justify-between ${
                            weightVarianceKg < 0 
                                ? 'bg-amber-50 dark:bg-amber-950/30 border-amber-200 text-amber-900 dark:text-amber-200' 
                                : 'bg-emerald-50 dark:bg-emerald-950/30 border-emerald-200 text-emerald-900 dark:text-emerald-200'
                        }`}>
                            <div className="flex items-center gap-2">
                                <Scale size={18} />
                                <span className="font-bold">
                                    {weightVarianceKg < 0 ? 'کسری وزن باسکول نسبت به پیش‌فاکتور:' : 'اضافه وزن باسکول نسبت به پیش‌فاکتور:'}
                                </span>
                            </div>
                            <span className="font-mono font-black text-sm">
                                {Math.abs(weightVarianceKg).toLocaleString('fa-IR')} کیلوگرم
                            </span>
                        </div>
                    )}
                </div>
            )}

            {/* SUB-TAB 5: Costing Analysis */}
            {subTab === 'costing' && (
                <div className="bg-white dark:bg-gray-800 p-6 rounded-3xl border border-gray-200 dark:border-gray-700 space-y-6 shadow-sm">
                    <div className="flex items-center justify-between border-b border-gray-100 dark:border-gray-700 pb-4">
                        <div>
                            <h3 className="text-base font-bold text-gray-900 dark:text-gray-100">تحلیل و محاسبه بهای تمام‌شده خرید پتروشیمی (Cost Per Kg)</h3>
                            <p className="text-xs text-gray-500 mt-0.5">تفکیک هزینه‌های خرید، ارزش افزوده، کارمزد بورس، کرایه حمل و قیمت تمام‌شده</p>
                        </div>
                        <span className="px-3 py-1 bg-rose-50 text-rose-700 dark:bg-rose-950/40 dark:text-rose-300 text-xs font-bold rounded-xl border border-rose-200">
                            آنالیز بهای تمام‌شده
                        </span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                        <div className="p-4 bg-slate-50 dark:bg-gray-900/40 rounded-2xl border border-slate-200 dark:border-gray-700">
                            <span className="text-xs text-gray-500 block mb-1">ارزش خرید کالا</span>
                            <span className="font-mono text-lg font-black text-gray-900 dark:text-gray-100">
                                {(petroData.totalGoodsPrice || 0).toLocaleString('fa-IR')} <span className="text-xs font-normal">ریال</span>
                            </span>
                        </div>

                        <div className="p-4 bg-slate-50 dark:bg-gray-900/40 rounded-2xl border border-slate-200 dark:border-gray-700">
                            <span className="text-xs text-gray-500 block mb-1">مالیات و کارمزد بورس</span>
                            <span className="font-mono text-lg font-black text-gray-900 dark:text-gray-100">
                                {((petroData.vatAmount || 0) + (petroData.brokerageFee || 0)).toLocaleString('fa-IR')} <span className="text-xs font-normal">ریال</span>
                            </span>
                        </div>

                        <div className="p-4 bg-slate-50 dark:bg-gray-900/40 rounded-2xl border border-slate-200 dark:border-gray-700">
                            <span className="text-xs text-gray-500 block mb-1">کرایه حمل ناوگان</span>
                            <span className="font-mono text-lg font-black text-gray-900 dark:text-gray-100">
                                {totalFreight.toLocaleString('fa-IR')} <span className="text-xs font-normal">ریال</span>
                            </span>
                        </div>

                        <div className="p-4 bg-emerald-50 dark:bg-emerald-950/40 rounded-2xl border border-emerald-200 dark:border-emerald-800">
                            <span className="text-xs text-emerald-800 dark:text-emerald-300 font-bold block mb-1">بهای تمام‌شده هر کیلوگرم</span>
                            <span className="font-mono text-xl font-black text-emerald-900 dark:text-emerald-100">
                                {costPerKg.toLocaleString('fa-IR')} <span className="text-xs font-normal">ریال/kg</span>
                            </span>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};
export default DomesticPetrochemicalTab;
