import {Fragment, useEffect, useMemo, useState} from 'react';
import {
    CaretDownIcon,
    CaretRightIcon,
    CheckCircleIcon,
    MicrosoftExcelLogoIcon,
    WarningCircleIcon,
    XIcon,
} from '@phosphor-icons/react';
import {FieldType} from '@airtable/blocks/interface/models';
import {
    canUpdateCampusField,
    formatMw,
    getFieldType,
    getSelectChoices,
    isWritableCampusField,
    MVA_THRESHOLD,
    readFormValueForField,
    writeValueForField,
} from './utils';

export function SetupPanel({missing}) {
    return (
        <div className="w-full h-full min-h-screen flex items-center justify-center p-6 bg-white dark:bg-gray-gray800">
            <div className="max-w-lg w-full rounded-lg border border-orange-orangeLight1 bg-white dark:bg-gray-gray700 p-6 shadow-sm">
                <div className="flex items-start gap-3">
                    <WarningCircleIcon
                        size={28}
                        className="text-orange-orange shrink-0 mt-0.5"
                        weight="fill"
                    />
                    <div>
                        <h1 className="text-lg font-semibold text-gray-gray800 dark:text-gray-gray100">
                            Configure STACK Land Book
                        </h1>
                        <p className="mt-2 text-sm text-gray-gray600 dark:text-gray-gray300">
                            Open the properties panel and set the campus table, the Load Ramp
                            table, and the Date field on Load Ramp. Quarter and year columns
                            are derived from that Date. Gross MW / MVA is the sum of those
                            incremental loads (including long-term delivery past the horizon).
                            Choose Director or Enterprise reporting, map Campus and Supplier
                            Building Code, the verification Status field (Verified / Not
                            Verified), and for Director view map the Director lookup on Load
                            Ramp. Enable any formula or lookup fields in Interface Data first
                            so they appear in the dropdowns.
                        </p>
                        {missing?.length ? (
                            <>
                                <p className="mt-4 text-sm font-medium text-gray-gray700 dark:text-gray-gray200">
                                    Still missing:
                                </p>
                                <ul className="mt-2 space-y-1.5 text-sm text-gray-gray700 dark:text-gray-gray200 list-disc list-inside">
                                    {missing.map((item) => (
                                        <li key={item}>{item}</li>
                                    ))}
                                </ul>
                            </>
                        ) : null}
                    </div>
                </div>
            </div>
        </div>
    );
}

export function DashboardHeader({
    logoSrc,
    grossMw,
    pipelineMw,
    campusCount,
    view,
    timelineStart,
    onTimelineStartChange,
    horizonLabel,
    isDownloading = false,
    onDownloadExcel,
}) {
    const isEnterprise = view !== 'director';
    return (
        <header className="mb-6 w-full max-w-full">
            <div className="flex items-start justify-between gap-4 w-full max-w-full">
                <div className="min-w-0">
                    <h1 className="text-3xl sm:text-4xl font-display font-bold text-blue-blue dark:text-blue-blueLight1 tracking-tight">
                        STACK LAND BOOK
                    </h1>
                    <div className="mt-2 h-0.5 w-full max-w-3xl bg-blue-blue" />
                    <div className="mt-3">
                        {isEnterprise ? (
                            <>
                                <p className="text-[16px] font-semibold text-blue-blueDark1 dark:text-blue-blueLight1">
                                    Enterprise reporting
                                </p>
                                <p className="mt-0.5 text-[13px] text-gray-gray600 dark:text-gray-gray300">
                                    Only Load Ramp records with status Verified are included in this
                                    view.
                                </p>
                            </>
                        ) : null}
                        <p
                            className={`${isEnterprise ? 'mt-1.5' : ''} text-[13px] text-gray-gray600 dark:text-gray-gray300`}
                        >
                            Quarterly values represent incremental MW additions.
                        </p>
                    </div>
                </div>
                <div className="flex flex-wrap items-center gap-3 justify-end shrink-0">
                    <DownloadExcelButton
                        isDownloading={isDownloading}
                        onClick={onDownloadExcel}
                    />
                    {logoSrc ? (
                        <div className="h-12 sm:h-14 w-[160px] sm:w-[180px] shrink-0 flex items-center justify-end">
                            <img
                                src={logoSrc}
                                alt="STACK Infrastructure"
                                className="max-h-full max-w-full h-12 sm:h-14 w-auto object-contain"
                                style={{maxHeight: 56, maxWidth: 180, height: 48, width: 'auto'}}
                            />
                        </div>
                    ) : null}
                </div>
            </div>

            <div className="mt-4 flex flex-wrap items-stretch gap-3">
                <SummaryCard
                    label="Gross MW"
                    value={formatMw(grossMw, {suffix: true, empty: '—'})}
                />
                <SummaryCard
                    label="Pipeline MW"
                    value={formatMw(pipelineMw, {suffix: true, empty: '—'})}
                />
                <SummaryCard label="Campuses" value={campusCount} />
                <label className="inline-flex rounded-lg border border-gray-gray200 dark:border-gray-gray600 bg-white dark:bg-gray-gray700 px-4 py-3 flex-col gap-1 min-w-[180px] shadow-md justify-center">
                    <span className="text-sm font-semibold text-gray-gray700 dark:text-gray-gray200">
                        Select Start Month
                    </span>
                    <input
                        type="month"
                        value={timelineStart || ''}
                        onChange={(event) => onTimelineStartChange?.(event.target.value)}
                        className="text-sm font-semibold text-gray-gray800 dark:text-gray-gray100 bg-transparent outline-none"
                    />
                    {horizonLabel ? (
                        <span className="text-xs text-gray-gray500 dark:text-gray-gray400">
                            {horizonLabel}
                        </span>
                    ) : null}
                </label>
            </div>
        </header>
    );
}

function SummaryCard({label, value}) {
    return (
        <div className="rounded-lg border border-blue-blueLight1 bg-blue-blueLight3 dark:bg-blue-blueDark1/40 dark:border-blue-blueLight1/40 px-4 py-3 min-w-[140px]">
            <div className="text-sm font-medium text-blue-blueDark1 dark:text-blue-blueLight1">
                {label}
            </div>
            <div className="text-2xl font-semibold tabular-nums text-blue-blueDark1 dark:text-blue-blueLight2">
                {value}
            </div>
        </div>
    );
}

export function ViewBanner({
    view,
    userName,
    pendingCount = 0,
    canAutoUnverify,
    canVerify,
    isVerifying = false,
    onVerifyAll,
}) {
    const isDirector = view === 'director';
    if (!isDirector) return null;
    return (
        <div className="mb-4 rounded-lg border border-blue-blueLight1 bg-blue-blueLight3 dark:bg-blue-blueDark1/30 dark:border-blue-blueLight1/40 px-4 py-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
                <p className="text-[14px] font-semibold text-blue-blueDark1 dark:text-blue-blueLight1">
                    Director view
                </p>
                <div className="flex flex-wrap items-center gap-2">
                    {pendingCount > 0 ? (
                        <span className="rounded-full bg-white dark:bg-gray-gray700 px-2.5 py-0.5 text-[12px] font-semibold text-blue-blueDark1 dark:text-blue-blueLight1">
                            {pendingCount} pending verification
                        </span>
                    ) : null}
                    {canVerify ? (
                        <VerifyButton
                            label={isVerifying ? 'Verifying…' : 'Verify'}
                            disabled={isVerifying || pendingCount === 0}
                            onClick={onVerifyAll}
                        />
                    ) : null}
                </div>
            </div>
            <p className="mt-1 text-[13px] text-gray-gray700 dark:text-gray-gray200">
                {`Showing Load Ramp for campuses assigned to ${
                    userName || 'you'
                }. Verify marks all of your Load Ramp records as Verified.`}
            </p>
            {isDirector && !canVerify ? (
                <p className="mt-1 text-[13px] text-orange-orange">
                    Verify is unavailable until the Status field has a Verified choice and you
                    have permission to edit Load Ramp.
                </p>
            ) : null}
            {!canAutoUnverify ? (
                <p className="mt-1 text-[13px] text-orange-orange">
                    Map Verified At and Last Modified Time in the properties panel so edits
                    after verification automatically return status to Not Verified.
                </p>
            ) : null}
        </div>
    );
}

export function MvaTabs({
    activeTab,
    onChange,
    overCount = 0,
    underCount = 0,
}) {
    function tabClass(isActive) {
        return `rounded-t-md border border-b-0 px-5 py-3 text-[15px] font-semibold transition-colors ${
            isActive
                ? 'border-blue-blue bg-white dark:bg-gray-gray700 text-blue-blue dark:text-blue-blueLight1'
                : 'border-transparent bg-gray-gray100 dark:bg-gray-gray600 text-gray-gray600 dark:text-gray-gray300 hover:bg-gray-gray50 dark:hover:bg-gray-gray500'
        }`;
    }

    return (
        <div className="mb-4">
            <div className="flex flex-wrap gap-1 border-b border-gray-gray200 dark:border-gray-gray600">
                <button
                    type="button"
                    onClick={() => onChange('over')}
                    className={tabClass(activeTab === 'over')}
                >
                    Over {MVA_THRESHOLD} MVA
                    <span className="ml-2 tabular-nums text-[13px] font-bold opacity-80">
                        ({overCount})
                    </span>
                </button>
                <button
                    type="button"
                    onClick={() => onChange('under')}
                    className={tabClass(activeTab === 'under')}
                >
                    Under {MVA_THRESHOLD} MVA
                    <span className="ml-2 tabular-nums text-[13px] font-bold opacity-80">
                        ({underCount})
                    </span>
                </button>
            </div>
        </div>
    );
}

export function DownloadExcelButton({disabled, isDownloading, onClick}) {
    return (
        <button
            type="button"
            disabled={disabled || isDownloading}
            onClick={onClick}
            className="inline-flex items-center gap-2 rounded-lg px-4 py-2.5 text-sm font-semibold transition-colors shadow-sm border border-blue-blueLight1 bg-white text-blue-blue hover:bg-blue-blueLight3 dark:bg-gray-gray700 dark:text-blue-blueLight1 dark:hover:bg-gray-gray600 disabled:opacity-50 disabled:cursor-not-allowed"
        >
            <MicrosoftExcelLogoIcon size={18} weight="bold" className="text-green-greenDark1" />
            {isDownloading ? 'Preparing…' : 'Download Excel'}
        </button>
    );
}

const COL = {
    expand: 44,
    campus: 176,
    metro: 108,
    supplier: 160,
    land: 124,
    rezoning: 132,
    water: 132,
    power: 156,
    gross: 96,
    live: 88,
    pipeline: 108,
    quarter: 56,
    longTerm: 124,
};

function colStyle(width) {
    return {width};
}

function WrappedLabel({lines}) {
    return lines.map((line) => (
        <span key={line} className="block">
            {line}
        </span>
    ));
}

function HeaderCell({children, className = '', colSpan, rowSpan, title, width, style}) {
    const label = title || (typeof children === 'string' ? children : undefined);
    return (
        <th
            colSpan={colSpan}
            rowSpan={rowSpan}
            title={label}
            style={{...(width ? colStyle(width) : null), ...style}}
            className={`px-2 py-2.5 text-[12px] leading-snug font-semibold uppercase text-white text-center align-middle box-border ${className}`}
        >
            <span className="block whitespace-normal break-words">{children}</span>
        </th>
    );
}

function Cell({children, className = '', title, colSpan, width, style}) {
    return (
        <td
            colSpan={colSpan}
            title={title}
            style={{...(width ? colStyle(width) : null), ...style}}
            className={`px-2 py-2.5 text-[14px] leading-snug align-middle text-center border-b border-gray-gray200 dark:border-gray-gray600 ${className}`}
        >
            {children}
        </td>
    );
}

function StatusBadge({label, colorHex}) {
    if (!label) return <span className="text-gray-gray400">—</span>;
    if (!colorHex) return <span className="whitespace-normal break-words">{label}</span>;
    return (
        <span
            className="inline-flex max-w-full items-center rounded-sm px-1.5 py-0.5 text-[12px] font-semibold leading-tight whitespace-normal break-words"
            style={{backgroundColor: `${colorHex}22`, color: colorHex}}
        >
            {label}
        </span>
    );
}

function VerifyButton({onClick, disabled, label = 'Verify'}) {
    return (
        <button
            type="button"
            onClick={onClick}
            disabled={disabled}
            className="inline-flex items-center gap-1 rounded-md bg-blue-blue px-2.5 py-1 text-[13px] font-semibold text-white hover:bg-blue-blueDark1 disabled:cursor-not-allowed disabled:opacity-50"
        >
            <CheckCircleIcon size={14} weight="bold" />
            {label}
        </button>
    );
}

function RampEntriesPanel({
    row,
    canExpandRamp,
    onRampClick,
}) {
    const entries = row.rampEntries || [];

    return (
        <div className="px-3 py-3 text-left">
            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                <p className="text-[12px] font-semibold uppercase tracking-wide text-gray-gray600 dark:text-gray-gray300">
                    Load Ramp entries
                </p>
            </div>
            {entries.length === 0 ? (
                <p className="text-[14px] text-gray-gray500">No Load Ramp records for this campus.</p>
            ) : (
                <table className="w-full max-w-3xl border-collapse text-[14px]">
                    <thead>
                        <tr className="text-left text-[12px] uppercase text-gray-gray500 dark:text-gray-gray400">
                            <th className="px-2 py-1 font-semibold">Date</th>
                            <th className="px-2 py-1 font-semibold">MW</th>
                            <th className="px-2 py-1 font-semibold">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {entries.map((entry) => (
                            <tr
                                key={entry.id}
                                className="border-t border-gray-gray200 dark:border-gray-gray600"
                            >
                                <td className="px-2 py-1.5">
                                    {canExpandRamp && onRampClick && entry.record ? (
                                        <button
                                            type="button"
                                            onClick={() => onRampClick(entry.record)}
                                            className="text-blue-blue dark:text-blue-blueLight1 underline-offset-2 hover:underline"
                                        >
                                            {entry.dateLabel || '—'}
                                        </button>
                                    ) : (
                                        entry.dateLabel || '—'
                                    )}
                                </td>
                                <td className="px-2 py-1.5 tabular-nums">
                                    {formatMw(entry.mw)}
                                </td>
                                <td className="px-2 py-1.5">
                                    {entry.verified ? (
                                        <span className="font-semibold text-green-green">
                                            Verified
                                        </span>
                                    ) : (
                                        <span className="text-gray-gray600 dark:text-gray-gray300">
                                            {entry.statusLabel || 'Not Verified'}
                                        </span>
                                    )}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            )}
        </div>
    );
}

export function LandBankTable({
    rows,
    years,
    canExpandRamp,
    onCampusClick,
    onRampClick,
    emptyMessage,
    view = 'enterprise',
}) {
    const [expandedIds, setExpandedIds] = useState(() => new Set());
    const showDetails = view === 'director' || view === 'enterprise';
    const quarterCols = years.flatMap((y) => y.quarters);
    const leftColCount = showDetails ? 11 : 10;
    const colCount = leftColCount + quarterCols.length + 1;

    const totals = {
        grossMw: sum(rows, (r) => r.grossMw),
        liveMw: sum(rows, (r) => r.liveMw),
        pipelineMw: sum(rows, (r) => r.pipelineMw),
        longTermMw: sum(rows, (r) => r.longTermMw),
        quarters: {},
    };
    for (const q of quarterCols) {
        totals.quarters[q.key] = sum(rows, (r) => r.quarters?.[q.key]);
    }

    function toggleExpanded(id) {
        setExpandedIds((current) => {
            const next = new Set(current);
            if (next.has(id)) next.delete(id);
            else next.add(id);
            return next;
        });
    }

    return (
        <div className="rounded-lg border border-gray-gray200 dark:border-gray-gray600 bg-white dark:bg-gray-gray700 shadow-sm">
            <table className="border-collapse table-fixed text-[14px] w-full">
                <colgroup>
                    {showDetails ? <col style={colStyle(COL.expand)} /> : null}
                    <col style={colStyle(COL.campus)} />
                    <col style={colStyle(COL.metro)} />
                    <col style={colStyle(COL.supplier)} />
                    <col style={colStyle(COL.land)} />
                    <col style={colStyle(COL.rezoning)} />
                    <col style={colStyle(COL.water)} />
                    <col style={colStyle(COL.power)} />
                    <col style={colStyle(COL.gross)} />
                    <col style={colStyle(COL.live)} />
                    <col style={colStyle(COL.pipeline)} />
                    {quarterCols.map((q) => (
                        <col key={q.key} style={colStyle(COL.quarter)} />
                    ))}
                    <col style={colStyle(COL.longTerm)} />
                </colgroup>
                <thead>
                    <tr className="bg-blue-blueDark1">
                        {showDetails ? (
                            <HeaderCell
                                rowSpan={2}
                                width={COL.expand}
                                title="Load Ramp details"
                                className="border-b border-blue-blue sticky left-0 z-20 bg-blue-blueDark1"
                            >
                                {' '}
                            </HeaderCell>
                        ) : null}
                        <HeaderCell
                            rowSpan={2}
                            width={COL.campus}
                            title="Campus"
                            className={`border-b border-blue-blue sticky z-20 bg-blue-blueDark1 ${
                                showDetails ? '' : 'left-0'
                            }`}
                            style={showDetails ? {left: COL.expand} : undefined}
                        >
                            Campus
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.metro}
                            title="Metro"
                            className="border-b border-blue-blue"
                        >
                            Metro
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.supplier}
                            title="Supplier Building Code"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Supplier', 'Building', 'Code']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.land}
                            title="Land Status"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Land', 'Status']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.rezoning}
                            title="Rezoning Status"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Rezoning', 'Status']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.water}
                            title="Water Capacity"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Water', 'Capacity']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.power}
                            title="Power Confidence Level"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Power', 'Confidence', 'Level']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.gross}
                            title="Gross MW"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Gross', 'MW']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.live}
                            title="Live MW"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Live', 'MW']} />
                        </HeaderCell>
                        <HeaderCell
                            rowSpan={2}
                            width={COL.pipeline}
                            title="Pipeline MW"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Pipeline', 'MW']} />
                        </HeaderCell>
                        {years.map((y) => (
                            <HeaderCell
                                key={y.year}
                                colSpan={y.quarters.length}
                                className="border-b border-blue-blue border-l border-blue-blue/40"
                            >
                                {y.label}
                            </HeaderCell>
                        ))}
                        <HeaderCell
                            rowSpan={2}
                            width={COL.longTerm}
                            title="Long-term Delivery"
                            className="border-b border-blue-blue"
                        >
                            <WrappedLabel lines={['Long-term', 'Delivery']} />
                        </HeaderCell>
                    </tr>
                    <tr className="bg-blue-blue">
                        {years.map((y) =>
                            y.quarters.map((q) => (
                                <HeaderCell
                                    key={q.key}
                                    className="border-b border-blue-blueDark1 border-l border-blue-blueDark1/40"
                                >
                                    {q.label}
                                </HeaderCell>
                            )),
                        )}
                    </tr>
                </thead>
                <tbody>
                    {rows.length === 0 ? (
                        <tr>
                            <td
                                colSpan={colCount}
                                className="px-4 py-8 text-center text-base text-gray-gray500 dark:text-gray-gray400"
                            >
                                {emptyMessage}
                            </td>
                        </tr>
                    ) : (
                        rows.map((row) => {
                            const expanded = expandedIds.has(row.id);
                            return (
                                <Fragment key={row.id}>
                                    <tr
                                        key={row.id}
                                        className="group hover:bg-blue-blueLight3/40 dark:hover:bg-gray-gray600/40"
                                    >
                                        {showDetails ? (
                                            <Cell
                                                width={COL.expand}
                                                className="sticky left-0 z-10 bg-white dark:bg-gray-gray700 group-hover:bg-blue-blueLight3/40 dark:group-hover:bg-gray-gray600/40"
                                            >
                                                <button
                                                    type="button"
                                                    onClick={() => toggleExpanded(row.id)}
                                                    className="inline-flex items-center justify-center rounded p-1 text-blue-blue dark:text-blue-blueLight1 hover:bg-blue-blueLight3"
                                                    title={
                                                        expanded
                                                            ? 'Hide Load Ramp entries'
                                                            : 'Show Load Ramp entries'
                                                    }
                                                >
                                                    {expanded ? (
                                                        <CaretDownIcon size={16} weight="bold" />
                                                    ) : (
                                                        <CaretRightIcon size={16} weight="bold" />
                                                    )}
                                                </button>
                                            </Cell>
                                        ) : null}
                                        <Cell
                                            width={COL.campus}
                                            className={`sticky z-10 bg-white dark:bg-gray-gray700 group-hover:bg-blue-blueLight3/40 dark:group-hover:bg-gray-gray600/40 text-left ${
                                                showDetails ? '' : 'left-0'
                                            }`}
                                            style={showDetails ? {left: COL.expand} : undefined}
                                        >
                                            {onCampusClick && row.record ? (
                                                <button
                                                    type="button"
                                                    onClick={() => onCampusClick(row.record)}
                                                    className="text-blue-blue dark:text-blue-blueLight1 underline-offset-2 hover:underline text-left whitespace-normal break-words [overflow-wrap:anywhere]"
                                                    title="Edit campus details"
                                                >
                                                    {row.campus || '—'}
                                                </button>
                                            ) : (
                                                <span className="text-gray-gray800 dark:text-gray-gray100 whitespace-normal break-words [overflow-wrap:anywhere]">
                                                    {row.campus || '—'}
                                                </span>
                                            )}
                                            {view === 'director' && row.pendingCount > 0 ? (
                                                <span className="mt-0.5 block text-[12px] font-medium text-orange-orange">
                                                    {row.pendingCount} to verify
                                                </span>
                                            ) : null}
                                        </Cell>
                                        <Cell
                                            width={COL.metro}
                                            className="text-gray-gray800 dark:text-gray-gray100 font-medium"
                                        >
                                            {row.metro || '—'}
                                        </Cell>
                                        <Cell
                                            width={COL.supplier}
                                            className="text-left whitespace-normal break-words [overflow-wrap:anywhere]"
                                        >
                                            {row.supplierBuildingCode || '—'}
                                        </Cell>
                                        <Cell>
                                            <StatusBadge
                                                label={row.landStatus}
                                                colorHex={row.landStatusColor}
                                            />
                                        </Cell>
                                        <Cell>
                                            <StatusBadge
                                                label={row.rezoningStatus}
                                                colorHex={row.rezoningStatusColor}
                                            />
                                        </Cell>
                                        <Cell className="whitespace-normal break-words">
                                            {row.waterCapacity || '—'}
                                        </Cell>
                                        <Cell className="whitespace-normal break-words text-left">
                                            <StatusBadge
                                                label={row.powerConfidence}
                                                colorHex={row.powerConfidenceColor}
                                            />
                                        </Cell>
                                        <Cell className="tabular-nums font-semibold whitespace-nowrap">
                                            {formatMw(row.grossMw)}
                                        </Cell>
                                        <Cell className="tabular-nums whitespace-nowrap">
                                            {formatMw(row.liveMw)}
                                        </Cell>
                                        <Cell className="tabular-nums whitespace-nowrap">
                                            {formatMw(row.pipelineMw)}
                                        </Cell>
                                        {quarterCols.map((q) => (
                                            <Cell
                                                key={q.key}
                                                className="tabular-nums whitespace-nowrap border-l border-gray-gray100 dark:border-gray-gray600"
                                            >
                                                {formatMw(row.quarters?.[q.key], {empty: ''})}
                                            </Cell>
                                        ))}
                                        <Cell className="tabular-nums font-medium whitespace-nowrap">
                                            {formatMw(row.longTermMw, {empty: ''})}
                                        </Cell>
                                    </tr>
                                    {showDetails && expanded ? (
                                        <tr key={`${row.id}-details`}>
                                            <td
                                                colSpan={colCount}
                                                className="bg-blue-blueLight3/50 dark:bg-gray-gray800 border-b border-gray-gray200 dark:border-gray-gray600"
                                            >
                                                <RampEntriesPanel
                                                    row={row}
                                                    canExpandRamp={canExpandRamp}
                                                    onRampClick={onRampClick}
                                                />
                                            </td>
                                        </tr>
                                    ) : null}
                                </Fragment>
                            );
                        })
                    )}
                </tbody>
                {rows.length > 0 ? (
                    <tfoot>
                        <tr className="bg-gray-gray50 dark:bg-gray-gray800 font-semibold">
                            <Cell
                                colSpan={showDetails ? 8 : 7}
                                className="text-right font-semibold text-gray-gray800 dark:text-gray-gray100"
                            >
                                Total:
                            </Cell>
                            <Cell className="tabular-nums font-semibold whitespace-nowrap">
                                {formatMw(totals.grossMw)}
                            </Cell>
                            <Cell className="tabular-nums whitespace-nowrap">
                                {formatMw(totals.liveMw)}
                            </Cell>
                            <Cell className="tabular-nums whitespace-nowrap">
                                {formatMw(totals.pipelineMw)}
                            </Cell>
                            {quarterCols.map((q) => (
                                <Cell
                                    key={q.key}
                                    className="tabular-nums whitespace-nowrap border-l border-gray-gray100 dark:border-gray-gray600"
                                >
                                    {formatMw(totals.quarters[q.key], {empty: '—'})}
                                </Cell>
                            ))}
                            <Cell className="tabular-nums font-semibold whitespace-nowrap">
                                {formatMw(totals.longTermMw, {empty: '—'})}
                            </Cell>
                        </tr>
                    </tfoot>
                ) : null}
            </table>
        </div>
    );
}

function sum(rows, getter) {
    let total = 0;
    let hasValue = false;
    for (const row of rows) {
        const value = getter(row);
        if (value == null || value === '') continue;
        total += Number(value) || 0;
        hasValue = true;
    }
    return hasValue ? total : null;
}

const FIELD_INPUT_CLASS =
    'mt-1 w-full rounded-md border border-gray-gray200 dark:border-gray-gray600 bg-white dark:bg-gray-gray800 px-3 py-2 text-sm text-gray-gray800 dark:text-gray-gray100 outline-none focus:border-blue-blue disabled:opacity-60 disabled:cursor-not-allowed';

function CampusFieldInput({field, value, onChange, disabled}) {
    if (!field) return null;
    const type = getFieldType(field);
    const choices = getSelectChoices(field);

    if (type === FieldType.SINGLE_SELECT) {
        const hasCurrent =
            !value || choices.some((choice) => choice.id === value || choice.name === value);
        return (
            <select
                value={value || ''}
                onChange={(event) => onChange(event.target.value)}
                disabled={disabled}
                className={FIELD_INPUT_CLASS}
            >
                <option value="">Select…</option>
                {hasCurrent ? null : <option value={value}>{value}</option>}
                {choices.map((choice) => (
                    <option key={choice.id || choice.name} value={choice.id || choice.name}>
                        {choice.name}
                    </option>
                ))}
            </select>
        );
    }

    if (type === FieldType.MULTIPLE_SELECTS) {
        const selected = Array.isArray(value) ? value : [];
        return (
            <div className="mt-1 max-h-40 overflow-auto rounded-md border border-gray-gray200 dark:border-gray-gray600 bg-white dark:bg-gray-gray800 px-3 py-2">
                {choices.length === 0 ? (
                    <p className="text-sm text-gray-gray500 dark:text-gray-gray400">No options</p>
                ) : (
                    choices.map((choice) => {
                        const choiceValue = choice.id || choice.name;
                        const checked = selected.includes(choice.id) || selected.includes(choice.name);
                        return (
                            <label
                                key={choiceValue}
                                className="flex items-center gap-2 py-1 text-sm text-gray-gray800 dark:text-gray-gray100"
                            >
                                <input
                                    type="checkbox"
                                    checked={checked}
                                    disabled={disabled}
                                    onChange={() => {
                                        if (checked) {
                                            onChange(
                                                selected.filter(
                                                    (item) => item !== choice.id && item !== choice.name,
                                                ),
                                            );
                                        } else {
                                            onChange([...selected, choiceValue]);
                                        }
                                    }}
                                />
                                {choice.name}
                            </label>
                        );
                    })
                )}
            </div>
        );
    }

    if (type === FieldType.MULTILINE_TEXT || type === FieldType.RICH_TEXT) {
        return (
            <textarea
                rows={3}
                value={value || ''}
                onChange={(event) => onChange(event.target.value)}
                disabled={disabled}
                className={FIELD_INPUT_CLASS}
            />
        );
    }

    return (
        <input
            type="text"
            value={value || ''}
            onChange={(event) => onChange(event.target.value)}
            disabled={disabled}
            className={FIELD_INPUT_CLASS}
        />
    );
}

export function CampusEditForm({
    record,
    table,
    campusName,
    metroField,
    supplierField,
    landStatusField,
    rezoningStatusField,
    onClose,
}) {
    const fields = useMemo(
        () =>
            [
                {key: 'metro', label: 'Metro', field: metroField},
                {key: 'supplier', label: 'Supplier Building Code', field: supplierField},
                {key: 'landStatus', label: 'Land Status', field: landStatusField},
                {key: 'rezoningStatus', label: 'Rezoning Status', field: rezoningStatusField},
            ].filter((item) => item.field),
        [metroField, supplierField, landStatusField, rezoningStatusField],
    );

    const [values, setValues] = useState(() => {
        const next = {};
        for (const item of fields) {
            next[item.key] = readFormValueForField(record, item.field);
        }
        return next;
    });
    const [isSaving, setIsSaving] = useState(false);
    const [error, setError] = useState('');

    const editableFields = fields.filter((item) =>
        canUpdateCampusField(table, record, item.field),
    );
    const canSave = editableFields.length > 0 && !isSaving;

    useEffect(() => {
        function onKeyDown(event) {
            if (event.key === 'Escape' && !isSaving) onClose?.();
        }
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [isSaving, onClose]);

    async function handleSave(event) {
        event.preventDefault();
        if (!canSave || !table || !record) return;
        const updates = {};
        for (const item of editableFields) {
            updates[item.field.id] = writeValueForField(item.field, values[item.key]);
        }
        if (!Object.keys(updates).length) return;
        setIsSaving(true);
        setError('');
        try {
            await table.updateRecordAsync(record, updates);
            onClose?.();
        } catch (saveError) {
            setError(saveError?.message || 'Could not save campus details.');
            setIsSaving(false);
        }
    }

    return (
        <div
            className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-gray-gray800/50"
            onClick={isSaving ? undefined : onClose}
            role="presentation"
        >
            <div
                role="dialog"
                aria-modal="true"
                aria-labelledby="campus-edit-title"
                className="w-full max-w-md rounded-lg border border-gray-gray200 dark:border-gray-gray600 bg-white dark:bg-gray-gray700 shadow-lg"
                onClick={(event) => event.stopPropagation()}
            >
                <form onSubmit={handleSave}>
                    <div className="flex items-start justify-between gap-3 border-b border-gray-gray200 dark:border-gray-gray600 px-5 py-4">
                        <div className="min-w-0">
                            <h2
                                id="campus-edit-title"
                                className="text-lg font-semibold text-gray-gray800 dark:text-gray-gray100"
                            >
                                Edit campus
                            </h2>
                            <p className="mt-0.5 text-sm text-gray-gray600 dark:text-gray-gray300 truncate">
                                {campusName || 'Campus'}
                            </p>
                        </div>
                        <button
                            type="button"
                            onClick={onClose}
                            disabled={isSaving}
                            className="rounded p-1 text-gray-gray500 hover:bg-gray-gray100 dark:hover:bg-gray-gray600 disabled:opacity-50"
                            title="Close"
                        >
                            <XIcon size={18} weight="bold" />
                        </button>
                    </div>

                    <div className="px-5 py-4 space-y-4">
                        {fields.length === 0 ? (
                            <p className="text-sm text-gray-gray600 dark:text-gray-gray300">
                                These campus fields are not mapped yet. Open the properties panel
                                to set Metro, Supplier Building Code, Land Status, and Rezoning
                                Status.
                            </p>
                        ) : (
                            fields.map((item) => {
                                const canEdit = canUpdateCampusField(table, record, item.field);
                                return (
                                    <label key={item.key} className="block">
                                        <span className="text-sm font-semibold text-gray-gray700 dark:text-gray-gray200">
                                            {item.label}
                                        </span>
                                        <CampusFieldInput
                                            field={item.field}
                                            value={values[item.key]}
                                            disabled={!canEdit || isSaving}
                                            onChange={(nextValue) =>
                                                setValues((current) => ({
                                                    ...current,
                                                    [item.key]: nextValue,
                                                }))
                                            }
                                        />
                                        {!isWritableCampusField(item.field) ? (
                                            <span className="mt-1 block text-xs text-gray-gray500 dark:text-gray-gray400">
                                                This field cannot be edited here.
                                            </span>
                                        ) : null}
                                    </label>
                                );
                            })
                        )}
                        {fields.length > 0 && editableFields.length === 0 ? (
                            <p className="text-sm text-orange-orange">
                                You do not have permission to edit these campus fields.
                            </p>
                        ) : null}
                        {error ? (
                            <p className="text-sm text-red-red">{error}</p>
                        ) : null}
                    </div>

                    <div className="flex items-center justify-end gap-2 border-t border-gray-gray200 dark:border-gray-gray600 px-5 py-4">
                        <button
                            type="button"
                            onClick={onClose}
                            disabled={isSaving}
                            className="rounded-md border border-gray-gray200 dark:border-gray-gray600 px-3 py-2 text-sm font-semibold text-gray-gray700 dark:text-gray-gray200 hover:bg-gray-gray50 dark:hover:bg-gray-gray600 disabled:opacity-50"
                        >
                            Cancel
                        </button>
                        <button
                            type="submit"
                            disabled={!canSave}
                            className="rounded-md bg-blue-blue px-3 py-2 text-sm font-semibold text-white hover:bg-blue-blueDark1 disabled:cursor-not-allowed disabled:opacity-50"
                        >
                            {isSaving ? 'Saving…' : 'Save'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}
