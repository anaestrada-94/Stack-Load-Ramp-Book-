import {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
    initializeBlock,
    useBase,
    useRecords,
    useCustomProperties,
    useSession,
    expandRecord,
    colorUtils,
} from '@airtable/blocks/interface/ui';
import {
    CampusEditForm,
    DashboardHeader,
    LandBankTable,
    MvaTabs,
    SetupPanel,
    ViewBanner,
} from './components';
import {downloadExcel} from './exportDashboard';
import {
    buildQuarterColumns,
    compareCampusRows,
    currentYearMonthString,
    findFieldByExactNames,
    findFieldByNameHints,
    findTableByNameHints,
    formatShortDate,
    getCellTimestamp,
    getDateParts,
    getLinkedRecordIds,
    getSelectColorToken,
    getSelectName,
    isCampusNameField,
    isDateAfterHorizon,
    isDateBeforeStart,
    isDateLikeField,
    isDateTimeField,
    isDirectorLikeField,
    isLastModifiedField,
    isLinkField,
    isNumericLikeField,
    isOver100Mva,
    isRampRecordVerified,
    isSelectField,
    isSingleSelectField,
    isVerifiedStatusLabel,
    horizonRangeLabel,
    lastQuarterOfHorizon,
    parseYearMonth,
    isTextLikeField,
    quarterFromMonth,
    quarterKey,
    readCellDisplayName,
    readNumericOrNull,
    recordAssignedToUser,
    resolveField,
    updateRecordsInChunks,
    unverifiedChoice,
    verifiedChoice,
    writeSelectChoice,
    writeTimestampForField,
} from './utils';
import stackLogo from './assets/stackLogoData.js';
import './style.css';

function fieldProperty(key, label, table, shouldFieldBeAllowed, hints, extra = {}) {
    return {
        key,
        label,
        type: 'field',
        table,
        shouldFieldBeAllowed,
        defaultValue: findFieldByNameHints(table, hints, {
            typePredicate: shouldFieldBeAllowed,
            ...extra,
        }),
    };
}

function getCustomProperties(base) {
    const campusesTable =
        findTableByNameHints(base, [
            '9_campus',
            '9_ campuses',
            'land bank',
            'campuses',
            'campus',
        ]) || base.tables[0];
    const loadRampTable =
        findTableByNameHints(base, [
            'load ramp',
            'load-ramp',
            'loadramp',
            'quarterly mw',
            'mw ramp',
        ]) ||
        base.tables.find((table) => {
            if (!table || table.id === campusesTable?.id) return false;
            const names = table.fields.map((field) => field.name.toLowerCase());
            return names.includes('date') && names.some((name) => name.includes('campus'));
        }) ||
        base.tables.find((table) => table.id !== campusesTable?.id) ||
        null;
    const loadRampFieldsTable = loadRampTable;

    const isText = (field) => isTextLikeField(field) || isSelectField(field);
    const isName = (field) => isCampusNameField(field) || isTextLikeField(field);
    const isNum = (field) => isNumericLikeField(field);
    const isSelectOrText = (field) => isSelectField(field) || isTextLikeField(field);

    return [
        {
            key: 'campusesTable',
            label: 'Campuses / Land Bank Table',
            type: 'table',
            defaultValue: campusesTable,
        },
        {
            key: 'loadRampTable',
            label: 'Load Ramp Table',
            type: 'table',
            defaultValue: loadRampTable,
        },
        {
            key: 'dashboardView',
            label: 'Dashboard view',
            type: 'enum',
            possibleValues: [
                {
                    value: 'enterprise',
                    label: 'Enterprise reporting (verified Load Ramp only)',
                },
                {
                    value: 'director',
                    label: 'Director (my campuses + Verify)',
                },
            ],
            defaultValue: 'enterprise',
        },
        fieldProperty(
            'campusDirectorField',
            'Director / Owner (on Campuses — optional fallback)',
            campusesTable,
            isDirectorLikeField,
            ['director', 'owner', 'project director', 'assigned to', 'lead'],
        ),
        fieldProperty(
            'metroField',
            'Metro',
            campusesTable,
            isText,
            ['metro', 'market', 'region'],
        ),
        {
            key: 'campusField',
            label: 'Campus',
            type: 'field',
            table: campusesTable,
            shouldFieldBeAllowed: isName,
            defaultValue:
                findFieldByExactNames(campusesTable, ['Campus', 'Campus Name']) ||
                findFieldByNameHints(campusesTable, ['campus name', 'campus'], {
                    excludeHints: ['supplier', 'building code'],
                }),
        },
        {
            key: 'campusNameField',
            label: 'Supplier Building Code',
            type: 'field',
            table: campusesTable,
            shouldFieldBeAllowed: isName,
            defaultValue:
                findFieldByExactNames(campusesTable, [
                    'Supplier Building Code',
                    'Building Code',
                ]) ||
                findFieldByNameHints(campusesTable, [
                    'supplier building code',
                    'building code',
                    'supplier code',
                ]),
        },
        fieldProperty(
            'landStatusField',
            'Land Status',
            campusesTable,
            isSelectOrText,
            ['land status', 'site control', 'psa'],
        ),
        fieldProperty(
            'rezoningStatusField',
            'Rezoning Status',
            campusesTable,
            isSelectOrText,
            ['rezoning status', 'rezoning', 'entitlement', 'by-right', 'by right'],
        ),
        fieldProperty(
            'waterCapacityField',
            'Water Capacity If Available',
            campusesTable,
            (field) => isText(field) || isNum(field),
            ['water capacity', 'water'],
            {excludeHints: ['confidence']},
        ),
        fieldProperty(
            'powerConfidenceField',
            'Power Utility Confidence Level',
            campusesTable,
            isSelectOrText,
            [
                'power utility confidence',
                'power confidence',
                'utility confidence',
                'confidence level',
            ],
        ),
        fieldProperty(
            'liveMwField',
            'Current Power — Live MW',
            campusesTable,
            isNum,
            ['live mw', 'current power', 'energized mw'],
        ),
        fieldProperty(
            'pipelineMwField',
            'Forecasted Power — Pipeline MW',
            campusesTable,
            isNum,
            ['pipeline mw', 'forecasted power', 'forecast mw'],
        ),
        ...(loadRampFieldsTable
            ? [
                  {
                      key: 'loadRampCampusField',
                      label: 'Campus (on Load Ramp)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: (field) =>
                          isLinkField(field) || isCampusNameField(field),
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, ['Campus']) ||
                          findFieldByNameHints(loadRampFieldsTable, ['campus']),
                  },
                  {
                      key: 'loadRampMwField',
                      label: 'MW (on Load Ramp)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isNum,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, ['MW', 'MVA']) ||
                          findFieldByNameHints(loadRampFieldsTable, ['mw', 'mva', 'mwg', 'load'], {
                              typePredicate: isNum,
                          }),
                  },
                  {
                      key: 'loadRampDateField',
                      label: 'Date (on Load Ramp — quarter and year are derived from this)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isDateLikeField,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, ['Date']) ||
                          findFieldByNameHints(loadRampFieldsTable, [
                              'date',
                              'delivery date',
                              'energization',
                              'cod',
                          ], {
                              typePredicate: isDateLikeField,
                          }),
                  },
                  {
                      key: 'loadRampDirectorField',
                      label: 'Director (on Load Ramp — user field or lookup)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isDirectorLikeField,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, [
                              'Director',
                              'Directors',
                              'Owner',
                              'Project Director',
                          ]) ||
                          findFieldByNameHints(loadRampFieldsTable, [
                              'director',
                              'owner',
                              'assigned to',
                              'lead',
                          ], {
                              typePredicate: isDirectorLikeField,
                          }),
                  },
                  {
                      key: 'loadRampStatusField',
                      label: 'Verification Status (on Load Ramp — Verified / Not Verified)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isSingleSelectField,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, [
                              'Status',
                              'Verification Status',
                              'Verified',
                          ]) ||
                          findFieldByNameHints(loadRampFieldsTable, [
                              'verification status',
                              'status',
                              'verified',
                          ], {
                              typePredicate: isSingleSelectField,
                          }),
                  },
                  {
                      key: 'loadRampVerifiedAtField',
                      label: 'Verified At (on Load Ramp)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isDateTimeField,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, [
                              'Verified At',
                              'Verified Time',
                          ]) ||
                          findFieldByNameHints(loadRampFieldsTable, [
                              'verified at',
                              'verified time',
                              'verification date',
                          ], {
                              typePredicate: isDateTimeField,
                          }),
                  },
                  {
                      key: 'loadRampLastModifiedField',
                      label: 'Last Modified Time (on Load Ramp — auto-unverify)',
                      type: 'field',
                      table: loadRampFieldsTable,
                      shouldFieldBeAllowed: isLastModifiedField,
                      defaultValue:
                          findFieldByExactNames(loadRampFieldsTable, [
                              'Last Modified Time',
                              'Last Modified',
                          ]) ||
                          findFieldByNameHints(loadRampFieldsTable, [
                              'last modified time',
                              'last modified',
                          ], {
                              typePredicate: isLastModifiedField,
                          }),
                  },
              ]
            : []),
    ];
}

function selectLabelAndColor(record, field) {
    if (!record || !field) return {label: '', colorHex: null};
    const cell = record.getCellValue(field);
    const label = getSelectName(cell) || readCellDisplayName(record, field);
    const token = getSelectColorToken(cell);
    let colorHex = null;
    if (token) {
        try {
            colorHex = colorUtils.getHexForColor(token);
        } catch {
            colorHex = null;
        }
    }
    return {label, colorHex};
}

function emptyRampEntry() {
    return {quarters: {}, longTerm: null, entries: []};
}

function compareRampEntries(a, b) {
    const ay = a.dateParts?.year || 0;
    const by = b.dateParts?.year || 0;
    if (ay !== by) return ay - by;
    const am = a.dateParts?.month || 0;
    const bm = b.dateParts?.month || 0;
    if (am !== bm) return am - bm;
    return (a.dateParts?.day || 0) - (b.dateParts?.day || 0);
}

function pushRampEntry(existing, entry) {
    if (!existing.entries.some((item) => item.id === entry.id)) {
        existing.entries.push(entry);
    }
}

function sumIncrementalLoads(quarters, longTermMw) {
    let total = 0;
    let hasValue = false;
    for (const value of Object.values(quarters || {})) {
        if (value == null || value === '') continue;
        total += Number(value) || 0;
        hasValue = true;
    }
    if (longTermMw != null && longTermMw !== '') {
        total += Number(longTermMw) || 0;
        hasValue = true;
    }
    return hasValue ? total : null;
}

function readRampDateParts(record, dateField) {
    if (!record || !dateField) return null;
    return (
        getDateParts(record.getCellValue(dateField)) ||
        getDateParts(record.getCellValueAsString(dateField))
    );
}

function buildRampLookup({
    loadRampRecords,
    loadRampCampusField,
    loadRampMwField,
    loadRampDateField,
    loadRampStatusField,
    loadRampLastModifiedField,
    loadRampVerifiedAtField,
    years,
    startYear,
    startMonth,
    verifiedOnly,
}) {
    const byCampus = new Map();
    if (!loadRampRecords || !loadRampCampusField || !loadRampMwField || !loadRampDateField) {
        return byCampus;
    }

    const validKeys = new Set(years.flatMap((y) => y.quarters.map((q) => q.key)));
    const lastQuarter = lastQuarterOfHorizon(years);

    for (const record of loadRampRecords) {
        const campusIds = getLinkedRecordIds(record.getCellValue(loadRampCampusField));
        const campusName = readCellDisplayName(record, loadRampCampusField).toLowerCase();
        const mw = readNumericOrNull(record, loadRampMwField);
        if (!campusIds.length && !campusName) continue;

        const verified = isRampRecordVerified(
            record,
            loadRampStatusField,
            loadRampLastModifiedField,
            loadRampVerifiedAtField,
        );
        if (verifiedOnly && !verified) continue;

        const statusLabel = loadRampStatusField
            ? getSelectName(record.getCellValue(loadRampStatusField)) ||
              readCellDisplayName(record, loadRampStatusField)
            : '';
        const parts = readRampDateParts(record, loadRampDateField);
        const beforeStart = isDateBeforeStart(parts, startYear, startMonth);
        const afterHorizon = isDateAfterHorizon(parts, lastQuarter);
        const quarter = quarterFromMonth(parts?.month);
        const key = quarter && parts?.year ? quarterKey(parts.year, quarter) : null;
        const isInHorizon = Boolean(!beforeStart && key && validKeys.has(key));
        const entry = {
            id: record.id,
            record,
            mw,
            dateLabel: formatShortDate(parts) || readCellDisplayName(record, loadRampDateField),
            dateParts: parts,
            statusLabel,
            verified,
            afterHorizon,
            inHorizon: isInHorizon,
        };

        const addTo = (mapKey) => {
            if (!mapKey) return;
            const existing = byCampus.get(mapKey) || emptyRampEntry();
            pushRampEntry(existing, entry);
            if (mw != null && !beforeStart && (isInHorizon || afterHorizon)) {
                if (afterHorizon) {
                    existing.longTerm = (existing.longTerm || 0) + mw;
                } else {
                    existing.quarters[key] = (existing.quarters[key] || 0) + mw;
                }
            }
            byCampus.set(mapKey, existing);
        };
        campusIds.forEach(addTo);
        if (campusName) addTo(`name:${campusName}`);
    }

    for (const value of byCampus.values()) {
        value.entries.sort(compareRampEntries);
    }

    return byCampus;
}

function LandBankApp() {
    const base = useBase();
    const fallbackTable = base.tables[0];
    const {customPropertyValueByKey, errorState} = useCustomProperties(getCustomProperties);

    const campusesTable = customPropertyValueByKey.campusesTable;
    const loadRampTable = customPropertyValueByKey.loadRampTable;
    const campusRecords = useRecords(campusesTable || fallbackTable);
    const loadRampRecords = useRecords(
        loadRampTable && loadRampTable.id !== campusesTable?.id
            ? loadRampTable
            : fallbackTable,
    );

    const metroField = resolveField(campusesTable, customPropertyValueByKey.metroField, {
        exactNames: ['Metro'],
    });
    const campusField = resolveField(campusesTable, customPropertyValueByKey.campusField, {
        exactNames: ['Campus', 'Campus Name'],
    });
    const campusNameField = resolveField(
        campusesTable,
        customPropertyValueByKey.campusNameField,
        {exactNames: ['Supplier Building Code', 'Building Code']},
    );
    const landStatusField = resolveField(
        campusesTable,
        customPropertyValueByKey.landStatusField,
        {exactNames: ['Land Status']},
    );
    const rezoningStatusField = resolveField(
        campusesTable,
        customPropertyValueByKey.rezoningStatusField,
        {exactNames: ['Rezoning Status']},
    );
    const waterCapacityField = resolveField(
        campusesTable,
        customPropertyValueByKey.waterCapacityField,
        {exactNames: ['Water Capacity', 'Water Capacity If Available']},
    );
    const powerConfidenceField = resolveField(
        campusesTable,
        customPropertyValueByKey.powerConfidenceField,
        {exactNames: ['Power Confidence Level', 'Power Utility Confidence Level']},
    );
    const liveMwField = resolveField(campusesTable, customPropertyValueByKey.liveMwField, {
        exactNames: ['Live MW'],
    });
    const pipelineMwField = resolveField(
        campusesTable,
        customPropertyValueByKey.pipelineMwField,
        {exactNames: ['Pipeline MW']},
    );

    const loadRampCampusField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampCampusField,
        {exactNames: ['Campus']},
    );
    const loadRampMwField = resolveField(loadRampTable, customPropertyValueByKey.loadRampMwField, {
        exactNames: ['MW', 'MVA'],
    });
    const loadRampDateField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampDateField,
        {exactNames: ['Date'], typePredicate: isDateLikeField},
    );
    const campusDirectorField = resolveField(
        campusesTable,
        customPropertyValueByKey.campusDirectorField,
        {
            exactNames: ['Director', 'Owner', 'Project Director', 'Assigned To'],
            typePredicate: isDirectorLikeField,
        },
    );
    const loadRampDirectorField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampDirectorField,
        {
            exactNames: ['Director', 'Directors', 'Owner', 'Project Director'],
            typePredicate: isDirectorLikeField,
        },
    );
    const loadRampStatusField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampStatusField,
        {
            exactNames: ['Status', 'Verification Status', 'Verified'],
            typePredicate: isSingleSelectField,
        },
    );
    const loadRampVerifiedAtField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampVerifiedAtField,
        {
            exactNames: ['Verified At', 'Verified Time'],
            typePredicate: isDateTimeField,
        },
    );
    const loadRampLastModifiedField = resolveField(
        loadRampTable,
        customPropertyValueByKey.loadRampLastModifiedField,
        {
            exactNames: ['Last Modified Time', 'Last Modified'],
            typePredicate: isLastModifiedField,
        },
    );

    const session = useSession();
    const currentUser = session?.currentUser || null;
    const dashboardView =
        customPropertyValueByKey.dashboardView === 'director' ? 'director' : 'enterprise';
    const isDirectorView = dashboardView === 'director';

    const [timelineStart, setTimelineStart] = useState(() => currentYearMonthString());
    const years = useMemo(() => {
        const parts = parseYearMonth(timelineStart);
        if (!parts) return buildQuarterColumns();
        return buildQuarterColumns(parts.year, parts.month);
    }, [timelineStart]);
    const timelineStartParts = parseYearMonth(timelineStart) || {
        year: new Date().getFullYear(),
        month: 1,
    };

    const [activeTab, setActiveTab] = useState('over');
    const [editingCampus, setEditingCampus] = useState(null);
    const [isVerifying, setIsVerifying] = useState(false);
    const [isDownloading, setIsDownloading] = useState(false);
    const unverifyAttemptedRef = useRef(new Set());

    const missing = [];
    if (!campusesTable) missing.push('Campuses / Land Bank table');
    if (!campusField) missing.push('Campus field');
    if (!campusNameField) missing.push('Supplier Building Code field');
    if (!loadRampTable || loadRampTable.id === campusesTable?.id) {
        missing.push('Load Ramp table');
    }
    if (!loadRampCampusField) missing.push('Campus field on Load Ramp');
    if (!loadRampMwField) missing.push('MW field on Load Ramp');
    if (!loadRampDateField) {
        missing.push('Date field on Load Ramp (quarter and year are derived from this)');
    }
    if (!loadRampStatusField) {
        missing.push(
            'Verification Status field on Load Ramp (single select with Verified and Not Verified)',
        );
    }
    if (isDirectorView && !loadRampDirectorField && !campusDirectorField) {
        missing.push(
            'Director field on Load Ramp (user field or lookup of the Airtable user)',
        );
    }
    const isConfigured = missing.length === 0;
    const canAutoUnverify = Boolean(loadRampLastModifiedField && loadRampVerifiedAtField);

    const rampByCampus = useMemo(
        () =>
            isConfigured
                ? buildRampLookup({
                      loadRampRecords,
                      loadRampCampusField,
                      loadRampMwField,
                      loadRampDateField,
                      loadRampStatusField,
                      loadRampLastModifiedField,
                      loadRampVerifiedAtField,
                      years,
                      startYear: timelineStartParts.year,
                      startMonth: timelineStartParts.month,
                      verifiedOnly: !isDirectorView,
                  })
                : new Map(),
        [
            isConfigured,
            loadRampRecords,
            loadRampCampusField,
            loadRampMwField,
            loadRampDateField,
            loadRampStatusField,
            loadRampLastModifiedField,
            loadRampVerifiedAtField,
            years,
            timelineStartParts.year,
            timelineStartParts.month,
            isDirectorView,
        ],
    );

    const assignedCampuses = useMemo(() => {
        const ids = new Set();
        const names = new Set();
        if (
            !isDirectorView ||
            !currentUser ||
            !loadRampDirectorField ||
            !loadRampCampusField ||
            !loadRampRecords
        ) {
            return {ids, names};
        }
        for (const record of loadRampRecords) {
            if (!recordAssignedToUser(record, loadRampDirectorField, currentUser)) continue;
            getLinkedRecordIds(record.getCellValue(loadRampCampusField)).forEach((id) =>
                ids.add(id),
            );
            const campusName = readCellDisplayName(record, loadRampCampusField).toLowerCase();
            if (campusName) names.add(campusName);
        }
        return {ids, names};
    }, [
        isDirectorView,
        currentUser,
        loadRampDirectorField,
        loadRampCampusField,
        loadRampRecords,
    ]);

    const rows = useMemo(() => {
        if (!isConfigured || !campusRecords) return [];

        return campusRecords
            .filter((record) => {
                if (!isDirectorView) return true;
                if (!currentUser) return false;
                if (
                    campusDirectorField &&
                    recordAssignedToUser(record, campusDirectorField, currentUser)
                ) {
                    return true;
                }
                if (assignedCampuses.ids.has(record.id)) return true;
                const campusName = campusField
                    ? readCellDisplayName(record, campusField).toLowerCase()
                    : '';
                const supplierCode = campusNameField
                    ? readCellDisplayName(record, campusNameField).toLowerCase()
                    : '';
                return Boolean(
                    (campusName && assignedCampuses.names.has(campusName)) ||
                        (supplierCode && assignedCampuses.names.has(supplierCode)),
                );
            })
            .map((record) => {
                const campus = campusField ? readCellDisplayName(record, campusField) : '';
                const supplierBuildingCode = campusNameField
                    ? readCellDisplayName(record, campusNameField)
                    : '';
                const metro = metroField ? readCellDisplayName(record, metroField) : '';
                const land = selectLabelAndColor(record, landStatusField);
                const rezoning = selectLabelAndColor(record, rezoningStatusField);
                const power = selectLabelAndColor(record, powerConfidenceField);
                const liveMw = readNumericOrNull(record, liveMwField);
                const pipelineMw = readNumericOrNull(record, pipelineMwField);
                const rampValues =
                    rampByCampus.get(record.id) ||
                    (campus
                        ? rampByCampus.get(`name:${campus.toLowerCase()}`)
                        : null) ||
                    (supplierBuildingCode
                        ? rampByCampus.get(`name:${supplierBuildingCode.toLowerCase()}`)
                        : null) ||
                    emptyRampEntry();

                const quarters = {};
                for (const year of years) {
                    for (const q of year.quarters) {
                        const fromRamp = rampValues.quarters[q.key];
                        quarters[q.key] = fromRamp != null ? fromRamp : null;
                    }
                }
                const longTermMw = rampValues.longTerm;
                const grossMw = sumIncrementalLoads(quarters, longTermMw);
                const mva = grossMw;
                const rampEntries = rampValues.entries || [];
                const pendingCount = rampEntries.filter((entry) => !entry.verified).length;

                return {
                    id: record.id,
                    record,
                    metro,
                    campus,
                    supplierBuildingCode,
                    landStatus: land.label,
                    landStatusColor: land.colorHex,
                    rezoningStatus: rezoning.label,
                    rezoningStatusColor: rezoning.colorHex,
                    waterCapacity: waterCapacityField
                        ? readCellDisplayName(record, waterCapacityField)
                        : '',
                    powerConfidence: power.label,
                    powerConfidenceColor: power.colorHex,
                    grossMw,
                    liveMw,
                    pipelineMw,
                    mva,
                    longTermMw,
                    quarters,
                    rampEntries,
                    pendingCount,
                    verifiedCount: rampEntries.length - pendingCount,
                };
            })
            .filter((row) => (isDirectorView ? true : row.rampEntries.length > 0))
            .sort(compareCampusRows);
    }, [
        isConfigured,
        campusRecords,
        campusField,
        campusNameField,
        metroField,
        landStatusField,
        rezoningStatusField,
        waterCapacityField,
        powerConfidenceField,
        liveMwField,
        pipelineMwField,
        rampByCampus,
        years,
        isDirectorView,
        campusDirectorField,
        assignedCampuses,
        currentUser,
    ]);

    const overRows = useMemo(() => rows.filter((row) => isOver100Mva(row.mva)), [rows]);
    const underRows = useMemo(() => rows.filter((row) => !isOver100Mva(row.mva)), [rows]);
    const visibleRows = activeTab === 'over' ? overRows : underRows;
    const pendingRecords = useMemo(() => {
        const seen = new Set();
        const records = [];
        for (const row of rows) {
            for (const entry of row.rampEntries || []) {
                if (entry.verified || !entry.record || seen.has(entry.record.id)) continue;
                seen.add(entry.record.id);
                records.push(entry.record);
            }
        }
        return records;
    }, [rows]);
    const pendingCount = pendingRecords.length;

    const canExpandRamp = Boolean(loadRampTable?.hasPermissionToExpandRecords?.());
    const canVerify = Boolean(
        isDirectorView &&
            loadRampTable &&
            loadRampStatusField &&
            verifiedChoice(loadRampStatusField) &&
            loadRampTable.hasPermissionToUpdateRecords?.(),
    );

    useEffect(() => {
        if (
            !isConfigured ||
            !loadRampTable ||
            !loadRampStatusField ||
            !canAutoUnverify ||
            !loadRampRecords?.length
        ) {
            return undefined;
        }
        if (!loadRampTable.hasPermissionToUpdateRecords?.()) return undefined;

        const stale = loadRampRecords.filter((record) => {
            const label = getSelectName(record.getCellValue(loadRampStatusField));
            if (!isVerifiedStatusLabel(label)) return false;
            if (
                isRampRecordVerified(
                    record,
                    loadRampStatusField,
                    loadRampLastModifiedField,
                    loadRampVerifiedAtField,
                )
            ) {
                return false;
            }
            const attemptKey = `${record.id}:${getCellTimestamp(record, loadRampLastModifiedField) || 0}`;
            if (unverifyAttemptedRef.current.has(attemptKey)) return false;
            return true;
        });
        if (!stale.length) return undefined;

        stale.forEach((record) => {
            unverifyAttemptedRef.current.add(
                `${record.id}:${getCellTimestamp(record, loadRampLastModifiedField) || 0}`,
            );
        });

        const nextStatus = writeSelectChoice(unverifiedChoice(loadRampStatusField));
        const updates = stale.map((record) => ({
            id: record.id,
            fields: {
                [loadRampStatusField.id]: nextStatus,
            },
        }));

        updateRecordsInChunks(loadRampTable, updates).catch((error) => {
            void error;
        });
        return undefined;
    }, [
        isConfigured,
        loadRampTable,
        loadRampRecords,
        loadRampStatusField,
        loadRampLastModifiedField,
        loadRampVerifiedAtField,
        canAutoUnverify,
    ]);

    const handleVerifyAll = useCallback(async () => {
        if (!canVerify || !pendingRecords.length) return;
        const choice = writeSelectChoice(verifiedChoice(loadRampStatusField));
        if (!choice) return;
        setIsVerifying(true);
        const fields = {
            [loadRampStatusField.id]: choice,
        };
        if (loadRampVerifiedAtField) {
            fields[loadRampVerifiedAtField.id] = writeTimestampForField(loadRampVerifiedAtField);
        }
        try {
            await updateRecordsInChunks(
                loadRampTable,
                pendingRecords.map((record) => ({id: record.id, fields})),
            );
        } finally {
            setIsVerifying(false);
        }
    }, [
        canVerify,
        pendingRecords,
        loadRampTable,
        loadRampStatusField,
        loadRampVerifiedAtField,
    ]);

    const handleDownloadExcel = useCallback(async () => {
        setIsDownloading(true);
        try {
            await downloadExcel({
                overRows,
                underRows,
                years,
                view: dashboardView,
            });
        } catch (error) {
            void error;
        } finally {
            setIsDownloading(false);
        }
    }, [overRows, underRows, years, dashboardView]);

    if (errorState) {
        return (
            <div className="w-full h-full min-h-screen p-6 bg-white dark:bg-gray-gray800">
                <p className="text-red-red">Error loading properties: {String(errorState)}</p>
            </div>
        );
    }

    if (!isConfigured) {
        return <SetupPanel missing={missing} />;
    }

    return (
        <div className="w-full h-full min-h-screen max-w-full overflow-hidden bg-white dark:bg-gray-gray800 flex flex-col">
            <div className="w-full max-w-full min-h-0 flex-1 flex flex-col p-4 sm:p-6 lg:p-8 pb-16 sm:pb-20 lg:pb-24">
                <div className="w-full max-w-full shrink-0">
                    <DashboardHeader
                        logoSrc={stackLogo}
                        view={dashboardView}
                        timelineStart={timelineStart}
                        onTimelineStartChange={setTimelineStart}
                        horizonLabel={horizonRangeLabel(years)}
                        isDownloading={isDownloading}
                        onDownloadExcel={handleDownloadExcel}
                        grossMw={visibleRows.reduce(
                            (sum, row) => sum + (row.grossMw == null ? 0 : row.grossMw),
                            0,
                        )}
                        pipelineMw={visibleRows.reduce(
                            (sum, row) => sum + (row.pipelineMw == null ? 0 : row.pipelineMw),
                            0,
                        )}
                        campusCount={visibleRows.length}
                    />
                    <ViewBanner
                        view={dashboardView}
                        userName={currentUser?.name || currentUser?.email || ''}
                        pendingCount={pendingCount}
                        canAutoUnverify={canAutoUnverify}
                        canVerify={canVerify}
                        isVerifying={isVerifying}
                        onVerifyAll={handleVerifyAll}
                    />
                    <MvaTabs
                        activeTab={activeTab}
                        onChange={setActiveTab}
                        overCount={overRows.length}
                        underCount={underRows.length}
                    />
                </div>
                <div className="w-full max-w-full min-h-0 flex-1 overflow-auto">
                    <LandBankTable
                        rows={visibleRows}
                        years={years}
                        canExpandRamp={canExpandRamp}
                        onCampusClick={(record) => setEditingCampus(record)}
                        onRampClick={(record) => expandRecord(record)}
                        view={dashboardView}
                        emptyMessage={
                            isDirectorView && !currentUser
                                ? 'Sign in to see campuses assigned to you.'
                                : isDirectorView && rows.length === 0
                                  ? 'No campuses are assigned to you on the Load Ramp Director lookup.'
                                  : activeTab === 'over'
                                    ? isDirectorView
                                        ? 'No assigned campuses at or above 100 MVA.'
                                        : 'No campuses at or above 100 MVA with verified Load Ramp.'
                                    : isDirectorView
                                      ? 'No assigned campuses under 100 MVA.'
                                      : 'No campuses under 100 MVA with verified Load Ramp.'
                        }
                    />
                </div>
            </div>
            {editingCampus ? (
                <CampusEditForm
                    key={editingCampus.id}
                    record={editingCampus}
                    table={campusesTable}
                    campusName={
                        campusField
                            ? readCellDisplayName(editingCampus, campusField)
                            : ''
                    }
                    metroField={metroField}
                    supplierField={campusNameField}
                    landStatusField={landStatusField}
                    rezoningStatusField={rezoningStatusField}
                    onClose={() => setEditingCampus(null)}
                />
            ) : null}
        </div>
    );
}

initializeBlock({interface: () => <LandBankApp />});
