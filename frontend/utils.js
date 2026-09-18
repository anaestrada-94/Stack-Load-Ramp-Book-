import {FieldType} from '@airtable/blocks/interface/models';

export const MVA_THRESHOLD = 100;
export const HORIZON_YEARS = 6;
export const QUARTERS = [1, 2, 3, 4];

export function getFieldType(field) {
    return field?.config?.type || field?.type || null;
}

export function findTableByNameHints(base, hints) {
    const lowerHints = hints.map((h) => h.toLowerCase());
    return (
        base.tables.find((table) => {
            const name = table.name.toLowerCase();
            return lowerHints.some((hint) => name.includes(hint));
        }) || null
    );
}

export function findFieldByNameHints(
    table,
    hints,
    {excludeHints = [], typePredicate} = {},
) {
    if (!table) return undefined;
    const lowerHints = hints.map((h) => h.toLowerCase());
    const lowerExclude = excludeHints.map((h) => h.toLowerCase());

    const candidates = table.fields.filter((field) => {
        if (typePredicate && !typePredicate(field)) return false;
        const name = field.name.toLowerCase();
        if (lowerExclude.some((ex) => name.includes(ex))) return false;
        return lowerHints.some((hint) => name.includes(hint));
    });

    if (candidates.length === 0) return undefined;

    const exactMatches = candidates.filter((field) =>
        lowerHints.some((hint) => field.name.toLowerCase() === hint),
    );
    return (exactMatches.length > 0 ? exactMatches : candidates)[0];
}

export function findFieldByExactNames(table, exactNames = []) {
    if (!table || !exactNames.length) return null;
    const lowered = exactNames.map((n) => String(n).toLowerCase());
    for (const want of lowered) {
        const match = table.fields.find((field) => field.name.toLowerCase() === want);
        if (match) return match;
    }
    return null;
}

export function resolveField(table, configuredField, {exactNames = [], typePredicate} = {}) {
    if (!table) return configuredField || null;

    if (configuredField?.id) {
        const live = table.fields.find((field) => field.id === configuredField.id) || null;
        if (live && (!typePredicate || typePredicate(live))) return live;
    }

    if (configuredField?.name) {
        const byName = findFieldByExactNames(table, [configuredField.name]);
        if (byName && (!typePredicate || typePredicate(byName))) return byName;
    }

    const exact = findFieldByExactNames(table, exactNames);
    if (exact && (!typePredicate || typePredicate(exact))) return exact;

    return null;
}

export function isLinkField(field) {
    return getFieldType(field) === FieldType.MULTIPLE_RECORD_LINKS;
}

export function isSelectField(field) {
    const type = getFieldType(field);
    return type === FieldType.SINGLE_SELECT || type === FieldType.MULTIPLE_SELECTS;
}

export function isSingleSelectField(field) {
    return getFieldType(field) === FieldType.SINGLE_SELECT;
}

export function isCollaboratorField(field) {
    const type = getFieldType(field);
    return type === FieldType.SINGLE_COLLABORATOR || type === FieldType.MULTIPLE_COLLABORATORS;
}

export function isLookupField(field) {
    return getFieldType(field) === FieldType.MULTIPLE_LOOKUP_VALUES;
}

export function isDirectorLikeField(field) {
    const type = getFieldType(field);
    return (
        isCollaboratorField(field) ||
        isLookupField(field) ||
        type === FieldType.EMAIL ||
        type === FieldType.FORMULA ||
        type === FieldType.ROLLUP ||
        type === FieldType.SINGLE_LINE_TEXT ||
        type === FieldType.MULTILINE_TEXT
    );
}

export function isLastModifiedField(field) {
    return getFieldType(field) === FieldType.LAST_MODIFIED_TIME;
}

export function isDateTimeField(field) {
    const type = getFieldType(field);
    return type === FieldType.DATE_TIME || type === FieldType.DATE;
}

export function isTextLikeField(field) {
    const type = getFieldType(field);
    return (
        type === FieldType.SINGLE_LINE_TEXT ||
        type === FieldType.MULTILINE_TEXT ||
        type === FieldType.RICH_TEXT ||
        type === FieldType.EMAIL ||
        type === FieldType.FORMULA ||
        type === FieldType.ROLLUP ||
        type === FieldType.MULTIPLE_LOOKUP_VALUES ||
        type === FieldType.MULTIPLE_RECORD_LINKS
    );
}

export function isNumericLikeField(field) {
    const type = getFieldType(field);
    return (
        type === FieldType.NUMBER ||
        type === FieldType.CURRENCY ||
        type === FieldType.PERCENT ||
        type === FieldType.FORMULA ||
        type === FieldType.ROLLUP ||
        type === FieldType.COUNT ||
        type === FieldType.DURATION ||
        type === FieldType.MULTIPLE_LOOKUP_VALUES ||
        type === FieldType.SINGLE_LINE_TEXT
    );
}

export function isDateLikeField(field) {
    const type = getFieldType(field);
    if (type === FieldType.DATE || type === FieldType.DATE_TIME) return true;
    if (
        type !== FieldType.FORMULA &&
        type !== FieldType.ROLLUP &&
        type !== FieldType.MULTIPLE_LOOKUP_VALUES
    ) {
        return false;
    }
    const result = field?.options?.result || field?.config?.options?.result;
    return result?.type === FieldType.DATE || result?.type === FieldType.DATE_TIME;
}

export function isCampusNameField(field) {
    const type = getFieldType(field);
    return (
        type === FieldType.MULTIPLE_RECORD_LINKS ||
        type === FieldType.MULTIPLE_LOOKUP_VALUES ||
        type === FieldType.SINGLE_LINE_TEXT ||
        type === FieldType.FORMULA ||
        type === FieldType.ROLLUP
    );
}

export function toNumber(value) {
    if (value == null || value === '') return 0;
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (typeof value === 'string') {
        const cleaned = value.replace(/[$,\s]/g, '').replace(/\((.*)\)/, '-$1');
        const parsed = Number(cleaned);
        return Number.isFinite(parsed) ? parsed : 0;
    }
    if (Array.isArray(value)) {
        return value.reduce((sum, item) => {
            if (item == null) return sum;
            if (typeof item === 'number') return sum + (Number.isFinite(item) ? item : 0);
            if (typeof item === 'object' && item !== null && 'value' in item) {
                return sum + toNumber(item.value);
            }
            return sum + toNumber(item);
        }, 0);
    }
    if (typeof value === 'object' && value !== null) {
        if (typeof value.value === 'number' || typeof value.value === 'string') {
            return toNumber(value.value);
        }
    }
    return 0;
}

export function readNumericOrNull(record, field) {
    if (!record || !field) return null;
    const raw = record.getCellValue(field);
    if (raw == null || raw === '') return null;
    if (Array.isArray(raw) && raw.length === 0) return null;
    const n = toNumber(raw);
    return Number.isFinite(n) ? n : null;
}

export function readCellDisplayName(record, field) {
    if (!record || !field) return '';
    const asString = record.getCellValueAsString(field);
    if (asString && asString.trim()) return asString.trim();

    const value = record.getCellValue(field);
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'Yes' : 'No';
    if (Array.isArray(value)) {
        return value
            .map((item) => {
                if (item == null) return '';
                if (typeof item === 'string' || typeof item === 'number') return String(item);
                return item.name || item.displayName || item.id || '';
            })
            .filter(Boolean)
            .join(', ');
    }
    if (typeof value === 'object') {
        return value.name || value.displayName || '';
    }
    return '';
}

export function getSelectName(cellValue) {
    if (!cellValue) return '';
    if (typeof cellValue === 'string') return cellValue;
    if (Array.isArray(cellValue)) {
        return cellValue
            .map((v) => (typeof v === 'string' ? v : v?.name))
            .filter(Boolean)
            .join(', ');
    }
    return cellValue.name || '';
}

export function getSelectColorToken(cellValue) {
    if (!cellValue) return null;
    const first = Array.isArray(cellValue) ? cellValue[0] : cellValue;
    if (!first || typeof first === 'string') return null;
    return first.color || null;
}

export function getLinkedRecordIds(cellValue) {
    if (!cellValue) return [];
    if (Array.isArray(cellValue)) {
        return cellValue.map((item) => (typeof item === 'string' ? item : item?.id)).filter(Boolean);
    }
    if (typeof cellValue === 'object' && cellValue.id) return [cellValue.id];
    if (typeof cellValue === 'string') return [cellValue];
    return [];
}

export function currentHorizonStartYear(now = new Date()) {
    return now.getFullYear();
}

export function currentYearMonthString(now = new Date()) {
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

export function parseYearMonth(value) {
    if (!value || typeof value !== 'string') return null;
    const match = /^(\d{4})-(\d{2})$/.exec(value);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (!year || month < 1 || month > 12) return null;
    return {year, month};
}

export function quarterKey(year, quarter) {
    return `${year}-Q${quarter}`;
}

const QUARTER_END = {
    1: {month: 3, day: 31},
    2: {month: 6, day: 30},
    3: {month: 9, day: 30},
    4: {month: 12, day: 31},
};

export function buildQuarterColumns(
    startYear = currentHorizonStartYear(),
    startMonth = 1,
) {
    const startQuarter = quarterFromMonth(startMonth) || 1;
    const years = [];
    let year = startYear;
    let quarter = startQuarter;

    for (let i = 0; i < HORIZON_YEARS * QUARTERS.length; i++) {
        let group = years.find((item) => item.year === year);
        if (!group) {
            group = {
                year,
                label: String(year),
                isCurrentYear: years.length === 0,
                quarters: [],
            };
            years.push(group);
        }
        group.quarters.push({
            key: quarterKey(year, quarter),
            year,
            quarter,
            label: `Q${quarter}`,
        });
        quarter += 1;
        if (quarter > 4) {
            quarter = 1;
            year += 1;
        }
    }
    return years;
}

export function lastQuarterOfHorizon(years) {
    const lastYear = years?.[years.length - 1];
    if (!lastYear?.quarters?.length) return null;
    return lastYear.quarters[lastYear.quarters.length - 1];
}

export function horizonRangeLabel(years) {
    const first = years?.[0]?.quarters?.[0];
    const last = lastQuarterOfHorizon(years);
    if (!first || !last) return '';
    return `${first.label} ${first.year} – ${last.label} ${last.year}`;
}

export function ymdValue(year, month = 1, day = 1) {
    return Number(year) * 10000 + Number(month || 1) * 100 + Number(day || 1);
}

export function isDateBeforeStart(parts, startYear, startMonth) {
    if (!parts?.year || startYear == null || startMonth == null) return false;
    return ymdValue(parts.year, parts.month, parts.day || 1) < ymdValue(startYear, startMonth, 1);
}

export function isDateAfterHorizon(parts, lastQuarter) {
    if (!parts?.year || !lastQuarter?.year || !lastQuarter?.quarter) return false;
    const end = QUARTER_END[lastQuarter.quarter] || QUARTER_END[4];
    return (
        ymdValue(parts.year, parts.month, parts.day) >
        ymdValue(lastQuarter.year, end.month, end.day)
    );
}

export function getDateParts(dateValue) {
    if (dateValue == null || dateValue === '') return null;

    if (Array.isArray(dateValue)) {
        return dateValue.length ? getDateParts(dateValue[0]) : null;
    }

    if (typeof dateValue === 'object' && !(dateValue instanceof Date)) {
        if (typeof dateValue.date === 'string') return getDateParts(dateValue.date);
        if (typeof dateValue.iso === 'string') return getDateParts(dateValue.iso);
        if (typeof dateValue.value === 'string' || typeof dateValue.value === 'number') {
            return getDateParts(dateValue.value);
        }
    }

    if (typeof dateValue === 'string') {
        const trimmed = dateValue.trim();
        if (!trimmed) return null;
        const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (iso) {
            return {year: Number(iso[1]), month: Number(iso[2]), day: Number(iso[3])};
        }
        const ym = trimmed.match(/^(\d{4})-(\d{2})$/);
        if (ym) return {year: Number(ym[1]), month: Number(ym[2]), day: 1};
        const us = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (us) {
            return {year: Number(us[3]), month: Number(us[1]), day: Number(us[2])};
        }
        const parsed = new Date(trimmed);
        if (!Number.isNaN(parsed.getTime())) {
            return {
                year: parsed.getUTCFullYear(),
                month: parsed.getUTCMonth() + 1,
                day: parsed.getUTCDate(),
            };
        }
        return null;
    }

    if (dateValue instanceof Date) {
        if (Number.isNaN(dateValue.getTime())) return null;
        return {
            year: dateValue.getUTCFullYear(),
            month: dateValue.getUTCMonth() + 1,
            day: dateValue.getUTCDate(),
        };
    }

    if (typeof dateValue === 'number') {
        const d = new Date(dateValue);
        if (Number.isNaN(d.getTime())) return null;
        return {
            year: d.getUTCFullYear(),
            month: d.getUTCMonth() + 1,
            day: d.getUTCDate(),
        };
    }

    return null;
}

export function quarterFromMonth(month) {
    if (!month) return null;
    return Math.ceil(month / 3);
}

export function formatMw(value, {suffix = false, empty = '—'} = {}) {
    if (value == null || value === '') return empty;
    const n = toNumber(value);
    if (!Number.isFinite(n)) return empty;
    const abs = Math.abs(n).toLocaleString('en-US', {
        maximumFractionDigits: n % 1 === 0 ? 0 : 1,
    });
    const signed = n < 0 ? `-${abs}` : abs;
    return suffix ? `${signed} MW` : signed;
}

export function formatShortDate(parts) {
    if (!parts?.year) return '';
    const month = String(parts.month || 1).padStart(2, '0');
    const day = String(parts.day || 1).padStart(2, '0');
    return `${month}/${day}/${parts.year}`;
}

export function flattenCellItems(value) {
    const out = [];
    const visit = (node) => {
        if (node == null || node === '') return;
        if (Array.isArray(node)) {
            node.forEach(visit);
            return;
        }
        if (typeof node === 'object') {
            if (Object.prototype.hasOwnProperty.call(node, 'value')) {
                visit(node.value);
            }
            out.push(node);
            return;
        }
        out.push(node);
    };
    visit(value);
    return out;
}

export function cellHasCurrentUser(cellValue, user, asString = '') {
    if (!user) return false;
    const email = String(user.email || '').toLowerCase();
    const name = String(user.name || '').toLowerCase();
    const items = flattenCellItems(cellValue);
    const matched = items.some((item) => {
        if (item == null) return false;
        if (typeof item === 'string') {
            const text = item.toLowerCase();
            return item === user.id || (email && text === email) || (name && text === name);
        }
        if (typeof item !== 'object') return false;
        if (item.id && item.id === user.id) return true;
        if (item.email && email && String(item.email).toLowerCase() === email) return true;
        if (item.name && name && String(item.name).toLowerCase() === name) return true;
        return false;
    });
    if (matched) return true;
    const text = String(asString || '').toLowerCase();
    if (!text) return false;
    if (email && text.includes(email)) return true;
    if (name && name.length > 2 && text.includes(name)) return true;
    return false;
}

export function recordAssignedToUser(record, field, user) {
    if (!record || !field || !user) return false;
    return cellHasCurrentUser(
        record.getCellValue(field),
        user,
        record.getCellValueAsString(field),
    );
}

export function getCellTimestamp(record, field) {
    if (!record || !field) return null;
    const value = record.getCellValue(field);
    if (value == null || value === '') return null;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (value instanceof Date) {
        const t = value.getTime();
        return Number.isNaN(t) ? null : t;
    }
    if (typeof value === 'string') {
        const t = Date.parse(value);
        return Number.isNaN(t) ? null : t;
    }
    if (typeof value === 'object') {
        if (typeof value.iso === 'string') {
            const t = Date.parse(value.iso);
            return Number.isNaN(t) ? null : t;
        }
        if (typeof value.date === 'string') {
            const t = Date.parse(value.date);
            return Number.isNaN(t) ? null : t;
        }
    }
    return null;
}

const VERIFIED_NAMES = ['verified', 'verificado'];
const UNVERIFIED_NAMES = [
    'not verified',
    'unverified',
    'not-verified',
    'no verificado',
    'pending',
];

export function isVerifiedStatusLabel(label) {
    const text = String(label || '')
        .trim()
        .toLowerCase();
    if (!text) return false;
    if (UNVERIFIED_NAMES.some((name) => text === name || text.includes(name))) {
        return false;
    }
    return VERIFIED_NAMES.some((name) => text === name);
}

export function getSelectChoices(field) {
    return field?.config?.options?.choices || field?.options?.choices || [];
}

export function isWritableTextField(field) {
    const type = getFieldType(field);
    return (
        type === FieldType.SINGLE_LINE_TEXT ||
        type === FieldType.MULTILINE_TEXT ||
        type === FieldType.RICH_TEXT ||
        type === FieldType.EMAIL ||
        type === FieldType.URL ||
        type === FieldType.PHONE_NUMBER
    );
}

export function isWritableCampusField(field) {
    if (!field || field.isComputed) return false;
    return isWritableTextField(field) || isSelectField(field);
}

export function readFormValueForField(record, field) {
    if (!field) return '';
    const type = getFieldType(field);
    if (!record) {
        return type === FieldType.MULTIPLE_SELECTS ? [] : '';
    }
    let cell = null;
    try {
        cell = record.getCellValue(field);
    } catch {
        cell = null;
    }
    if (type === FieldType.SINGLE_SELECT) {
        if (!cell) return '';
        return cell.id || cell.name || '';
    }
    if (type === FieldType.MULTIPLE_SELECTS) {
        if (!Array.isArray(cell)) return [];
        return cell.map((item) => item?.id || item?.name).filter(Boolean);
    }
    return readCellDisplayName(record, field);
}

export function writeValueForField(field, formValue) {
    const type = getFieldType(field);
    if (type === FieldType.SINGLE_SELECT) {
        if (!formValue) return null;
        const choice = getSelectChoices(field).find(
            (item) => item.id === formValue || item.name === formValue,
        );
        return writeSelectChoice(choice) || {name: String(formValue)};
    }
    if (type === FieldType.MULTIPLE_SELECTS) {
        const values = Array.isArray(formValue) ? formValue : formValue ? [formValue] : [];
        const choices = getSelectChoices(field);
        return values.map((value) => {
            const choice = choices.find((item) => item.id === value || item.name === value);
            return writeSelectChoice(choice) || {name: String(value)};
        });
    }
    return formValue == null ? '' : String(formValue);
}

export function canUpdateCampusField(table, record, field) {
    if (!table || !record || !isWritableCampusField(field)) return false;
    const sample = writeValueForField(field, readFormValueForField(record, field));
    if (typeof table.hasPermissionToUpdateRecord === 'function') {
        return Boolean(table.hasPermissionToUpdateRecord(record, {[field.id]: sample}));
    }
    if (typeof table.hasPermissionToUpdateRecords !== 'function') return true;
    return Boolean(
        table.hasPermissionToUpdateRecords([
            {
                id: record.id,
                fields: {[field.id]: sample},
            },
        ]),
    );
}

export function findSelectChoice(field, nameHints) {
    const choices = getSelectChoices(field);
    const hints = nameHints.map((n) => String(n).toLowerCase());
    return (
        choices.find((choice) => hints.includes(String(choice.name || '').toLowerCase())) ||
        choices.find((choice) =>
            hints.some((hint) => String(choice.name || '').toLowerCase().includes(hint)),
        ) ||
        null
    );
}

export function verifiedChoice(field) {
    return findSelectChoice(field, ['verified', 'verificado']);
}

export function unverifiedChoice(field) {
    return findSelectChoice(field, [
        'not verified',
        'unverified',
        'no verificado',
        'pending',
    ]);
}

export function writeSelectChoice(choice) {
    if (!choice) return null;
    if (choice.id) return {id: choice.id, name: choice.name};
    if (choice.name) return {name: choice.name};
    return null;
}

export function writeTimestampForField(field, date = new Date()) {
    const type = getFieldType(field);
    if (type === FieldType.DATE) {
        return date.toISOString().slice(0, 10);
    }
    return date.toISOString();
}

export const VERIFY_STALE_MS = 4000;

export function isRampRecordVerified(record, statusField, lastModifiedField, verifiedAtField) {
    if (!record || !statusField) return false;
    if (!isVerifiedStatusLabel(getSelectName(record.getCellValue(statusField)))) return false;
    if (!lastModifiedField || !verifiedAtField) return true;
    const modified = getCellTimestamp(record, lastModifiedField);
    const verifiedAt = getCellTimestamp(record, verifiedAtField);
    if (modified == null || verifiedAt == null) return true;
    return modified <= verifiedAt + VERIFY_STALE_MS;
}

export async function updateRecordsInChunks(table, updates) {
    if (!table || !updates?.length) return;
    for (let i = 0; i < updates.length; i += 50) {
        await table.updateRecordsAsync(updates.slice(i, i + 50));
    }
}

export function isOver100Mva(mva) {
    return toNumber(mva) >= MVA_THRESHOLD;
}

export function compareCampusRows(a, b) {
    const campus = String(a.campus || '').localeCompare(String(b.campus || ''), undefined, {
        sensitivity: 'base',
    });
    if (campus !== 0) return campus;
    return String(a.metro || '').localeCompare(String(b.metro || ''), undefined, {
        sensitivity: 'base',
    });
}
