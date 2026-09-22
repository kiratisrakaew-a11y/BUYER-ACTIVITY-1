/**
 * Schema.js — THE single place where every table and column is defined (SPEC §10.1).
 *
 * A future module adds a table or a column by editing this file only.
 * Setup.js walks this definition to create sheets and to append columns that are
 * missing from an existing sheet, so re-running setup() after a schema change is safe.
 *
 * Column descriptor:
 *   name      — header text written into row 1 of the sheet
 *   type      — 'string' | 'text' | 'code' | 'email' | 'url' | 'date' | 'datetime'
 *               | 'number' | 'int' | 'bool'
 *   required  — server-side validation rejects empty values
 *   list      — Config_Lists.List_Name this code must belong to
 *   parent    — the column whose value filters `list` via Config_Lists.Parent_Code
 *   reserved  — column exists in the sheet but no M1 UI writes it (SPEC §10.6)
 *   max       — maximum string length accepted from the client
 *   pdpa      — personal data; never included in notification e-mails (SPEC §11)
 *
 * Table descriptor:
 *   sheet     — sheet name inside the DB spreadsheet
 *   pk        — primary key column
 *   id        — how IdGenerator mints the pk; omitted when the pk is user-supplied
 *   audit     — append the six audit columns (SPEC §5.0)
 *   caseIdCol — column linking a row back to its Case, used by Change_Log and by
 *               Repository's indexed lookup
 *   module    — the module that owns the table, for documentation
 */
var Schema = (function () {

  /** SPEC §5.0 — present on every transaction table, always at the far right. */
  var AUDIT_COLUMNS = [
    { name: 'Created_At', type: 'datetime' },
    { name: 'Created_By', type: 'email' },
    { name: 'Updated_At', type: 'datetime' },
    { name: 'Updated_By', type: 'email' },
    { name: 'Version', type: 'int' },
    { name: 'Is_Deleted', type: 'bool' }
  ];

  var AUDIT_COLUMN_NAMES = AUDIT_COLUMNS.map(function (c) { return c.name; });

  var TABLES = {

    /* ---------------------------------------------------------- transaction */

    Cases: {
      sheet: 'Cases',
      pk: 'Case_ID',
      id: { prefix: 'SRC', yearly: true, pad: 4 },
      audit: true,
      caseIdCol: 'Case_ID',
      module: 'M1',
      columns: [
        { name: 'Case_ID', type: 'string', required: true },
        { name: 'Request_Date', type: 'date', required: true },
        { name: 'Request_Ref', type: 'string', required: true, max: 200 },
        { name: 'Requester_Name', type: 'string', required: true, max: 150 },
        { name: 'Requester_Email', type: 'email', max: 150 },
        { name: 'Department_Code', type: 'code', required: true, list: 'DEPARTMENT' },
        { name: 'Method', type: 'code', required: true, list: 'METHOD' },
        { name: 'Budget_Type', type: 'code', required: true, list: 'BUDGET_TYPE' },
        { name: 'Sub_Type', type: 'code', required: true, list: 'SUB_TYPE', parent: 'Budget_Type' },
        { name: 'Description', type: 'text', required: true, max: 2000 },
        { name: 'Required_Date', type: 'date' },
        { name: 'Intake_Complete', type: 'bool', required: true },
        { name: 'Intake_Note', type: 'text', max: 2000 },
        { name: 'Buyer_Owner', type: 'email', required: true },
        { name: 'Status', type: 'code', required: true },
        { name: 'PR_No', type: 'string', max: 50 },
        { name: 'PR_Date', type: 'date' },
        { name: 'Company_Code', type: 'code', list: 'COMPANY' },
        { name: 'Department_Feedback', type: 'text', reserved: true, max: 2000 },
        { name: 'Drive_Folder_ID', type: 'string' },
        { name: 'Closed_At', type: 'datetime' }
      ]
    },

    Activities: {
      sheet: 'Activities',
      pk: 'Activity_ID',
      id: { prefix: 'ACT', pad: 6 },
      audit: true,
      caseIdCol: 'Case_ID',
      module: 'M1',
      columns: [
        { name: 'Activity_ID', type: 'string', required: true },
        { name: 'Case_ID', type: 'string', required: true },
        { name: 'Vendor_ID', type: 'string' },
        { name: 'Module', type: 'code', required: true },
        { name: 'Activity_Date', type: 'datetime', required: true },
        { name: 'Activity_Type', type: 'code', required: true, list: 'ACTIVITY_TYPE' },
        { name: 'Channel', type: 'code', list: 'CHANNEL' },
        { name: 'Activity_Description', type: 'text', required: true, max: 2000 },
        { name: 'Performed_By', type: 'email', required: true },
        { name: 'Next_Action', type: 'text', max: 1000 },
        { name: 'Next_Action_Date', type: 'date' },
        { name: 'Next_Action_Done', type: 'bool' },
        { name: 'Attachment_URL', type: 'url' }
      ]
    },

    /* --------------------------------------------------------------- master */

    Vendors: {
      sheet: 'Vendors',
      pk: 'Vendor_ID',
      id: { prefix: 'VEN', pad: 5 },
      audit: true,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'Vendor_ID', type: 'string', required: true },
        { name: 'Vendor_No', type: 'string', max: 50 },
        { name: 'Vendor_Name', type: 'string', required: true, max: 250 },
        { name: 'Tax_ID', type: 'string', required: true, max: 13 },
        { name: 'Address', type: 'text', max: 1000 },
        { name: 'Contact_Name', type: 'string', max: 150, pdpa: true },
        { name: 'Contact_Phone', type: 'string', max: 50, pdpa: true },
        { name: 'Contact_Email', type: 'email', max: 150, pdpa: true },
        { name: 'Categories', type: 'string', max: 500 },
        { name: 'Vendor_Status', type: 'code', required: true },
        { name: 'Directors', type: 'text', reserved: true, max: 2000 },
        { name: 'Remark', type: 'text', max: 1000 }
      ]
    },

    Users: {
      sheet: 'Users',
      pk: 'Email',
      audit: false,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'Email', type: 'email', required: true },
        { name: 'Name', type: 'string', required: true, max: 150 },
        { name: 'Role', type: 'code', required: true },
        { name: 'Responsible_Scope', type: 'string', max: 200 },
        { name: 'Is_Active', type: 'bool', required: true }
      ]
    },

    /* --------------------------------------------------------------- config */

    Config_Settings: {
      sheet: 'Config_Settings',
      pk: 'Key',
      audit: false,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'Key', type: 'string', required: true },
        { name: 'Value', type: 'string' },
        { name: 'Description', type: 'text' }
      ]
    },

    Config_Lists: {
      sheet: 'Config_Lists',
      pk: null,
      audit: false,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'List_Name', type: 'string', required: true },
        { name: 'Code', type: 'string', required: true },
        { name: 'Label_TH', type: 'string', required: true },
        { name: 'Parent_Code', type: 'string' },
        { name: 'Sort_Order', type: 'int' },
        { name: 'Is_Active', type: 'bool' }
      ]
    },

    Status_Master: {
      sheet: 'Status_Master',
      pk: 'Status_Code',
      audit: false,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'Status_Code', type: 'string', required: true },
        { name: 'Label_TH', type: 'string', required: true },
        { name: 'Sequence', type: 'int', required: true },
        { name: 'Module', type: 'string' },
        { name: 'Is_Terminal', type: 'bool' },
        { name: 'Allowed_Next', type: 'string' }
      ]
    },

    /* --------------------------------------------------------------- system */

    /** Append-only. No api_* function may ever update or delete a row here (SPEC §5.4). */
    Change_Log: {
      sheet: 'Change_Log',
      pk: 'Log_ID',
      id: { prefix: 'LOG', pad: 8 },
      audit: false,
      caseIdCol: 'Case_ID',
      appendOnly: true,
      module: 'M1',
      columns: [
        { name: 'Log_ID', type: 'string', required: true },
        { name: 'Timestamp', type: 'datetime', required: true },
        { name: 'User', type: 'string', required: true },
        { name: 'Table_Name', type: 'string', required: true },
        { name: 'Record_ID', type: 'string' },
        { name: 'Case_ID', type: 'string' },
        { name: 'Action', type: 'string', required: true },
        { name: 'Field', type: 'string' },
        { name: 'Old_Value', type: 'text' },
        { name: 'New_Value', type: 'text' },
        { name: 'Reason', type: 'text' }
      ]
    },

    Counters: {
      sheet: 'Counters',
      pk: null,
      audit: false,
      caseIdCol: null,
      module: 'M1',
      columns: [
        { name: 'Counter_Name', type: 'string', required: true },
        { name: 'Year', type: 'int' },
        { name: 'Last_No', type: 'int', required: true }
      ]
    }
  };

  /** Sheets are created in this order so the DB spreadsheet reads top-down. */
  var TABLE_ORDER = [
    'Cases', 'Activities', 'Vendors', 'Users',
    'Config_Settings', 'Config_Lists', 'Status_Master', 'Change_Log', 'Counters'
  ];

  function getTable(name) {
    var t = TABLES[name];
    if (!t) throw new Error('Unknown table in Schema: ' + name);
    return t;
  }

  /** Declared columns plus the audit columns, in sheet order. */
  function getColumns(name) {
    var t = getTable(name);
    return t.audit ? t.columns.concat(AUDIT_COLUMNS) : t.columns.slice();
  }

  function getColumnNames(name) {
    return getColumns(name).map(function (c) { return c.name; });
  }

  function getColumn(name, columnName) {
    var cols = getColumns(name);
    for (var i = 0; i < cols.length; i++) {
      if (cols[i].name === columnName) return cols[i];
    }
    return null;
  }

  function isAuditColumn(columnName) {
    return AUDIT_COLUMN_NAMES.indexOf(columnName) !== -1;
  }

  function tableNames() {
    return TABLE_ORDER.slice();
  }

  return {
    TABLES: TABLES,
    TABLE_ORDER: TABLE_ORDER,
    AUDIT_COLUMNS: AUDIT_COLUMNS,
    AUDIT_COLUMN_NAMES: AUDIT_COLUMN_NAMES,
    getTable: getTable,
    getColumns: getColumns,
    getColumnNames: getColumnNames,
    getColumn: getColumn,
    isAuditColumn: isAuditColumn,
    tableNames: tableNames
  };
})();
