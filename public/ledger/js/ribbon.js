/* Ledger — the ribbon of the Paper 2016 look (common/ribbon.js): Excel's tabs and groups over Ledger's commands.
 * Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { autoCorrectDlg: 'options', taskPane: 'taskpane', pasteValues: 'paste', fillSeries: 'fillColor', clearAll: 'eraser', goTo: 'forward', insertSheet: 'insertSheet', moveCopySheet: 'insertSheet', nameDefine: 'nameBox', pasteName: 'nameBox',
      nameCreate: 'nameBox', nameApply: 'nameBox', functionsHelp: 'help', calcSheet: 'calcNow', errorChecking: 'traceError', showAll: 'filter', advancedFilter: 'filterOn', consolidate: 'subtotal',
      groupRows: 'group', ungroupRows: 'ungroup', showDetail: 'expand', hideDetail: 'collapse', autoOutline: 'group', clearOutline: 'ungroup', importText: 'textToCols', refreshData: 'redo',
      protectWorkbook: 'protect', allowEditRanges: 'protect', deleteComment: 'delete', showHideComment: 'comment', showComments: 'comment', viewNormal: 'normalGrid', newWindow: 'newWindow',
      arrangeWindows: 'arrangeAll', macros: 'macros', formulaBar: 'fx', statusBar: 'layout', gridlinesView: 'grid', headingsView: 'headingRows', setPrintArea: 'pageSetup', insertPageBreak: 'pageBreak',
      resetPageBreaks: 'pageBreak', sheetBackground: 'background', createList: 'table', totalRow: 'autoSum', convertToRange: 'tableToText', newSheetChart: 'chart', exportCSV: 'excel', conditionalFormat: 'condFormat', styleDlg: 'stylesPane', headerFooter: 'headerFooter', formatObject: 'formatPicture' },
    qat: ['save', 'undo', 'redo'],
    file: () => ['newBook', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'properties', 'links', '-', 'pageSetup', 'printPreview', 'print', '-', 'exportPDF', 'exportCSV', '-', 'options', 'addIns', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '&Home', groups: [
        { label: 'Clipboard', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteValues', 'pasteSpecial', 'pasteHyperlink']) }, 'cut', 'copy', 'formatPainter'] },
        { label: 'Font', icon: 'font', launcher: 'formatCells', items: [
          { row: [{ make: () => A().parts.font({ width: 132, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 44, id: 'rb-size' }) }] },
          { row: ['bold', 'italic', 'underline', 'strike', '|', { split: 'bordersApply', menu: (r, b) => A().menus.borders(r, b) }, { split: 'fillColorApply', menu: (r, b) => A().menus.fill(r, b) }, { split: 'fontColorApply', menu: (r, b) => A().menus.fontColor(r, b) }] }] },
        { label: 'Alignment', launcher: 'formatCells', items: [
          { row: ['alignLeft', 'alignCenter', 'alignRight', '|', 'decIndent', 'incIndent'] },
          { row: [{ cmd: 'wrapText', size: 'medium', label: 'Wrap Text' }, { cmd: 'mergeCenter', size: 'medium', label: 'Merge & Center' }] }] },
        { label: 'Number', launcher: 'formatCells', items: [
          { row: [{ drop: 'Number Format', icon: 'formatCells', size: 'medium', menu: open(['fmtGeneral', 'fmtNumber', 'fmtCurrency', 'fmtDate', 'fmtTime', 'fmtExp', '-', 'formatCells']) }] },
          { row: ['currencyStyle', 'percentStyle', 'commaStyle', '|', 'incDecimal', 'decDecimal'] }] },
        { label: 'Styles', items: [{ cmd: 'conditionalFormat', size: 'large', label: 'Conditional Formatting' }, { cmd: 'autoFormat', size: 'large', label: 'Format as Table' }, { cmd: 'styleDlg', size: 'large', label: 'Cell Styles' }] },
        { label: 'Cells', items: [
          { drop: 'Insert', icon: 'insertCells', size: 'large', menu: open(['insertCells', 'insertRows', 'insertCols', 'insertSheet']) },
          { drop: 'Delete', icon: 'deleteSheet', size: 'large', menu: open(['deleteCells', 'deleteSheet']) },
          { drop: 'Format', icon: 'formatCells', size: 'large', menu: open([{ label: 'Row', sub: ['rowHeight', 'rowAutofit', 'rowHide', 'rowUnhide'] }, { label: 'Column', sub: ['colWidth', 'colAutofit', 'colHide', 'colUnhide', 'colStandard'] }, { label: 'Sheet', sub: ['sheetRename', 'sheetHide', 'sheetUnhide', 'moveCopySheet', 'tabColor'] }, '-', 'protectSheet', 'formatCells']) }] },
        { label: 'Editing', items: [
          { split: 'autoSum', label: 'AutoSum', menu: (r) => A().menus.autoSum(r) },
          { drop: 'Fill', icon: 'fillColor', menu: open(['fillDown', 'fillRight', 'fillUp', 'fillLeft', 'fillAcross', 'fillSeries', 'fillJustify']) },
          { drop: 'Clear', icon: 'eraser', menu: open(['clearAll', 'clearFormats', 'clearContents', 'clearComments', 'clearHyperlinks']) },
          { drop: 'Sort & Filter', icon: 'sortAsc', size: 'large', menu: open(['sortAsc', 'sortDesc', 'sortDlg', '-', 'autoFilter', 'showAll', 'advancedFilter']) },
          { drop: 'Find & Select', icon: 'find', size: 'large', menu: open(['find', 'replace', 'goTo', 'goToSpecial']) }] },
      ] },
      { id: 'insert', label: '&Insert', groups: [
        { label: 'Tables', items: [{ cmd: 'pivotTable', size: 'large', label: 'PivotTable' }, { cmd: 'createList', size: 'large', label: 'Table' }] },
        { label: 'Illustrations', items: [{ cmd: 'insertPicture', size: 'large', label: 'Pictures' }, 'insertClipArt', { cmd: 'insertShape', label: 'Shapes', icon: 'autoshapes' }] },
        { label: 'Charts', items: [{ cmd: 'chartWizard', size: 'large', label: 'Chart' }, 'insertChart', 'newSheetChart'] },
        { label: 'Links', items: [{ cmd: 'hyperlink', size: 'large', label: 'Hyperlink' }] },
        { label: 'Text', items: [{ cmd: 'insertTextBox', size: 'large', label: 'Text Box' }, 'headerFooter'] },
        { label: 'Symbols', items: [{ cmd: 'insertSymbol', size: 'large', label: 'Symbol' }] },
      ] },
      { id: 'pageLayout', label: 'Page Layout', groups: [
        { label: 'Page Setup', launcher: 'pageSetup', items: [{ cmd: 'pageSetup', size: 'large', label: 'Margins' }, { drop: 'Print Area', icon: 'pageSetup', size: 'large', menu: open(['setPrintArea', 'clearPrintArea']) },
          { drop: 'Breaks', icon: 'pageBreak', size: 'large', menu: open(['insertPageBreak', 'resetPageBreaks']) }, { cmd: 'sheetBackground', size: 'large', label: 'Background' }] },
        { label: 'Sheet Options', items: [{ cmd: 'gridlinesView', label: 'Gridlines' }, { cmd: 'headingsView', label: 'Headings' }] },
        { label: 'Arrange', items: [{ drop: 'Bring Forward / Send Backward', icon: 'bringFront', size: 'large', label: 'Arrange', menu: (r) => A().menus.draw(r) }] },
      ] },
      { id: 'formulas', label: 'Formulas', groups: [
        { label: 'Function Library', items: [{ cmd: 'insertFunction', size: 'large', label: 'Insert Function' }, { split: 'autoSum', size: 'large', label: 'AutoSum', menu: (r) => A().menus.autoSum(r) }, 'functionsHelp'] },
        { label: 'Defined Names', items: [{ cmd: 'nameDefine', size: 'large', label: 'Define Name' }, 'pasteName', 'nameCreate', 'nameApply'] },
        { label: 'Formula Auditing', items: ['tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas', 'traceError', 'evaluateFormula', { cmd: 'watchWindow', size: 'large', label: 'Watch Window' }] },
        { label: 'Calculation', items: [{ cmd: 'calcNow', size: 'large', label: 'Calculate Now' }, 'calcSheet'] },
      ] },
      { id: 'data', label: '&Data', groups: [
        { label: 'Get External Data', items: [{ cmd: 'importText', size: 'large', label: 'From Text' }, { cmd: 'refreshData', size: 'large', label: 'Refresh All' }] },
        { label: 'Sort & Filter', items: [{ row: ['sortAsc', 'sortDesc'] }, { cmd: 'sortDlg', size: 'large', label: 'Sort' }, { cmd: 'autoFilter', size: 'large', label: 'Filter' }, 'showAll', 'advancedFilter'] },
        { label: 'Data Tools', items: [{ cmd: 'textToColumns', size: 'large', label: 'Text to Columns' }, { split: 'validation', menu: open(['validation', 'circleInvalid', 'clearCircles']) }, 'consolidate', 'dataForm'] },
        { label: 'Forecast', items: [{ drop: 'What-If Analysis', icon: 'goalSeek', size: 'large', menu: open(['scenarios', 'goalSeek', 'dataTable']) }] },
        { label: 'Outline', launcher: 'outlineSettings', items: [{ cmd: 'groupRows', size: 'large', label: 'Group' }, { split: 'ungroupRows', size: 'large', label: 'Ungroup', menu: open(['ungroupRows', 'clearOutline']) }, { cmd: 'subtotals', size: 'large', label: 'Subtotal' }, 'showDetail', 'hideDetail', 'autoOutline'] },
      ] },
      { id: 'review', label: 'Re&view', groups: [
        { label: 'Proofing', items: [{ cmd: 'spelling', size: 'large', label: 'Spelling' }, 'errorChecking', 'autoCorrectDlg'] },
        { label: 'Comments', items: [{ cmd: 'insertComment', size: 'large', label: 'New Comment' }, 'deleteComment', 'showHideComment', 'showComments'] },
        { label: 'Changes', items: [{ cmd: 'protectSheet', size: 'large', label: 'Protect Sheet' }, { cmd: 'protectWorkbook', size: 'large', label: 'Protect Workbook' }, 'allowEditRanges'] },
      ] },
      { id: 'view', label: 'Vie&w', groups: [
        { label: 'Workbook Views', items: [{ cmd: 'viewNormal', size: 'large', label: 'Normal' }, { cmd: 'viewPageBreak', size: 'large', label: 'Page Break Preview' }, 'fullScreen'] },
        { label: 'Show', items: ['formulaBar', { cmd: 'gridlinesView', label: 'Gridlines' }, { cmd: 'headingsView', label: 'Headings' }, 'statusBar', 'taskPane'] },
        { label: 'Zoom', items: [{ cmd: 'zoomDlg', size: 'large', label: 'Zoom' }, { make: () => A().parts.zoom({ width: 70, id: 'rb-zoom' }) }] },
        { label: 'Window', items: [{ cmd: 'newWindow', size: 'large', label: 'New Window' }, 'arrangeWindows', { cmd: 'freezePanes', size: 'large', label: 'Freeze Panes' }, 'splitWindow'] },
        { label: 'Macros', items: [{ cmd: 'macros', size: 'large', label: 'Macros' }] },
      ] },
    ],
    contextual: [
      { id: 'chart', set: 'Chart Tools', when: ctx('chart'), tabs: [
        { id: 'chartDesign', label: 'Design', groups: [
          { label: 'Type', items: [{ drop: 'Change Chart Type', icon: 'chart', size: 'large', menu: (r) => L.chartWizard && L.chartWizard.typeMenu(r) }] },
          { label: 'Data', items: [{ cmd: 'chartWizard', size: 'large', label: 'Chart Wizard' }] },
        ] },
        { id: 'chartFormat', label: 'Format', groups: [{ label: 'Current Selection', items: [{ cmd: 'formatObject', size: 'large', label: 'Format Selection' }] }, { label: 'Arrange', items: [{ drop: 'Arrange', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] }] },
      ] },
      { id: 'picture', set: 'Picture Tools', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: 'Format', groups: [{ label: 'Picture', items: [{ cmd: 'formatObject', size: 'large', label: 'Format Picture' }] }, { label: 'Arrange', items: [{ drop: 'Arrange', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] }] },
      ] },
      { id: 'drawing', set: 'Drawing Tools', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: 'Format', groups: [
          { label: 'Insert Shapes', items: [{ cmd: 'insertShape', size: 'large', label: 'Shapes', icon: 'autoshapes' }, 'insertTextBox'] },
          { label: 'Shape Styles', launcher: 'formatObject', items: [{ cmd: 'fillColorApply', size: 'large', label: 'Shape Fill' }] },
          { label: 'Arrange', items: [{ drop: 'Arrange', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] },
        ] },
      ] },
      { id: 'list', set: 'Table Tools', when: ctx('list'), tabs: [
        { id: 'listDesign', label: 'Design', groups: [
          { label: 'Tools', items: [{ cmd: 'convertToRange', size: 'large', label: 'Convert to Range' }, 'sortDlg', 'autoFilter'] },
          { label: 'Table Style Options', items: [{ cmd: 'totalRow', label: 'Total Row' }, 'insertRows', 'insertCols'] },
        ] },
      ] },
    ],
  });
})();
