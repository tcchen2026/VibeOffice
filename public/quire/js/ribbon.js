/* Quire — the ribbon of the Paper 2016 look (common/ribbon.js): Word's tabs and groups over Quire's commands.
 * Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { autocorrectDlg: 'options', clearFormats: 'eraser', goTo: 'forward', insertBreak: 'pageBreak', bookmark: 'bookmark', crossReference: 'crossRef', pageNumbers: 'pageNum', insertField: 'field',
      insertObject: 'object', insertFile: 'insertFile', themeDlg: 'themes', autoFormat: 'tableAutoFormat', printedWatermark: 'watermark', hyphenation: 'hyphenation', tabsDlg: 'tabStops',
      indexTables: 'toc', insertFootnoteNow: 'footnote', insertEndnoteNow: 'endnote', footnotes: 'footnote', insertCaption: 'caption', letterWizard: 'templates', thesaurus: 'thesaurus',
      setLanguage: 'language', deleteAllComments: 'delete', showComments: 'comment', compareDocs: 'compare', newWindow: 'newWindow', arrangeAll: 'arrangeAll', splitWindow: 'split',
      macros: 'macros', headingRows: 'headingRows', splitTable: 'splitTable', tableToText: 'tableToText', formulaDlg: 'fx', wordartEdit: 'editText', hfClose: 'close' },
    qat: ['save', { split: 'undo', menu: (r) => A().historyMenu(r, 'undo') }, { split: 'redo', menu: (r) => A().historyMenu(r, 'redo') }],
    file: () => ['new', 'open', 'close', 'templates', 'fileSearch', '-', 'save', 'saveAs', 'saveWeb', 'webPreview', '-', 'properties', 'permission', 'links', '-', 'pageSetup', 'printPreview', 'print', '-', 'exportPDF', 'sendMail', '-', 'optionsDlg', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '&Home', groups: [
        { label: 'Clipboard', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteSpecial', 'pasteHyperlink']) }, 'cut', 'copy', 'painter'] },
        { label: 'Font', icon: 'font', launcher: 'fontDlg', items: [
          { row: [{ make: () => A().parts.font({ width: 128, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 42, id: 'rb-size' }) }, 'growFont', 'shrinkFont', 'changeCase', 'clearFormats'] },
          { row: ['bold', 'italic', 'underline', 'strike', 'subscript', 'superscript', '|', { split: 'highlightBtn', menu: (r, b) => A().menus.highlight(r, b) }, { split: 'fontColorBtn', menu: (r, b) => A().menus.fontColor(r, b) }] }] },
        { label: 'Paragraph', icon: 'paragraph', launcher: 'paragraphDlg', items: [
          { row: [{ split: 'bullets', menu: open(['bulletsDlg']) }, 'numbering', 'decIndent', 'incIndent', '|', 'sortDlg', 'showMarks'] },
          { row: ['alignLeft', 'alignCenter', 'alignRight', 'justify', '|', { drop: 'Line Spacing', icon: 'lineSpacing', menu: (r) => A().menus.lineSpacing(r) }, { split: 'borderBtn', menu: (r, b) => A().menus.borders(r, b) }] }] },
        { label: 'Styles', launcher: 'stylesPane', items: [{ make: () => A().parts.style({ width: 150, id: 'rb-style' }) }, 'stylesPane', 'revealFormatting'] },
        { label: 'Editing', items: ['find', 'replace', { drop: 'Select', icon: 'selectAll', menu: open(['selectAll', 'goTo']) }] },
      ] },
      { id: 'insert', label: '&Insert', groups: [
        { label: 'Pages', items: [{ cmd: 'insertBreak', size: 'large', label: 'Break' }] },
        { label: 'Tables', items: [{ drop: 'Table', icon: 'table', size: 'large', menu: (r) => A().menus.tableGrid(r, ['-', 'insertTable', 'drawTable', 'textToTable']) }] },
        { label: 'Illustrations', items: [{ cmd: 'insertPicture', size: 'large', label: 'Pictures' }, 'insertClipArt', { drop: 'Shapes', icon: 'autoshapes', menu: (r) => A().openAutoShapes(r) }, 'insertDiagram', 'insertChart'] },
        { label: 'Links', items: ['hyperlink', 'bookmark', 'crossReference'] },
        { label: 'Comments', items: [{ cmd: 'insertComment', size: 'large', label: 'Comment' }] },
        { label: 'Header & Footer', items: ['headerFooter', 'pageNumbers'] },
        { label: 'Text', items: [{ cmd: 'textBox', size: 'large', label: 'Text Box' }, 'insertWordArt', 'dropCap', 'dateTime', 'insertField', 'insertObject', 'insertFile'] },
        { label: 'Symbols', items: [{ cmd: 'insertSymbol', size: 'large', label: 'Symbol' }] },
      ] },
      { id: 'design', label: '&Design', groups: [
        { label: 'Document Formatting', items: [{ cmd: 'themeDlg', size: 'large', label: 'Themes' }, 'autoFormat'] },
        { label: 'Page Background', items: [{ cmd: 'printedWatermark', size: 'large', label: 'Watermark' }, { cmd: 'bordersDlg', size: 'large', label: 'Page Borders' }] },
      ] },
      { id: 'layout', label: 'Layout', groups: [
        { label: 'Page Setup', launcher: 'pageSetup', items: [{ cmd: 'pageSetup', size: 'large', label: 'Margins' }, { drop: 'Columns', icon: 'columns', size: 'large', menu: (r) => A().menus.columns(r, ['-', 'columnsDlg']) }, 'insertBreak', 'hyphenation', 'textDirection'] },
        { label: 'Paragraph', icon: 'paragraph', launcher: 'paragraphDlg', items: ['incIndent', 'decIndent', 'tabsDlg'] },
        { label: 'Arrange', items: ['bringFront', 'sendBack', { drop: 'Wrap Text', icon: 'textWrap', menu: (r) => A().wrapMenu(r) }, 'groupObj', 'ungroupObj', 'rotateRight'] },
      ] },
      { id: 'references', label: 'Refe&rences', groups: [
        { label: 'Table of Contents', items: [{ cmd: 'indexTables', size: 'large', label: 'Table of Contents' }, 'updateTOC'] },
        { label: 'Footnotes', launcher: 'insertFootnote', items: [{ cmd: 'insertFootnoteNow', size: 'large', label: 'Insert Footnote' }, 'insertEndnoteNow', 'footnotes'] },
        { label: 'Captions', items: [{ cmd: 'insertCaption', size: 'large', label: 'Insert Caption' }, 'crossReference'] },
      ] },
      { id: 'mailings', label: 'Mailin&gs', groups: [
        { label: 'Create', items: [{ cmd: 'envelopes', size: 'large', label: 'Envelopes' }, 'letterWizard'] },
        { label: 'Start Mail Merge', items: [{ cmd: 'mailMerge', size: 'large', label: 'Start Mail Merge' }, 'mmOpenSource', 'mmRecipients'] },
        { label: 'Write & Insert Fields', items: ['mmAddress', 'mmGreeting', 'mmInsertField'] },
        { label: 'Preview Results', items: [{ cmd: 'mmViewData', size: 'large', label: 'Preview Results' }, { row: ['mmFirst', 'mmPrev', 'mmNext', 'mmLast'] }] },
        { label: 'Finish', items: [{ cmd: 'mmMergeNew', size: 'large', label: 'Finish & Merge' }, 'mmMergePrint'] },
      ] },
      { id: 'review', label: 'Re&view', groups: [
        { label: 'Proofing', items: [{ cmd: 'spelling', size: 'large', label: 'Spelling & Grammar' }, 'thesaurus', 'wordCount', 'autocorrectDlg', 'autoSummarize'] },
        { label: 'Insights', items: [{ cmd: 'research', size: 'large', label: 'Research' }] },
        { label: 'Language', items: [{ cmd: 'setLanguage', size: 'large', label: 'Language' }] },
        { label: 'Comments', items: [{ cmd: 'insertComment', size: 'large', label: 'New Comment' }, 'deleteAllComments', 'showComments'] },
        { label: 'Tracking', items: [{ cmd: 'trackChanges', size: 'large', label: 'Track Changes' },
          { drop: 'Final Showing Markup', icon: 'reveal', menu: open(['markupFinalMarkup', 'markupFinal', 'markupOriginalMarkup', 'markupOriginal']), tip: 'Display for Review' },
          { drop: 'Show Markup', icon: 'comment', menu: open(['markup', '-', 'showComments', 'showInsDel', 'showFormatting', '-', 'reviewersMenu', '-', 'balloonsToggle']) }, 'reviewingPane'] },
        { label: 'Changes', items: [{ split: 'acceptChange', size: 'large', label: 'Accept', menu: open(['acceptChange', 'acceptShown', 'acceptAll']) }, { split: 'rejectChange', label: 'Reject', menu: open(['rejectChange', 'rejectShown', 'rejectAll', '-', 'deleteAllComments']) }, 'prevChange', 'nextChange'] },
        { label: 'Compare', items: [{ cmd: 'compareDocs', size: 'large', label: 'Compare' }] },
        { label: 'Protect', items: [{ cmd: 'protectDoc', size: 'large', label: 'Protect Document' }] },
      ] },
      { id: 'view', label: 'Vie&w', groups: [
        { label: 'Views', items: [{ cmd: 'viewReading', size: 'large', label: 'Read Mode' }, { cmd: 'viewPrint', size: 'large', label: 'Print Layout' }, { cmd: 'viewWeb', size: 'large', label: 'Web Layout' }, 'viewOutline', 'viewNormal'] },
        { label: 'Show', items: ['ruler', 'docMap', 'thumbnails', 'taskPane'] },
        { label: 'Zoom', items: [{ cmd: 'zoomDlg', size: 'large', label: 'Zoom' }, { make: () => A().parts.zoom({ width: 90, id: 'rb-zoom' }) }, 'fullScreen'] },
        { label: 'Window', items: [{ cmd: 'newWindow', size: 'large', label: 'New Window' }, 'arrangeAll', 'splitWindow'] },
        { label: 'Macros', items: [{ cmd: 'macros', size: 'large', label: 'Macros' }] },
      ] },
    ],
    contextual: [
      { id: 'table', set: 'Table Tools', when: ctx('table'), tabs: [
        { id: 'tableDesign', label: 'Design', groups: [
          { label: 'Table Styles', items: [{ cmd: 'tableAutoFormat', size: 'large', label: 'Table Styles' }, 'headingRows'] },
          { label: 'Borders', launcher: 'bordersDlg', items: [{ drop: 'Border Styles', icon: 'dashStyle', menu: (r) => A().borderStyleMenu(r) }, { drop: 'Pen Weight', icon: 'lineStyle', menu: (r) => A().borderWeightMenu(r) }, { split: 'borderBtn', size: 'large', label: 'Borders', menu: (r, b) => A().menus.borders(r, b) }] },
        ] },
        { id: 'tableLayout', label: 'Layout', groups: [
          { label: 'Table', items: [{ drop: 'Select', icon: 'selectTable', menu: open(['selectCell', 'selectCol', 'selectRow', 'selectTable']) }, 'gridlines', { cmd: 'tableProps', size: 'large', label: 'Properties' }] },
          { label: 'Draw', items: [{ cmd: 'drawTable', size: 'large', label: 'Draw Table' }, { cmd: 'eraser', size: 'large', label: 'Eraser' }] },
          { label: 'Rows & Columns', launcher: 'insertCells', items: [{ drop: 'Delete', icon: 'deleteTable', size: 'large', menu: open(['deleteCells', 'deleteCols', 'deleteRows', 'deleteTable']) }, 'rowAbove', 'rowBelow', 'colLeft', 'colRight'] },
          { label: 'Merge', items: ['mergeCells', 'splitCells', 'splitTable'] },
          { label: 'Cell Size', items: [{ drop: 'AutoFit', icon: 'autofit', menu: open(['autofitContents', 'autofitWindow', 'fixedWidth']) }, 'distRows', 'distCols'] },
          { label: 'Alignment', items: ['cellTop', 'cellMiddle', 'cellBottom', 'textDirection'] },
          { label: 'Data', items: [{ cmd: 'sortDlg', size: 'large', label: 'Sort' }, 'headingRows', 'tableToText', 'formulaDlg'] },
        ] },
      ] },
      { id: 'picture', set: 'Picture Tools', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: 'Format', groups: [
          { label: 'Adjust', items: [{ row: ['picMoreBright', 'picLessBright', 'picMoreContrast', 'picLessContrast'] }, { drop: 'Color', icon: 'colorPic', menu: open([{ label: '&Automatic', run: () => A().picMode('') }, { label: '&Grayscale', run: () => A().picMode('gray') }, { label: '&Black & White', run: () => A().picMode('bw') }, { label: '&Washout', run: () => A().picMode('washout') }]) }, 'picReset'] },
          { label: 'Arrange', items: [{ drop: 'Wrap Text', icon: 'textWrap', size: 'large', menu: (r) => A().wrapMenu(r) }, 'bringFront', 'sendBack', 'picRotate'] },
          { label: 'Size', launcher: 'formatObject', items: [{ cmd: 'picCrop', size: 'large', label: 'Crop' }] },
        ] },
      ] },
      { id: 'drawing', set: 'Drawing Tools', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: 'Format', groups: [
          { label: 'Insert Shapes', items: [{ drop: 'Shapes', icon: 'autoshapes', size: 'large', menu: (r) => A().openAutoShapes(r) }, 'textBox'] },
          { label: 'Shape Styles', launcher: 'formatObject', items: [{ split: 'fillColor', menu: (r, b) => A().objColorMenu(b, 'fill') }, { split: 'lineColor', menu: (r, b) => A().objColorMenu(b, 'line') }, { drop: 'Shadow', icon: 'shadowStyle', menu: (r) => A().objShadowMenu(r) }] },
          { label: 'LettersArt Styles', items: [{ cmd: 'wordartGallery', size: 'large', label: 'LettersArt Styles' }, 'wordartEdit'] },
          { label: 'Arrange', items: [{ drop: 'Wrap Text', icon: 'textWrap', menu: (r) => A().wrapMenu(r) }, 'bringFront', 'sendBack', 'groupObj', 'ungroupObj', 'rotateRight'] },
        ] },
      ] },
      { id: 'headerFooter', set: 'Header & Footer Tools', open: true, when: ctx('headerFooter'), tabs: [
        { id: 'hfDesign', label: 'Design', groups: [
          { label: 'Header & Footer', items: [{ cmd: 'hfPageNum', size: 'large', label: 'Page Number' }, 'hfNumPages', 'hfFormatPageNum'] },
          { label: 'Insert', items: [{ cmd: 'hfDate', size: 'large', label: 'Date' }, 'hfTime'] },
          { label: 'Navigation', items: [{ cmd: 'hfSwitch', size: 'large', label: 'Go to Footer' }, 'hfPrev', 'hfNext', 'hfLinkPrev'] },
          { label: 'Options', items: ['hfShowText', 'pageSetup'] },
          { label: 'Close', items: [{ cmd: 'hfClose', size: 'large', label: 'Close Header and Footer' }] },
        ] },
      ] },
      { id: 'outlining', set: 'Outline', open: true, when: ctx('outlining'), tabs: [
        { id: 'outlining', label: 'Outlining', groups: [
          { label: 'Outline Tools', items: [{ row: ['promoteH1', 'promote', 'demote', 'demoteBody'] }, { row: ['moveUp', 'moveDown', 'expand', 'collapse'] }, 'olFirstLine', 'olFormatting'] },
          { label: 'Table of Contents', items: ['updateTOC', 'gotoTOC'] },
          { label: 'Close', items: [{ cmd: 'viewPrint', size: 'large', label: 'Close Outline View' }] },
        ] },
      ] },
    ],
  });
})();
