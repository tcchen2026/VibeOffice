/* Lectern — the ribbon of the Paper 2016 look (common/ribbon.js): PowerPoint's tabs and groups over Lectern's
 * commands. Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { clear: 'delete', duplicate: 'duplicate', deleteSlide: 'delete', vTextBox: 'vtextbox', slidesFromFiles: 'insertSlides', slidesFromOutline: 'outlineTxt', replaceFonts: 'font',
      viewNotes: 'notesView', viewColor: 'grayscale', viewBW: 'grayscale', zoomFit: 'zoom', closeMaster: 'close', insertTitleMaster: 'master', deleteTitleMaster: 'delete', speakerNotes: 'notesView',
      selectTable: 'selectTable', tableBordersFill: 'tablesBorders', regroup: 'group', alignToSlide: 'alignC', editPoints: 'freeform', summarySlide: 'summary', slidesPane: 'thumbnails',
      gridGuides: 'grid', autocorrectDlg: 'options', duplicateSlide: 'duplicate', subscript: 'subscript', superscript: 'superscript', headerFooter: 'headerFooter',
      pageSetup: 'pageSetup', animSchemes: 'animation' },
    qat: ['save', 'undo', 'redo', 'showFromStart'],
    file: () => ['new', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'properties', '-', 'pageSetup', 'printPreview', 'print', '-', 'optionsDlg', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '&Home', groups: [
        { label: 'Clipboard', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteSpecial']) }, 'cut', { split: 'copy', menu: open(['copy', 'duplicate']) }, 'painter'] },
        { label: 'Slides', items: [{ split: 'newSlide', size: 'large', label: 'New Slide', menu: open(['newSlide', 'duplicateSlide', '-', 'slidesFromFiles', 'slidesFromOutline']) }, { cmd: 'slideLayout', label: 'Layout' }, 'duplicateSlide', 'deleteSlide'] },
        { label: 'Font', icon: 'font', launcher: 'fontDlg', items: [
          { row: [{ make: () => A().parts.font({ width: 130, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 44, id: 'rb-size' }) }, 'growFont', 'shrinkFont', 'changeCase'] },
          { row: ['bold', 'italic', 'underline', 'shadowText', 'subscript', 'superscript', '|', { split: 'fontColor', menu: (r, b) => A().menus.color('font')(r, b) }] }] },
        { label: 'Paragraph', icon: 'paragraph', launcher: 'bulletsDlg', items: [
          { row: ['bullets', 'numbering', '|', 'promote', 'demote', '|', 'lineSpacing'] },
          { row: ['alignLeft', 'alignCenter', 'alignRight', 'justify'] }] },
        { label: 'Drawing', launcher: 'formatObject', items: [
          { drop: 'Shapes', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) },
          { drop: 'Arrange', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) },
          { split: 'fillColor', label: 'Shape Fill', menu: (r, b) => A().menus.color('fill')(r, b) },
          { split: 'lineColor', label: 'Shape Outline', menu: (r, b) => A().menus.color('line')(r, b) },
          { drop: 'Shape Effects', icon: 'shadowStyle', menu: (r) => A().menus.shadow(r) }] },
        { label: 'Editing', items: ['find', 'replace', 'selectAll'] },
      ] },
      { id: 'insert', label: '&Insert', groups: [
        { label: 'Slides', items: [{ split: 'newSlide', size: 'large', label: 'New Slide', menu: open(['newSlide', 'duplicateSlide', '-', 'slidesFromFiles', 'slidesFromOutline']) }] },
        { label: 'Tables', items: [{ cmd: 'insertTable', size: 'large', label: 'Table' }] },
        { label: 'Images', items: [{ cmd: 'insertPicture', size: 'large', label: 'Pictures' }, 'insertClipArt', 'photoAlbum'] },
        { label: 'Illustrations', items: [{ drop: 'Shapes', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) }, { cmd: 'insertDiagram', size: 'large', label: 'Diagram' }, { cmd: 'insertChart', size: 'large', label: 'Chart' }] },
        { label: 'Links', items: [{ cmd: 'hyperlink', size: 'large', label: 'Hyperlink' }, { cmd: 'actionSettings', size: 'large', label: 'Action' }] },
        { label: 'Text', items: [{ cmd: 'textBox', size: 'large', label: 'Text Box' }, { cmd: 'headerFooter', size: 'large', label: 'Header & Footer' }, { cmd: 'insertWordArt', size: 'large', label: 'WordArt' }, 'insertDateTime', 'insertSlideNumber', 'vTextBox'] },
        { label: 'Symbols', items: [{ cmd: 'insertSymbol', size: 'large', label: 'Symbol' }] },
        { label: 'Media', items: [{ cmd: 'insertMedia', size: 'large', label: 'Media' }] },
      ] },
      { id: 'design', label: 'Desi&gn', groups: [
        { label: 'Themes', items: [{ cmd: 'slideDesign', size: 'large', label: 'Themes' }, 'replaceFonts'] },
        { label: 'Customize', items: [{ cmd: 'pageSetup', size: 'large', label: 'Slide Size' }, { cmd: 'background', size: 'large', label: 'Format Background' }] },
      ] },
      { id: 'transitions', label: 'Transitions', groups: [
        { label: 'Transition to This Slide', items: [{ cmd: 'transitionPane', size: 'large', label: 'Transitions' }] },
        { label: 'Timing', items: [{ cmd: 'rehearse', size: 'large', label: 'Rehearse Timings' }] },
      ] },
      { id: 'animations', label: 'Animations', groups: [
        { label: 'Animation', items: [{ cmd: 'animSchemes', size: 'large', label: 'Animation Schemes' }] },
        { label: 'Advanced Animation', items: [{ cmd: 'customAnim', size: 'large', label: 'Animation Pane' }] },
      ] },
      { id: 'slideShow', label: 'Slide Sho&w', groups: [
        { label: 'Start Slide Show', items: [{ cmd: 'showFromStart', size: 'large', label: 'From Beginning' }, { cmd: 'showFromCurrent', size: 'large', label: 'From Current Slide' }] },
        { label: 'Set Up', items: [{ cmd: 'setupShow', size: 'large', label: 'Set Up Slide Show' }, 'hideSlide', 'rehearse'] },
      ] },
      { id: 'review', label: 'Re&view', groups: [
        { label: 'Proofing', items: [{ cmd: 'spelling', size: 'large', label: 'Spelling' }, 'autocorrectDlg'] },
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Presentation Views', items: [{ cmd: 'viewNormal', size: 'large', label: 'Normal' }, { cmd: 'viewSorter', size: 'large', label: 'Slide Sorter' }, { cmd: 'viewNotes', size: 'large', label: 'Notes Page' }] },
        { label: 'Master Views', items: [{ cmd: 'viewMaster', size: 'large', label: 'Slide Master' }] },
        { label: 'Show', items: ['ruler', 'showGrid', 'gridGuides', { cmd: 'speakerNotes', label: 'Notes' }, 'slidesPane', 'taskPane'] },
        { label: 'Zoom', items: [{ cmd: 'zoomDlg', size: 'large', label: 'Zoom' }, { cmd: 'zoomFit', size: 'large', label: 'Fit to Window' }, { make: () => A().parts.zoom({ width: 64, id: 'rb-zoom' }) }] },
        { label: 'Color/Grayscale', items: ['viewColor', 'viewGray', 'viewBW'] },
      ] },
    ],
    contextual: [
      { id: 'master', set: 'Slide Master', open: true, when: ctx('master'), tabs: [
        { id: 'masterTab', label: 'Slide Master', groups: [
          { label: 'Edit Master', items: [{ cmd: 'insertTitleMaster', size: 'large', label: 'Insert Title Master' }, 'deleteTitleMaster'] },
          { label: 'Close', items: [{ cmd: 'closeMaster', size: 'large', label: 'Close Master View' }] },
        ] },
      ] },
      { id: 'drawing', set: 'Drawing Tools', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: 'Format', groups: [
          { label: 'Insert Shapes', items: [{ drop: 'Shapes', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) }, 'textBox', 'editPoints'] },
          { label: 'Shape Styles', launcher: 'formatObject', items: [{ split: 'fillColor', label: 'Shape Fill', menu: (r, b) => A().menus.color('fill')(r, b) }, { split: 'lineColor', label: 'Shape Outline', menu: (r, b) => A().menus.color('line')(r, b) },
            { drop: 'Line Style', icon: 'lineStyle', menu: (r) => A().menus.lineStyle(r) }, { drop: 'Dash Style', icon: 'dashStyle', menu: (r) => A().menus.dash(r) }, { drop: 'Arrow Style', icon: 'arrowStyle', menu: (r) => A().menus.arrow(r) }, { drop: 'Shadow Style', icon: 'shadowStyle', menu: (r) => A().menus.shadow(r) }] },
          { label: 'Arrange', items: ['bringFront', 'sendBack', 'group', 'ungroup', { drop: 'Align', icon: 'alignC', menu: open(['distH', 'distV', 'alignToSlide']) }, 'rotateRight', 'flipH', 'flipV'] },
        ] },
      ] },
      { id: 'picture', set: 'Picture Tools', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: 'Format', groups: [
          { label: 'Adjust', items: [{ row: ['picMoreBright', 'picLessBright', 'picMoreContrast', 'picLessContrast'] }, { drop: 'Color', icon: 'colorPic', menu: (r) => A().menus.picColor(r) }, 'picTransparent', 'picReset'] },
          { label: 'Picture Styles', launcher: 'formatObject', items: [{ drop: 'Picture Border', icon: 'lineStyle', menu: (r) => A().menus.lineStyle(r) }] },
          { label: 'Arrange', items: ['bringFront', 'sendBack', 'rotateLeft'] },
          { label: 'Size', items: [{ cmd: 'picCrop', size: 'large', label: 'Crop' }] },
        ] },
      ] },
      { id: 'table', set: 'Table Tools', when: ctx('table'), tabs: [
        { id: 'tableDesign', label: 'Design', groups: [
          { label: 'Table Styles', launcher: 'tableBordersFill', items: [{ drop: 'Shading', icon: 'fillColor', size: 'large', menu: (r, b) => A().menus.tableFill(r, b) }, { drop: 'Borders', icon: 'bordersAll', size: 'large', menu: (r) => A().menus.tableBorders(r) }] },
        ] },
        { id: 'tableLayout', label: 'Layout', groups: [
          { label: 'Table', items: [{ cmd: 'selectTable', size: 'large', label: 'Select' }] },
          { label: 'Rows & Columns', items: [{ drop: 'Delete', icon: 'deleteTable', size: 'large', menu: open(['deleteRows', 'deleteCols']) }, 'rowAbove', 'rowBelow', 'colLeft', 'colRight'] },
          { label: 'Merge', items: ['mergeCells', 'splitCell'] },
          { label: 'Cell Size', items: ['distRows', 'distCols'] },
          { label: 'Alignment', items: ['cellTop', 'cellMiddle', 'cellBottom'] },
        ] },
      ] },
      { id: 'wordart', set: 'WordArt Tools', when: ctx('wordart'), tabs: [
        { id: 'wordartFormat', label: 'Format', groups: [
          { label: 'Text', items: [{ cmd: 'wordartEdit', size: 'large', label: 'Edit Text' }, 'wordartGallery'] },
          { label: 'WordArt Styles', launcher: 'formatObject', items: [{ drop: 'WordArt Shape', icon: 'waShape', size: 'large', menu: (r) => A().menus.waShape(r) }, 'waSameHeight', 'waVertical'] },
        ] },
      ] },
    ],
  });
})();
