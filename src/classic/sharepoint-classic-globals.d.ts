/**
 * SharePoint Classic browser globals.
 *
 * These objects are supplied at runtime by SharePoint's classic pages and
 * clientpeoplepicker.js. TypeScript does not know about them automatically,
 * so we declare their names here without implementing them.
 */
declare var SP: any;
declare var SPClientPeoplePicker: any;
declare function SPClientPeoplePicker_InitStandaloneControlWrapper(
  peoplePickerElementId: string,
  value: any,
  schema: any
): void;
