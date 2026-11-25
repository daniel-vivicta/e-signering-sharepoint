import { ClassicEntry } from './ClassicEntry';
import {
  ClassicSigningForm,
  IClassicDocumentContext
} from './ClassicSigningForm';
import { ClassicHttpClientAdapter } from '../http/ClassicHttpClientAdapter';
import { ISigningDocument } from '../models/ISigningRequest';

function createEntry(webUrl: string): ClassicEntry {
  return new ClassicEntry(
    new ClassicHttpClientAdapter(),
    webUrl
  );
}

export function open(
  itemId: string,
  itemUrl: string,
  listId: string,
  siteUrl: string
): void {

  const entry: ClassicEntry = createEntry(siteUrl);

  const initialDocument: ISigningDocument = entry.open(
    itemId,
    itemUrl,
    listId,
    siteUrl
  );

  const documentContext: IClassicDocumentContext = {
    itemId: itemId,
    itemUrl: itemUrl,
    listId: listId,
    siteUrl: siteUrl
  };

  entry.getDocument(initialDocument)
    .then((document: ISigningDocument): void => {
      const form: ClassicSigningForm =
        new ClassicSigningForm(
          entry,
          documentContext,
          document
        );

      form.open();
    })
    .catch((error: any): void => {
      if (window.console && window.console.error) {
        window.console.error(
          '[ESIGN][CLASSIC DOCUMENT ERROR]',
          error
        );
      }

      alert(
        'Det gick inte att hämta dokumentinformationen.'
      );
    });
}

export function create(webUrl: string): ClassicEntry {
  return createEntry(webUrl);
}
