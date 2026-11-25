import { ISigner } from './ISigner';

/**
 * Identifies the SharePoint document that is sent for signing.
 */
export interface ISigningDocument {
  webUrl: string;
  listId: string;
  itemId: string;
  uniqueId: string;
  fileName: string;
  fileRef: string;
}

/**
 * Represents the user who starts the signing request.
 *
 * This is kept separate from ISigner because the initiator does not
 * necessarily have to be one of the people who signs the document.
 */
export interface ISigningInitiator {
  id?: number;
  displayName: string;
  email: string;
  accountName: string;
  userName: string;
  userPrincipalName: string;
  source: string;
  identifier: string;
}

/**
 * Complete request created by the SPFx frontend before the request is
 * written to the signing queue.
 */
export interface ISigningRequest {
  correlationId: string;

  document: ISigningDocument;

  signers: ISigner[];

  message: string;
  deadline: string;

  initiator: ISigningInitiator;
}
