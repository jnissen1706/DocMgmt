export interface IDocMgmtState {
  disaplyState: ComponentState;
  errorMessage: string;
}



//Component state
export enum ComponentState {
  //Loading message
  loadingSpinner = 0,
  //Error
  error = 1,
  //Archive Options
  archiveAttachments = 2,
  //No Message
  noMessageID = 3,
}