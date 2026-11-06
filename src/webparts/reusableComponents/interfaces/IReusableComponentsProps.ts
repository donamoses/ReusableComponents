import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "@fluentui/react";

export interface IReusableComponentsProps {
  context: WebPartContext;
  wpproperties: IReusableComponentsWebPartProps;
}
export interface IReusableComponentsWebPartProps {
  WebpartTitle: string;
  divstyle: boolean;
  uploadfiles: boolean;
  multidragdropfiles: boolean;
  singledragdropfile: boolean;
  griddragdropfile: boolean;
}
export interface IReusableComponentsState {
  currentUser: {
    id: number;
    email: string;
    title: string;
  };
  modaloverlay: {
    isOpen: boolean;
    modalText: string;
  };
  fileUploaded: IUploadedFile[];
  // Single document uploaded through the drag & drop control
  singleFile?: IUploadedFile;
  // Rows of the upload grid (one document per row, via GridDragDropFile)
  gridRows: IGridDocumentRow[];
  departmentOptions: IDropdownOption[];
  selectedDepartment?: string;
  title: string;
  dueDate?: Date;
  assignedTo?: {
    id: number;
    email: string;
    title: string;
    loginName: string;
  };
}
export interface IUploadedFile {
  SlNo: number;
  file: File;
  DocumentName: string;
  DocumentType?: string;
  UploadedBy: string;
  UploadedOn: string;
  FileName?: string;
  DocumentID?: string;
  FileLink?: string;
}
export interface IGridDocumentRow {
  // Stable key for React (row index changes when rows are removed)
  RowId: number;
  file?: IUploadedFile;
}
