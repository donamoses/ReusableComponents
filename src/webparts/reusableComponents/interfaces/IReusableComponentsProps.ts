import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "@fluentui/react";

export interface IReusableComponentsProps {
  context: WebPartContext;
  wpproperties: IReusableComponentsWebPartProps;
}
export interface IReusableComponentsWebPartProps {
  WebpartTitle: string;
  divstyle: boolean;
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