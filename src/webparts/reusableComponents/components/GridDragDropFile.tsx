import * as React from 'react';
import { DragDropFiles as PnPDragDropFiles, IDragDropFilesProps } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker as PnPFilePicker, IFilePickerProps, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'ReusableComponentsWebPartStrings';
import ToastService from '../../../shared/controls/Toast/Toast';
import styles from './DragDropFile.module.scss';

// @pnp/spfx-controls-react 3.x ships React 17 typings; this project is on React 18 (type-only mismatch)
const DragDropFiles = PnPDragDropFiles as unknown as React.ComponentType<React.PropsWithChildren<IDragDropFilesProps>>;
const FilePicker = PnPFilePicker as unknown as React.ComponentType<IFilePickerProps>;

export interface IGridDragDropFileProps {
    context: WebPartContext;
    // Called with the single dropped/browsed file. selectionType is strings.DragFileSelection or strings.BrowseFileSelection
    returnFileData: (file: File, selectionType: string) => void;
}

export default class GridDragDropFile extends React.Component<IGridDragDropFileProps> {

    private _getDropFiles = (files: File[]): void => {
        if (files && files.length > 0) {
            if (files.length > 1) {
                ToastService.warning("Multiple file upload not supported.");
            }
            else {
                this.props.returnFileData(files[0], strings.DragFileSelection);
            }
        }
    }

    private _getBrowsedFile = async (filesReceived: IFilePickerResult[]): Promise<void> => {
        const picked = filesReceived && filesReceived[0];
        // Local uploads have no absolute url; read the content as a File
        if (picked && picked.fileAbsoluteUrl === null) {
            const file = await picked.downloadFileContent();
            this.props.returnFileData(file, strings.BrowseFileSelection);
        }
    }

    public render(): React.ReactElement<IGridDragDropFileProps> {

        return (
            <div>
                <DragDropFiles
                    dropEffect="copy"
                    enable={true}
                    onDrop={this._getDropFiles}
                    iconName="Upload"
                    labelMessage=""
                >
                    <div className={styles.filepickerdiv}>
                        <FilePicker
                            buttonIcon=""
                            buttonClassName={styles.filepickerbutton}
                            buttonLabel="Drag & Drop / Browse"
                            hideWebSearchTab={true}
                            hideStockImages={true}
                            hideOrganisationalAssetTab={true}
                            hideOneDriveTab={true}
                            hideSiteFilesTab={true}
                            hideLocalMultipleUploadTab={true}
                            hideLinkUploadTab={true}
                            onSave={(filePickerResult: IFilePickerResult[]) => {
                                this._getBrowsedFile(filePickerResult).catch(err => {
                                    console.error("Could not read the selected file", err);
                                    ToastService.error("Could not read the selected file");
                                });
                            }}
                            context={this.props.context as unknown as IFilePickerProps["context"]}
                        />
                    </div>
                </DragDropFiles>
            </div>
        );
    }
}
