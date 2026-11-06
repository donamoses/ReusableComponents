import * as React from 'react';
import { DragDropFiles as PnPDragDropFiles, IDragDropFilesProps } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker as PnPFilePicker, IFilePickerProps, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DragDropFile.module.scss';
import * as strings from 'ReusableComponentsWebPartStrings';

// @pnp/spfx-controls-react 3.x ships React 17 typings; this project is on React 18 (type-only mismatch)
const DragDropFiles = PnPDragDropFiles as unknown as React.ComponentType<React.PropsWithChildren<IDragDropFilesProps>>;
const FilePicker = PnPFilePicker as unknown as React.ComponentType<IFilePickerProps>;

export interface IDragDropFileProps {
    context: WebPartContext;
    // Dropped files arrive as File[]; browsed files as IFilePickerResult[]. selectionType is strings.DragFileSelection or strings.BrowseFileSelection
    returnFileData: (files: File[] | IFilePickerResult[], selectionType: string) => void;
}

export default class DragDropFile extends React.Component<IDragDropFileProps> {

    private _getDropFiles = (files: File[]): void => {
        if (files && files.length > 0) {
            this.props.returnFileData(files, strings.DragFileSelection);
        }
    }

    private _getBrowsedFile = (filesReceived: IFilePickerResult[]): void => {
        this.props.returnFileData(filesReceived, strings.BrowseFileSelection);
    }

    public render(): React.ReactElement<IDragDropFileProps> {

        return (
            <div className={styles.DragDropFile}>
                <DragDropFiles
                    dropEffect="copy"
                    enable={true}
                    onDrop={this._getDropFiles}
                    iconName="Upload"
                    labelMessage="Drop file here..."
                >
                    <div className={styles.uploadArea}>
                        <span>Drag and drop files here...</span><br />
                        <span>OR</span>
                        <FilePicker
                            buttonIcon="FileImage"
                            buttonLabel="Browse"
                            hideWebSearchTab={true}
                            hideLocalUploadTab={true}
                            hideStockImages={true}
                            hideOrganisationalAssetTab={true}
                            hideOneDriveTab={true}
                            hideSiteFilesTab={true}
                            hideLocalMultipleUploadTab={false}
                            hideLinkUploadTab={true}
                            onSave={(filePickerResult: IFilePickerResult[]) => {
                                this._getBrowsedFile(filePickerResult);
                            }}
                            context={this.props.context as unknown as IFilePickerProps["context"]}
                        />
                    </div>
                </DragDropFiles>
            </div>
        );
    }
}
