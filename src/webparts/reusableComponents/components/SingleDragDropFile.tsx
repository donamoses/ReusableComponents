import * as React from 'react';
import { IconButton, Label, Link, Icon } from '@fluentui/react';
import { DragDropFiles as PnPDragDropFiles, IDragDropFilesProps } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker as PnPFilePicker, IFilePickerProps, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DragDropFile.module.scss';
import ToastService from '../../../shared/controls/Toast/Toast';
import { IUploadedFile } from '../interfaces/IReusableComponentsProps';

// @pnp/spfx-controls-react 3.x ships React 17 typings; this project is on React 18 (type-only mismatch)
const DragDropFiles = PnPDragDropFiles as unknown as React.ComponentType<React.PropsWithChildren<IDragDropFilesProps>>;
const FilePicker = PnPFilePicker as unknown as React.ComponentType<IFilePickerProps>;

export interface ISingleDragDropFileProps {
    context: WebPartContext;
    label?: string;
    // The file already uploaded to the library (shows view/delete instead of the drop area)
    uploadedFile?: IUploadedFile;
    disabled?: boolean;
    // Called with the single selected/dropped file
    onFileSelected: (file: File) => void;
    onView: (file: IUploadedFile) => void;
    onDelete: (file: IUploadedFile) => void;
}

export default class SingleDragDropFile extends React.Component<ISingleDragDropFileProps> {

    private _getDropFiles = (files: File[]): void => {
        if (this.props.disabled || !files || files.length === 0) return;
        if (files.length > 1) {
            ToastService.warning("Multiple file upload not supported.");
            return;
        }
        this.props.onFileSelected(files[0]);
    }

    private _getBrowsedFile = async (filesReceived: IFilePickerResult[]): Promise<void> => {
        const picked = filesReceived && filesReceived[0];
        if (!picked) return;
        try {
            // Local uploads have no absolute url; read the content as a File
            const file = await picked.downloadFileContent();
            this.props.onFileSelected(file);
        } catch (err) {
            console.error("Could not read the selected file", err);
            ToastService.error("Could not read the selected file");
        }
    }

    public render(): React.ReactElement<ISingleDragDropFileProps> {
        const { label, uploadedFile, disabled, context } = this.props;
        return (
            <div className={styles.DragDropFile}>
                {label && <Label>{label}</Label>}
                {uploadedFile ? (
                    <div className={styles.fileCard}>
                        <Icon iconName="Page" className={styles.fileIcon} />
                        <Link className={styles.fileName} onClick={() => this.props.onView(uploadedFile)} underline title="View document">
                            {uploadedFile.DocumentName}
                        </Link>
                        <IconButton iconProps={{ iconName: 'View' }} title="View" ariaLabel="View" onClick={() => this.props.onView(uploadedFile)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" disabled={disabled} onClick={() => this.props.onDelete(uploadedFile)} />
                    </div>
                ) : (
                    <DragDropFiles
                        dropEffect="copy"
                        enable={!disabled}
                        onDrop={this._getDropFiles}
                        iconName="Upload"
                        labelMessage="Drop file here..."
                    >
                        <div className={styles.uploadArea}>
                            <span>Drag and drop a file here...</span><br />
                            <span>OR</span>
                            <FilePicker
                                buttonIcon="FileImage"
                                buttonLabel="Browse"
                                disabled={disabled}
                                hideWebSearchTab={true}
                                hideStockImages={true}
                                hideOrganisationalAssetTab={true}
                                hideOneDriveTab={true}
                                hideSiteFilesTab={true}
                                hideRecentTab={true}
                                hideLocalMultipleUploadTab={true}
                                hideLinkUploadTab={true}
                                onSave={(results: IFilePickerResult[]) => { this._getBrowsedFile(results).catch(e => console.error(e)); }}
                                context={context as unknown as IFilePickerProps["context"]}
                            />
                        </div>
                    </DragDropFiles>
                )}
            </div>
        );
    }
}
