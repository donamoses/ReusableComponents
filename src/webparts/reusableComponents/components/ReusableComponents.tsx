import * as React from 'react';
import styles from './ReusableComponents.module.scss';
import type { IReusableComponentsProps, IReusableComponentsState, IUploadedFile } from '../interfaces/IReusableComponentsProps';
import { DatePicker, Dropdown, IconButton, IDropdownOption, Link, PrimaryButton, TextField } from '@fluentui/react';
import CustomFileInput from './CustomFileInput';
import { ReusableComponentService, UPLOAD_SITE_URL } from '../services/ReusableComponentService';
import ModalOverlay from '../../../shared/controls/Overlay/Overlay';
import ToastService from '../../../shared/controls/Toast/Toast';
import { PeoplePicker as PnPPeoplePicker, PrincipalType, IPeoplePickerContext, IPeoplePickerProps } from '@pnp/spfx-controls-react/lib/PeoplePicker';

// @pnp/spfx-controls-react 3.x ships React 17 / SPFx 1.23 typings; this project is on React 18 / SPFx 1.24.
// The mismatch is type-only (SPFx loads react and sp-* from the page at runtime), so re-type the control.
const PeoplePicker = PnPPeoplePicker as unknown as React.ComponentType<IPeoplePickerProps>;


export default class ReusableComponents extends React.Component<IReusableComponentsProps, IReusableComponentsState> {
  private service: ReusableComponentService;
  private peoplePickerContext: IPeoplePickerContext;
  constructor(props: IReusableComponentsProps) {
    super(props);
    this.peoplePickerContext = {
      absoluteUrl: props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: props.context.msGraphClientFactory as unknown as IPeoplePickerContext["msGraphClientFactory"],
      spHttpClient: props.context.spHttpClient as unknown as IPeoplePickerContext["spHttpClient"]
    };
    this.service = new ReusableComponentService(props.context, props.context.pageContext.web.absoluteUrl);
    this.state = {
      currentUser: {
        id: 0,
        email: "",
        title: ""
      },
      modaloverlay: { isOpen: false, modalText: "" },
      fileUploaded: [],
      departmentOptions: [],
      selectedDepartment: undefined,
      title: "",
      dueDate: undefined,
      assignedTo: undefined
    };
  }
  //On Load
  public async componentDidMount() {
    this.loadDepartments().catch(err => console.error(err));
    const user = await this.service.getCurrentUser();
    if (user) {
      this.setState({
        currentUser: {
          id: user.Id,
          email: user.Email,
          title: user.Title
        },
        modaloverlay: { isOpen: false, modalText: "" },
      });
    }
  }
  // Bind the Departments list to the dropdown (Title as both key and text)
  private loadDepartments = async (): Promise<void> => {
    try {
      const items = await this.service.getDepartments();
      const departmentOptions: IDropdownOption[] = items
        .filter(i => !!i.Title)
        .map(i => ({ key: i.Title, text: i.Title }));
      this.setState({ departmentOptions });
    } catch (err) {
      console.error("Failed to load departments", err);
      ToastService.error("Could not load departments");
    }
  }
  private onDepartmentChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ selectedDepartment: option ? String(option.key) : undefined });
  }
  private onDueDateChange = (date: Date | null | undefined): void => {
    this.setState({ dueDate: date || undefined });
  }
  private onTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ title: newValue || "" });
  }
  // Document Name is "<Title>_<ReusableId>.<document type>", e.g. "My Title_12.docx"
  // Characters SharePoint doesn't allow in file names are replaced with "_"
  private getDocumentName = (title: string, reusableId: number, documentType?: string): string => {
    const safeTitle = title.trim().replace(/["*:<>?/\\|#%~&{}]/g, "_").replace(/^\.+|\.+$/g, "");
    const name = `${safeTitle}_${reusableId}`+ "_" + Date.now();
    return documentType ? `${name}.${documentType}` : name;
  }
  // File name for each uploaded file; files after the first with the same name get a "_2", "_3"... suffix
  private getLibraryFileNames = (title: string, reusableId: number, files: IUploadedFile[]): string[] => {
    const used: Record<string, number> = {};
    return files.map(f => {
      const base = this.getDocumentName(title, reusableId);
      const key = `${base}.${f.DocumentType}`.toLowerCase();
      used[key] = (used[key] || 0) + 1;
      const name = used[key] === 1 ? base : `${base}_${used[key]}`;
      return f.DocumentType ? `${name}.${f.DocumentType}` : name;
    });
  }
  // Store the selected user (Id is the SharePoint user id, usable as AssignedToId)
  private onAssignedToChange = (items: any[]): void => {
    const person = items && items[0];
    this.setState({
      assignedTo: person
        ? { id: Number(person.id), email: person.secondaryText, title: person.text, loginName: person.loginName }
        : undefined
    });
  }
  // Save: create the Reusable list item, then link every uploaded file in the ReusableUpload library to it
  private onSave = async (): Promise<void> => {
    const { fileUploaded, title, selectedDepartment, dueDate, assignedTo } = this.state;
    const linkedFiles = fileUploaded.filter(f => {
      const id = Number(f.DocumentID);
      return !isNaN(id) && id > 0;
    });
    if (linkedFiles.length === 0) {
      ToastService.error("Please upload at least one document before saving");
      return;
    }
    if (!title.trim()) {
      ToastService.error("Please enter a title");
      return;
    }
    this.setState({ modaloverlay: { isOpen: true, modalText: "Saving..." } });
    let reusableId: number | undefined;
    try {
      // Reusable list column internal names — adjust here if the list columns are named differently
      // User ids are per site, so resolve the person on the Learning site where the list lives
      const assignedToId = assignedTo ? await this.service.ensureUploadSiteUser(assignedTo.loginName) : null;
      const metadata: Record<string, unknown> = {
        Title: title.trim(),
        Department: selectedDepartment || null,
        DueDate: dueDate ? dueDate.toISOString() : null,
        AssignedToId: assignedToId
      };
      // 1. Add the request to the Reusable list
      reusableId = await this.service.addReusableItem(metadata);
      // 2. DocumentName needs the new item id: "<Title>_<ReusableId>.<type of first file>"
      const documentName = this.getDocumentName(title, reusableId, linkedFiles[0].DocumentType);
      await this.service.setReusableDocumentName(reusableId, documentName);
      // 3. Copy the same metadata to the uploaded files, set the Reusable list item as their lookup
      //    and rename each file to "<Title>_<ReusableId>.<ext>"
      //    (the service resolves the lookup column's REST name, e.g. ReusableIDId)
      const fileNames = this.getLibraryFileNames(title, reusableId, linkedFiles);
      const libraryItems = linkedFiles.map((f, i) => ({ id: Number(f.DocumentID), fileUrl: f.FileLink, fileName: fileNames[i] }));
      await this.service.linkLibraryItemsToReusable(libraryItems, reusableId, metadata);
      ToastService.success(`Saved and linked ${libraryItems.length} document(s)`);
      this.resetForm();
    } catch (err) {
      console.error("Save failed", err);
      // Don't leave an orphan list item behind if linking the documents failed
      if (reusableId) {
        await this.service.deleteReusableItem(reusableId).catch(e => console.error("Rollback failed", e));
      }
      ToastService.error("Could not save the details");
    } finally {
      this.setState({ modaloverlay: { isOpen: false, modalText: "" } });
    }
  }
  // Submit button
  private onsubmit = (): void => {
    this.onSave().catch(err => console.error(err));
  }
  // Close button: clear the form
  private onCancel = (): void => {
    this.resetForm();
  }
  private resetForm = (): void => {
    this.setState({
      fileUploaded: [],
      title: "",
      selectedDepartment: undefined,
      dueDate: undefined,
      assignedTo: undefined
    });
  }
  // Display dates as DD/MM/YYYY
  private formatDate = (date?: Date): string => {
    if (!date) return "";
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${date.getFullYear()}`;
  }
  public uploadFile = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const tempFile: IUploadedFile[] = [];
    const files = event.target.files;
    if (files && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];

        const duplicate = this.state.fileUploaded.some(tempItem => tempItem.DocumentName === file.name);
        const duplicateFromList = this.state.fileUploaded.some(listItem => listItem.FileName === file.name);
        const duplicateInBatch = tempFile.some(tempItem => tempItem.DocumentName === file.name);
        const fileExtension = (file.name.split('.').pop() || '').toLowerCase();
        const isZipFile = fileExtension === 'zip';
        const isAudioFile = ['mp3', 'wav', 'ogg', 'aac'].indexOf(fileExtension) !== -1;
        const isVideoFile = ['mp4', 'avi', 'mkv', 'mov'].indexOf(fileExtension) !== -1;
        if (isZipFile || isAudioFile || isVideoFile) {
          // Ignore ZIP, audio, and video files
          // Optionally, you can display an error message or handle these files differently
          continue;
        }
        if (!duplicate && !duplicateFromList && !duplicateInBatch) {
          tempFile.push({
            SlNo: this.state.fileUploaded.length + tempFile.length + 1,
            file: file,
            DocumentName: file.name,
            DocumentType: fileExtension,
            FileName: file.name,
            UploadedBy: this.state.currentUser.title,
            UploadedOn: new Date().toLocaleDateString(),
          });
        }
      }
    }

    // Reset the input so the same file can be selected again after removal
    event.target.value = '';

    if (tempFile.length > 0) {
      this.uploadToLibrary(tempFile).catch(err => console.error(err));
    }
  }
  // Upload the selected files to the ReusableUpload library and add them to the grid
  private uploadToLibrary = async (files: IUploadedFile[]): Promise<void> => {
    this.setState({ modaloverlay: { isOpen: true, modalText: "Uploading..." } });
    const uploaded: IUploadedFile[] = [];
    const failed: string[] = [];
    for (const item of files) {
      try {
        const res = await this.service.uploadToReusableLibrary(item.file);
        uploaded.push({ ...item, DocumentID: String(res.itemId), FileLink: res.serverRelativeUrl });
      } catch (err) {
        console.error(`Upload failed for ${item.DocumentName}`, err);
        failed.push(item.DocumentName);
      }
    }
    this.setState(prevState => {
      const fileUploaded = [...prevState.fileUploaded, ...uploaded];
      return {
        fileUploaded,
        modaloverlay: { isOpen: false, modalText: "" }
      };
    });
    if (uploaded.length > 0) ToastService.success(`${uploaded.length} file(s) uploaded`);
    if (failed.length > 0) ToastService.error(`Upload failed: ${failed.join(", ")}`);
  }
  // Open an uploaded document in a new tab
  private onDocumentOpen = (fileLink?: string): void => {
    if (fileLink) window.open(`${new URL(UPLOAD_SITE_URL).origin}${fileLink}?web=1`, "_blank");
  }
  // Table File Item Delete Grid Item
  public gridItemDelete = async (key: number, _e: React.MouseEvent<HTMLElement | HTMLAnchorElement | HTMLButtonElement, MouseEvent>) => {
    const deleteitem = this.state.fileUploaded[key];
    if (deleteitem?.FileLink) {
      try {
        await this.service.deleteFromReusableLibrary(deleteitem.FileLink);
      } catch (err) {
        console.error(err);
        ToastService.error(`Could not delete ${deleteitem.DocumentName}`);
        return;
      }
    }
    this.setState(prev => {
      const fileUploaded = prev.fileUploaded.filter((_item, index) => index !== key);
      return { fileUploaded };
    });
  }
  public render(): React.ReactElement<IReusableComponentsProps> {
    

    return (
      <section className={`${styles.container}`}>
        <ModalOverlay isModalOpen={this.state.modaloverlay.isOpen} modalText={this.state.modaloverlay.modalText} />
        {ToastService.container()}
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.WebpartTitle}</div>
          </div>
          <div className={styles.formbody}>
            {this.props.wpproperties.divstyle&&(<div>
            {/* One Column */}
            <div className={styles.row}>
              <div className={styles.col12}>
                <TextField label="Column 1"  />
              </div>
            </div>
           {/* 3/4 column + 1/4 column */}
            <div className={styles.row}>
              <div className={styles.col8}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col4}>
                <TextField label="Column 2" value=""/>
              </div>
            </div>
            {/* 2 columns */}
            <div className={styles.row}>
              <div className={styles.col6}>
                <TextField label="Column 1" value=""/>
              </div>
               <div className={styles.col6}>
                <TextField label="Column 2" value=""/>
              </div>
            </div>
            {/* 3 columns */}
            <div className={styles.row}>
              <div className={styles.col4}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col4}>
                <TextField label="Column 2" value=""  />
              </div>
              <div className={styles.col4}>
                <TextField label="Column 3" value=""/>
              </div>
            </div>
            {/* 4 columns */}
            <div className={styles.row}>
              <div className={styles.col2}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col2}>
                <TextField label="Column 2" value=""  />
              </div>
              <div className={styles.col2}>
                <TextField label="Column 3" value=""/>
              </div>
              <div className={styles.col2}>
                <TextField label="Column 4" value=""/>
              </div>
            </div>
            </div>)}
            {/* Title */}
            <div className={styles.row}>
              <div className={styles.col12}>
                <TextField
                  label="Title"
                  placeholder="Enter a title"
                  value={this.state.title}
                  onChange={this.onTitleChange}
                />
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.col6}>
                <Dropdown
                  label="Department"
                  placeholder="Select a department"
                  options={this.state.departmentOptions}
                  selectedKey={this.state.selectedDepartment}
                  onChange={this.onDepartmentChange}
                />
                </div>
               <div className={styles.col6}>
                  <DatePicker
                    label="Due Date"
                    placeholder="Select a due date"
                    value={this.state.dueDate}
                    onSelectDate={this.onDueDateChange}
                    formatDate={this.formatDate}
                    minDate={new Date()}
                    allowTextInput={false}
                  />
                </div>
            </div>
            {/* Assigned To */}
            <div className={styles.row}>
              <div className={styles.col6}>
                <PeoplePicker
                  context={this.peoplePickerContext}
                  titleText="Assigned To"
                  placeholder="Enter a name or email"
                  personSelectionLimit={1}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={300}
                  ensureUser={true}
                  defaultSelectedUsers={this.state.assignedTo ? [this.state.assignedTo.email] : []}
                  onChange={this.onAssignedToChange}
                />
              </div>
            </div>
            {/* Upload File */}
            <div className={styles.row}>
              <div className={styles.col12}>
                <CustomFileInput onChange={this.uploadFile} />
              </div>
            </div>
            {/* table */}
            {this.state.fileUploaded.length > 0 &&
              <div className={styles.doctable}>
                <table className={styles.table} >
                  <tr className={styles.tr}>
                    <th className={styles.th}>Sl No</th>
                    <th className={styles.th}>Document Name</th>
                    <th className={styles.th}>Document Type</th>
                    <th className={styles.th}>Uploaded By</th>
                    <th className={styles.th}>Uploaded On</th>
                    <th className={styles.th}>Action</th>
                  </tr>
                  {this.state.fileUploaded.map((item, key) => {
                    return (
                      <tr key={key} className={styles.tr}>
                        <td className={styles.th}>{key + 1}</td>
                        <td className={styles.th}>{item.FileLink ? <Link onClick={() => this.onDocumentOpen(item.FileLink)} underline>{item.DocumentName}</Link> : item.DocumentName}</td>
                        <td className={styles.th}>{item.DocumentType}</td>
                        <td className={styles.th} >{item.UploadedBy}</td>
                        <td className={styles.th}>{item.UploadedOn}</td>
                        <td className={styles.th}>
                          <IconButton iconProps={{ iconName: 'Delete' }} onClick={(e) => this.gridItemDelete(key, e as React.MouseEvent<HTMLButtonElement>)} /></td>
                      </tr>
                    );
                  })}
                </table>
              </div>
            }
            <div className={styles.row}>
              <div className={styles.col12}>
                <div className={styles.rgtalign}>
                  
                  <PrimaryButton className={styles.btn} onClick={this.onsubmit}>Submit</PrimaryButton >
                  <PrimaryButton className={styles.btn} onClick={this.onCancel}>Close</PrimaryButton >
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}

  
