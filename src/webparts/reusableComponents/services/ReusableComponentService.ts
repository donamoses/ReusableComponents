import { BaseService } from "../../../shared/services/BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { getSP } from "../../../shared/PnP/pnpjsConfig";
import { IAddUsingPathProps, IFileUploadProgressData } from "@pnp/sp/files";
import { Web, IWeb } from "@pnp/sp/webs";
import "@pnp/sp/fields";

export const UPLOAD_SITE_URL = "https://ccsdev01.sharepoint.com/sites/LearningTest";
export const UPLOAD_LIBRARY_URL = "/sites/LearningTest/ReusableUpload";
export const DEPARTMENTS_LIST_URL = "/sites/LearningTest/Lists/Department";
export const REUSABLE_LIST_URL = "/sites/LearningTest/Lists/Reusable";
// Internal name of the lookup column in the ReusableUpload library that points to the Reusable list
// (REST expects the "Id" suffix, e.g. ReusableID -> ReusableIDId)
export const REUSABLE_LOOKUP_FIELD = "ReusableID";

export class ReusableComponentService extends BaseService {
  private spfi: SPFI;
  constructor(context: WebPartContext, siteUrl: string) {
    super(context, siteUrl);
    this.spfi = getSP(context);
  }

  public getMyTasksSelect(queryurl: string, select: string): Promise<any> {
    return this.spfi.web.getList(queryurl).items
      .select(select)()
  }

  public getCurrentUser() {
    return this.spfi.web.currentUser();
  }
  public getLibraryItem(queryurl: string, id: number): Promise<any> {
    return this.spfi.web.getList(queryurl).items.getById(id).select("FileRef,FileLeafRef")()
  }

  public getLibraryItemFilter(queryurl: string, select: string, filter: string): Promise<any> {
    return this.spfi.web.getList(queryurl).items.select(select).filter(filter)()
  }



  public async createFolderInSourceDocumentLibrary(libraryPath: string, folderName: string) {
    try {
      const folderPath = `${libraryPath}/${folderName}`;

      // Check if the folder exists
      const folder = this.spfi.web.getFolderByServerRelativePath(folderPath);
      await folder();  // Try to get the folder to see if it exists

    } catch (error) {
      console.error(`Folder ${folderName} does not exist in ${libraryPath}. Creating it...`, error);
      // Folder does not exist, so we create it
      await this.spfi.web.folders.addUsingPath(`${libraryPath}/${folderName}`, true);  // Use true for overwrite
    }
  }

  // Method to upload the document to the folder and replace if it exists
  public async uploadDocumentToFolder(folderUrl: string, file: File) {
    const fileName = file.name;
    const fileBuffer = await file.arrayBuffer();

    // Specify the options using IAddUsingPathProps to set Overwrite to true
    const addOptions: IAddUsingPathProps = { Overwrite: true };

    // Upload the file to the folder, setting overwrite to true to replace if it exists
    await this.spfi.web.getFolderByServerRelativePath(folderUrl).files.addUsingPath(fileName, fileBuffer, addOptions);
  }

  public async getFileContent(fileUrl: string): Promise<any> {
    return this.spfi.web.getFileByServerRelativePath(fileUrl).getItem();
  }
  public async uploadSourceDocument(filename: string, filedata: any, libraryname: string, metadata: any): Promise<any> {
    let uploadedfile: any;
    if (filedata.size <= 10485760) {
      uploadedfile = await this.spfi.web.getFolderByServerRelativePath(libraryname)
        .files.addUsingPath(filename, filedata, { Overwrite: true });
    } else {
      uploadedfile = await this.spfi.web.getFolderByServerRelativePath(libraryname)
        .files.addChunked(filename, filedata, {
          Overwrite: true,
          progress: (data: IFileUploadProgressData) => {
            console.log(`Uploaded ${JSON.stringify(data)}`);
          }
        });
    }

    const item = await this.spfi.web.getFileByServerRelativePath(uploadedfile.ServerRelativeUrl).getItem<{ Id: number }>("Id");

    try {
      await item.update({
        DocumentType: metadata.DocumentType,
        UploadedById: metadata.UploadedById,
        UploadedOn: metadata.UploadedOn
      });
    } catch (error) {
      console.error("Metadata update failed:", error);
    }

    const itemId = item.Id;
    return { uploadedfile, itemId };
  }
  public async uploadFileandUpdateProperty(libraryname: string, filename: string, filedata: any, metadata: any): Promise<any> {
   
let uploadedfile: any;
const FILE_SIZE_THRESHOLD = 10 * 1024 * 1024;
// const CHUNK_SIZE = 10 * 1024 * 1024;

if (filedata.size <= FILE_SIZE_THRESHOLD) {
  // Small file upload
  uploadedfile = await this.spfi.web
    .getFolderByServerRelativePath(libraryname)
    .files.addUsingPath(filename, filedata, { Overwrite: true });

  const item = await this.spfi.web
    .getFileByServerRelativePath(uploadedfile.ServerRelativeUrl)
    .getItem();
  await item.update(metadata);

  return uploadedfile;

} else {
  // Large file chunked upload
  uploadedfile = await this.spfi.web
    .getFolderByServerRelativePath(libraryname)
    .files.addChunked(filename, filedata, {
      // chunkSize: CHUNK_SIZE,
      Overwrite: true,
      progress: (data: IFileUploadProgressData) => {
        // ✅ Only use progress for UI updates, NOT for metadata  
           console.log(`Uploading chunk... ${JSON.stringify(data)}`);      
      
        // e.g. this.setState({ uploadProgress: data.blockNumber });
      }
    });

  // ✅ Update metadata AFTER chunked upload completes
  const item = await this.spfi.web
    .getFileByServerRelativePath(uploadedfile.ServerRelativeUrl)
    .getItem();
  await item.update(metadata);
  return uploadedfile;
}

}
  public async deletefileItemURL(fileUrl: string): Promise<any> {
    return this.spfi.web.getFileByServerRelativePath(fileUrl).delete();
  }

  // ---- ReusableUpload library (https://ccsdev01.sharepoint.com/sites/Learning/ReusableUpload) ----

  // Web bound to the Learning site, so uploads work no matter which site hosts the web part
  private getUploadWeb(): IWeb {
    return Web([this.spfi.web, UPLOAD_SITE_URL]);
  }

  // Uploads a file to the ReusableUpload library (chunked above 10 MB) and returns its URL and item id
  public async uploadToReusableLibrary(file: File): Promise<{ serverRelativeUrl: string; itemId: number }> {
    const web = this.getUploadWeb();
    const folder = web.getFolderByServerRelativePath(UPLOAD_LIBRARY_URL);
    const result = file.size <= 10 * 1024 * 1024
      ? await folder.files.addUsingPath(file.name, file, { Overwrite: true })
      : await folder.files.addChunked(file.name, file, { Overwrite: true });

    const item = await web.getFileByServerRelativePath(result.ServerRelativeUrl).getItem<{ Id: number }>("Id");
    return { serverRelativeUrl: result.ServerRelativeUrl, itemId: item.Id };
  }

  // Gets Id and Title of every item in the Departments list on the Learning site
  public async getDepartments(): Promise<{ Id: number; Title: string }[]> {
    return this.getUploadWeb().getList(DEPARTMENTS_LIST_URL).items
      .select("Id", "Title")
      .orderBy("Title", true)
      .top(5000)();
  }

  // Ensures a user on the Learning site (user ids differ per site) and returns their site user id
  public async ensureUploadSiteUser(loginName: string): Promise<number> {
    const user = await this.getUploadWeb().ensureUser(loginName);
    return user.Id;
  }

  // Creates an item in the Reusable list and returns its id
  public async addReusableItem(data: Record<string, unknown>): Promise<number> {
    const item = await this.getUploadWeb().getList(REUSABLE_LIST_URL).items.add(data);
    return item.Id;
  }

  // Deletes an item from the Reusable list (used to roll back a failed save)
  public async deleteReusableItem(id: number): Promise<void> {
    await this.getUploadWeb().getList(REUSABLE_LIST_URL).items.getById(id).recycle();
  }

  private documentNameFields: Record<string, string> = {};

  // Finds the internal name of the "Document Name" column in a list/library
  // (matches internal name DocumentName or display name "Document Name"/"DocumentName")
  private async getDocumentNameField(listUrl: string): Promise<string> {
    if (this.documentNameFields[listUrl]) return this.documentNameFields[listUrl];
    const fields = await this.getUploadWeb().getList(listUrl).fields
      .select("InternalName", "Title")
      .filter("Hidden eq false")<{ InternalName: string; Title: string }[]>();
    const clean = (s: string): string => (s || "").replace(/[\s_]|x0020/g, "").toLowerCase();
    const field = fields.find(f => f.InternalName === "DocumentName") ||
      fields.find(f => clean(f.Title) === "documentname" || clean(f.InternalName) === "documentname");
    if (!field) {
      throw new Error(`No "Document Name" column found in ${listUrl}. ` +
        `Available columns: ${fields.map(f => `${f.Title} (${f.InternalName})`).join(", ")}`);
    }
    this.documentNameFields[listUrl] = field.InternalName;
    return field.InternalName;
  }

  // Sets the Document Name column of an item in the Reusable list
  public async setReusableDocumentName(id: number, documentName: string): Promise<void> {
    const field = await this.getDocumentNameField(REUSABLE_LIST_URL);
    await this.getUploadWeb().getList(REUSABLE_LIST_URL).items.getById(id).update({ [field]: documentName });
  }

  // Sets the Reusable list lookup (plus any extra metadata) on every given item in the ReusableUpload library.
  // When fileName is given, the Document Name column is set to it and the file itself is renamed to it.
  public async linkLibraryItemsToReusable(items: { id: number; fileUrl?: string; fileName?: string }[], reusableId: number, metadata: Record<string, unknown> = {}): Promise<void> {
    const web = this.getUploadWeb();
    const [property, docNameField] = await Promise.all([
      this.getReusableLinkProperty(),
      this.getDocumentNameField(UPLOAD_LIBRARY_URL)
    ]);
    const list = web.getList(UPLOAD_LIBRARY_URL);
    for (const i of items) {
      // 1. Metadata + lookup + Document Name
      await list.items.getById(i.id).update({
        ...metadata,
        [property]: reusableId,
        ...(i.fileName ? { [docNameField]: i.fileName } : {})
      });
      // 2. Rename the file (same folder, new name); the item id stays the same
      if (i.fileName) {
        const fileUrl = i.fileUrl ||
          (await list.items.getById(i.id).select("FileRef")<{ FileRef: string }>()).FileRef;
        const folder = fileUrl.substring(0, fileUrl.lastIndexOf("/"));
        const newUrl = `${folder}/${i.fileName}`;
        if (newUrl.toLowerCase() !== fileUrl.toLowerCase()) {
          await web.getFileByServerRelativePath(fileUrl).moveByPath(newUrl, false);
        }
      }
    }
  }

  private reusableLinkProperty: string | undefined;

  // Works out the REST property name of the ReusableUpload column that links to the Reusable list.
  // Prefers a lookup column pointing at the Reusable list; falls back to a column whose title or
  // internal name matches REUSABLE_LOOKUP_FIELD. Lookups need the "Id" suffix, plain number/text columns don't.
  private async getReusableLinkProperty(): Promise<string> {
    if (this.reusableLinkProperty) return this.reusableLinkProperty;

    const web = this.getUploadWeb();
    const [reusableList, fields] = await Promise.all([
      web.getList(REUSABLE_LIST_URL).select("Id")<{ Id: string }>(),
      web.getList(UPLOAD_LIBRARY_URL).fields
        .select("InternalName", "Title", "TypeAsString", "LookupList")
        .filter("Hidden eq false")<{ InternalName: string; Title: string; TypeAsString: string; LookupList?: string }[]>()
    ]);

    const normalize = (s: string | undefined): string => (s || "").replace(/[{}]/g, "").toLowerCase();
    const target = REUSABLE_LOOKUP_FIELD.toLowerCase();
    const field =
      fields.find(f => f.TypeAsString.indexOf("Lookup") === 0 && normalize(f.LookupList) === normalize(reusableList.Id)) ||
      fields.find(f => f.InternalName.toLowerCase() === target || f.Title.toLowerCase() === target);

    if (!field) {
      throw new Error(`No column linking to the Reusable list was found in ${UPLOAD_LIBRARY_URL}. ` +
        `Available columns: ${fields.map(f => `${f.Title} (${f.InternalName}, ${f.TypeAsString})`).join(", ")}`);
    }
    if (field.TypeAsString === "LookupMulti") {
      throw new Error(`Column ${field.InternalName} is a multi-value lookup; set it with { results: [id] } instead.`);
    }

    this.reusableLinkProperty = field.TypeAsString === "Lookup" ? `${field.InternalName}Id` : field.InternalName;
    return this.reusableLinkProperty;
  }

  // Writes the same metadata to every given item in the ReusableUpload library
  public async updateReusableLibraryItems(itemIds: number[], metadata: Record<string, unknown>): Promise<void> {
    const list = this.getUploadWeb().getList(UPLOAD_LIBRARY_URL);
    await Promise.all(itemIds.map(id => list.items.getById(id).update(metadata)));
  }

  // Deletes a file from the ReusableUpload library (sends it to the recycle bin)
  public async deleteFromReusableLibrary(serverRelativeUrl: string): Promise<void> {
    await this.getUploadWeb().getFileByServerRelativePath(serverRelativeUrl).recycle();
  }
}
