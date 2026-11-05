import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../PnP/pnpjsConfig";
import "@pnp/sp/attachments";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";
import "@pnp/sp/folders";

export class BaseService {
    private sp: SPFI;
    constructor(context: WebPartContext, siteUrl: string) {
        this.sp = getSP(context);
    }
    // Method to get the current user's information
    public getCurrentUser() {
        return this.sp.web.currentUser();
    }
    // Method to get the groups the current user belongs to
    public getCurrentUserGroups() {
        return this.sp.web.currentUser.groups();
    }
    // Method to get all lisr items in the list
    public getListItems(url: string): Promise<any> {
        return this.sp.web.getList(url).items();
    }
    // Method to get list items with pagination
    public async getPagedListItems(queryurl: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items with pagination and filter
    public async getPagedFilterListItems(queryurl: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items with pagination, filter and select
    public async getPagedSelectFilterExpandListItems(queryurl: string, select: string, filter: string, expand: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).filter(filter).expand(expand).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items with select
    public getItemsSelect(queryurl: string, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)()
    }
    // Method to get list items with select and expand
    public getItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }
    // Method to get list items by ID
    public getItemsById(queryurl: string, id: any): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)()
    }
    // Method to get list items by ID with select
    public getItemsByIdSelect(queryurl: string, id: any, select: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)()
    }
    // Method to get list items with filter
    public getItemsFilter(queryurl: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .filter(filter)()
    }
    // Method to get list items with select and filter
    public getItemsSelectFilter(queryurl: string, select: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .filter(filter)()
    }
    // Method to get list items with select, expand and filter
    public getItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }
    // Method to get list items with pagination , select and expand
    public async getPagedItemsSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items with pagination , select, expand and filter
    public async getPagedItemsSelectExpandFilter(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select(select).expand(expand).filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items with pagination and filter
    public async getPagedItemsFilter(queryurl: string, filter: string): Promise<any> {
        let collection: any = []
        for await (const items of this.sp.web.getList(queryurl).items.select('*').filter(filter).top(250) as any) {
            collection = collection.concat(items)
        }
        return collection
    }
    // Method to get list items by ID with select and expand
    public getItemsByIdSelectExpand(queryurl: string, id: any, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .getById(id)
            .select(select)
            .expand(expand)()
    }
    // Method to create new list item
    public createNewItem(url: string, data: any): Promise<any> {
        return this.sp.web.getList(url).items.add(data);
    }
    // Method to update list item by ID
    public updateItem(url: string, data: any, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).update(data);
    }
    // Method to delete list item by ID
    public DeleteItem(url: string, id: number): Promise<any> {
        return this.sp.web.getList(url).items.getById(id).delete();
    }
    // Method to create list item inside folder
    public async createListItemInsideFolder(): Promise<any> {
        // Set the list URL and folder name
        const listUrl = "/sites/ProcurementDMS/Lists/RequestMaster";
        const folderName = "MARKETING";
        return this.sp.web.getList(listUrl).items.add({
            Title: "Item inside folder",
            FileSystemObjectType: 0,
            FileLeafRef: "Item inside folder",
            'FolderPath': { 'DecodedUrl': `${listUrl}/${folderName}` }
            // Add other fields as needed
        });
    }
    // Method to get list items with select and expand
    public getSelectExpand(queryurl: string, select: string, expand: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)()
    }
    // Method to get user information by ID
    public async getUser(userId: number): Promise<any> {
        return this.sp.web.getUserById(userId)();
    }
    // Method to get file content by file URL
    public async getFileContent(fileUrl: string): Promise<ArrayBuffer> {
        return this.sp.web.getFileByServerRelativePath(fileUrl).getBuffer();
    }
    // Method to get document libraries in the site
    public getDocLibrary(url: string): Promise<any> {
        return this.sp.site.getDocumentLibraries(url);
    }
    // Method to upload document to a library
    public async uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        const response = await this.sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
        return response
    }
    // Method to get library item by ID
    public getLibraryItem(queryurl: string, id: number): Promise<any> {
        return this.sp.web.getList(queryurl).items.getById(id).select("FileRef,FileLeafRef")()
    }
    // Method to get folder content by folder URL
    public async getFolderContent(FolderName: string): Promise<ArrayBuffer> {
        return this.sp.web.getFolderByServerRelativePath(FolderName).files();
    }
    // Method to get list items with select, expand and filter
    public getItemsSelectFilterExpand(queryurl: string, select: string, expand: string, filter: string): Promise<any> {
        return this.sp.web.getList(queryurl).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }
}