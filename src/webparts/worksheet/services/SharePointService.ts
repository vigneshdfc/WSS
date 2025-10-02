import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IWorksheetData } from "../components/IWorksheetProps";

export class SharePointService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    // Initialize PnPjs with the web part context
    sp.setup({
      spfxContext: context
    });
  }

  // Get all worksheets from SharePoint list
  public async getWorksheets(): Promise<IWorksheetData[]> {
    try {
      const items = await sp.web.lists.getByTitle("Worksheets").items.getAll();
      return items.map(item => this.mapListItemToWorksheet(item));
    } catch (error) {
      console.error("Error getting worksheets:", error);
      throw new Error("Failed to retrieve worksheets");
    }
  }

  // Save worksheet to SharePoint list
  public async saveWorksheet(worksheetData: IWorksheetData): Promise<void> {
    try {
      const list = sp.web.lists.getByTitle("Worksheets");
      
      if (worksheetData.worksheetNumber) {
        // Update existing item
        await list.items.getById(parseInt(worksheetData.worksheetNumber)).update({
          Title: worksheetData.standardName,
          StandardName: worksheetData.standardName,
          LEG: worksheetData.leg,
          UPG: worksheetData.upg,
          EOShikeishoNo: worksheetData.eoShikeishoNo,
          Project: worksheetData.project,
          SubProject: worksheetData.subProject,
          Model: worksheetData.model,
          SubModel: worksheetData.subModel,
          PassedQualityGate: worksheetData.passedQualityGate,
          Keyword: worksheetData.keyword,
          Remarks: worksheetData.remarks,
          PersonInCharge: worksheetData.personInCharge,
          PersonInChargeDate: worksheetData.personInChargeDate ? new Date(worksheetData.personInChargeDate) : null,
          Checker1: worksheetData.checker1,
          Checker1Date: worksheetData.checker1Date ? new Date(worksheetData.checker1Date) : null,
          Checker2: worksheetData.checker2,
          Checker2Date: worksheetData.checker2Date ? new Date(worksheetData.checker2Date) : null,
          Checker3: worksheetData.checker3,
          Checker3Date: worksheetData.checker3Date ? new Date(worksheetData.checker3Date) : null,
          Approver: worksheetData.approver,
          ApproverDate: worksheetData.approverDate ? new Date(worksheetData.approverDate) : null,
          Creator: worksheetData.creator,
          DeputyOfCreator: worksheetData.deputyOfCreator,
          DistributionList: worksheetData.distributionList
        });
      } else {
        // Create new item
        const result = await list.items.add({
          Title: worksheetData.standardName,
          StandardName: worksheetData.standardName,
          LEG: worksheetData.leg,
          UPG: worksheetData.upg,
          EOShikeishoNo: worksheetData.eoShikeishoNo,
          Project: worksheetData.project,
          SubProject: worksheetData.subProject,
          Model: worksheetData.model,
          SubModel: worksheetData.subModel,
          PassedQualityGate: worksheetData.passedQualityGate,
          Keyword: worksheetData.keyword,
          Remarks: worksheetData.remarks,
          PersonInCharge: worksheetData.personInCharge,
          PersonInChargeDate: worksheetData.personInChargeDate ? new Date(worksheetData.personInChargeDate) : null,
          Checker1: worksheetData.checker1,
          Checker1Date: worksheetData.checker1Date ? new Date(worksheetData.checker1Date) : null,
          Checker2: worksheetData.checker2,
          Checker2Date: worksheetData.checker2Date ? new Date(worksheetData.checker2Date) : null,
          Checker3: worksheetData.checker3,
          Checker3Date: worksheetData.checker3Date ? new Date(worksheetData.checker3Date) : null,
          Approver: worksheetData.approver,
          ApproverDate: worksheetData.approverDate ? new Date(worksheetData.approverDate) : null,
          Creator: worksheetData.creator,
          DeputyOfCreator: worksheetData.deputyOfCreator,
          DistributionList: worksheetData.distributionList
        });
      }
    } catch (error) {
      console.error("Error saving worksheet:", error);
      throw new Error("Failed to save worksheet");
    }
  }

  // Get dropdown options for projects
  public async getProjects(): Promise<string[]> {
    try {
      const items = await sp.web.lists.getByTitle("Projects").items.select("Title").getAll();
      return items.map(item => item.Title);
    } catch (error) {
      console.error("Error getting projects:", error);
      return ["SMEC"]; // Default fallback
    }
  }

  // Get dropdown options for models
  public async getModels(): Promise<string[]> {
    try {
      const items = await sp.web.lists.getByTitle("Models").items.select("Title").getAll();
      return items.map(item => item.Title);
    } catch (error) {
      console.error("Error getting models:", error);
      return [];
    }
  }

  // Create the required SharePoint lists if they don't exist
  public async ensureListsExist(): Promise<void> {
    try {
      // Check if Worksheets list exists, create if not
      const lists = await sp.web.lists.get();
      const worksheetListExists = lists.some(list => list.Title === "Worksheets");
      
      if (!worksheetListExists) {
        await sp.web.lists.add("Worksheets", "Worksheets list for storing worksheet data", 100, false);
        
        // Add fields to the list
        const list = sp.web.lists.getByTitle("Worksheets");
        await list.fields.addText("StandardName", 255, { Required: true });
        await list.fields.addText("LEG", 100);
        await list.fields.addText("UPG", 100);
        await list.fields.addText("EOShikeishoNo", 255);
        await list.fields.addText("Project", 100);
        await list.fields.addText("SubProject", 100);
        await list.fields.addText("Model", 100);
        await list.fields.addText("SubModel", 100);
        await list.fields.addText("PassedQualityGate", 100);
        await list.fields.addText("Keyword", 255);
        await list.fields.addMultilineText("Remarks", 6, false, false, false);
        await list.fields.addText("PersonInCharge", 100);
        await list.fields.addDateTime("PersonInChargeDate");
        await list.fields.addText("Checker1", 100);
        await list.fields.addDateTime("Checker1Date");
        await list.fields.addText("Checker2", 100);
        await list.fields.addDateTime("Checker2Date");
        await list.fields.addText("Checker3", 100);
        await list.fields.addDateTime("Checker3Date");
        await list.fields.addText("Approver", 100);
        await list.fields.addDateTime("ApproverDate");
        await list.fields.addText("Creator", 100);
        await list.fields.addText("DeputyOfCreator", 100);
        await list.fields.addMultilineText("DistributionList", 6, false, false, false);
      }
    } catch (error) {
      console.error("Error ensuring lists exist:", error);
      throw new Error("Failed to create required SharePoint lists");
    }
  }

  private mapListItemToWorksheet(item: any): IWorksheetData {
    return {
      worksheetNumber: item.Id?.toString() || "",
      standardName: item.StandardName || "",
      leg: item.LEG || "",
      upg: item.UPG || "",
      eoShikeishoNo: item.EOShikeishoNo || "",
      project: item.Project || "",
      subProject: item.SubProject || "",
      model: item.Model || "",
      subModel: item.SubModel || "",
      passedQualityGate: item.PassedQualityGate || "",
      keyword: item.Keyword || "",
      remarks: item.Remarks || "",
      personInCharge: item.PersonInCharge || "",
      personInChargeDate: item.PersonInChargeDate ? new Date(item.PersonInChargeDate).toISOString().split('T')[0] : "",
      checker1: item.Checker1 || "",
      checker1Date: item.Checker1Date ? new Date(item.Checker1Date).toISOString().split('T')[0] : "",
      checker2: item.Checker2 || "",
      checker2Date: item.Checker2Date ? new Date(item.Checker2Date).toISOString().split('T')[0] : "",
      checker3: item.Checker3 || "",
      checker3Date: item.Checker3Date ? new Date(item.Checker3Date).toISOString().split('T')[0] : "",
      approver: item.Approver || "",
      approverDate: item.ApproverDate ? new Date(item.ApproverDate).toISOString().split('T')[0] : "",
      creator: item.Creator || "",
      deputyOfCreator: item.DeputyOfCreator || "",
      distributionList: item.DistributionList || ""
    };
  }
}
