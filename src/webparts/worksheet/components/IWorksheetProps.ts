import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWorksheetData {
  worksheetNumber: string;
  standardName: string;
  leg: string;
  upg: string;
  eoShikeishoNo: string;
  project: string;
  subProject: string;
  model: string;
  subModel: string;
  passedQualityGate: string;
  keyword: string;
  remarks: string;
  personInCharge: string;
  personInChargeDate: string;
  checker1: string;
  checker1Date: string;
  checker2: string;
  checker2Date: string;
  checker3: string;
  checker3Date: string;
  approver: string;
  approverDate: string;
  creator: string;
  deputyOfCreator: string;
  distributionList: string;
}

export interface IWorksheetProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
