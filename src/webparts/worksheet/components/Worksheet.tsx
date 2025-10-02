import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { PrimaryButton, DefaultButton, Spinner, MessageBar, MessageBarType } from "@fluentui/react";
import styles from "./Worksheet.module.scss";
import { IWorksheetProps, IWorksheetData } from "./IWorksheetProps";
import { SharePointService } from "../services/SharePointService";

const Worksheet: React.FC<IWorksheetProps> = (props: IWorksheetProps) => {
  const [activeTab, setActiveTab] = useState<"basic" | "worksheet" | "approval">("basic");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [success, setSuccess] = useState<string>("");
  const [projects, setProjects] = useState<string[]>([]);
  const [models, setModels] = useState<string[]>([]);
  
  const [worksheetData, setWorksheetData] = useState<IWorksheetData>({
    worksheetNumber: "",
    standardName: "",
    leg: "",
    upg: "",
    eoShikeishoNo: "",
    project: "",
    subProject: "",
    model: "",
    subModel: "",
    passedQualityGate: "",
    keyword: "",
    remarks: "",
    personInCharge: "",
    personInChargeDate: "",
    checker1: "",
    checker1Date: "",
    checker2: "",
    checker2Date: "",
    checker3: "",
    checker3Date: "",
    approver: "",
    approverDate: "",
    creator: props.userDisplayName,
    deputyOfCreator: "",
    distributionList: ""
  });

  const sharePointService = new SharePointService(props.context);

  // Initialize data on component mount
  useEffect(() => {
    initializeData();
  }, []);

  const initializeData = async () => {
    try {
      setIsLoading(true);
      setError("");
      
      // Ensure SharePoint lists exist
      await sharePointService.ensureListsExist();
      
      // Load dropdown options
      const [projectsData, modelsData] = await Promise.all([
        sharePointService.getProjects(),
        sharePointService.getModels()
      ]);
      
      setProjects(projectsData);
      setModels(modelsData);
    } catch (err) {
      setError(`Failed to initialize data: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleInputChange = useCallback((field: keyof IWorksheetData, value: string) => {
    setWorksheetData(prev => ({
      ...prev,
      [field]: value
    }));
  }, []);

  const handleSave = async () => {
    try {
      setIsLoading(true);
      setError("");
      setSuccess("");
      
      await sharePointService.saveWorksheet(worksheetData);
      setSuccess("Worksheet saved successfully!");
    } catch (err) {
      setError(`Failed to save worksheet: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleClose = () => {
    // Reset form or navigate away
    setWorksheetData({
      worksheetNumber: "",
      standardName: "",
      leg: "",
      upg: "",
      eoShikeishoNo: "",
      project: "",
      subProject: "",
      model: "",
      subModel: "",
      passedQualityGate: "",
      keyword: "",
      remarks: "",
      personInCharge: "",
      personInChargeDate: "",
      checker1: "",
      checker1Date: "",
      checker2: "",
      checker2Date: "",
      checker3: "",
      checker3Date: "",
      approver: "",
      approverDate: "",
      creator: props.userDisplayName,
      deputyOfCreator: "",
      distributionList: ""
    });
    setActiveTab("basic");
  };

  if (isLoading && !projects.length) {
    return (
      <div className={styles.newWbsRequest}>
        <div style={{ textAlign: 'center', padding: '20px' }}>
          <Spinner label="Loading..." />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.newWbsRequest}>
      {/* Error/Success Messages */}
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError("")}>
          {error}
        </MessageBar>
      )}
      {success && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccess("")}>
          {success}
        </MessageBar>
      )}

      {/* Header */}
      <div className={styles.header}>
        <h2>New WSS Request</h2>
      </div>

      {/* Action Buttons */}
      <div className={styles.actions}>
        <PrimaryButton 
          text="Save" 
          onClick={handleSave}
          disabled={isLoading}
        />
        <DefaultButton 
          text="Close" 
          onClick={handleClose}
          disabled={isLoading}
        />
        <DefaultButton 
          text="Go to Home page" 
          disabled={isLoading}
        />
        <DefaultButton 
          text="Open design standard" 
          disabled={isLoading}
        />
      </div>

      {/* Tabs */}
      <div className={styles.tabs}>
        <button
          className={activeTab === "basic" ? styles.activeTab : ""}
          onClick={() => setActiveTab("basic")}
        >
          Basic Info.
        </button>
        <button
          className={activeTab === "worksheet" ? styles.activeTab : ""}
          onClick={() => setActiveTab("worksheet")}
        >
          Worksheet
        </button>
        <button
          className={activeTab === "approval" ? styles.activeTab : ""}
          onClick={() => setActiveTab("approval")}
        >
          Check/Approval
        </button>
      </div>

      {/* Tab Content */}
      <div className={styles.tabContent}>
        {/* Basic Info Tab */}
        {activeTab === "basic" && (
          <div className={styles.tabContainer}>
            <div className={styles.formSection}>
              <div className={styles.row}>
                <label>Worksheet Number</label>
                <input 
                  type="text" 
                  value={worksheetData.worksheetNumber}
                  onChange={(e) => handleInputChange('worksheetNumber', e.target.value)}
                  disabled={isLoading}
                />
              </div>
              <div className={styles.row}>
                <label>Standard Name</label>
                <input 
                  type="text" 
                  value={worksheetData.standardName}
                  onChange={(e) => handleInputChange('standardName', e.target.value)}
                  disabled={isLoading}
                />
              </div>
              <div className={styles.row}>
                <label>LEG</label>
                <input 
                  type="text" 
                  value={worksheetData.leg}
                  onChange={(e) => handleInputChange('leg', e.target.value)}
                  disabled={isLoading}
                />
              </div>
              <div className={styles.row}>
                <label>UPG</label>
                <input 
                  type="text" 
                  value={worksheetData.upg}
                  onChange={(e) => handleInputChange('upg', e.target.value)}
                  disabled={isLoading}
                />
              </div>
              <div className={styles.row}>
                <label>EO/Shikeisho No. <span>required</span></label>
                <input 
                  type="text" 
                  placeholder="Comma separated values"
                  value={worksheetData.eoShikeishoNo}
                  onChange={(e) => handleInputChange('eoShikeishoNo', e.target.value)}
                  disabled={isLoading}
                />
              </div>
              <div className={styles.row}>
                <label>Project</label>
                <select 
                  value={worksheetData.project}
                  onChange={(e) => handleInputChange('project', e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">Select</option>
                  {projects.map(project => (
                    <option key={project} value={project}>{project}</option>
                  ))}
                </select>
              </div>
              <div className={styles.row}>
                <label>Sub Project</label>
                <select 
                  value={worksheetData.subProject}
                  onChange={(e) => handleInputChange('subProject', e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Model</label>
                <select 
                  value={worksheetData.model}
                  onChange={(e) => handleInputChange('model', e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">Select</option>
                  {models.map(model => (
                    <option key={model} value={model}>{model}</option>
                  ))}
                </select>
              </div>
              <div className={styles.row}>
                <label>Sub Model</label>
                <select 
                  value={worksheetData.subModel}
                  onChange={(e) => handleInputChange('subModel', e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Passed Quality Gate</label>
                <select 
                  value={worksheetData.passedQualityGate}
                  onChange={(e) => handleInputChange('passedQualityGate', e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Keyword</label>
                <input 
                  type="text" 
                  value={worksheetData.keyword}
                  onChange={(e) => handleInputChange('keyword', e.target.value)}
                  disabled={isLoading}
                />
              </div>
            </div>
          </div>
        )}

        {/* Worksheet Tab */}
        {activeTab === "worksheet" && (
          <div className={styles.tabContainer}>
            <div className={styles.gridHeader}>
              <div>Item</div>
              <div>Check</div>
              <div>Comment</div>
              <div>Attachment/Link</div>
            </div>
            <div className={styles.gridRow}>
              <label>Remarks</label>
              <textarea 
                placeholder="Enter remarks here..."
                value={worksheetData.remarks}
                onChange={(e) => handleInputChange('remarks', e.target.value)}
                disabled={isLoading}
              ></textarea>
            </div>
          </div>
        )}

        {/* Check/Approval Tab */}
        {activeTab === "approval" && (
          <div className={`${styles.tabContainer} ${styles.checkApproval}`}>
            <div className={styles.gridHeader}>
              <div>Responsibility</div>
              <div>Person</div>
              <div>Date</div>
            </div>

            <div className={styles.gridRow}>
              <label>Person in charge</label>
              <input 
                type="text" 
                placeholder="Enter name"
                value={worksheetData.personInCharge}
                onChange={(e) => handleInputChange('personInCharge', e.target.value)}
                disabled={isLoading}
              />
              <input 
                type="date" 
                value={worksheetData.personInChargeDate}
                onChange={(e) => handleInputChange('personInChargeDate', e.target.value)}
                disabled={isLoading}
              />
            </div>

            <div className={styles.gridRow}>
              <label>Checker1 *</label>
              <input 
                type="text" 
                placeholder="Enter name"
                value={worksheetData.checker1}
                onChange={(e) => handleInputChange('checker1', e.target.value)}
                disabled={isLoading}
              />
              <input 
                type="date" 
                value={worksheetData.checker1Date}
                onChange={(e) => handleInputChange('checker1Date', e.target.value)}
                disabled={isLoading}
              />
            </div>

            <div className={styles.gridRow}>
              <label>Checker2</label>
              <input 
                type="text" 
                placeholder="Enter name"
                value={worksheetData.checker2}
                onChange={(e) => handleInputChange('checker2', e.target.value)}
                disabled={isLoading}
              />
              <input 
                type="date" 
                value={worksheetData.checker2Date}
                onChange={(e) => handleInputChange('checker2Date', e.target.value)}
                disabled={isLoading}
              />
            </div>

            <div className={styles.gridRow}>
              <label>Checker3</label>
              <input 
                type="text" 
                placeholder="Enter name"
                value={worksheetData.checker3}
                onChange={(e) => handleInputChange('checker3', e.target.value)}
                disabled={isLoading}
              />
              <input 
                type="date" 
                value={worksheetData.checker3Date}
                onChange={(e) => handleInputChange('checker3Date', e.target.value)}
                disabled={isLoading}
              />
            </div>

            <div className={styles.gridRow}>
              <label>Approver *</label>
              <input 
                type="text" 
                placeholder="Enter name"
                value={worksheetData.approver}
                onChange={(e) => handleInputChange('approver', e.target.value)}
                disabled={isLoading}
              />
              <input 
                type="date" 
                value={worksheetData.approverDate}
                onChange={(e) => handleInputChange('approverDate', e.target.value)}
                disabled={isLoading}
              />
            </div>
          </div>
        )}
      </div>

      {/* Creator Section (always visible) */}
      <div className={styles.creatorSection}>
        <div className={styles.row}>
          <label>Creator</label>
          <input
            type="text"
            value={worksheetData.creator}
            readOnly
            disabled={isLoading}
          />
        </div>
        <div className={styles.row}>
          <label>Deputy of Creator</label>
          <input 
            type="text" 
            value={worksheetData.deputyOfCreator}
            onChange={(e) => handleInputChange('deputyOfCreator', e.target.value)}
            disabled={isLoading}
          />
        </div>
        <div className={styles.row}>
          <label>Distribution list at approval</label>
          <input
            type="text"
            placeholder="Enter users separated with semicolons."
            value={worksheetData.distributionList}
            onChange={(e) => handleInputChange('distributionList', e.target.value)}
            disabled={isLoading}
          />
        </div>
      </div>
    </div>
  );
};

export default Worksheet;
