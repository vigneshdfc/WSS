import * as React from "react";
import { useState, useEffect } from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import styles from "./Worksheet.module.scss";
import { IWorksheetProps } from "./IWorksheetProps";

const Worksheet: React.FC<IWorksheetProps> = (props: IWorksheetProps): JSX.Element => {
  const [activeTab, setActiveTab] = useState<"basic" | "worksheet" | "approval">("basic");

  // Dropdown states
  const [standards, setStandards] = useState<string[]>([]);
  const [projects, setProjects] = useState<string[]>([]);
  const [subProjects, setSubProjects] = useState<string[]>([]);
  const [models, setModels] = useState<string[]>([]);
  const [subModels, setSubModels] = useState<string[]>([]);
  const [qualityGates, setQualityGates] = useState<string[]>([]);

  // Header fields
  const [standardNumber, setStandardNumber] = useState<string>("eAxle Gearbox for electric vehicle Design standard");
  const [title, setTitle] = useState<string>("");

  const siteUrl = "https://corptb.sharepoint.com/sites/apac-04929-WSS";

  // Fetch SharePoint list items helper
  const fetchListItems = async (listName: string, field: string = "Title"): Promise<string[]> => {
    try {
      const response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${field}`,
        {
          method: "GET",
          headers: {
            Accept: "application/json;odata=nometadata",
          },
        }
      );

      const data = await response.json();
      return data.value.map((i: any) => i[field]);
    } catch (error) {
      console.error(`Error fetching list ${listName}:`, error);
      return [];
    }
  };

  // Fetch all data for dropdowns
  const fetchAllData = async (): Promise<void> => {
    const [
      standardsData,
      projectsData,
      subProjectsData,
      modelsData,
      subModelsData,
      gatesData,
    ] = await Promise.all([
      fetchListItems("DesignStandard", "Title"),
      fetchListItems("Project Configuration", "Project"),
      fetchListItems("Project Configuration", "Title"),
      fetchListItems("Model Configuration", "Model"),
      fetchListItems("Model Configuration", "Title"),
      fetchListItems("WSS", "Gate"),
    ]);

    setStandards(standardsData);
    setProjects(projectsData);
    setSubProjects(subProjectsData);
    setModels(modelsData);
    setSubModels(subModelsData);
    setQualityGates(gatesData);
  };

  useEffect(() => {
    void fetchAllData();
  }, []);

  return (
    <div className={styles.newWbsRequest}>
      {/* Top Title Bar */}
      <div className={styles.header}>
        <h2>New WSS Request</h2>
      </div>

      {/* Action Buttons */}
      <div className={styles.actions}>
        <button className={styles.btn}>Save</button>
        <button className={styles.btn}>Close</button>
        <button className={styles.btn}>Go to Home page</button>
        <button className={styles.btnPrimary}>Open design standard</button>
      </div>

      {/* Header Fields */}
      <div className={styles.headerFieldsShell}>
        <div className={styles.headerFieldsCard}>
          {/* Standard Number */}
          <div className={styles.headerRow}>
            <div className={styles.headerLabelCol}>
              <span className={styles.headerLabelText}>Standard Number</span>
            </div>
            <div className={styles.headerFieldCol}>
              <div className={styles.withIcon}>
                <Icon iconName="List" className={styles.inputIcon} />
                <select
                  className={styles.input}
                  value={standardNumber}
                  onChange={(e) => setStandardNumber(e.target.value)}
                >
                  <option>eAxle Gearbox for electric vehicle Design standard</option>
                  {standards.map((s, i) => (
                    <option key={i} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Title */}
          <div className={styles.headerRow}>
            <div className={styles.headerLabelCol}>
              <span className={styles.headerLabelText}>Title</span>
            </div>
            <div className={styles.headerFieldCol}>
              <div className={styles.withIcon}>
                <Icon iconName="Edit" className={styles.inputIcon} />
                <input
                  className={styles.input}
                  type="text"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                />
              </div>
            </div>
          </div>

          {/* Status */}
          <div className={styles.headerRow}>
            <div className={styles.headerLabelCol}>
              <span className={styles.headerLabelText}>Status</span>
            </div>
            <div className={styles.headerFieldCol}>
              <div className={styles.withIcon}>
                <Icon iconName="Blocked" className={styles.inputIcon} />
                <input className={styles.input} type="text" disabled />
              </div>
            </div>
          </div>
        </div>
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
        {/* === BASIC INFO === */}
        {activeTab === "basic" && (
          <div className={styles.tabContainer}>
            <div className={styles.headerFieldsShell}>
              {/* Worksheet Number */}
              <div className={styles.formRow}>
                <label>Worksheet Number</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Blocked" className={styles.inputIcon} />
                  <input type="text" disabled />
                </div>
              </div>

              {/* Standard Name */}
              <div className={styles.formRow}>
                <label>Standard Name</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Blocked" className={styles.inputIcon} />
                  <input type="text" value="MSW1-9658S" disabled />
                </div>
              </div>

              {/* LEG */}
              <div className={styles.formRow}>
                <label>LEG</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Blocked" className={styles.inputIcon} />
                  <input type="text" disabled />
                </div>
              </div>

              {/* UPG */}
              <div className={styles.formRow}>
                <label>UPG</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Blocked" className={styles.inputIcon} />
                  <input type="text" disabled />
                </div>
              </div>

              {/* EO/Shikeisho */}
              <div className={styles.formRow}>
                <label>EO/Shikeisho No.</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Edit" className={styles.inputIcon} />
                  <input type="text" />
                </div>
                <div className={styles.helpText}>EO or Shikeisyo No. are separated by comma.</div>
              </div>

              {/* Project */}
              <div className={styles.formRow}>
                <label>Project</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="List" className={styles.inputIcon} />
                  <select>
                    <option>Select</option>
                    {projects.map((p, i) => (
                      <option key={i} value={p}>
                        {p}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Sub Project */}
              <div className={styles.formRow}>
                <label>Sub Project</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="List" className={styles.inputIcon} />
                  <select>
                    <option>Select</option>
                    {subProjects.map((sp, i) => (
                      <option key={i} value={sp}>
                        {sp}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Model */}
              <div className={styles.formRow}>
                <label>Model</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="List" className={styles.inputIcon} />
                  <select>
                    <option>Select</option>
                    {models.map((m, i) => (
                      <option key={i} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Sub Model */}
              <div className={styles.formRow}>
                <label>Sub Model</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="List" className={styles.inputIcon} />
                  <select>
                    <option>Select</option>
                    {subModels.map((sm, i) => (
                      <option key={i} value={sm}>
                        {sm}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Passed Quality Gate */}
              <div className={styles.formRow}>
                <label>Passed Quality Gate</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="List" className={styles.inputIcon} />
                  <select>
                    <option>Select</option>
                    {qualityGates.map((qg, i) => (
                      <option key={i} value={qg}>
                        {qg}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Keyword */}
              <div className={styles.formRow}>
                <label>Keyword</label>
                <div className={styles.inputWithIcon}>
                  <Icon iconName="Edit" className={styles.inputIcon} />
                  <input type="text" />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* === WORKSHEET === */}
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
              <textarea placeholder="Enter remarks here..." />
            </div>
          </div>
        )}

        {/* === APPROVAL === */}
        {activeTab === "approval" && (
          <div className={styles.tabContainer}>
            <div className={styles.gridHeader}>
              <div>Responsibility</div>
              <div>Person</div>
              <div>Date</div>
            </div>

            <div className={styles.gridRow}>
              <label>Person in charge</label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>

            <div className={styles.gridRow}>
              <label>Checker1 <span className={styles.req}>*</span></label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>

            <div className={styles.gridRow}>
              <label>Checker2</label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>

            <div className={styles.gridRow}>
              <label>Checker3</label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>

            <div className={styles.gridRow}>
              <label>Approver <span className={styles.req}>*</span></label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>
          </div>
        )}
      </div>

      {/* === CREATOR SECTION === */}
      <div className={styles.creatorSection}>
        <div className={styles.formRow}>
          <label>Creator</label>
          <div className={styles.inputWithIcon}>
            <Icon iconName="Contact" className={styles.inputIcon} />
            <input type="text" value="Duraisamy, Ragunath (575) (EXT)" disabled />
          </div>
        </div>
        <div className={styles.formRow}>
          <label>Deputy of Creator</label>
          <div className={styles.inputWithIcon}>
            <Icon iconName="Contact" className={styles.inputIcon} />
            <input type="text" />
          </div>
        </div>
        <div className={styles.formRow}>
          <label>Distribution list at approval</label>
          <div className={styles.inputWithIcon}>
            <Icon iconName="Contact" className={styles.inputIcon} />
            <input type="text" placeholder="Enter users separated with semicolons." />
          </div>
        </div>
      </div>
    </div>
  );
};

export default Worksheet;
