import * as React from "react";
import { useState } from "react";
import styles from "./Worksheet.module.scss";
import { IWorksheetProps } from "./IWorksheetProps";

const Worksheet: React.FC<IWorksheetProps> = (props: IWorksheetProps) => {
  const [activeTab, setActiveTab] = useState<"basic" | "worksheet" | "approval">("basic");

  return (
    <div className={styles.newWbsRequest}>
      {/* Header */}
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
                <input type="text" />
              </div>
              <div className={styles.row}>
                <label>Standard Name</label>
                <input type="text" />
              </div>
              <div className={styles.row}>
                <label>LEG</label>
                <input type="text" />
              </div>
              <div className={styles.row}>
                <label>UPG</label>
                <input type="text" />
              </div>
              <div className={styles.row}>
                <label>EO/Shikeisho No. <span>required</span></label>
                <input type="text" placeholder="Comma separated values" />
              </div>
              <div className={styles.row}>
                <label>Project</label>
                <select>
                  <option>Select</option>
                  <option>SMEC</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Sub Project</label>
                <select>
                  <option>Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Model</label>
                <select>
                  <option>Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Sub Model</label>
                <select>
                  <option>Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Passed Quality Gate</label>
                <select>
                  <option>Select</option>
                </select>
              </div>
              <div className={styles.row}>
                <label>Keyword</label>
                <input type="text" />
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
              <textarea placeholder="Enter remarks here..."></textarea>
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
              <input type="text" placeholder="Enter name" />
              <input type="date" />
            </div>

            <div className={styles.gridRow}>
              <label>Checker1 *</label>
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
              <label>Approver *</label>
              <input type="text" placeholder="Enter name" />
              <input type="date" />
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
            value="Duraisamy, Ragunath (575) (EXT)"
            readOnly
          />
        </div>
        <div className={styles.row}>
          <label>Deputy of Creator</label>
          <input type="text" />
        </div>
        <div className={styles.row}>
          <label>Distribution list at approval</label>
          <input
            type="text"
            placeholder="Enter users separated with semicolons."
          />
        </div>
      </div>
    </div>
  );
};

export default Worksheet;
