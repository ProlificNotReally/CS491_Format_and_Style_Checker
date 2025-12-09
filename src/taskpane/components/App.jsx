import React, { useState } from "react";
import { analyzeFormatting, goToError } from "../wordChecks";
import { checkDocument } from "../docChecks";
import { checkStyles } from "../checkStyles";
import {
  checkHeaderFooterFormatting,
  fixHeaderFooterIssue,
  syncHeaderFooterByOrientation,
} from "../checkHeaderFooterFormatting";
import { fixGeneralDocumentIssue } from "../fixGeneralDocIssues"


export default function App() {
  const [results, setResults] = useState([]);
  const [docResults, setDocResults] = useState([]);
  const [styleResults, setStyleResults] = useState([]);
  const [compResults, setCompResults] = useState([]);
  const [headerFooterResults, setHeaderFooterResults] = useState([]);

  const [isChecking, setIsChecking] = useState(false);
  const [isCheckingDocument, setIsCheckingDocument] = useState(false);
  const [isCheckingStyles, setIsCheckingStyles] = useState(false);
  const [isCheckingComp, setIsCheckingComp] = useState(false);
  const [isCheckingHeaderFooter, setIsCheckingHeaderFooter] = useState(false);
  const [isFixingHeaderFooter, setIsFixingHeaderFooter] = useState(false);
  const [isFixingGeneralIssue, setIsFixingGeneralIssue] = useState(false);
  const [isFixingAllGeneral, setIsFixingAllGeneral] = useState(false);

  //Run formatting analysis
  const handleRunCheck = async () => {
    try {
      setIsChecking(true);
      const issues = await analyzeFormatting();
      setResults(issues);
    } catch (err) {
      console.error("Error running formatting checks:", err);
    } finally {
      setIsChecking(false);
    }
  };

  // Run header/footer checker independently (unchanged)
  const handleRunHeaderFooterCheck = async () => {
    try {
      setIsCheckingHeaderFooter(true);
      const issues = await checkHeaderFooterFormatting();
      setHeaderFooterResults(issues);
    } catch (err) {
      console.error("Error running header/footer checks:", err);
    } finally {
      setIsCheckingHeaderFooter(false);
    }
  };

  // Fix a single header/footer issue (one row)
  const handleFixHeaderFooterIssue = async (issue) => {
    try {
      setIsFixingHeaderFooter(true);
      await fixHeaderFooterIssue(issue);          // normalize headers/footers
      const refreshed = await checkHeaderFooterFormatting(); // re-check
      setHeaderFooterResults(refreshed);
    } catch (err) {
      console.error("Error fixing single header/footer issue:", err);
    } finally {
      setIsFixingHeaderFooter(false);
    }
  };


const handleFixAllHeaderFooter = async () => {
  try {
    setIsFixingHeaderFooter(true);

    // 1. Get the current list of header/footer issues
    const initialIssues = await Word.run(async (context) => {
      return await checkHeaderFooterFormatting(context);
    });

    // 2. Fix each issue individually (same as clicking each row's GoFix)
    for (const issue of initialIssues) {
      try {
        await fixHeaderFooterIssue(issue); // your existing per-issue fix
      } catch (e) {
        console.error("Error fixing header/footer issue", issue, e);
      }
    }

    // 3. Run the majority-based sync *after* individual fixes
    await syncHeaderFooterByOrientation();

    // 4. Re-run the checker so the UI shows what's left (if anything)
    const finalIssues = await Word.run(async (context) => {
      return await checkHeaderFooterFormatting(context);
    });
    setHeaderFooterResults(finalIssues);
  } catch (err) {
    console.error("Error fixing all header/footer issues:", err);
  } finally {
    setIsFixingHeaderFooter(false);
  }
};



  const handleRunDocumentCheck = async () => {
    try {
      setIsCheckingDocument(true);
      const issues = await checkDocument();
      setDocResults(issues);
    } catch (err) {
      console.error("Error running formatting checks:", err);
    } finally {
      setIsCheckingDocument(false);
    }
  };

  const handleFixGeneralIssue = async (issue) => {
    try {
      setIsFixingGeneralIssue(true);
      await fixGeneralDocumentIssue(issue);          // normalize headers/footers
      const refreshed = await checkDocument(); // re-check
      setDocResults(refreshed);
    } catch (err) {
      console.error("Error fixing single general document issue:", err);
    } finally {
      setIsFixingGeneralIssue(false);
    }
  };

  const handleFixAllGeneralIssues = async () => {
  try {
    setIsFixingAllGeneral(true);

    // 1. Get the current list of header/footer issues
    const initialIssues = await Word.run(async (context) => {
      return await checkDocument(context);
    });

    // 2. Fix each issue individually (same as clicking each row's GoFix)
    for (const issue of initialIssues) {
      try {
        await fixGeneralDocumentIssue(issue); // your existing per-issue fix
      } catch (e) {
        console.error("Error fixing general document issue", issue, e);
      }
    }

    // 4. Re-run the checker so the UI shows what's left (if anything)
    const finalIssues = await Word.run(async (context) => {
      return await checkDocument(context);
    });
    setDocResults(finalIssues);
  } catch (err) {
    console.error("Error fixing all header/footer issues:", err);
  } finally {
    setIsFixingAllGeneral(false);
  }
};

  const handleRunStylesCheck = async () => {
    try {
      setIsCheckingStyles(true);
      const issues = await checkStyles();
      setStyleResults(issues);
    } catch (err) {
      console.error("Error running style checks:", err);
    } finally {
      setIsCheckingStyles(false);
    }
  };

  const handleRunComprehensiveCheck = async() => {
    try {
      setIsCheckingComp(true);
      const formatting_issues = await analyzeFormatting();
      const header_footer_issues = await handleRunHeaderFooterCheck();
      const general_doc_issues = await checkDocument();
      const style_issues = await checkStyles();
      const all_issues = [...formatting_issues, ...header_footer_issues, ...general_doc_issues, ...style_issues]
      setCompResults(all_issues)

    } catch (err) {
      console.error("Error running comprehensive checks:", err)
    } finally {
      setIsCheckingComp(false);
    }
  }

  //Jump to a specific error in Word
  const handleGoTo = async (issue) => {
    if (issue.canLocate && issue.location) {
      await goToError(issue.location);
    }
  };

  return (
    <div>
      <div style={styles.container}>
        <h1 style={styles.title}>Comprehensive Checker</h1>

        <button onClick = {handleRunComprehensiveCheck} style={styles.button} disabled={isCheckingComp}>
          {isCheckingComp ? "Checking..." : "Run Comprehensive Check"}
        </button>

       <div style={styles.resultsContainer}>
          {compResults.length === 0 && !isCheckingComp && (
            <p style={styles.placeholder}>No results yet. Click “Run Comprehensive Check”.</p>
          )}

          {compResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
              }}
            >
              <b>{r.type}</b>
              <p style={styles.message}>{r.message}</p>
            </div>
          ))}
        </div>

      </div>
      <div style={styles.container}>
        <h1 style={styles.title}>Formatting Checker</h1>

        <button onClick={handleRunCheck} style={styles.button} disabled={isChecking}>
          {isChecking ? "Checking..." : "Run Formatting Check"}
        </button>

        <div style={styles.resultsContainer}>
          {results.length === 0 && !isChecking && (
            <p style={styles.placeholder}>No results yet. Click “Run Formatting Check”.</p>
          )}

          {results.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
              }}
            >
              <b>{r.type}</b>
              <p style={styles.message}>{r.message}</p>
            </div>
          ))}
        </div>
      </div>

{/* HEADER/FOOTER CHECKER */}
<div style={styles.container}>
  <h2 style={styles.title}>Header/Footer Checker</h2>

  <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
    <button
      onClick={handleRunHeaderFooterCheck}
      style={styles.button}
      disabled={isCheckingHeaderFooter || isFixingHeaderFooter}
    >
      {isCheckingHeaderFooter ? "Checking..." : "Run Header/Footer Check"}
    </button>

    <button
      onClick={handleFixAllHeaderFooter}
      style={styles.secondaryButton}
      disabled={
        isCheckingHeaderFooter ||
        isFixingHeaderFooter ||
        headerFooterResults.length === 0
      }
    >
      {isFixingHeaderFooter ? "Fixing..." : "Fix All Header/Footer"}
    </button>
  </div>

  <div style={styles.resultsContainer}>
    {headerFooterResults.length === 0 &&
      !isCheckingHeaderFooter &&
      !isFixingHeaderFooter && (
        <p style={styles.placeholder}>
          No results yet. Click “Run Header/Footer Check”.
        </p>
      )}

    {headerFooterResults.map((r) => (
      <div key={r.id} style={styles.resultRow}>
        <div
          style={{
            ...styles.resultBox,
            flex: 1,
            cursor: r.canLocate ? "pointer" : "default",
            backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
          }}
          onClick={() => r.canLocate && handleGoTo(r)}
        >
          <b>{r.type}</b>
          <p style={styles.message}>{r.message}</p>
        </div>

        <div style={styles.resultActions}>
          {r.canLocate && (
            <button
              style={styles.smallButton}
              onClick={() => handleGoTo(r)}
            >
              Go
            </button>
          )}
          <button
            style={styles.smallButton}
            disabled={isFixingHeaderFooter}
            onClick={() => handleFixHeaderFooterIssue(r)}
          >
            {isFixingHeaderFooter ? "Fixing..." : "Fix"}
          </button>
        </div>
      </div>
    ))}
  </div>
</div>


      <div style={styles.container}>
        <h2 style={styles.title}>Document Checker</h2>

        {/* <button onClick={handleRunDocumentCheck} style={styles.button} disabled={isCheckingDocument}>
          {isCheckingDocument ? "Checking..." : "Run Document Check"}
        </button> */}

       <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button
            onClick={handleRunDocumentCheck}
            style={styles.button}
            disabled={isCheckingDocument || isFixingGeneralIssue || isFixingAllGeneral}
          >
            {isCheckingDocument? "Checking..." : "Run Document Check"}
          </button>

          <button
            onClick={handleFixAllGeneralIssues}
            style={styles.secondaryButton}
            disabled={
              isCheckingDocument ||
               isFixingGeneralIssue ||
               isFixingAllGeneral ||
               docResults.length === 0
            }
          >
            {isFixingAllGeneral ? "Fixing..." : "Fix All General Document Issues"}
          </button>
        </div>

        <div style={styles.resultsContainer}>
          {docResults.length === 0 && !isCheckingDocument &&  !isFixingGeneralIssue && !isFixingAllGeneral && (
            <p style={styles.placeholder}>No results yet. Click “Run Document Check”.</p>
          )}

          {docResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
              }}
            >
              <b>{r.type}</b>
              <p style={styles.message}>{r.message}</p>

              <div style={styles.resultActions}>
                {r.canLocate && (
                  <button
                    style={styles.smallButton}
                    onClick={() => handleGoTo(r)}
                  >
                    Go
                  </button>
                )}
                <button
                  style={styles.smallButton}
                  disabled={ isCheckingDocument || isFixingGeneralIssue || isFixingAllGeneral}
                  onClick={() => handleFixGeneralIssue(r)}
                >
                  {isFixingGeneralIssue ? "Fixing..." : "Fix"}
                </button>
              </div>
            </div>
          ))}

        </div>
      </div>

      <div style={styles.container}>
        <h2 style={styles.title}>Styles Checker</h2>

        <button onClick={handleRunStylesCheck} style={styles.button} disabled={isCheckingStyles}>
          {isCheckingStyles ? "Checking..." : "Run Styles Check"}
        </button>

        <div style={styles.resultsContainer}>
          {styleResults.length === 0 && !isCheckingStyles && (
            <p style={styles.placeholder}>No results yet. Click “Run Styles Check”.</p>
          )}

          {styleResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
              }}
            >
              <b>{r.type}</b>
              <p style={styles.message}>{r.message}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  
  );
}

//Inline styling
const styles = {
  container: {
    padding: "16px",
    fontFamily: "Segoe UI, sans-serif",
  },
  title: {
    fontSize: "20px",
    fontWeight: 600,
    marginBottom: "12px",
  },
  button: {
    backgroundColor: "#2b579a",
    color: "#fff",
    padding: "8px 12px",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    marginBottom: "14px",
  },
  resultsContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "70vh",
    overflowY: "auto",
  },
  resultBox: {
    border: "1px solid #ccc",
    borderRadius: "6px",
    padding: "10px",
    transition: "background 0.2s ease",
  },
  message: {
    margin: 0,
    fontSize: "13px",
  },
  placeholder: {
    color: "#666",
    fontStyle: "italic",
  },
};
