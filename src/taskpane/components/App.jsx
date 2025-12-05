import React, { useState } from "react";
import { analyzeFormatting, goToError } from "../wordChecks";
import { checkDocument } from "../docChecks";
import { checkStyles } from "../checkStyles";

export default function App() {
  const [results, setResults] = useState([]);
  const [docResults, setDocResults] = useState([]);
  const [styleResults, setStyleResults] = useState([]);
  const [compResults, setCompResults] = useState([]);
  const [isChecking, setIsChecking] = useState(false);
  const [isCheckingDocument, setIsCheckingDocument] = useState(false);
  const [isCheckingStyles, setIsCheckingStyles] = useState(false);
  const [isCheckingComp, setIsCheckingComp] = useState(false);

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
      const general_doc_issues = await checkDocument();
      const style_issues = await checkStyles();
      const all_issues = [...formatting_issues, ...general_doc_issues, ...style_issues]
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

      <div style={styles.container}>
        <h2 style={styles.title}>Document Checker</h2>

        <button onClick={handleRunDocumentCheck} style={styles.button} disabled={isCheckingDocument}>
          {isCheckingDocument ? "Checking..." : "Run Document Check"}
        </button>

        <div style={styles.resultsContainer}>
          {docResults.length === 0 && !isCheckingDocument && (
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
