import React, { useState, useEffect } from "react";
import { analyzeFormatting, checkDocument, checkStyles, checkHeaderFooterFormatting, checkSymbols } from "../services/checks";
import { autoFixIssues } from "../services/fixes";
import { goToError } from "../services/navigation";

// Add CSS animation for spinner
const spinnerAnimation = `
  @keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
  }
`;

export default function App() {
  // Inject the animation styles
  useEffect(() => {
    const styleSheet = document.createElement("style");
    styleSheet.textContent = spinnerAnimation;
    document.head.appendChild(styleSheet);
    return () => {
      document.head.removeChild(styleSheet);
    };
  }, []);

  const [results, setResults] = useState([]);
  const [docResults, setDocResults] = useState([]);
  const [styleResults, setStyleResults] = useState([]);
  const [symbolResults, setSymbolResults] = useState([]);
  const [compResults, setCompResults] = useState([]);
  const [headerFooterResults, setHeaderFooterResults] = useState([]);

  const [isChecking, setIsChecking] = useState(false);
  const [isCheckingDocument, setIsCheckingDocument] = useState(false);
  const [isCheckingStyles, setIsCheckingStyles] = useState(false);
  const [isCheckingSymbols, setIsCheckingSymbols] = useState(false);
  const [isCheckingComp, setIsCheckingComp] = useState(false);
  const [isCheckingHeaderFooter, setIsCheckingHeaderFooter] = useState(false);
  
  const [autoFixEnabled, setAutoFixEnabled] = useState(false);
  const [fixResults, setFixResults] = useState(null);
  const [isFixing, setIsFixing] = useState(false);

  // Track completed checks with no issues
  const [checkComplete, setCheckComplete] = useState({
    formatting: false,
    document: false,
    styles: false,
    symbols: false,
    comprehensive: false,
    headerFooter: false
  });

  //Run formatting analysis
  const handleRunCheck = async () => {
    try {
      setIsChecking(true);
      setCheckComplete({...checkComplete, formatting: false});
      const issues = await analyzeFormatting();
      setResults(issues);
      
      // Set complete flag if no issues
      if (issues.length === 0) {
        setCheckComplete(prev => ({...prev, formatting: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && issues.length > 0) {
        await handleAutoFix(issues, setResults);
      }
    } catch (err) {
      console.error("Error running formatting checks:", err);
    } finally {
      setIsChecking(false);
    }
  };

  const handleRunDocumentCheck = async () => {
    try {
      setIsCheckingDocument(true);
      setCheckComplete(prev => ({...prev, document: false}));
      const issues = await checkDocument();
      setDocResults(issues);
      
      // Set complete flag if no issues
      if (issues.length === 0) {
        setCheckComplete(prev => ({...prev, document: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && issues.length > 0) {
        await handleAutoFix(issues, setDocResults);
      }
    } catch (err) {
      console.error("Error running formatting checks:", err);
    } finally {
      setIsCheckingDocument(false);
    }
  };

  const handleRunStylesCheck = async () => {
    try {
      setIsCheckingStyles(true);
      setCheckComplete(prev => ({...prev, styles: false}));
      const issues = await checkStyles();
      setStyleResults(issues);
      
      // Set complete flag if no issues
      if (issues.length === 0) {
        setCheckComplete(prev => ({...prev, styles: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && issues.length > 0) {
        await handleAutoFix(issues, setStyleResults);
      }
    } catch (err) {
      console.error("Error running style checks:", err);
    } finally {
      setIsCheckingStyles(false);
    }
  };

  const handleRunSymbolsCheck = async () => {
    try {
      setIsCheckingSymbols(true);
      setCheckComplete(prev => ({...prev, symbols: false}));
      const issues = await checkSymbols();
      setSymbolResults(issues);
      
      // Set complete flag if no issues
      if (issues.length === 0) {
        setCheckComplete(prev => ({...prev, symbols: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && issues.length > 0) {
        await handleAutoFix(issues, setSymbolResults);
      }
    } catch (err) {
      console.error("Error running symbols checks:", err);
    } finally {
      setIsCheckingSymbols(false);
    }
  };

  const handleRunComprehensiveCheck = async() => {
    try {
      setIsCheckingComp(true);
      setCheckComplete(prev => ({...prev, comprehensive: false}));
      const formatting_issues = await analyzeFormatting();
      const general_doc_issues = await checkDocument();
      const style_issues = await checkStyles();
      const symbol_issues = await checkSymbols();
      const all_issues = [...formatting_issues, ...general_doc_issues, ...style_issues, ...symbol_issues]
      setCompResults(all_issues)
      
      // Set complete flag if no issues
      if (all_issues.length === 0) {
        setCheckComplete(prev => ({...prev, comprehensive: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && all_issues.length > 0) {
        await handleAutoFix(all_issues, setCompResults);
      }

    } catch (err) {
      console.error("Error running comprehensive checks:", err)
    } finally {
      setIsCheckingComp(false);
    }
  }

  // Run header/footer checker independently
  const handleRunHeaderFooterCheck = async () => {
    try {
      setIsCheckingHeaderFooter(true);
      setCheckComplete(prev => ({...prev, headerFooter: false}));
      const issues = await checkHeaderFooterFormatting();
      setHeaderFooterResults(issues);
      
      // Set complete flag if no issues
      if (issues.length === 0) {
        setCheckComplete(prev => ({...prev, headerFooter: true}));
      }
      
      // Auto-fix if enabled
      if (autoFixEnabled && issues.length > 0) {
        await handleAutoFix(issues, setHeaderFooterResults);
      }
    } catch (err) {
      console.error("Error running header/footer checks:", err);
    } finally {
      setIsCheckingHeaderFooter(false);
    }
  };

  // Auto-fix handler
  const handleAutoFix = async (issues, setResultsFn) => {
    try {
      setIsFixing(true);
      const fixableIssues = issues.filter(issue => issue.canLocate && issue.location);
      
      if (fixableIssues.length === 0) {
        setFixResults({
          fixed: [],
          unfixed: issues,
          summary: "No fixable issues found"
        });
        return;
      }

      const result = await autoFixIssues(fixableIssues);
      setFixResults(result);

      // Update the results to show only unfixed issues
      if (setResultsFn) {
        setResultsFn(result.unfixed);
      }

      console.log("Auto-fix results:", result);
    } catch (err) {
      console.error("Error during auto-fix:", err);
      setFixResults({
        fixed: [],
        unfixed: issues,
        summary: `Error: ${err.message}`
      });
    } finally {
      setIsFixing(false);
    }
  };

  // Manual fix button for existing results
  const handleManualFix = async (resultsArray, setResultsFn) => {
    await handleAutoFix(resultsArray, setResultsFn);
  };

  // Color coding based on issue category
  const getCategoryColor = (issueType) => {
    const formattingTypes = ["Highlighting", "Hidden Text", "Font Color", "Font Size", "Font", "Justification", "Hyperlink", "Blank Page", "Table Header"];
    const documentTypes = ["Comment", "Revision", "Text Box", "Watermark", "Invalid Reference", "Page Size"];
    const headerFooterTypes = ["Header", "Footer"];
    const marginTypes = ["Margins"];
    const styleTypes = ["Caption", "Blank Paragraph Mark", "Section/Page Breaks", "Heading"];
    const symbolTypes = ["Symbols"];

    if (formattingTypes.includes(issueType)) {
      return "#e3f2fd"; // Light blue for formatting
    } else if (documentTypes.includes(issueType)) {
      return "#fff3e0"; // Light orange for general document
    } else if (headerFooterTypes.includes(issueType)) {
      return "#f3e5f5"; // Light purple for headers/footers
    } else if (marginTypes.includes(issueType)) {
      return "#fce4ec"; // Light pink for margins
    } else if (styleTypes.includes(issueType)) {
      return "#e8f5e9"; // Light green for styles
    } else if (symbolTypes.includes(issueType)) {
      return "#fff9c4"; // Light yellow for symbols
    } else {
      return "#f9f9f9"; // Default gray
    }
  };

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
        
        {/* Color Legend */}
        <div style={styles.legend}>
          <div style={styles.legendTitle}>Issue Categories:</div>
          <div style={styles.legendItems}>
            <span style={{...styles.legendItem, backgroundColor: "#e3f2fd"}}>Formatting</span>
            <span style={{...styles.legendItem, backgroundColor: "#fff3e0"}}>Document</span>
            <span style={{...styles.legendItem, backgroundColor: "#f3e5f5"}}>Header/Footer</span>
            <span style={{...styles.legendItem, backgroundColor: "#fce4ec"}}>Margins</span>
            <span style={{...styles.legendItem, backgroundColor: "#e8f5e9"}}>Styles</span>
            <span style={{...styles.legendItem, backgroundColor: "#fff9c4"}}>Symbols</span>
          </div>
        </div>
        
        {/* Auto-fix Toggle */}
        <div style={styles.autoFixToggle}>
          <label style={styles.toggleLabel}>
            <input
              type="checkbox"
              checked={autoFixEnabled}
              onChange={(e) => setAutoFixEnabled(e.target.checked)}
              style={styles.checkbox}
            />
            <span>Automatically fix issues</span>
          </label>
        </div>

        {/* Fix Results Display */}
        {fixResults && (
          <div style={styles.fixResultsBox}>
            <strong>Auto-Fix Results:</strong> {fixResults.summary}
            {fixResults.fixed.length > 0 && (
              <div style={styles.fixedList}>
                <p>✓ Fixed {fixResults.fixed.length} issues</p>
              </div>
            )}
            {fixResults.unfixed.length > 0 && (
              <div style={styles.unfixedList}>
                <p>⚠ {fixResults.unfixed.length} issues could not be fixed automatically</p>
              </div>
            )}
          </div>
        )}

        {/* Loading Message */}
        {(isCheckingComp || isFixing) && (
          <div style={styles.loadingMessage}>
            <div style={styles.loadingContent}>
              <div style={styles.spinner}></div>
              <p>{isFixing ? "Fixing issues, please hold..." : "Check is currently running, please hold..."}</p>
            </div>
          </div>
        )}

        <button onClick = {handleRunComprehensiveCheck} style={styles.button} disabled={isCheckingComp || isFixing}>
          {isChecking || isFixing ? "Processing..." : "Run Comprehensive Check"}
        </button>
        
        {/* Manual Fix Button */}
        {!autoFixEnabled && compResults.length > 0 && (
          <button 
            onClick={() => handleManualFix(compResults, setCompResults)} 
            style={{...styles.button, backgroundColor: "#28a745", marginLeft: "8px"}}
            disabled={isFixing}
          >
            {isFixing ? "Fixing..." : "Fix Issues Now"}
          </button>
        )}

        {/* Success Message - No Issues Found */}
        {checkComplete.comprehensive && compResults.length === 0 && (
          <div style={styles.successMessage}>
            <span style={styles.checkmark}>✓</span>
            <span>Check is complete - No issues found</span>
          </div>
        )}

       <div style={styles.resultsContainer}>
          {compResults.length === 0 && !isCheckingComp && !checkComplete.comprehensive && (
            <p style={styles.placeholder}>No results yet. Click "Run Comprehensive Check".</p>
          )}

          {compResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: getCategoryColor(r.type),
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

        {/* Loading Message */}
        {(isChecking || isFixing) && (
          <div style={styles.loadingMessage}>
            <div style={styles.loadingContent}>
              <div style={styles.spinner}></div>
              <p>{isFixing ? "Fixing issues, please hold..." : "Check is currently running, please hold..."}</p>
            </div>
          </div>
        )}

        <button onClick={handleRunCheck} style={styles.button} disabled={isChecking || isFixing}>
          {isChecking || isFixing ? "Processing..." : "Run Formatting Check"}
        </button>
        
        {/* Manual Fix Button */}
        {!autoFixEnabled && results.length > 0 && (
          <button 
            onClick={() => handleManualFix(results, setResults)} 
            style={{...styles.button, backgroundColor: "#28a745", marginLeft: "8px"}}
            disabled={isFixing}
          >
            {isFixing ? "Fixing..." : "Fix Issues Now"}
          </button>
        )}

        {/* Success Message - No Issues Found */}
        {checkComplete.formatting && results.length === 0 && (
          <div style={styles.successMessage}>
            <span style={styles.checkmark}>✓</span>
            <span>Check is complete - No issues found</span>
          </div>
        )}

        <div style={styles.resultsContainer}>
          {results.length === 0 && !isChecking && !checkComplete.formatting && (
            <p style={styles.placeholder}>No results yet. Click "Run Formatting Check".</p>
          )}

          {results.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: getCategoryColor(r.type),
              }}
            >
              <b>{r.type}</b>
              <p style={styles.message}>{r.message}</p>
            </div>
          ))}
        </div>
      </div>

      {/* NEW: Header/Footer Checker section */}
      <div style={styles.container}>
        <h2 style={styles.title}>Header/Footer Checker</h2>

        {/* Loading Message */}
        {isCheckingHeaderFooter && (
          <div style={styles.loadingMessage}>
            <div style={styles.loadingContent}>
              <div style={styles.spinner}></div>
              <p>Check is currently running, please hold...</p>
            </div>
          </div>
        )}

        <button
          onClick={handleRunHeaderFooterCheck}
          style={styles.button}
          disabled={isCheckingHeaderFooter}
        >
          {isCheckingHeaderFooter
            ? "Checking..."
            : "Run Header/Footer Check"}
        </button>

        {/* Manual Fix Button */}
        {!autoFixEnabled && headerFooterResults.length > 0 && (
          <button 
            onClick={() => handleManualFix(headerFooterResults, setHeaderFooterResults)} 
            style={{...styles.button, backgroundColor: "#28a745", marginLeft: "8px"}}
            disabled={isFixing}
          >
            {isFixing ? "Fixing..." : "Fix Issues Now"}
          </button>
        )}

        {/* Success Message - No Issues Found */}
        {checkComplete.headerFooter && headerFooterResults.length === 0 && (
          <div style={styles.successMessage}>
            <span style={styles.checkmark}>✓</span>
            <span>Check is complete - No issues found</span>
          </div>
        )}

        <div style={styles.resultsContainer}>
          {headerFooterResults.length === 0 && !isCheckingHeaderFooter && !checkComplete.headerFooter && (
            <p style={styles.placeholder}>
              No results yet. Click "Run Header/Footer Check".
            </p>
          )}

          {headerFooterResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: getCategoryColor(r.type),
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

        {/* Loading Message */}
        {(isCheckingDocument || isFixing) && (
          <div style={styles.loadingMessage}>
            <div style={styles.loadingContent}>
              <div style={styles.spinner}></div>
              <p>{isFixing ? "Fixing issues, please hold..." : "Check is currently running, please hold..."}</p>
            </div>
          </div>
        )}

        <button onClick={handleRunDocumentCheck} style={styles.button} disabled={isCheckingDocument || isFixing}>
          {isCheckingDocument || isFixing ? "Processing..." : "Run Document Check"}
        </button>
        
        {/* Manual Fix Button */}
        {!autoFixEnabled && docResults.length > 0 && (
          <button 
            onClick={() => handleManualFix(docResults, setDocResults)} 
            style={{...styles.button, backgroundColor: "#28a745", marginLeft: "8px"}}
            disabled={isFixing}
          >
            {isFixing ? "Fixing..." : "Fix Issues Now"}
          </button>
        )}

        {/* Success Message - No Issues Found */}
        {checkComplete.document && docResults.length === 0 && (
          <div style={styles.successMessage}>
            <span style={styles.checkmark}>✓</span>
            <span>Check is complete - No issues found</span>
          </div>
        )}

        <div style={styles.resultsContainer}>
          {docResults.length === 0 && !isCheckingDocument && !checkComplete.document && (
            <p style={styles.placeholder}>No results yet. Click "Run Document Check".</p>
          )}

          {docResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: getCategoryColor(r.type),
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

        {/* Loading Message */}
        {(isCheckingStyles || isFixing) && (
          <div style={styles.loadingMessage}>
            <div style={styles.loadingContent}>
              <div style={styles.spinner}></div>
              <p>{isFixing ? "Fixing issues, please hold..." : "Check is currently running, please hold..."}</p>
            </div>
          </div>
        )}

        <button onClick={handleRunStylesCheck} style={styles.button} disabled={isCheckingStyles || isFixing}>
          {isCheckingStyles || isFixing ? "Processing..." : "Run Styles Check"}
        </button>
        
        {/* Manual Fix Button */}
        {!autoFixEnabled && styleResults.length > 0 && (
          <button 
            onClick={() => handleManualFix(styleResults, setStyleResults)} 
            style={{...styles.button, backgroundColor: "#28a745", marginLeft: "8px"}}
            disabled={isFixing}
          >
          {isFixing ? "Fixing..." : "Fix Issues Now"}
          </button>
        )}

        {/* Success Message - No Issues Found */}
        {checkComplete.styles && styleResults.length === 0 && (
          <div style={styles.successMessage}>
            <span style={styles.checkmark}>✓</span>
            <span>Check is complete - No issues found</span>
          </div>
        )}

        <div style={styles.resultsContainer}>
          {styleResults.length === 0 && !isCheckingStyles && !checkComplete.styles && (
            <p style={styles.placeholder}>No results yet. Click "Run Styles Check".</p>
          )}

          {styleResults.map((r) => (
            <div
              key={r.id}
              onClick={() => handleGoTo(r)}
              style={{
                ...styles.resultBox,
                cursor: r.canLocate ? "pointer" : "default",
                backgroundColor: getCategoryColor(r.type),
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
  autoFixToggle: {
    marginBottom: "16px",
    padding: "12px",
    backgroundColor: "#f0f8ff",
    border: "1px solid #2b579a",
    borderRadius: "6px",
  },
  toggleLabel: {
    display: "flex",
    alignItems: "center",
    cursor: "pointer",
    fontSize: "14px",
    fontWeight: 500,
  },
  checkbox: {
    marginRight: "8px",
    width: "18px",
    height: "18px",
    cursor: "pointer",
  },
  fixResultsBox: {
    marginBottom: "16px",
    padding: "12px",
    backgroundColor: "#e8f5e9",
    border: "1px solid #4caf50",
    borderRadius: "6px",
    fontSize: "14px",
  },
  fixedList: {
    marginTop: "8px",
    color: "#2e7d32",
    fontSize: "13px",
  },
  unfixedList: {
    marginTop: "8px",
    color: "#f57c00",
    fontSize: "13px",
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
  loadingMessage: {
    position: "fixed",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    backgroundColor: "rgba(0, 0, 0, 0.85)",
    color: "#fff",
    padding: "24px 32px",
    borderRadius: "8px",
    zIndex: 1000,
    boxShadow: "0 4px 12px rgba(0, 0, 0, 0.3)",
    minWidth: "250px",
    textAlign: "center",
  },
  loadingContent: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
  },
  spinner: {
    width: "40px",
    height: "40px",
    border: "4px solid rgba(255, 255, 255, 0.3)",
    borderTop: "4px solid #fff",
    borderRadius: "50%",
    animation: "spin 1s linear infinite",
  },
  legend: {
    marginBottom: "16px",
    padding: "12px",
    backgroundColor: "#f5f5f5",
    border: "1px solid #ddd",
    borderRadius: "6px",
  },
  legendTitle: {
    fontSize: "13px",
    fontWeight: 600,
    marginBottom: "8px",
    color: "#333",
  },
  legendItems: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  legendItem: {
    padding: "4px 12px",
    borderRadius: "4px",
    fontSize: "12px",
    fontWeight: 500,
    border: "1px solid #ccc",
  },
  successMessage: {
    display: "flex",
    alignItems: "center",
    padding: "12px 16px",
    marginBottom: "12px",
    backgroundColor: "#d4edda",
    border: "1px solid #28a745",
    borderRadius: "6px",
    color: "#155724",
    fontSize: "14px",
    fontWeight: 500,
  },
  checkmark: {
    fontSize: "20px",
    fontWeight: "bold",
    color: "#28a745",
    marginRight: "10px",
    backgroundColor: "#fff",
    borderRadius: "50%",
    width: "24px",
    height: "24px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    border: "2px solid #28a745",
  },
};
