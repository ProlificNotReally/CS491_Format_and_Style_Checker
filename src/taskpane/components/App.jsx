import React, { useState } from "react";
import { analyzeFormatting, goToError } from "../wordChecks";
import { checkDocument } from "../docChecks";
import { checkStyles } from "../checkStyles";
import { checkSymbols, fixSymbolIssue, fixAllSymbolIssues } from "../symbolChecks";
import { checkHeaderFooterFormatting, fixHeaderFooterIssue, fixAllHeaderFooterIssues } from "../checkHeaderFooterFormatting";

export default function App() {
  const [results, setResults] = useState([]);
  const [docResults, setDocResults] = useState([]);
  const [styleResults, setStyleResults] = useState([]);
  const [compResults, setCompResults] = useState([]);
  const [headerFooterResults, setHeaderFooterResults] = useState([]);
  const [symbolResults, setSymbolResults] = useState([]);

  const [isChecking, setIsChecking] = useState(false);
  const [isCheckingDocument, setIsCheckingDocument] = useState(false);
  const [isCheckingStyles, setIsCheckingStyles] = useState(false);
  const [isCheckingComp, setIsCheckingComp] = useState(false);
  const [isCheckingHeaderFooter, setIsCheckingHeaderFooter] = useState(false);
  const [isFixingHeaderFooter, setIsFixingHeaderFooter] = useState(false);
  const [isCheckingSymbols, setIsCheckingSymbols] = useState(false);
  const [isFixingSymbols, setIsFixingSymbols] = useState(false);
  const [fixingItemId, setFixingItemId] = useState(null);

  const [hasRunFormatting, setHasRunFormatting] = useState(false);
  const [hasRunDocument, setHasRunDocument] = useState(false);
  const [hasRunStyles, setHasRunStyles] = useState(false);
  const [hasRunComp, setHasRunComp] = useState(false);
  const [hasRunHeaderFooter, setHasRunHeaderFooter] = useState(false);
  const [hasRunSymbols, setHasRunSymbols] = useState(false);

  //Run formatting analysis
  const handleRunCheck = async () => {
    try {
      setIsChecking(true);
      const issues = await analyzeFormatting();
      setResults(issues);
      setHasRunFormatting(true);
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
      setHasRunDocument(true);
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
      setHasRunStyles(true);
    } catch (err) {
      console.error("Error running style checks:", err);
    } finally {
      setIsCheckingStyles(false);
    }
  };

  const handleRunSymbolsCheck = async () => {
    try {
      setIsCheckingSymbols(true);
      const issues = await checkSymbols();
      setSymbolResults(issues);
      setHasRunSymbols(true);
    } catch (err) {
      console.error("Error running symbols check:", err);
    } finally {
      setIsCheckingSymbols(false);
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
      setHasRunComp(true);

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
      const issues = await checkHeaderFooterFormatting();
      setHeaderFooterResults(issues);
      setHasRunHeaderFooter(true);
    } catch (err) {
      console.error("Error running header/footer checks:", err);
    } finally {
      setIsCheckingHeaderFooter(false);
    }
  };

  // Fix individual header/footer issue
  const handleFixSingleIssue = async (issue) => {
    try {
      console.log("Fixing issue:", issue);
      const result = await fixHeaderFooterIssue(issue);
      console.log("Fix result:", result);
      
      if (result.success) {
        // Re-run check to update results
        const updatedIssues = await checkHeaderFooterFormatting();
        setHeaderFooterResults(updatedIssues);
        alert(`✅ ${result.message}`);
      } else {
        alert(`❌ Could not fix: ${result.message}\n\nIssue ID: ${issue.id}`);
      }
    } catch (err) {
      console.error("Error fixing issue:", err);
      alert(`❌ Error: ${err.message}\n\nIssue ID: ${issue.id}`);
    }
  };

  // Fix all header/footer issues at once
  const handleFixAllHeaderFooter = async () => {
    try {
      setIsFixingHeaderFooter(true);
      const fixableIssues = headerFooterResults.filter(
        r => r.type !== "Info" && !r.id.includes("draft") && !r.id.includes("inconsistent")
      );
      
      if (fixableIssues.length === 0) {
        alert("No fixable issues found.");
        return;
      }

      console.log("Fixing issues:", fixableIssues.map(i => i.id));
      const results = await fixAllHeaderFooterIssues(fixableIssues);
      console.log("Fix results:", results);
      
      const successCount = results.filter(r => r.success).length;
      const failedIssues = results.filter(r => !r.success);
      
      // Re-run check to update results
      const updatedIssues = await checkHeaderFooterFormatting();
      setHeaderFooterResults(updatedIssues);
      
      let message = `✅ Fixed ${successCount} out of ${fixableIssues.length} issues.`;
      if (failedIssues.length > 0) {
        message += `\n\n❌ Failed to fix ${failedIssues.length} issues:\n`;
        message += failedIssues.slice(0, 3).map(f => `- ${f.message}`).join('\n');
        if (failedIssues.length > 3) {
          message += `\n... and ${failedIssues.length - 3} more`;
        }
      }
      alert(message);
    } catch (err) {
      console.error("Error fixing all issues:", err);
      alert(`❌ Error: ${err.message}`);
    } finally {
      setIsFixingHeaderFooter(false);
    }
  };

  // Fix individual symbol issue
  const handleFixSingleSymbol = async (issue) => {
    setFixingItemId(issue.id);
    await new Promise(resolve => setTimeout(resolve, 10)); // Ensure UI updates
    
    try {
      console.log("Fixing symbol issue:", issue);
      const result = await fixSymbolIssue(issue);
      console.log("Fix result:", result);
      
      if (result.success) {
        const updatedIssues = await checkSymbols();
        setSymbolResults(updatedIssues);
        alert(`✅ Symbol font applied successfully!`);
      } else {
        alert(`❌ Could not fix: ${result.error}\n\nIssue ID: ${issue.id}`);
      }
    } catch (err) {
      console.error("Error fixing symbol issue:", err);
      alert(`❌ Error: ${err.message}\n\nIssue ID: ${issue.id}`);
    } finally {
      setFixingItemId(null);
    }
  };

  // Fix all symbol issues at once
  const handleFixAllSymbols = async () => {
    try {
      setIsFixingSymbols(true);
      
      if (symbolResults.length === 0) {
        alert("No fixable issues found.");
        return;
      }

      console.log("Fixing symbol issues:", symbolResults.map(i => i.id));
      const results = await fixAllSymbolIssues(symbolResults);
      console.log("Fix results:", results);
      
      const successCount = results.filter(r => r.success).length;
      const failedIssues = results.filter(r => !r.success);
      
      const updatedIssues = await checkSymbols();
      setSymbolResults(updatedIssues);
      
      let message = `✅ Fixed ${successCount} out of ${symbolResults.length} issues.`;
      if (failedIssues.length > 0) {
        message += `\n\n❌ Failed to fix ${failedIssues.length} issues:\n`;
        message += failedIssues.slice(0, 3).map(f => `- ${f.error}`).join('\n');
        if (failedIssues.length > 3) {
          message += `\n... and ${failedIssues.length - 3} more`;
        }
      }
      alert(message);
    } catch (err) {
      console.error("Error fixing all symbol issues:", err);
      alert(`❌ Error: ${err.message}`);
    } finally {
      setIsFixingSymbols(false);
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
      {/* 1. Comprehensive Checker */}
      <div style={styles.container}>
        <h1 style={styles.title}>Comprehensive Checker</h1>

        <button onClick = {handleRunComprehensiveCheck} style={{ ...styles.button, backgroundColor: "#914137" }} disabled={isCheckingComp}>
          {isCheckingComp ? "Checking..." : "Run Comprehensive Check"}
        </button>

       <div style={styles.resultsContainer}>
          {isCheckingComp && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {compResults.length === 0 && !isCheckingComp && !hasRunComp && (
            <p style={styles.placeholder}>No results yet. Click "Run Comprehensive Check".</p>
          )}

          {compResults.length === 0 && !isCheckingComp && hasRunComp && (
            <p style={styles.successMessage}>🎉 Congrats! No errors found.</p>
          )}

          {!isCheckingComp && compResults.map((r) => (
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

      {/* 2. Formatting Checker */}
      <div style={styles.container}>
        <h2 style={styles.title}>Formatting Checker</h2>

        <button onClick={handleRunCheck} style={{ ...styles.button, backgroundColor: "#29423f" }} disabled={isChecking}>
          {isChecking ? "Checking..." : "Run Formatting Check"}
        </button>

        <div style={styles.resultsContainer}>
          {isChecking && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {results.length === 0 && !isChecking && !hasRunFormatting && (
            <p style={styles.placeholder}>No results yet. Click "Run Formatting Check".</p>
          )}

          {results.length === 0 && !isChecking && hasRunFormatting && (
            <p style={styles.successMessage}>🎉 Congrats! No errors found.</p>
          )}

          {!isChecking && results.map((r) => (
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

      {/* 3. General Document Checker */}
      <div style={styles.container}>
        <h2 style={styles.title}>General Document Checker</h2>

        <button onClick={handleRunDocumentCheck} style={{ ...styles.button, backgroundColor: "#7c152d" }} disabled={isCheckingDocument}>
          {isCheckingDocument ? "Checking..." : "Run Document Check"}
        </button>

        <div style={styles.resultsContainer}>
          {isCheckingDocument && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {docResults.length === 0 && !isCheckingDocument && !hasRunDocument && (
            <p style={styles.placeholder}>No results yet. Click "Run Document Check".</p>
          )}

          {docResults.length === 0 && !isCheckingDocument && hasRunDocument && (
            <p style={styles.successMessage}>🎉 Congrats! No errors found.</p>
          )}

          {!isCheckingDocument && docResults.map((r) => (
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

      {/* 4. Headers and Footers Checker */}
      <div style={styles.container}>
        <h2 style={styles.title}>Headers and Footers Checker</h2>

        <div style={{ display: "flex", gap: "10px", marginBottom: "14px" }}>
          <button
            onClick={handleRunHeaderFooterCheck}
            style={{ ...styles.button, backgroundColor: "#684e4e" }}
            disabled={isCheckingHeaderFooter}
          >
            {isCheckingHeaderFooter
              ? "Checking..."
              : "Run Header/Footer Check"}
          </button>

          {headerFooterResults.length > 0 && (
            <button
              onClick={handleFixAllHeaderFooter}
              style={{ ...styles.button, backgroundColor: "#107c10" }}
              disabled={isFixingHeaderFooter || isCheckingHeaderFooter}
            >
              {isFixingHeaderFooter ? "Fixing..." : "Fix All Issues"}
            </button>
          )}
        </div>

        <div style={styles.resultsContainer}>
          {isCheckingHeaderFooter && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {headerFooterResults.length === 0 && !isCheckingHeaderFooter && !hasRunHeaderFooter && (
            <p style={styles.placeholder}>
              No results yet. Click "Run Header/Footer Check".
            </p>
          )}

          {headerFooterResults.length === 0 && !isCheckingHeaderFooter && hasRunHeaderFooter && (
            <p style={styles.successMessage}>
              🎉 All header/footer checks complete!<br />
              ✓ Portrait Header style applied<br />
              ✓ Portrait Footer style applied<br />
              ✓ Landscape Header style applied<br />
              ✓ Landscape Footer style applied<br />
              ✓ "Draft" appears in headers<br />
              ✓ Header information consistent on each page<br />
              ✓ Footer information consistent on each page<br />
              ✓ Header/footer margin settings conform to 0.5"
            </p>
          )}

          {!isCheckingHeaderFooter && headerFooterResults.map((r) => {
            const canFix = r.type !== "Info" && !r.id.includes("draft") && !r.id.includes("inconsistent");
            
            return (
              <div
                key={r.id}
                style={{
                  ...styles.resultBox,
                  backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div
                    onClick={() => handleGoTo(r)}
                    style={{ 
                      flex: 1, 
                      cursor: r.canLocate ? "pointer" : "default" 
                    }}
                  >
                    <b>{r.type}</b>
                    <p style={styles.message}>{r.message}</p>
                  </div>
                  
                  {canFix && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        handleFixSingleIssue(r);
                      }}
                      style={styles.fixButton}
                    >
                      Fix
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 5. Margins Checker - Placeholder */}
      <div style={styles.container}>
        <h2 style={styles.title}>Margins Checker</h2>
        <p style={styles.placeholder}>Coming soon - Margin validation will be added here.</p>
      </div>

      {/* 6. Styles Checker */}
      <div style={styles.container}>
        <h2 style={styles.title}>Styles Checker</h2>

        <button onClick={handleRunStylesCheck} style={{ ...styles.button, backgroundColor: "#ef6641" }} disabled={isCheckingStyles}>
          {isCheckingStyles ? "Checking..." : "Run Styles Check"}
        </button>

        <div style={styles.resultsContainer}>
          {isCheckingStyles && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {styleResults.length === 0 && !isCheckingStyles && !hasRunStyles && (
            <p style={styles.placeholder}>No results yet. Click "Run Styles Check".</p>
          )}

          {styleResults.length === 0 && !isCheckingStyles && hasRunStyles && (
            <p style={styles.successMessage}>🎉 Congrats! No errors found.</p>
          )}

          {!isCheckingStyles && styleResults.map((r) => (
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

      {/* 7. Symbols Checker */}
      <div style={styles.container}>
        <h2 style={styles.title}>Symbols Checker</h2>

        <div style={{ display: "flex", gap: "10px", marginBottom: "14px" }}>
          <button
            onClick={handleRunSymbolsCheck}
            style={{ ...styles.button, backgroundColor: "#365d9f" }}
            disabled={isCheckingSymbols}
          >
            {isCheckingSymbols ? "Checking..." : "Run Symbols Check"}
          </button>

          {symbolResults.length > 0 && (
            <button
              onClick={handleFixAllSymbols}
              style={{ ...styles.button, backgroundColor: "#107c10" }}
              disabled={isFixingSymbols || isCheckingSymbols}
            >
              {isFixingSymbols ? "Fixing..." : "Fix All Issues"}
            </button>
          )}
        </div>

        <div style={styles.resultsContainer}>
          {isCheckingSymbols && (
            <p style={styles.loadingMessage}>Loading...</p>
          )}

          {symbolResults.length === 0 && !isCheckingSymbols && !hasRunSymbols && (
            <p style={styles.placeholder}>No results yet. Click "Run Symbols Check".</p>
          )}

          {symbolResults.length === 0 && !isCheckingSymbols && hasRunSymbols && (
            <p style={styles.successMessage}>🎉 Congrats! No errors found.</p>
          )}

          {!isCheckingSymbols && symbolResults.map((r) => (
            <div
              key={r.id}
              style={{
                ...styles.resultBox,
                backgroundColor: r.canLocate ? "#eef5ff" : "#f9f9f9",
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div
                  onClick={() => handleGoTo(r)}
                  style={{ 
                    flex: 1, 
                    cursor: r.canLocate ? "pointer" : "default" 
                  }}
                >
                  <b>{r.type}</b>
                  <p style={styles.message}>{r.message}</p>
                </div>
                
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    handleFixSingleSymbol(r);
                  }}
                  style={styles.fixButton}
                  disabled={fixingItemId === r.id || isFixingSymbols}
                >
                  {fixingItemId === r.id ? "Fixing..." : "Fix"}
                </button>
              </div>
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
    fontFamily: "'Times New Roman', Times, serif",
  },
  title: {
    fontSize: "20px",
    fontWeight: 600,
    marginBottom: "12px",
    fontFamily: "'Times New Roman', Times, serif",
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
    color: "#2b579a",
    fontStyle: "italic",
    fontSize: "14px",
    textAlign: "center",
    padding: "10px",
  },
  successMessage: {
    color: "#107c10",
    fontWeight: "600",
    fontSize: "16px",
    textAlign: "left",
    padding: "20px",
    backgroundColor: "#f0f9f0",
    borderRadius: "6px",
    border: "2px solid #107c10",
  },
  fixButton: {
    backgroundColor: "#107c10",
    color: "#fff",
    padding: "6px 12px",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    fontSize: "12px",
    marginLeft: "10px",
    flexShrink: 0,
  },
};
