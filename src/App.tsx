import { useState } from "react";
import {
  Shield,
  Search,
  AlertTriangle,
  Download,
  Lock,
  User,
  FileText,
  LogIn,
  RefreshCw,
  CheckCircle2,
} from "lucide-react";
import {
  useMsal,
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import {
  getSites,
  getDrives,
  getDriveItems,
  getFileContent,
} from "./services/graphService";
import {
  scanForBSN,
  analyzePermissions,
  getSeverity,
} from "./utils/auditEngine";
import type { AuditResult } from "./utils/auditEngine";
import "./App.css";

function Dashboard({
  isDemo,
  onLogout,
}: {
  isDemo?: boolean;
  onLogout: () => void;
}) {
  const { instance, accounts } = useMsal();
  const [isScanning, setIsScanning] = useState(false);
  const [scanScope, setScanScope] = useState("all");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [sites, setSites] = useState<any[]>([]);
  const [results, setResults] = useState<AuditResult[]>([]);
  const [progress, setProgress] = useState({
    current: 0,
    total: 0,
    status: "",
  });
  const [validatorText, setValidatorText] = useState("");
  const [isPiiDetected, setIsPiiDetected] = useState(false);

  const user = isDemo
    ? { name: "Demo Auditor", username: "demo@example.com" }
    : accounts[0];

  const handleLogout = () => {
    if (isDemo) onLogout();
    else instance.logoutPopup();
  };

  const runScan = async () => {
    setIsScanning(true);
    setResults([]);
    setProgress({ current: 0, total: 0, status: "Initializing scan..." });

    if (isDemo) {
      // Mock Scan logic for Demo purposes
      const mockSites = [
        { id: "1", displayName: "Finance & HR" },
        { id: "2", displayName: "Marketing Public" },
        { id: "3", displayName: "IT Operations" },
        { id: "4", displayName: "Executive Board" },
      ];

      setSites(mockSites);
      setProgress({
        current: 0,
        total: mockSites.length,
        status: "Simulating data retrieval...",
      });

      for (let i = 0; i < mockSites.length; i++) {
        await new Promise((r) => setTimeout(r, 800)); // Simulation delay
        setProgress((p) => ({
          ...p,
          current: i + 1,
          status: `Auditing: ${mockSites[i].displayName}`,
        }));
      }

      const mockResults: AuditResult[] = [
        {
          id: "m1",
          file: "Salaries_2025.xlsx",
          site: "Finance & HR",
          type: "Anonymous Link",
          severity: "High",
          bsnFound: true,
          linkUrl: "#",
        },
        {
          id: "m2",
          file: "Brand_Guidelines.pdf",
          site: "Marketing Public",
          type: "Everyone Link",
          severity: "Low",
          bsnFound: false,
          linkUrl: "#",
        },
        {
          id: "m3",
          file: "Network_Topology.vsdx",
          site: "IT Operations",
          type: "Everyone Link",
          severity: "Medium",
          bsnFound: false,
          linkUrl: "#",
        },
        {
          id: "m4",
          file: "Board_Meeting_Minutes.docx",
          site: "Executive Board",
          type: "Anonymous Link",
          severity: "High",
          bsnFound: true,
          linkUrl: "#",
        },
      ];

      setResults(mockResults);
      setProgress((p) => ({ ...p, status: "Demo Scan Complete!" }));
      setIsScanning(false);
      return;
    }

    try {
      const resp = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      const token = resp.accessToken;

      // 1. Get Sites
      setProgress((p) => ({ ...p, status: "Fetching SharePoint sites..." }));
      const allSites = await getSites(token);
      setSites(allSites);
      setProgress((p) => ({
        ...p,
        total: allSites.length,
        status: `Found ${allSites.length} sites. Analyzing...`,
      }));

      const auditResults: AuditResult[] = [];

      for (let i = 0; i < allSites.length; i++) {
        const site = allSites[i];
        setProgress((p) => ({
          ...p,
          current: i + 1,
          status: `Scanning site: ${site.displayName}`,
        }));

        try {
          // 2. Get Drives
          const drives = await getDrives(site.id, token);

          for (const drive of drives) {
            // 3. Get Items
            const items = await getDriveItems(drive.id, token);

            for (const item of items) {
              // 4. Analyze Permissions
              const violationType = analyzePermissions(item.permissions || []);

              if (violationType) {
                // 5. Scan Content if text-based
                let bsnFound = false;
                const isText = /\.(txt|csv|md|json)$/i.test(item.name);

                if (isText) {
                  const content = await getFileContent(
                    drive.id,
                    item.id,
                    token,
                  );
                  if (content) bsnFound = scanForBSN(content);
                }

                auditResults.push({
                  id: item.id,
                  file: item.name,
                  site: site.displayName,
                  type: violationType,
                  severity: getSeverity(violationType, bsnFound),
                  bsnFound,
                  linkUrl: item.webUrl,
                });
              }
            }
          }
        } catch (err) {
          console.error(`Error scanning site ${site.displayName}:`, err);
        }
      }

      setResults(auditResults);
      setProgress((p) => ({ ...p, status: "Scan complete!" }));
    } catch (error) {
      console.error("Scan failed:", error);
      setProgress((p) => ({ ...p, status: "Scan failed. Check permissions." }));
    } finally {
      setIsScanning(false);
    }
  };

  const handleValidatorChange = (text: string) => {
    setValidatorText(text);
    setIsPiiDetected(scanForBSN(text));
  };

  const exportReport = () => {
    const blob = new Blob([JSON.stringify(results, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `m365-security-audit-${new Date().toISOString().split("T")[0]}.json`;
    a.click();
  };

  const stats = [
    { label: "Sites Found", value: sites.length.toString(), icon: Shield },
    {
      label: "Files Analyzed",
      value: progress.current > 0 ? `${progress.current} sites` : "-",
      icon: FileText,
    },
    {
      label: "Vulnerabilities",
      value: results.length.toString(),
      icon: AlertTriangle,
      color:
        results.length > 0 ? "var(--accent-warning)" : "var(--accent-success)",
    },
    {
      label: "Critical Risks",
      value: results.filter((r) => r.severity === "High").length.toString(),
      icon: Lock,
      color: "var(--accent-danger)",
    },
  ];

  return (
    <div className="dashboard-container animate-fade">
      {isDemo && (
        <div
          style={{
            position: "absolute",
            top: "-1.5rem",
            left: "2rem",
            background: "var(--accent-warning)",
            color: "black",
            padding: "0.2rem 1rem",
            borderRadius: "1rem",
            fontSize: "0.7rem",
            fontWeight: 700,
          }}
        >
          DEMO MODE ACTIVE
        </div>
      )}
      <header className="header">
        <div className="brand">
          <div className="logo-icon">
            <Shield size={24} color="white" />
          </div>
          <div>
            <h1>
              SecureAudit M365{" "}
              {isDemo && (
                <span style={{ opacity: 0.5, fontSize: "0.9rem" }}>(Demo)</span>
              )}
            </h1>
            <p className="stat-label" style={{ margin: 0 }}>
              AI Governance & Data Security
            </p>
          </div>
        </div>

        <div
          className="user-profile"
          onClick={handleLogout}
          title="Click to Logout"
        >
          <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
            <div style={{ textAlign: "right" }}>
              <p style={{ fontWeight: 600 }}>{user?.name || user?.username}</p>
              <p className="stat-label" style={{ fontSize: "0.65rem" }}>
                Security Auditor
              </p>
            </div>
            <div style={{ position: "relative" }}>
              <User
                size={32}
                style={{
                  padding: "6px",
                  backgroundColor: "var(--bg-accent)",
                  borderRadius: "50%",
                }}
              />
              <div
                style={{
                  position: "absolute",
                  bottom: 0,
                  right: 0,
                  width: 10,
                  height: 10,
                  backgroundColor: isDemo
                    ? "var(--accent-warning)"
                    : "var(--accent-success)",
                  borderRadius: "50%",
                  border: "2px solid var(--bg-secondary)",
                }}
              ></div>
            </div>
          </div>
        </div>
      </header>

      <main>
        <div className="stats-grid">
          {stats.map((stat, i) => (
            <div key={i} className="stat-card">
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  marginBottom: "1rem",
                }}
              >
                <stat.icon
                  size={20}
                  color={stat.color || "var(--accent-blue)"}
                />
              </div>
              <div className="stat-label">{stat.label}</div>
              <div className="stat-value">{stat.value}</div>
            </div>
          ))}
        </div>

        <section className="controls-panel">
          <div className="form-group">
            <label>Scan Scope</label>
            <select
              value={scanScope}
              onChange={(e) => setScanScope(e.target.value)}
              disabled={isScanning}
            >
              <option value="all">All SharePoint Sites</option>
              <option value="selective" disabled>
                Selected Sites Only (Premium)
              </option>
              <option value="high-risk" disabled>
                High-Risk Sites Only (AI-Driven)
              </option>
            </select>
          </div>

          <div style={{ display: "flex", gap: "1rem" }}>
            <button className="primary" onClick={runScan} disabled={isScanning}>
              {isScanning ? (
                <RefreshCw
                  size={18}
                  className="animate-spin"
                  style={{ marginRight: "8px", verticalAlign: "middle" }}
                />
              ) : (
                <Search
                  size={18}
                  style={{ marginRight: "8px", verticalAlign: "middle" }}
                />
              )}
              {isScanning ? "Running Audit..." : "Start Security Scan"}
            </button>

            <button
              className="secondary"
              onClick={exportReport}
              disabled={results.length === 0 || isScanning}
            >
              <Download
                size={18}
                style={{ marginRight: "8px", verticalAlign: "middle" }}
              />
              Export JSON
            </button>
          </div>
        </section>

        <section className="validator-card animate-fade">
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: "1.5rem",
            }}
          >
            <div>
              <h3 style={{ fontSize: "1.1rem", marginBottom: "0.25rem" }}>
                Real-time PII Lab
              </h3>
              <p className="stat-label" style={{ fontSize: "0.7rem" }}>
                Paste test content below to verify security compliance
              </p>
            </div>
            <div className="pii-status-bar" style={{ marginTop: 0 }}>
              {validatorText.length === 0 ? (
                <div
                  className="status-badge"
                  style={{
                    background: "var(--bg-accent)",
                    color: "var(--text-secondary)",
                  }}
                >
                  Awaiting Input
                </div>
              ) : isPiiDetected ? (
                <div className="status-badge status-detected">
                  <AlertTriangle size={14} /> PII Detected (BSN)
                </div>
              ) : (
                <div className="status-badge status-secure">
                  <CheckCircle2 size={14} /> Secure Content
                </div>
              )}
            </div>
          </div>
          <textarea
            className="validator-input"
            placeholder="Enter or paste text here (e.g., meeting notes, customer records)..."
            value={validatorText}
            onChange={(e) => handleValidatorChange(e.target.value)}
          />
        </section>

        {isScanning && (
          <div
            style={{
              marginBottom: "2rem",
              padding: "1rem",
              background: "var(--bg-accent)",
              borderRadius: "12px",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                marginBottom: "0.5rem",
              }}
            >
              <span className="stat-label">{progress.status}</span>
              <span className="stat-label">
                {progress.total > 0
                  ? Math.round((progress.current / progress.total) * 100)
                  : 0}
                %
              </span>
            </div>
            <div
              style={{
                width: "100%",
                height: "6px",
                background: "var(--bg-primary)",
                borderRadius: "3px",
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  width: `${progress.total > 0 ? (progress.current / progress.total) * 100 : 0}%`,
                  height: "100%",
                  background:
                    "linear-gradient(90deg, var(--accent-blue), var(--accent-purple))",
                  transition: "width 0.3s ease",
                }}
              ></div>
            </div>
          </div>
        )}

        <section className="results-table-container">
          <table>
            <thead>
              <tr>
                <th>File Name</th>
                <th>Site</th>
                <th>Violation</th>
                <th>Severity</th>
                <th>PII (BSN)</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {results.length > 0 ? (
                results.map((res) => (
                  <tr key={res.id} className="animate-fade">
                    <td style={{ fontWeight: 500 }}>{res.file}</td>
                    <td>{res.site}</td>
                    <td>
                      <span
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "6px",
                        }}
                      >
                        {res.type === "Anonymous Link" ? (
                          <AlertTriangle
                            size={14}
                            color="var(--accent-danger)"
                          />
                        ) : (
                          <User size={14} />
                        )}
                        {res.type}
                      </span>
                    </td>
                    <td>
                      <span
                        className={
                          res.severity === "High"
                            ? "severity-high"
                            : "severity-low"
                        }
                      >
                        {res.severity}
                      </span>
                    </td>
                    <td>
                      {res.bsnFound ? (
                        <span
                          style={{
                            color: "var(--accent-danger)",
                            display: "flex",
                            alignItems: "center",
                            gap: "4px",
                          }}
                        >
                          <Lock size={14} /> DETECTED
                        </span>
                      ) : (
                        <span style={{ color: "var(--text-secondary)" }}>
                          Clean
                        </span>
                      )}
                    </td>
                    <td>
                      <a
                        href={res.linkUrl}
                        target={isDemo ? "_self" : "_blank"}
                        rel="noreferrer"
                        style={{
                          color: "var(--accent-blue)",
                          textDecoration: "none",
                          fontSize: "0.75rem",
                          fontWeight: 600,
                        }}
                      >
                        {isDemo ? "Review (Mock)" : "Review Link"}
                      </a>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td
                    colSpan={6}
                    style={{
                      textAlign: "center",
                      padding: "3rem",
                      color: "var(--text-secondary)",
                    }}
                  >
                    {isScanning
                      ? "Analyzing data..."
                      : "No vulnerabilities found. Run a scan to begin."}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </section>
      </main>
    </div>
  );
}

function App() {
  const { instance } = useMsal();
  const [isDemoMode, setIsDemoMode] = useState(false);

  const handleLogin = () => {
    instance.loginPopup(loginRequest).catch((e) => console.error(e));
  };

  if (isDemoMode) {
    return <Dashboard isDemo onLogout={() => setIsDemoMode(false)} />;
  }

  return (
    <>
      <AuthenticatedTemplate>
        <Dashboard onLogout={() => {}} />
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <div
          className="dashboard-container"
          style={{ textAlign: "center", padding: "5rem 2rem" }}
        >
          <div
            className="logo-icon"
            style={{ margin: "0 auto 2rem", width: "64px", height: "64px" }}
          >
            <Shield size={32} color="white" />
          </div>
          <h1 style={{ fontSize: "2.5rem", marginBottom: "1rem" }}>
            M365 Security Auditor
          </h1>
          <p
            style={{
              color: "var(--text-secondary)",
              maxWidth: "500px",
              margin: "0 auto 2rem",
            }}
          >
            Identify over-privileged sharing links and sensitive PII across your
            SharePoint infrastructure. Ensure your AI Copilot is safe by
            auditing your data governance.
          </p>

          <div
            style={{
              display: "flex",
              justifyContent: "center",
              gap: "1rem",
              marginBottom: "4rem",
            }}
          >
            <button
              className="primary"
              style={{ padding: "1rem 2rem", fontSize: "1rem" }}
              onClick={handleLogin}
            >
              <LogIn
                size={20}
                style={{ marginRight: "10px", verticalAlign: "middle" }}
              />
              Login with M365
            </button>
            <button
              className="secondary"
              style={{ padding: "1rem 2rem", fontSize: "1rem" }}
              onClick={() => setIsDemoMode(true)}
            >
              <Search
                size={20}
                style={{ marginRight: "10px", verticalAlign: "middle" }}
              />
              View Demo
            </button>
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(3, 1fr)",
              gap: "2rem",
            }}
          >
            <div>
              <Search
                size={24}
                color="var(--accent-blue)"
                style={{ marginBottom: "1rem" }}
              />
              <h4>Site Discovery</h4>
              <p style={{ fontSize: "0.8rem", color: "var(--text-secondary)" }}>
                Automated inventory of all SharePoint sites and drives.
              </p>
            </div>
            <div>
              <AlertTriangle
                size={24}
                color="var(--accent-warning)"
                style={{ marginBottom: "1rem" }}
              />
              <h4>Link Audit</h4>
              <p style={{ fontSize: "0.8rem", color: "var(--text-secondary)" }}>
                Detection of Anonymous and "Everyone" sharing links.
              </p>
            </div>
            <div>
              <Lock
                size={24}
                color="var(--accent-danger)"
                style={{ marginBottom: "1rem" }}
              />
              <h4>PII Scanner</h4>
              <p style={{ fontSize: "0.8rem", color: "var(--text-secondary)" }}>
                Deep scan for sensitive Dutch BSN numbers in file contents.
              </p>
            </div>
          </div>
        </div>
      </UnauthenticatedTemplate>

      <footer
        style={{
          marginTop: "2rem",
          textAlign: "center",
          color: "var(--text-secondary)",
          fontSize: "0.8rem",
          paddingBottom: "2rem",
        }}
      >
        <p>&copy; 2026 SecureAudit M365 • Technical Demonstration</p>
      </footer>
    </>
  );
}

export default App;
