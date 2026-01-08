import React, { useState } from "react";
import "./App.css";
import { BrowserRouter as Router, Routes, Route, useNavigate } from "react-router-dom";
import { generateSkillFactorySlides, skillFactoryName } from "./pptx/skillFactory";
import { getFirstSlide, getLastSlide } from "./pptx/templateEditor";
import PreviewPage from "./PreviewPage";

/**
 * Returns the formatted date in YYYY-MM-DD.
 */
function getTodayString() {
  const today = new Date();
  return today.toISOString().slice(0, 10);
}

// Slide names for the sidebar
function getSidebarSlideNames(slides, factories) {
  const items = [];
  if (slides.length >= 1) items.push("1st Slide");
  factories.forEach((factory, i) => {
    items.push(`${factory}`);
  });
  if (slides.length >= 2) items.push("Last Slide");
  return items;
}

/**
 * PUBLIC_INTERFACE
 * App component for PPT generator builder.
 * - Handles routing for editor (/) and preview carousel (/preview).
 * - Keeps the fixed navbar/sidebar/panel behaviors and layout as before.
 */
function App() {
  // Initial Slides: just 1st and last
  const [slides, setSlides] = useState([
    getFirstSlide(getTodayString()),
    getLastSlide(),
  ]);
  // Sidebar factory names (unique)
  const [factories, setFactories] = useState([]);
  // Keep editable date
  const [firstSlideDate, setFirstSlideDate] = useState(getTodayString());
  // Sidebar selection
  const [selectedSidebarIdx, setSelectedSidebarIdx] = useState(0);
  // Export dropdown
  const [exportDropdown, setExportDropdown] = useState(false);

  // Track which navbar option is active: "preview", "skillFactory", "export"
  // This is used for button highlight only (not for routing anymore)
  const [activeNav, setActiveNav] = useState("preview");

  // React Router navigation hook (inside Router context)
  function EditorRouter() {
    const navigate = useNavigate();

    // Add Skill Factory inserts 4 slides before last slide.
    // Each factory is named e.g. Skill Factory 1, Skill Factory 2, ...
    const handleAddFactory = () => {
      const factoryNum = factories.length + 1;
      const factoryLabel = `${skillFactoryName} ${factoryNum}`;
      // Generate factory slides
      const factorySlides = generateSkillFactorySlides(factoryLabel);
      setFactories([...factories, factoryLabel]);
      // Insert before the last slide
      setSlides((prev) => [
        ...prev.slice(0, prev.length - 1),
        ...factorySlides,
        prev[prev.length - 1],
      ]);
      setSelectedSidebarIdx(slides.length - 1); // jump to factory
      setActiveNav("skillFactory");
      // IMPORTANT: Remain on the same route (main/editor)
    };

    // Only allow editing the date of the first slide
    const handleFirstSlideDateChange = (e) => {
      const dateStr = e.target.value;
      setFirstSlideDate(dateStr);
      // Regenerate first slide with the edited date
      setSlides(([first, ...rest]) => [
        getFirstSlide(dateStr),
        ...rest,
      ]);
    };

    // Handle Preview button - route to /preview
    const handlePreview = () => {
      setActiveNav("preview");
      navigate("/preview");
    };

    // Export handlers - stub (TODO wire up to pptx/pdf generators if needed)
    const handleExport = (format) => {
      setActiveNav("export");
      if (format === "pptx") {
        if (window.exportPPTX) window.exportPPTX(slides);
      } else if (format === "pdf") {
        if (window.exportPDF) window.exportPDF(slides);
      }
      setExportDropdown(false);
    };

    // Sidebar click behavior based on sidebar item.
    const handleSidebarSelect = (idx) => {
      setSelectedSidebarIdx(idx);

      // Sidebar items: ["1st Slide", factory1..., ..., "Last Slide"]
      if (idx === 0) {
        // 1st Slide: Stay on "/" (editor)
        navigate("/");
      } else if (idx === sidebarItems.length - 1) {
        // Last Slide: Go to new view-only route /last
        navigate("/last");
      } else {
        // Factory slide: select in list, stay on "/" (could expand with more editing in future)
        navigate("/");
      }
    };

    // Sidebar names reflect default+factory rules
    const sidebarItems = getSidebarSlideNames(slides, factories);

    // Last slide view-only page
    function LastSlideView() {
      // Get slide object for last slide (safe fallback)
      const lastSlide = slides.length >= 2 ? slides[slides.length - 1] : { title: "Last Slide", content: "Thank you!" };
      return (
        <main className="main-content">
          <div className="panel" style={{ pointerEvents: "none", opacity: 0.95 }}>
            <h2>{lastSlide.title}</h2>
            <pre style={{
              color: "#1a1a1a",
              fontSize: "1.13em",
              background: "#F3F6FC",
              borderRadius: "9px",
              padding: "20px 14px",
              minHeight: "90px",
              userSelect: "text",
            }}>
              {lastSlide.content}
            </pre>
            <div aria-label="This page is view-only for the final slide." style={{
              marginTop: 14, color: "#818cf8", fontWeight: 500, fontSize: "0.98em"
            }}>
              View-only. <span style={{color:"#E87A41"}}>No edits allowed.</span>
            </div>
          </div>
        </main>
      );
    }

    return (
      <div className="app-root ocean-pro">
        {/* Fixed Navbar */}
        <nav className="fixed-navbar">
          <div className="navbar-title">Weekly Statistic Report</div>
          <div className="navbar-actions">
            <button
              className={`navbar-btn primary${activeNav === "preview" ? " selected" : ""}`}
              onClick={handlePreview}
              aria-pressed={activeNav === "preview"}
            >
              Preview
            </button>
            <button
              className={`navbar-btn accent${activeNav === "skillFactory" ? " selected" : ""}`}
              onClick={handleAddFactory}
              aria-pressed={activeNav === "skillFactory"}
            >
              Add Skill Factory
            </button>
            <div className="export-dropdown">
              <button
                className={`navbar-btn${activeNav === "export" ? " selected" : ""}`}
                onClick={() => {
                  setActiveNav("export");
                  setExportDropdown((v) => !v);
                }}
                aria-pressed={activeNav === "export" || exportDropdown}
              >
                Export <span className="dropdown-arrow">▼</span>
              </button>
              {exportDropdown && (
                <div className="dropdown-menu">
                  <button onClick={() => handleExport("pptx")}>Export .pptx</button>
                  <button onClick={() => handleExport("pdf")}>Export .pdf</button>
                </div>
              )}
            </div>
          </div>
        </nav>

        {/* Fixed Sidebar */}
        <aside className="fixed-sidebar">
          {sidebarItems.map((item, idx) => (
            <div
              key={idx}
              className={`sidebar-item${selectedSidebarIdx === idx ? " selected" : ""}`}
              onClick={() => handleSidebarSelect(idx)}
              tabIndex={0}
              role="button"
              aria-current={selectedSidebarIdx === idx ? "page" : undefined}
            >
              {item}
            </div>
          ))}
        </aside>

        {/* Routes for main editor, preview carousel, and last slide view */}
        <Routes>
          <Route
            path="/"
            element={
              <main className="main-content">
                {/* Editing panel for first slide date only */}
                <div className="panel">
                  <h2>Report Configuration</h2>
                  <label>
                    <span className="input-label">Report Date:</span>
                    <input
                      type="date"
                      value={firstSlideDate}
                      onChange={handleFirstSlideDateChange}
                    />
                  </label>
                </div>
              </main>
            }
          />
          <Route
            path="/last"
            element={
              <LastSlideView />
            }
          />
          <Route
            path="/preview"
            element={
              <PreviewPage
                slides={slides}
                selectedSidebarIdx={selectedSidebarIdx}
                setSelectedSidebarIdx={setSelectedSidebarIdx}
                onBack={() => {
                  setActiveNav(""); // clear highlight, will be set when navigating
                  // Go back to editor page (main)
                  navigate("/");
                }}
              />
            }
          />
        </Routes>
      </div>
    );
  }

  // App outer structure: Provides Router context.
  return (
    <Router>
      <EditorRouter />
    </Router>
  );
}

export default App;
