import React, { useState } from "react";
import "./App.css";
import PptxPreview from "./pptx/PptxPreview";
import { generateSkillFactorySlides, skillFactoryName } from "./pptx/skillFactory";
import { getFirstSlide, getLastSlide } from "./pptx/templateEditor";

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
 * - Conditionally renders the slide preview only when Preview is active in the navbar.
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
  const [activeNav, setActiveNav] = useState("preview");

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

  // Handle Preview button - mark preview active and scroll into view
  const handlePreview = () => {
    setActiveNav("preview");
    setTimeout(() => {
      // Smooth scroll to preview panel only if it exists
      const section = document.getElementById("preview-section");
      if (section) section.scrollIntoView({ behavior: "smooth" });
    }, 10);
  };

  // Export handlers - stub, TODO wire up to pptx/pdf generators
  const handleExport = (format) => {
    setActiveNav("export");
    if (format === "pptx") {
      // Trigger pptx export (to be implemented by PptxPreview)
      if (window.exportPPTX) window.exportPPTX(slides);
    } else if (format === "pdf") {
      // Trigger pdf export (to be implemented by PptxPreview)
      if (window.exportPDF) window.exportPDF(slides);
    }
    setExportDropdown(false);
  };

  // Sidebar click: jumps to preview slide (carousel is handled in PptxPreview)
  const handleSidebarSelect = (idx) => {
    setSelectedSidebarIdx(idx);
    // Always set preview active when clicking a slide in sidebar
    setActiveNav("preview");
    setTimeout(() => {
      const section = document.getElementById("preview-section");
      if (section) section.scrollIntoView({ behavior: "smooth" });
    }, 10);
  };

  // Sidebar names reflect default+factory rules
  const sidebarItems = getSidebarSlideNames(slides, factories);

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

      {/* Main Content - leave space for navbar/sidebar */}
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

        {/* Show slide preview only when Preview option is active in the navbar */}
        {activeNav === "preview" && (
          <div id="preview-section" className="preview-block">
            <PptxPreview
              slides={slides}
              selectedIdx={selectedSidebarIdx}
              setSlideIdx={setSelectedSidebarIdx}
            />
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
