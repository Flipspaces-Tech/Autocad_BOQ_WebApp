// src/pages/BoqUploadPage.jsx
import React, { useState, useRef, useCallback } from "react";
import "./BoqUploadPage.css";
import { BACKEND_URL } from "../config";

export default function BoqUploadPage() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [status, setStatus] = useState("");
  const [sheetUrl, setSheetUrl] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [uploadId, setUploadId] = useState("");
  const inputRef = useRef(null);

  // ✅ Settings collapse (closed by default)
  const [showSettings, setShowSettings] = useState(false);

  const [renderSettings, setRenderSettings] = useState({
    gsheetId: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
    gsheetTab: "TEST",

    previewTargetSize: 128,
    previewDpi: 240,

    previewInkThreshold: 100,
    previewThickenPx: 2,
    previewThickenIter: 2,
    previewCloseGapsKsize: 1,

    previewEdgeBgThresh: 250,
    previewMinVisibleAlpha: 8,

    previewPadPct: 0.04,
    previewMarginPct: 0.10,
    previewSupersample: 2,
  });

  /** Reset BOQ result state */
  const resetState = () => {
    setStatus("");
    setSheetUrl("");
    setSheetName("");
    setUploadId("");
  };

  /** Validate and store selected file */
  const handleFileSelect = (f) => {
    if (!f) return;

    const name = (f.name || "").toLowerCase();
    const isDxf = name.endsWith(".dxf");
    const isDwg = name.endsWith(".dwg");

    if (!isDxf && !isDwg) {
      setStatus("❌ Please upload a .dwg or .dxf file.");
      return;
    }

    setFile(f);
    resetState();
  };

  /** Drag & drop handlers */
  const onDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const droppedFiles = e.dataTransfer.files;
    if (droppedFiles && droppedFiles[0]) {
      handleFileSelect(droppedFiles[0]);
    }
  }, []);

  const onDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isDragging) setIsDragging(true);
  };

  const onDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleBrowseClick = () => {
    inputRef.current?.click();
  };

  /** Submit handler */
  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!file) {
      setStatus("Please choose a Dwg or DXF file first.");
      return;
    }

    setIsUploading(true);
    setStatus("Uploading and generating BOQ…");
    setSheetUrl("");
    setSheetName("");
    setUploadId("");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("settings", JSON.stringify(renderSettings));

      const url = `${BACKEND_URL}/process-cad`;
      console.log("Calling backend:", url);

      const res = await fetch(url, {
        method: "POST",
        body: formData,
      });

      const data = await res.json().catch(() => ({}));
      console.log("BOQ result:", res.status, data);

      if (!res.ok || !data.ok) {
        const msg =
          data.detail || data.error || data.message || "Unknown backend error.";
        setStatus("❌ Error: " + msg);
        return;
      }

      setSheetUrl(data.sheetUrl || "");
      setSheetName(data.sheetName || "");
      setUploadId(data.uploadId || "");
      setStatus(data.message || "✅ BOQ generated successfully.");
    } catch (err) {
      console.error(err);
      setStatus("❌ Network error: " + err.message);
    } finally {
      setIsUploading(false);
    }
  };

  return (
    <div className="boq-page">
      <div className="boq-card">
        {/* Header */}
        <div className="boq-header">
          <div className="boq-pill">Vizdom · AutoCAD BOQ</div>
          <h1>BOQ Generator</h1>
          <p>
            Upload an <strong>ASCII DWG</strong> file and get an auto-generated
            BOQ in Google Sheets.
          </p>
        </div>

        {/* Upload form */}
        <form className="boq-form" onSubmit={handleSubmit}>
          <div
            className={`boq-dropzone ${isDragging ? "boq-dropzone--drag" : ""}`}
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDragEnd={onDragLeave}
          >
            <input
              type="file"
              ref={inputRef}
              style={{ display: "none" }}
              accept=".dwg,.dxf"
              onChange={(e) => handleFileSelect(e.target.files[0])}
            />

            <div className="boq-dropzone-icon">📐</div>

            {file ? (
              <>
                <div className="boq-file-name">{file.name}</div>
                <div className="boq-file-meta">
                  {(file.size / (1024 * 1024)).toFixed(2)} MB ·{" "}
                  {file.name.toLowerCase().endsWith(".dxf") ? "DXF" : "DWG"}
                </div>
                <button
                  type="button"
                  className="boq-secondary-btn"
                  onClick={handleBrowseClick}
                  disabled={isUploading}
                >
                  Change file
                </button>
              </>
            ) : (
              <>
                <p className="boq-dropzone-title">Drag &amp; drop your drawing here</p>
                <p className="boq-dropzone-sub">
                  or{" "}
                  <button
                    type="button"
                    className="boq-link-btn"
                    onClick={handleBrowseClick}
                    disabled={isUploading}
                  >
                    browse from your computer
                  </button>
                </p>
                <p className="boq-dropzone-hint">Supported: .dwg / .dxf</p>
              </>
            )}
          </div>

          {/* ✅ Collapsible settings */}
          <div className="boq-settings">
            <button
              type="button"
              className="boq-settings-toggle"
              onClick={() => setShowSettings((v) => !v)}
              aria-expanded={showSettings}
            >
              <span>Settings</span>
              <span className={`boq-caret ${showSettings ? "open" : ""}`}>▾</span>
            </button>

            {showSettings && (
              <div className="boq-settings-body">
                <div className="boq-settings-grid">
                  <label className="boq-setting full">
                    <span>Google Sheet ID</span>
                    <input
                      type="text"
                      value={renderSettings.gsheetId}
                      onChange={(e) =>
                        setRenderSettings((s) => ({ ...s, gsheetId: e.target.value }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Sheet Tab</span>
                    <input
                      type="text"
                      value={renderSettings.gsheetTab}
                      onChange={(e) =>
                        setRenderSettings((s) => ({ ...s, gsheetTab: e.target.value }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Preview Size</span>
                    <input
                      type="number"
                      min="32"
                      max="512"
                      value={renderSettings.previewTargetSize}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewTargetSize: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Preview DPI</span>
                    <input
                      type="number"
                      min="72"
                      max="600"
                      value={renderSettings.previewDpi}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewDpi: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Ink Threshold</span>
                    <input
                      type="number"
                      min="0"
                      max="255"
                      value={renderSettings.previewInkThreshold}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewInkThreshold: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Thicken (px)</span>
                    <input
                      type="number"
                      min="0"
                      max="10"
                      value={renderSettings.previewThickenPx}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewThickenPx: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Thicken Iter</span>
                    <input
                      type="number"
                      min="0"
                      max="10"
                      value={renderSettings.previewThickenIter}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewThickenIter: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>

                  <label className="boq-setting">
                    <span>Close Gaps K</span>
                    <input
                      type="number"
                      min="0"
                      max="10"
                      value={renderSettings.previewCloseGapsKsize}
                      onChange={(e) =>
                        setRenderSettings((s) => ({
                          ...s,
                          previewCloseGapsKsize: Number(e.target.value || 0),
                        }))
                      }
                    />
                  </label>
                </div>
              </div>
            )}
          </div>

          {/* Actions */}
          <div className="boq-actions">
            <button
              type="submit"
              className="boq-primary-btn"
              disabled={!file || isUploading}
            >
              {isUploading ? (
                <span className="boq-loader">
                  <span className="boq-loader-dot" />
                  <span className="boq-loader-dot" />
                  <span className="boq-loader-dot" />
                  Processing…
                </span>
              ) : (
                "Generate BOQ"
              )}
            </button>

            {sheetUrl && (
              <a
                href={sheetUrl}
                target="_blank"
                rel="noreferrer"
                className="boq-ghost-btn"
              >
                Open Google Sheet
              </a>
            )}
          </div>
        </form>

        {/* Footer / Status */}
        <div className="boq-footer">
          <div className="boq-status">
            {status && <span>{status}</span>}
            {sheetName && (
              <span className="boq-tag">
                Sheet: <strong>{sheetName}</strong>
              </span>
            )}
            {uploadId && <span className="boq-tag">Run ID: {uploadId}</span>}
          </div>

          <div className="boq-footnote">
            Auto-syncs to your BOQ master sheet. No GCP login required for users.
          </div>
        </div>
      </div>
    </div>
  );
}
